import sys
import re
import subprocess

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select


class NewBrowserConnection():
    def __init__(self):
        # Create Edge options object
        # Add the experimental option to specify the debugger address
        browser_options = Options()
        browser_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        # Initialize the Edge WebDriver, connecting to the running instance
        # Make sure to provide the path to your msedgedriver.exe if it's not in your system's PATH
        self.driver = webdriver.Edge(options=browser_options)
        self.main_handle = None
        self.leader_email = ''
  
    def teardown(self):
        self.driver.quit()

    def check_element(self, id):
        cur_element = self.driver.find_element(By.ID, id)
        if not cur_element.is_selected():
            cur_element.click()
    
    def waitForClickable(self, type, id):
        timeoutsec = 30

        try:
            match type:
                case "FRAME_SWITCH":
                    WebDriverWait(self.driver, timeoutsec).until(
                        EC.frame_to_be_available_and_switch_to_it((By.ID, id))
                    )
                case "ELEM_CLICKABLE":
                    WebDriverWait(self.driver, timeoutsec).until(
                        EC.element_to_be_clickable((By.ID, id))
                    )
            
        except TimeoutException:
            # Handle the case where the element is not found within 10 seconds
            print("Timeout: Element not found within the specified time.")
    
        except NoSuchElementException:
            # Some ExpectedConditions might throw NoSuchElementException (e.g. text_to_be_present_in_element)
            print("No such element found in the DOM.")

        except Exception as e:
            # Catch other potential exceptions
            print(f"An unexpected error occurred: {e}")


    def GetDefaultContext(self):

        # Make sure we are on the PHOC main tab
        all_handles = self.driver.window_handles
        for handle in all_handles:
            self.driver.switch_to.window(handle)
            if self.driver.title == 'Piedmont Hiking and Outing Club - Event details':
                self.main_handle = handle # Not sure I need this, but save it for now
                break

        self.driver.switch_to.default_content()


    def ClickEdit(self):
        self.driver.switch_to.window(self.main_handle)
        self.driver.switch_to.default_content()

        # Click on the Edit button and switch into contentFrame
        wait = WebDriverWait(self.driver, 10)
        edit_btn = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn-group > .btn:nth-child(1)"))
        )
        edit_btn.click()

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentFrame")))



    def ProcessEventdetailsTab(self):
        print("   Processing the Event details tab...")
        wait = WebDriverWait(self.driver, 10)

        # 1. Switch to contentFrame and open the Event Details tab
        self.driver.switch_to.default_content()
        self.waitForClickable("FRAME_SWITCH", "contentFrame")
        self.waitForClickable("ELEM_CLICKABLE", "1-link-id_InnerControl")
        self.driver.find_element(By.ID, "1-link-id_InnerControl").click()

        # 2. Check attendee visibility boxes
        self.waitForClickable("ELEM_CLICKABLE", "eventDetailsMain_editAttendeesSettings_showRegistrantsList")
        self.check_element("eventDetailsMain_editAttendeesSettings_showRegistrantsList")
        self.check_element("eventDetailsMain_editAttendeesSettings_visibilityMembers")

        # 3. Switch into the dynamic Event Description editor iframe and extract text
        editor_iframe = wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[id^='idEditorIFrame_']"))
        )
        text_element = wait.until(EC.presence_of_element_located((By.ID, "idPrimaryContentBlock1Content")))
        description_text = text_element.text

        email_match = re.search(r'Email:\s*(.*?)[\r\n]', description_text)
        self.leader_email = email_match.group(1).strip() if email_match else ""

        addtl_match = re.search(r'Move to "Additional event information" section below\s*\n(.*)', description_text, re.DOTALL)
        move_below_text = addtl_match.group(1).strip() if addtl_match else ""

        # 4. Open the HTML source editor modal
        text_element.click()
        self.driver.switch_to.parent_frame()
        wait.until(EC.element_to_be_clickable((By.ID, "idEditorToolbar_EditorEventDescriptionLocalToolbar_HTML_HTMLEdit"))).click()

        # 5. Wait for the HTML editor textarea to load and contain <strong>
        self.driver.switch_to.default_content()
        html_box_locator = (By.ID, "idBEditor_EditHTML_Dialog_HTMLCodeContainer")

        def html_box_ready(driver):
            try:
                el = driver.find_element(*html_box_locator)
                content = el.get_attribute("value") or el.text
                return el if "<strong>" in content.lower() else False
            except (NoSuchElementException, StaleElementReferenceException):
                return False

        html_edit_window = wait.until(html_box_ready)

        # 6. Clean up unwanted/empty fields from the HTML source
        html_text = (html_edit_window.get_attribute("value") or html_edit_window.text).replace('\n', '')
        strong_match = re.search(r'(?i)<strong>.*', html_text)
        if strong_match:
            html_text = strong_match.group(0)

        # Strip everything from 'Move to "Additional event information"' onward
        html_text = re.sub(re.escape('Move to "Additional event information"') + r'.*', '', html_text)

        # Remove empty optional fields
        empty_fields = [
            r'Distance:\s*</(?i:strong)>\s*miles\s*<br>',
            r'Difficulty:\s*</(?i:strong)>\s*<br>',
            r'Intended Pace:\s*</(?i:strong)>\s*<br>',
            r'Elevation Gain:\s*</(?i:strong)>\s*ft\.\s*<br>',
            r'Required Gear:\s*</(?i:strong)>\s*<br>',
            r'Max Participants:\s*</(?i:strong)>\s*<br>',
        ]
        for field_pattern in empty_fields:
            html_text = re.sub(r'<(?i:strong)>' + field_pattern, '', html_text)

        html_text = html_text.replace('R-,', '')
        html_text = re.sub(r' {2,}', ' ', html_text).strip()

        # 7. Inject the cleaned HTML directly (no clipboard or sleep needed) and save
        self.driver.execute_script(
            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));",
            html_edit_window,
            html_text,
        )
        wait.until(EC.element_to_be_clickable((By.ID, "idBEditor_EditHTML_Dialog_SaveButton"))).click()

        # 8. Populate the "Additional event information" box at the bottom
        self.driver.switch_to.default_content()
        self.waitForClickable("FRAME_SWITCH", "contentFrame")
        addtn_info_window = wait.until(EC.element_to_be_clickable((By.ID, "eventDetailsMain_editExtraEventInfo")))
        addtn_info_window.clear()
        if move_below_text:
            addtn_info_window.send_keys(move_below_text)


    def ProcessTickettypeTab(self):
        ##### Tab: Ticket types & settings
        # Check box: Enable waitlist when limit is reached
        # Wait for Multiple registrations checkbox first to confirm the tab finished loading,
        # because the waitlist checkbox is only shown when a registration limit is set.
        print("   Processing the Ticket types & settings tab...")

        self.waitForClickable("ELEM_CLICKABLE", "3-link-id_InnerControl")
        self.driver.find_element(By.ID, "3-link-id_InnerControl").click()
        self.waitForClickable("ELEM_CLICKABLE", "ctl03_cbMultipleRegistration")

        waitlist_boxes = self.driver.find_elements(By.ID, "ctl03_waitlistEnableCheckBox")
        if waitlist_boxes and waitlist_boxes[0].is_displayed():
            self.check_element("ctl03_waitlistEnableCheckBox")

    def ProcessWaitlistTab(self):
        print("   Processing the Waitlist tab...")

        # 1. Switch to the Waitlist & settings tab safely
        self.waitForClickable("ELEM_CLICKABLE", "6-link-id_InnerControl")
        self.driver.find_element(By.ID, "6-link-id_InnerControl").click()

        # 2. Wait for the tab content to load, then select both radio options
        radio_ids = [
            "eventWaitlistMain_eventWaitlistRegistrationTypeSelector_rbRegistrationTypeAuto",
            "eventWaitlistMain_informationToCollectSelector_rbInformationToCollectContactInformation",
        ]
        self.waitForClickable("ELEM_CLICKABLE", radio_ids[0])

        for radio_id in radio_ids:
            self.check_element(radio_id)

    def ProcessEmailsTab(self):
        print("   Processing the Emails details tab...")

        # 1. Switch to the Emails tab safely``
        self.waitForClickable("ELEM_CLICKABLE", "4-link-id_InnerControl")
        self.driver.find_element(By.ID, "4-link-id_InnerControl").click()

        # 2. Wait for the first checkbox to confirm the tab finished loading
        checkbox_ids = [
            "eventEmails_registrationConfirmedOffline_cbxAttendee",
            "eventEmails_registrationConfirmedOffline_cbxGuest",
            "eventEmails_registrationConfirmedOffline_cbxAdmin",
            "eventEmails_registrationPendingOffline_cbxAttendee",
            "eventEmails_registrationPendingOffline_cbxGuest",
            "eventEmails_registrationPendingOffline_cbxAdmin",
            "eventEmails_registrationCanceled_cbxAttendee",
            "eventEmails_registrationCanceled_cbxGuest",
            "eventEmails_registrationCanceled_cbxAdmin",
            "eventEmails_registrationNewWaitlistEntry_cbxAttendee",
            "eventEmails_registrationNewWaitlistEntry_cbxAdmin",
        ]
        self.waitForClickable("ELEM_CLICKABLE", checkbox_ids[0])

        # 3. Check all 11 email notification boxes
        for checkbox_id in checkbox_ids:
            self.check_element(checkbox_id)


    def ProcessEventOrganizer(self):
        print("   Finally, back to the Event details tab to get Event Organizer...")
        wait = WebDriverWait(
            self.driver, 
            15, 
            ignored_exceptions=(StaleElementReferenceException,)
        )

        # 1. Switch to the main content frame and open the Event Details tab
        self.driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentFrame")))
        wait.until(EC.element_to_be_clickable((By.ID, "1-link-id_InnerControl"))).click()
        
        # 2. Wait for the organizers container and click "+ Add organizer"
        wait.until(EC.element_to_be_clickable((By.ID, "eventDetailsMain_eventOrganizers_eventOrganizersEditor")))
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.typeButton[value="+ Add organizer"]'))).click()
        
        # 3. Switch into the outer modal iframe
        self.driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "idBaseIFrame_SelectRecipientDialog")))

        # 4. Safely enter the inner reloading iframe and wait for SearchBox to be ready
        def switch_to_reloaded_iframe_and_get_searchbox(driver):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame("idBaseIFrame_SelectRecipientDialog")
                driver.switch_to.frame("idReloadIFrame_SelectRecipientDialog")
                box = driver.find_element(By.ID, "ctl00_innerMainContainer_contactListDisplay_SearchBox")
                if box.is_displayed() and box.is_enabled():
                    return box
            except (NoSuchElementException, StaleElementReferenceException, Exception):
                pass
            return False

        email_lookup_element = wait.until(switch_to_reloaded_iframe_and_get_searchbox)
        email_lookup_element.clear()
        email_lookup_element.send_keys(self.leader_email)

        # 5. Select the matching row, or halt cleanly if not found
        row_locator = (By.XPATH, f'//table//tbody//tr[contains(., "{self.leader_email}")]')
        try:
            row = WebDriverWait(
                self.driver, 
                4, 
                ignored_exceptions=(StaleElementReferenceException,)
            ).until(EC.element_to_be_clickable(row_locator))
            row.click()
        except TimeoutException:
            print(f"\n\n   STOPPING: Could not find '{self.leader_email}' in the search table.\n")
            if hasattr(self, "driver") and self.driver.service.process:
                self.driver.service.stop()
            sys.exit(0)

        # 6. Click the "Select" button to close the modal
        wait.until(EC.element_to_be_clickable((By.ID, "ctl00_innerMainContainer_proceedButton"))).click()

        # 7. Switch back to contentFrame and select 'Event lead'
        self.driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentFrame")))
        dropdown_el = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "eventOrganizerLevelSelect")))
        Select(dropdown_el).select_by_visible_text("Event lead")


def prompt_new_browser():
    validresponse = False
    while not validresponse:
        response = input(f'\nAn Edge browser must be used with a debug port enabled. Do you wish to have that browser created now? (y/n/quit): ')
        if response in ('y', 'n', 'quit'):
            validresponse = True

    if response == 'y':
        print(f'Opening a new browser using the following CMD:\n' \
            '   "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe" "https://phoc.club" --remote-debugging-port=9222 --user-data-dir="%temp%\\SeleniumEdgeProfile"\n' \
            '\n\n')
        
        # This will invoke a new Edge browser, and it will create a profile in the user's temp directory. (C:\Users\<userid>\AppData\Local\Temp)
        # When this is done the first time, it is as if you are using a browser for the first time, and it takes you through all of the
        # browser setup. For this reason, I don't delete this profile when the program is done. That way, you only have to go through
        # the setup once.
        command_list = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", 
                        r"https://phoc.club",
                        r"--remote-debugging-port=9222",
                        r"--user-data-dir=%temp%\PHOCSeleniumEdgeProfile"]
        subprocess.Popen(command_list,
                        shell=True, stdin=None, stdout=None, stderr=None,
                        creationflags=subprocess.CREATE_NEW_CONSOLE)

        input(f'Once the Edge browser is invoked:\n' \
            '   * log into phoc.club as Admin\n' \
            '   * open your first event for processing\n' \
            '   * press "Enter" to continue processing the first event\n\n')
        return False
    elif response == 'n':
        input(f'From the Edge browser you already have invoked:\n' \
            '   * open your next event for processing\n' \
            '   * press "Enter" to continue processing the next event\n\n')
        return False
    return True  # 'quit' was the input


def main():

    # In order to be able to connect Selenium to the browser, it must be invoked from a cmd line with the following:
    # Start an Edge (or Chrome) browser from a cmd prompt using the following, before running this script:
    # "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" "https://phoc.club" --remote-debugging-port=9222 --user-data-dir="%temp%\SeleniumEdgeProfile"
    # "C:\Program Files\Google\Chrome\Application\chrome.exe" "https://phoc.club" --remote-debugging-port=9222 --user-data-dir="%temp%\SeleniumChromeProfile"

    if len(sys.argv) <= 1:
        quit_request = prompt_new_browser()

    elif sys.argv[1] == '-np':
        quit_request = False
        
    else:
        print(f'\nImproper command usage\n\n')
        quit_request = True

    if not quit_request:
        print(f'Processing the event... Processing completes on the search of the leader email.')
        
        browser = NewBrowserConnection()
        browser.GetDefaultContext()

        browser.ClickEdit()
        browser.ProcessEventdetailsTab()
        browser.ProcessTickettypeTab()
        browser.ProcessWaitlistTab()
        browser.ProcessEmailsTab()

        # This is most likely to fail, if we can't find the leader's email address.
        # Therefore, we will do this very last.
        browser.ProcessEventOrganizer()

        browser.teardown()

        print(f'Processing complete. Save the current event, open the next event,\n' \
                'and then re-launch this script with the following:\n' \
                '\n' \
                'phoc_event_fill.exe -np\n')
    
    sys.exit()


if __name__ == "__main__":
    main()