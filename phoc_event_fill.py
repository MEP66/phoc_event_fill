import sys
import time
import re
import subprocess
import pyperclip  # pip install pyperclip (into virtual env)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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


        # Click on the Edit button, and switch to contentFrame
        self.driver.find_element(By.CSS_SELECTOR, ".btn-group > .btn:nth-child(1)").click()


    def ProcessEventdetailsTab(self):

        def text_strong_present_ignore_case(driver):
            try:
                element = driver.find_element(By.ID, "idBEditor_EditHTML_Dialog_HTMLCodeContainer")
                return "<strong>" in element.text.lower()
            except (NoSuchElementException, StaleElementReferenceException):
                return False

        ##### Tab: Event details
        # Check box: Show registrants who want to be listed
        # Button: to members only
        print(f'   Processing the Event details tab...')
        
        self.driver.switch_to.default_content()
        self.waitForClickable("FRAME_SWITCH", "contentFrame")
        self.waitForClickable("ELEM_CLICKABLE", "1-link-id_InnerControl")
        self.driver.find_element(By.ID, "1-link-id_InnerControl").click()
        
        self.waitForClickable("ELEM_CLICKABLE", "eventDetailsMain_editAttendeesSettings_showRegistrantsList")
        self.check_element("eventDetailsMain_editAttendeesSettings_showRegistrantsList")
        self.check_element("eventDetailsMain_editAttendeesSettings_visibilityMembers")

        # Next find the Event Detail frame. There is a unique ID number in this frame,
        # so we have to wildcard it. Fetch the leader email (to be used later) as well
        # as the text that needs to be moved below to the Additional event information box.
        frameElement = self.driver.find_element(By.CSS_SELECTOR, "iframe[id^='idEditorIFrame_']")
        self.driver.switch_to.frame(frameElement)
        text_element = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "idPrimaryContentBlock1Content"))
            )
        description_text = text_element.text

        email_pattern = r'Email: (.*?)[\r\n]'
        self.leader_email = re.search(email_pattern, description_text).group(1)

        addtl_info_pattern = r'Move to "Additional event information" section below\n(.*)'
        move_below_text = re.search(addtl_info_pattern, description_text, re.DOTALL).group(1)
        
        # Clean up the HTML text box
        
        self.driver.find_element(By.ID, "idPrimaryContentBlock1Content").click()
        self.driver.switch_to.parent_frame()
        self.driver.find_element(By.ID, "idEditorToolbar_EditorEventDescriptionLocalToolbar_HTML_HTMLEdit").click()

        self.driver.switch_to.default_content()
        html_edit_window = self.driver.find_element(By.ID, "idBEditor_EditHTML_Dialog_HTMLCodeContainer")

        # Wait for element to appear. I've seen <STRONG> and <strong> in the
        # HTML code, so allow for both.
        WebDriverWait(self.driver, 10).until(text_strong_present_ignore_case)

        # Clean up the html description text of un-wanted text. (Those entries
        # where no value was provided by the submitter.)
        # Once again, look for the 'STRONG' tag with case insensitivity.
        html_text = html_edit_window.text
        html_text = html_text.replace('\n', '')
        html_text = re.search(r'<(?i:strong)>.*', html_text).group(0)
        
        substring_to_remove = 'Move to "Additional event information"'
        pattern = re.escape(substring_to_remove) + r'.*'
        html_text = re.sub(pattern, "", html_text)

        html_text = re.sub(r'<(?i:strong)>Distance: *</(?i:strong)> *miles *<br>', '', html_text)
        html_text = re.sub(r'R-,', '', html_text)
        html_text = re.sub(r'<(?i:strong)>Difficulty: *</(?i:strong)> *<br>', '', html_text)
        html_text = re.sub(r'<(?i:strong)>Intended Pace: *</(?i:strong)> *<br>', '', html_text)
        html_text = re.sub(r'<(?i:strong)>Elevation Gain: *</(?i:strong)> *ft. *<br>', '', html_text)
        html_text = re.sub(r'<(?i:strong)>Required Gear: *</(?i:strong)> *<br>', '', html_text)
        html_text = re.sub(r'<(?i:strong)>Max Participants: *</(?i:strong)> *<br>', '', html_text)
        html_text = re.sub(r' {2,}', ' ', html_text)  # Substitute double space for single space
        
        # Delete the existing html text.
        actions = ActionChains(self.driver)
        actions.click(html_edit_window) \
            .key_down(Keys.CONTROL) \
            .send_keys("a") \
            .key_up(Keys.CONTROL) \
            .send_keys(Keys.DELETE) \
            .perform()

        # Paste the new text
        pyperclip.copy(html_text)

        actions.click(html_edit_window) \
            .key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL) \
            .perform()

        # Save the result
        self.driver.find_element(By.ID, "idBEditor_EditHTML_Dialog_SaveButton").click()

        # Paste the Additional information in the text box at the bottom.
        self.driver.switch_to.default_content()
        self.waitForClickable("FRAME_SWITCH", "contentFrame")
        
        addtn_info_window = self.driver.find_element(By.ID, "eventDetailsMain_editExtraEventInfo")
        addtn_info_window.clear()
        time.sleep(1)

        pyperclip.copy(move_below_text)
        actions.click(addtn_info_window) \
            .key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL) \
            .perform()


    def ProcessTickettypeTab(self):
        ##### Tab: Ticket types & settings
        # Check box: Enable waitlist when limit is reached
        # This element is only displayed if the wait limit is set. Since this elemnet
        # may or may not be displayed, I have to wait on another element to be clickable
        # (the Multiple registrations checkbox) to make sure the page is loaded. 
        print(f'   Processing the Ticket types & settings tab...')
        
        self.driver.find_element(By.ID, "3-link-id_InnerControl").click()
        self.waitForClickable("ELEM_CLICKABLE", "ctl03_cbMultipleRegistration")

        element = self.driver.find_element(By.ID, "ctl03_waitlistEnableCheckBox")
        if element.is_displayed():
            self.check_element("ctl03_waitlistEnableCheckBox")

    def ProcessWaitlistTab(self):
        ##### Tab: Waitlist & settings
        # Button: Automatic registration
        # Button: All contact information
        print(f'   Processing the Waitlist tab...')

        self.driver.find_element(By.ID, "6-link-id_InnerControl").click()
        self.waitForClickable("ELEM_CLICKABLE", "eventWaitlistMain_eventWaitlistRegistrationTypeSelector_rbRegistrationTypeAuto")
        
        self.check_element("eventWaitlistMain_eventWaitlistRegistrationTypeSelector_rbRegistrationTypeAuto")
        self.check_element("eventWaitlistMain_informationToCollectSelector_rbInformationToCollectContactInformation")


    def ProcessEmailsTab(self):
        #### Tab: Emails
        # Check box: All 11 email boxes
        print(f'   Processing the Emails details tab...')

        self.driver.find_element(By.ID, "4-link-id_InnerControl").click()
        self.waitForClickable("ELEM_CLICKABLE", "eventEmails_registrationConfirmedOffline_cbxAttendee")
        
        self.check_element("eventEmails_registrationConfirmedOffline_cbxAttendee")
        self.check_element("eventEmails_registrationConfirmedOffline_cbxGuest")
        self.check_element("eventEmails_registrationConfirmedOffline_cbxAdmin")
        self.check_element("eventEmails_registrationPendingOffline_cbxAttendee")
        self.check_element("eventEmails_registrationPendingOffline_cbxGuest")
        self.check_element("eventEmails_registrationPendingOffline_cbxAdmin")
        self.check_element("eventEmails_registrationCanceled_cbxAttendee")
        self.check_element("eventEmails_registrationCanceled_cbxGuest")
        self.check_element("eventEmails_registrationCanceled_cbxAdmin")
        self.check_element("eventEmails_registrationNewWaitlistEntry_cbxAttendee")
        self.check_element("eventEmails_registrationNewWaitlistEntry_cbxAdmin")



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