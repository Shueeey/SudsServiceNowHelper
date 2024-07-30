import platform
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QComboBox, QLineEdit, QListWidget
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QTimer, Qt
import openpyxl
from datetime import datetime
import time
import getpass
from bs4 import BeautifulSoup



from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


from selenium.webdriver.common.by import By



# URL to open
url = "https://sydneytest.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D3c714f09dbe080502d38cae43a9619cd%26sysparm_link_parent%3D5fbc29844fba1fc05ad9d0311310c75d%26sysparm_catalog%3D09a851b34faadbc05ad9d0311310c7e7%26sysparm_catalog_view%3Dsm_cat_categories%26sysparm_view%3Dtext_search"
task_list_url = "https://sydneytest.service-now.com/nav_to.do?uri=%2F$interactive_analysis.do%3Fsysparm_field%3Dopened_at%26sysparm_table%3Dtask%26sysparm_from_list%3Dtrue%26sysparm_query%3Dactive%3Dtrue%5Estate!%3D6%5Eassignment_group%3Djavascript:getMyGroups()%5Eassigned_to%3D%26sysparm_list_view%3Dcatalog"

# Function to open the URL in Chrome
def open_url(url):
    if platform.system() == "Windows":
        try:
            import webbrowser
            from pywinauto import Application

            # Connect to the already open Chrome window
            app = Application(backend="uia").connect(title_re=".*Chrome.*")
            chrome = app.window(title_re=".*Chrome.*")

            # Open the URL in a new tab
            chrome.set_focus()
            webbrowser.open_new_tab(url)

            print("URL opened successfully in Chrome on Windows!")
        except Exception as e:
            print("An error occurred on Windows:", str(e))

    elif platform.system() == "Darwin":  # macOS
        try:
            import subprocess

            # Open the URL in a new tab of Google Chrome using AppleScript
            script = f'tell application "Google Chrome" to open location "{url}"'
            subprocess.call(["osascript", "-e", script])

            print("URL opened successfully in Chrome on macOS!")
        except Exception as e:
            print("An error occurred on macOS:", str(e))

    else:
        print("Unsupported operating system.")

class SnowSoftwareWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sud's SNow Software")
        self.setGeometry(100, 100, 400, 300)

        # Create the timer label
        self.timer_label = QLabel("00:00:00", self)
        self.timer_label.setFont(QFont("Helvetica", 24))
        self.timer_label.setAlignment(Qt.AlignCenter)

        # Create the start button
        self.start_button = QPushButton("Start", self)
        self.start_button.clicked.connect(self.start_timer)

        # Create the stop/resume button
        self.stop_resume_button = QPushButton("Stop", self)
        self.stop_resume_button.setEnabled(False)
        self.stop_resume_button.clicked.connect(self.stop_resume_timer)

        # Create the reset button
        self.reset_button = QPushButton("Reset", self)
        self.reset_button.setEnabled(False)
        self.reset_button.clicked.connect(self.reset_timer)

        # Create the task list button
        self.task_list_button = QPushButton("Open Task List", self)
        self.task_list_button.clicked.connect(self.open_task_list)

        # Create the refund tickets button
        self.refund_tickets_button = QPushButton("List Refund Tickets", self)
        self.refund_tickets_button.clicked.connect(self.filter_print_refund_tickets)



        # Create the Unikey input field
        self.unikey_input = QLineEdit(self)
        self.unikey_input.setPlaceholderText("Enter Unikey")

        # Create the counter location dropdown
        self.counter_location_dropdown = QComboBox(self)
        self.counter_location_dropdown.addItem("Fisher")
        self.counter_location_dropdown.addItem("Other Location")

        # Create the assistance category dropdown
        self.assistance_category_dropdown = QComboBox(self)
        self.assistance_category_dropdown.addItems([
            "Card Encoding", "General Assist", "Lab Computer", "Loaned Device",
            "Lost & Found", "Navigation", "Personal Device", "Printing",
            "Software Installation", "Unikey Assistance", "University Application Assistance",
            "WiFi", "Other"
        ])

        # Create a list widget to display the ticket links
        self.ticket_list = QListWidget(self)

        # Create a vertical layout and add the widgets
        layout = QVBoxLayout()
        layout.addWidget(self.timer_label)
        layout.addWidget(self.unikey_input)
        layout.addWidget(self.counter_location_dropdown)
        layout.addWidget(self.assistance_category_dropdown)
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_resume_button)
        layout.addWidget(self.reset_button)
        layout.addWidget(self.task_list_button)
        layout.addWidget(self.refund_tickets_button)
        layout.addWidget(self.ticket_list)
        self.setLayout(layout)

        self.elapsed_time = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_timer)

        # Initialize the Excel workbook and sheet
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "Assistance Data"
        self.sheet.append(["Date", "Unikey", "Counter Location", "Assistance Category", "Time Spent (minutes)", "Notes"])

    def start_timer(self):
        self.start_button.setEnabled(False)
        self.stop_resume_button.setEnabled(True)
        self.stop_resume_button.setText("Stop")
        self.reset_button.setEnabled(True)
        open_url(url)
        self.timer.start(1000)  # Update every second

    def stop_resume_timer(self):
        if self.timer.isActive():
            self.timer.stop()
            self.stop_resume_button.setText("Resume")
        else:
            self.timer.start(1000)
            self.stop_resume_button.setText("Stop")

    def reset_timer(self):
        self.elapsed_time = 0
        self.timer_label.setText("00:00:00")
        self.start_button.setEnabled(True)
        self.stop_resume_button.setEnabled(False)
        self.stop_resume_button.setText("Stop")
        self.reset_button.setEnabled(False)
        self.timer.stop()
        self.save_to_excel()

    def update_timer(self):
        self.elapsed_time += 1
        minutes, seconds = divmod(self.elapsed_time, 60)
        hours, minutes = divmod(minutes, 60)
        self.timer_label.setText(f"{hours:02d}:{minutes:02d}:{seconds:02d}")

    def save_to_excel(self):
        date = datetime.now().strftime("%Y-%m-%d")
        unikey = self.unikey_input.text()
        counter_location = self.counter_location_dropdown.currentText()
        assistance_category = self.assistance_category_dropdown.currentText()
        time_spent = self.elapsed_time
        notes = ""  # Add notes if needed

        self.sheet.append([date, unikey, counter_location, assistance_category, time_spent, notes])
        self.workbook.save("assistance_data.xlsx")

    import getpass

    def open_task_list(self):
        self.driver = webdriver.Chrome()

        try:
            # Navigate to the task list URL
            self.driver.get(task_list_url)

            self.driver.implicitly_wait(2)

            # Prompt user for username and password
            username_input = input("Enter your username: ")
            password_input = getpass.getpass("Enter your password: ")

            username = self.driver.find_element(By.ID, "input28")
            username.send_keys(username_input)

            self.driver.implicitly_wait(2)
            password = self.driver.find_element(By.ID, "input36")
            password.send_keys(password_input)
            print("Credentials entered")

            # Clicks the sign in button
            self.driver.implicitly_wait(2)

            sign_in = self.driver.find_element(By.CLASS_NAME, "o-form-button-bar")
            sign_in.click()

            self.driver.implicitly_wait(2)

            choose_push_nofif = self.driver.find_element(By.XPATH, '//*[@id="form61"]/div[2]/div/div[2]/div[2]/div[2]')
            choose_push_nofif.click()

            print(self.driver)

            # Open additional URLs in new tabs
            self.driver.execute_script("window.open('');")
            self.driver.switch_to.window(self.driver.window_handles[-1])
            self.driver.get("https://followme-print.sydney.edu.au:9192/app?service=page/UserList")

            sign_into_papercut = self.driver.find_element(By.XPATH, '//*[@id="inputUsername"]')
            sign_into_papercut.send_keys(username_input)

            sign_into_papercut1 = self.driver.find_element(By.XPATH, '//*[@id="inputPassword"]')
            sign_into_papercut1.send_keys(password_input)

            sign_into_papercut_btn = self.driver.find_element(By.XPATH, '//*[@id="login"]/input')
            sign_into_papercut_btn.click()

            self.driver.switch_to.window(self.driver.window_handles[0])
            # Switch back to the original tab
            time.sleep(10)
            self.driver.execute_script("window.open('');")
            self.driver.switch_to.window(self.driver.window_handles[-1])
            self.driver.get("https://iga.sydney.edu.au/ui/a/admin/identities/all-identities")
            self.driver.switch_to.window(self.driver.window_handles[0])

        except Exception as e:
            print("An error occurred while opening the task list:", str(e))

    def filter_print_refund_tickets(self):
        try:
            # Wait for the page to load completely
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # JavaScript to get all IDs, including those in iframes
            script = """
            function getAllIds(doc) {
                var elements = doc.getElementsByTagName('*');
                var ids = [];
                for (var i = 0; i < elements.length; i++) {
                    if (elements[i].id) {
                        ids.push(elements[i].id);
                    }
                }
                return ids;
            }

            var allIds = getAllIds(document);

            // Check iframes
            var iframes = document.getElementsByTagName('iframe');
            for (var i = 0; i < iframes.length; i++) {
                try {
                    allIds = allIds.concat(getAllIds(iframes[i].contentDocument));
                } catch(e) {
                    console.log('Cannot access iframe:', e);
                }
            }

            return allIds;
            """

            # Execute the JavaScript
            all_ids = self.driver.execute_script(script)

            # Process the IDs as per your requirement
            ids = []
            found_target = False
            for id in all_ids:
                if id == "snIAVisualList":
                    ids.append(id)
                    found_target = True
                    continue
                if found_target:
                    ids.append(id)
                    if len(ids) == 11:  # 1 for snIAVisualList + 10 following IDs
                        break

            # Print the results
            if ids:
                print("Found 'snIAVisualList' and the following 10 IDs:")
                for id in ids:
                    print(id)

                # Construct the XPath using the second-to-last element
                if len(ids) >= 2:
                    target_id = ids[-2]

                    # JavaScript to click all buttons under the hdr_{target_id} element, including in iframes
                    click_all_buttons_script = f"""
                    function clickAllButtons(doc, targetId) {{
                        var headerElement = doc.getElementById('hdr_' + targetId);
                        var clickedButtonIds = [];
                        if (headerElement) {{
                            var buttons = headerElement.getElementsByTagName('button');
                            for (var i = 0; i < buttons.length; i++) {{
                                if (buttons[i].id) {{
                                    buttons[i].click();
                                    clickedButtonIds.push(buttons[i].id);
                                }}
                            }}
                        }}
                        return clickedButtonIds;
                    }}

                    var allClickedButtonIds = clickAllButtons(document, '{target_id}');

                    // Check iframes
                    var iframes = document.getElementsByTagName('iframe');
                    for (var i = 0; i < iframes.length; i++) {{
                        try {{
                            allClickedButtonIds = allClickedButtonIds.concat(clickAllButtons(iframes[i].contentDocument, '{target_id}'));
                        }} catch(e) {{
                            console.log('Cannot access iframe:', e);
                        }}
                    }}

                    return allClickedButtonIds;
                    """

                    clicked_button_ids = self.driver.execute_script(click_all_buttons_script)

                    if clicked_button_ids:
                        print(f"Clicked {len(clicked_button_ids)} buttons under hdr_{target_id} (including in iframes)")
                        print("IDs of clicked buttons:")
                        for button_id in clicked_button_ids:
                            print(button_id)
                    else:
                        print(f"No buttons found under hdr_{target_id} (including in iframes)")

                    # Now click the specific button
                    drop_down_xpath = f'//*[@id="hdr_{target_id}"]/th[1]/span/button'
                    click_script = f"""
                    function clickElement(doc, xpath) {{
                        var element = doc.evaluate(xpath, doc, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        if (element) {{
                            element.click();
                            return element.id || 'no-id';
                        }}
                        return null;
                    }}

                    var clickedId = clickElement(document, '{drop_down_xpath}');
                    if (clickedId) {{
                        return clickedId;
                    }}

                    var iframes = document.getElementsByTagName('iframe');
                    for (var i = 0; i < iframes.length; i++) {{
                        try {{
                            clickedId = clickElement(iframes[i].contentDocument, '{drop_down_xpath}');
                            if (clickedId) {{
                                return clickedId;
                            }}
                        }} catch(e) {{
                            console.log('Cannot access iframe:', e);
                        }}
                    }}

                    return null;
                    """

                    clicked_id = self.driver.execute_script(click_script)

                    if clicked_id:
                        print(f"Clicked on element with XPath: {drop_down_xpath}")
                        print(f"ID of clicked element: {clicked_id}")
                    else:
                        print(f"Element with XPath: {drop_down_xpath} not found")

                    # JavaScript to count rows, list unique specific hrefs based on sys_id, and open them
                    count_list_and_open_script = f"""
                        function countRowsListAndOpenUniqueSpecificHrefs(doc, targetId) {{
                            var headerElement = doc.getElementById('hdr_' + targetId);
                            var result = {{rowCount: 0, specificHrefs: []}};
                            var uniqueSysIds = new Set();
                            if (headerElement) {{
                                var table = headerElement.closest('table');
                                if (table) {{
                                    // Count rows (excluding header row)
                                    var tbody = table.getElementsByTagName('tbody')[0];
                                    if (tbody) {{
                                        result.rowCount = tbody.getElementsByTagName('tr').length;
                                    }}

                                    // Find and open unique specific hrefs based on sys_id
                                    var links = table.getElementsByTagName('a');
                                    for (var i = 0; i < links.length; i++) {{
                                        var href = links[i].getAttribute('href');
                                        if (href && (href.includes('incident.do?sys_id=') || href.includes('sc_req_item.do?sys_id='))) {{
                                            var sysIdMatch = href.match(/(incident\\.do\\?sys_id=|sc_req_item\\.do\\?sys_id=)([^&]+)/);
                                            if (sysIdMatch && sysIdMatch[2]) {{
                                                var sysId = sysIdMatch[2];
                                                if (!uniqueSysIds.has(sysId)) {{
                                                    uniqueSysIds.add(sysId);
                                                    result.specificHrefs.push(href);
                                                    window.open(href, '_blank');
                                                }}
                                            }}
                                        }}
                                    }}
                                }}
                            }}
                            return result;
                        }}

                        var mainResult = countRowsListAndOpenUniqueSpecificHrefs(document, '{target_id}');
                        var totalRowCount = mainResult.rowCount;
                        var allSpecificHrefs = mainResult.specificHrefs;

                        // Check iframes
                        var iframes = document.getElementsByTagName('iframe');
                        for (var i = 0; i < iframes.length; i++) {{
                            try {{
                                var iframeResult = countRowsListAndOpenUniqueSpecificHrefs(iframes[i].contentDocument, '{target_id}');
                                totalRowCount += iframeResult.rowCount;
                                allSpecificHrefs = allSpecificHrefs.concat(iframeResult.specificHrefs);
                            }} catch(e) {{
                                console.log('Cannot access iframe:', e);
                            }}
                        }}

                        return {{rowCount: totalRowCount, specificHrefs: allSpecificHrefs}};
                        """

                    result = self.driver.execute_script(count_list_and_open_script)
                    row_count = result['rowCount']
                    specific_hrefs = result['specificHrefs']

                    print(f"Total number of rows in the table: {row_count}")

                    if specific_hrefs:
                        print(f"\nFound and opened {len(specific_hrefs)} unique specific links:")
                        for href in specific_hrefs:
                            print(href)
                    else:
                        print("\nNo unique incident or sc_req_item links with sys_id found in the table.")

                    # Store the original window handle
                    original_window = self.driver.current_window_handle

                    # Get all window handles
                    all_handles = self.driver.window_handles

                    # Enhanced JavaScript function to get all IDs, including those in iframes
                    get_all_ids_and_user_info_script = """
                    function getAllIdsAndUserInfo(doc) {
                        var elements = doc.getElementsByTagName('*');
                        var ids = [];
                        var callerIdValue = null;
                        var requestedForValue = null;
                        for (var i = 0; i < elements.length; i++) {
                            if (elements[i].id) {
                                ids.push(elements[i].id);
                                if (elements[i].id === 'sys_display.original.incident.caller_id') {
                                    callerIdValue = elements[i].value;
                                }
                                if (elements[i].id === 'sys_display.original.sc_req_item.request.requested_for') {
                                    requestedForValue = elements[i].value;
                                }
                            }
                        }
                        return {ids: ids, callerIdValue: callerIdValue, requestedForValue: requestedForValue};
                    }

                    function getAllIdsAndUserInfoIncludingIframes(doc) {
                        var result = getAllIdsAndUserInfo(doc);
                        var allIds = result.ids;
                        var callerIdValue = result.callerIdValue;
                        var requestedForValue = result.requestedForValue;

                        // Check iframes
                        var iframes = doc.getElementsByTagName('iframe');
                        for (var i = 0; i < iframes.length; i++) {
                            try {
                                var iframeResult = getAllIdsAndUserInfoIncludingIframes(iframes[i].contentDocument);
                                allIds = allIds.concat(iframeResult.ids);
                                if (!callerIdValue && iframeResult.callerIdValue) {
                                    callerIdValue = iframeResult.callerIdValue;
                                }
                                if (!requestedForValue && iframeResult.requestedForValue) {
                                    requestedForValue = iframeResult.requestedForValue;
                                }
                            } catch(e) {
                                console.log('Cannot access iframe:', e);
                            }
                        }

                        return {ids: allIds, callerIdValue: callerIdValue, requestedForValue: requestedForValue};
                    }

                    return getAllIdsAndUserInfoIncludingIframes(document);
                    """

                    def extract_user_id(full_string):
                        if full_string and " - " in full_string:
                            return full_string.split(" - ")[-1].strip()
                        return full_string  # Return the original string if " - " is not found

                    # Iterate through all opened tabs
                    for handle in all_handles:
                        if handle != original_window:
                            # Switch to the new tab
                            self.driver.switch_to.window(handle)

                            # Get the current URL
                            current_url = self.driver.current_url

                            # Wait for the page to load (you might need to adjust the timeout)
                            WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.TAG_NAME, "body"))
                            )

                            if "sc_req_item.do?sys_id=" in current_url or "incident.do?sys_id=" in current_url:
                                try:
                                    # Execute the JavaScript to get all IDs and user info
                                    result = self.driver.execute_script(get_all_ids_and_user_info_script)

                                    all_ids = result['ids']
                                    caller_id_value = result['callerIdValue']
                                    requested_for_value = result['requestedForValue']

                                    print(f"\nProcessing tab: {current_url}")
                                    print(f"Total IDs found: {len(all_ids)}")

                                    user_ids = []

                                    if caller_id_value:
                                        caller_user_id = extract_user_id(caller_id_value)
                                        user_ids.append(caller_user_id)
                                        print(f"Caller ID (user ID only): {caller_user_id}")

                                    if requested_for_value:
                                        requested_for_user_id = extract_user_id(requested_for_value)
                                        user_ids.append(requested_for_user_id)
                                        print(f"Requested For (user ID only): {requested_for_user_id}")

                                    if not user_ids:
                                        print("No user IDs found in this tab")
                                        continue

                                    for user_id in user_ids:
                                        try:
                                            # Open FollowMe Print UserList in a new tab
                                            self.driver.execute_script("window.open('');")
                                            self.driver.switch_to.window(self.driver.window_handles[-1])
                                            self.driver.get(
                                                "https://followme-print.sydney.edu.au:9192/app?service=page/UserList")

                                            # Wait for the page to load and the input field to be present
                                            WebDriverWait(self.driver, 10).until(
                                                EC.presence_of_element_located((By.XPATH, '//*[@id="quickFindAuto"]'))
                                            )

                                            # Find the input field and enter the user ID
                                            input_field = self.driver.find_element(By.XPATH, '//*[@id="quickFindAuto"]')
                                            input_field.send_keys(user_id)
                                            input_field.send_keys(Keys.RETURN)

                                            print(f"Opened FollowMe Print UserList for user: {user_id}")

                                            # Wait for the link to be clickable and then click it
                                            link_xpath = '//*[@id="content"]/div[1]/ul/li[4]/div/a/span[2]/span'
                                            WebDriverWait(self.driver, 10).until(
                                                EC.element_to_be_clickable((By.XPATH, link_xpath))
                                            )
                                            link = self.driver.find_element(By.XPATH, link_xpath)
                                            link.click()

                                            print(f"Clicked the specified link for user: {user_id}")

                                            # Wait for the print history to load
                                            WebDriverWait(self.driver, 10).until(
                                                EC.presence_of_element_located(
                                                    (By.XPATH, '//*[@id="content"]/div[2]/div[2]'))
                                            )

                                            # Get the print history content
                                            print_history_content = self.driver.find_element(By.XPATH,
                                                                                             '//*[@id="content"]/div[2]/div[2]').get_attribute(
                                                'outerHTML')

                                            # Open IGA page in a new tab
                                            self.driver.execute_script("window.open('');")
                                            self.driver.switch_to.window(self.driver.window_handles[-1])
                                            self.driver.get(
                                                "https://iga.sydney.edu.au/ui/a/admin/identities/all-identities")

                                            # Use JavaScript to find the input element and enter the user ID
                                            find_and_input_script = """
                                                                                    function findInputByPlaceholder(placeholder) {
                                                                                        var inputs = document.getElementsByTagName('input');
                                                                                        for (var i = 0; i < inputs.length; i++) {
                                                                                            if (inputs[i].placeholder === placeholder) {
                                                                                                return inputs[i];
                                                                                            }
                                                                                        }
                                                                                        return null;
                                                                                    }

                                                                                    var input = findInputByPlaceholder('Search Identities');
                                                                                    if (input) {
                                                                                        input.value = arguments[0];
                                                                                        input.dispatchEvent(new Event('input', { bubbles: true }));
                                                                                        input.dispatchEvent(new KeyboardEvent('keydown', {'key': 'Enter'}));
                                                                                        return true;
                                                                                    }
                                                                                    return false;
                                                                                    """

                                            # Wait for the page to load and the input to be available
                                            WebDriverWait(self.driver, 10).until(
                                                EC.presence_of_element_located((By.TAG_NAME, "body"))
                                            )

                                            # Execute the script to find and input the user ID
                                            success = self.driver.execute_script(find_and_input_script, user_id)

                                            if success:
                                                print(f"Entered user ID {user_id} in IGA search")
                                            else:
                                                print("Failed to find the search input in IGA")

                                            # Wait for 1 second
                                            time.sleep(1)

                                            # Click the specified element
                                            click_xpath = '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div/div[1]/div[2]/div[3]/div[2]/div/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span'
                                            WebDriverWait(self.driver, 10).until(
                                                EC.element_to_be_clickable((By.XPATH, click_xpath))
                                            )
                                            click_element = self.driver.find_element(By.XPATH, click_xpath)
                                            click_element.click()

                                            # Wait for the page to load after clicking the result
                                            WebDriverWait(self.driver, 10).until(
                                                EC.presence_of_element_located(
                                                    (By.XPATH, '//*[@id="single-spa-application:cloud-ui-admiral"]'))
                                            )
                                            time.sleep(3)
                                            print_all_text_script = """
                                            function getAllText() {
                                                return document.body.innerText;
                                            }
                                            return getAllText();
                                            """

                                            try:
                                                all_text = self.driver.execute_script(print_all_text_script)

                                                # Check if 'extroUID' exists in the text
                                                if 'extroUID' in all_text:
                                                    print("'extroUID' found in the page text.")
                                                else:
                                                    print("'extroUID' not found in the page text.")

                                            except Exception as e:
                                                print(f"Error extracting page text: {str(e)}")
                                            # First, check if ExtroUID exists on the page
                                            check_extro_uid_exists_script = """
                                            function checkExtroUIDExists() {
                                                var bodyText = document.body.innerText;
                                                return bodyText.toLowerCase().includes('extrouid');
                                            }
                                            return checkExtroUIDExists();
                                            """

                                            try:
                                                extro_uid_exists = self.driver.execute_script(
                                                    check_extro_uid_exists_script)
                                                if extro_uid_exists:
                                                    print("ExtroUID found on the page")

                                                    # If ExtroUID exists, proceed with extraction
                                                    extract_extro_uid_script = """
                                                    function findExtroUID() {
                                                        var attributes = document.getElementsByTagName('slpt-attribute');
                                                        for (var i = 0; i < attributes.length; i++) {
                                                            var attribute = attributes[i];
                                                            if (attribute.textContent.toLowerCase().includes('extrouid')) {
                                                                var span = attribute.querySelector('span');
                                                                if (span) {
                                                                    var number = span.textContent.trim().match(/\d+/);
                                                                    if (number) {
                                                                        return number[0];
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        return 'ExtroUID not found';
                                                    }
                                                    return findExtroUID();
                                                    """

                                                    extro_uid = self.driver.execute_script(extract_extro_uid_script)
                                                    print(f"ExtroUID for user {user_id}: {extro_uid}")
                                                else:
                                                    print("ExtroUID not found on the page")
                                                    extro_uid = "ExtroUID not present"
                                            except Exception as e:
                                                print(f"Error checking/extracting ExtroUID: {str(e)}")
                                                extro_uid = "ExtroUID extraction failed"

                                            additional_info = f"ExtroUID: {extro_uid}"

                                            # Switch back to the original ServiceNow tab
                                            self.driver.switch_to.window(handle)

                                            # Create a floating window with the print history content and additional info
                                            create_floating_window_script = f"""
                                            var floatingWindow = document.createElement('div');
                                            floatingWindow.style.position = 'fixed';
                                            floatingWindow.style.top = '20px';
                                            floatingWindow.style.right = '20px';
                                            floatingWindow.style.width = '600px';
                                            floatingWindow.style.height = '80%';
                                            floatingWindow.style.minWidth = '300px';
                                            floatingWindow.style.minHeight = '200px';
                                            floatingWindow.style.backgroundColor = '#f4f4f4';
                                            floatingWindow.style.border = '1px solid #ddd';
                                            floatingWindow.style.zIndex = '9999';
                                            floatingWindow.style.overflow = 'auto';
                                            floatingWindow.style.padding = '10px';
                                            floatingWindow.style.boxShadow = '-2px 0 5px rgba(0,0,0,0.1)';
                                            floatingWindow.style.fontFamily = 'Arial, sans-serif';
                                            floatingWindow.style.resize = 'both';

                                            floatingWindow.innerHTML = `
                                                <style>
                                                    .print-history-header {{
                                                        background-color: #2c3e50;
                                                        color: white;
                                                        padding: 10px;
                                                        margin: -10px -10px 10px -10px;
                                                        display: flex;
                                                        justify-content: space-between;
                                                        align-items: center;
                                                        cursor: move;
                                                    }}
                                                    .print-history-title {{
                                                        margin: 0;
                                                        font-size: 18px;
                                                    }}
                                                    .print-history-close {{
                                                        background: none;
                                                        border: none;
                                                        color: white;
                                                        font-size: 20px;
                                                        cursor: pointer;
                                                    }}
                                                    .print-history-table {{
                                                        width: 100%;
                                                        border-collapse: collapse;
                                                    }}
                                                    .print-history-table th {{
                                                        background-color: #34495e;
                                                        color: white;
                                                        text-align: left;
                                                        padding: 8px;
                                                    }}
                                                    .print-history-table td {{
                                                        background-color: white;
                                                        border-bottom: 1px solid #ddd;
                                                        padding: 8px;
                                                    }}
                                                    .print-history-table tr:hover td {{
                                                        background-color: #f5f5f5;
                                                    }}
                                                    .status-printed {{
                                                        color: #27ae60;
                                                    }}
                                                    .status-refund {{
                                                        color: #2980b9;
                                                    }}
                                                    .export-buttons {{
                                                        margin-top: 10px;
                                                    }}
                                                    .export-button {{
                                                        background-color: #3498db;
                                                        color: white;
                                                        border: none;
                                                        padding: 5px 10px;
                                                        margin-right: 5px;
                                                        cursor: pointer;
                                                    }}
                                                    .extro-uid {{
                                                        cursor: pointer;
                                                        text-decoration: underline;
                                                        color: #3498db;
                                                    }}
                                                    .extro-uid:hover {{
                                                        color: #2980b9;
                                                    }}
                                                    .tooltip {{
                                                        position: absolute;
                                                        background-color: #333;
                                                        color: #fff;
                                                        padding: 5px;
                                                        border-radius: 3px;
                                                        font-size: 12px;
                                                        display: none;
                                                    }}
                                                </style>
                                                <div class="print-history-header">
                                                    <h2 class="print-history-title">Print History for {user_id} - <span class="extro-uid" title="Click to copy">{extro_uid}</span></h2>
                                                    <button class="print-history-close" onclick="this.closest('div[style]').remove();">×</button>
                                                </div>
                                                <div class="tooltip">Copied!</div>
                                                <div class="print-history-content">
                                                    {print_history_content}
                                                </div>
                                                <div class="export-buttons">
                                                    <button class="export-button">PDF</button>
                                                    <button class="export-button">CSV</button>
                                                    <button class="export-button">Excel</button>
                                                </div>
                                            `;

                                            document.body.appendChild(floatingWindow);

                                            // Make the window draggable
                                            var header = floatingWindow.querySelector('.print-history-header');
                                            var isDragging = false;
                                            var dragOffsetX, dragOffsetY;

                                            header.addEventListener('mousedown', function(e) {{
                                                isDragging = true;
                                                dragOffsetX = e.clientX - floatingWindow.offsetLeft;
                                                dragOffsetY = e.clientY - floatingWindow.offsetTop;
                                            }});

                                            document.addEventListener('mousemove', function(e) {{
                                                if (isDragging) {{
                                                    floatingWindow.style.left = (e.clientX - dragOffsetX) + 'px';
                                                    floatingWindow.style.top = (e.clientY - dragOffsetY) + 'px';
                                                    floatingWindow.style.right = 'auto';
                                                }}
                                            }});

                                            document.addEventListener('mouseup', function() {{
                                                isDragging = false;
                                            }});

                                            // Ensure the window stays within the viewport
                                            function adjustPosition() {{
                                                var rect = floatingWindow.getBoundingClientRect();
                                                if (rect.right > window.innerWidth) {{
                                                    floatingWindow.style.left = (window.innerWidth - rect.width) + 'px';
                                                }}
                                                if (rect.bottom > window.innerHeight) {{
                                                    floatingWindow.style.top = (window.innerHeight - rect.height) + 'px';
                                                }}
                                                if (rect.left < 0) {{
                                                    floatingWindow.style.left = '0px';
                                                }}
                                                if (rect.top < 0) {{
                                                    floatingWindow.style.top = '0px';
                                                }}
                                            }}

                                            // Add event listener for resize
                                            window.addEventListener('resize', adjustPosition);

                                            // Periodically check and adjust position (for when the window is resized by the user)
                                            setInterval(adjustPosition, 100);

                                            // Apply additional styling to the existing content
                                            var table = floatingWindow.querySelector('table');
                                            if (table) {{
                                                table.className = 'print-history-table';
                                                var statusCells = table.querySelectorAll('td:last-child');
                                                statusCells.forEach(cell => {{
                                                    if (cell.textContent.trim().toLowerCase() === 'printed') {{
                                                        cell.className = 'status-printed';
                                                    }} else if (cell.textContent.trim().toLowerCase().includes('refund')) {{
                                                        cell.className = 'status-refund';
                                                    }}
                                                }});
                                            }}

                                            // Add click-to-copy functionality
                                            var extroUidSpan = floatingWindow.querySelector('.extro-uid');
                                            var tooltip = floatingWindow.querySelector('.tooltip');

                                            extroUidSpan.addEventListener('click', function() {{
                                                var textArea = document.createElement("textarea");
                                                textArea.value = this.textContent;
                                                document.body.appendChild(textArea);
                                                textArea.select();
                                                document.execCommand("Copy");
                                                textArea.remove();

                                                // Show tooltip
                                                tooltip.style.display = 'block';
                                                tooltip.style.left = (this.offsetLeft + this.offsetWidth / 2 - tooltip.offsetWidth / 2) + 'px';
                                                tooltip.style.top = (this.offsetTop + this.offsetHeight + 5) + 'px';

                                                // Hide tooltip after 2 seconds
                                                setTimeout(function() {{
                                                    tooltip.style.display = 'none';
                                                }}, 2000);
                                            }});
                                            """

                                            self.driver.execute_script(create_floating_window_script)

                                            print(f"Created floating window with print history for user: {user_id}")

                                            # Close the FollowMe Print and IGA tabs
                                            for _ in range(2):
                                                self.driver.switch_to.window(self.driver.window_handles[-1])
                                                self.driver.close()

                                            # Switch back to the original ServiceNow tab
                                            self.driver.switch_to.window(handle)

                                        except Exception as e:
                                            print(f"Error processing user: {user_id}")
                                            print(str(e))

                                except Exception as e:
                                    print(f"Error processing tab: {current_url}")
                                    print(str(e))

                    # Switch back to the original window
                    self.driver.switch_to.window(original_window)

                    print("Finished processing all tabs.")
                    return ids

        except Exception as e:
            print("An error occurred:", e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SnowSoftwareWindow()
    window.show()
    sys.exit(app.exec_())