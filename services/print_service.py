from playwright.sync_api import sync_playwright, expect
from PyQt5.QtWidgets import QInputDialog, QMessageBox
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time

def reprint(self):
    unikey, ok = QInputDialog.getText(self, "Enter Unikey", "Please enter the Unikey:")
    if not ok or not unikey:
        return

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else browser.new_context()

            reprint_page = context.new_page()
            reprint_page.goto(
                "https://sydneyuni.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D3c714f09dbe080502d38cae43a9619cd%26sysparm_link_parent%3D5fbc29844fba1fc05ad9d0311310c75d%26sysparm_catalog%3D09a851b34faadbc05ad9d0311310c7e7%26sysparm_catalog_view%3Dsm_cat_categories%26sysparm_view%3Dtext_search")

            reprint_page.wait_for_load_state("domcontentloaded")

            try:
                frame = reprint_page.frame_locator("iframe[name=\"gsft_main\"]").first
                select_locator = frame.locator("select[name=\"IO\\:1352c389dbe080502d38cae43a96194c\"]")
                expect(select_locator).to_be_visible(timeout=1000)
            except Exception as e:
                QMessageBox.warning(self, "Login Required", "You should be logged into SNow first. You may be brought to the SSO screen by the Counter Stuff button in the Main Menu")
                return

            unikey_input = frame.locator("#sys_display\\.IO\\:35028389dbe080502d38cae43a961977")
            unikey_input.fill(unikey)
            frame.locator("select[name=\"IO\\:1352c389dbe080502d38cae43a96194c\"]").select_option(
                "Printing")
            frame.locator("select[name=\"IO\\:9eae32dedb5dc8502d38cae43a961914\"]").select_option(
                "Reprint")

            additional_details = frame.get_by_role("textbox", name="Additional details")
            additional_details.fill(
                f"Assisted with reprinting for {unikey}.")
            print("Successfully interacted with all form elements")
            QMessageBox.information(self, "Process Complete",
                                    "The reprint form has been completed. The pages will remain open for your review.")

        except Exception as e:
            QMessageBox.warning(self, "Error",
                                f"An error occurred: {str(e)}")
        finally:
            context.close()

    print("Reprint process completed")

def reprint_hold_for_auth(self):
    unikey, ok = QInputDialog.getText(self, "Enter Unikey", "Please enter the Unikey:")
    if not ok or not unikey:
        return

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else browser.new_context()

            reprint_page = context.new_page()
            reprint_page.goto(
                "https://sydneyuni.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D3c714f09dbe080502d38cae43a9619cd%26sysparm_link_parent%3D5fbc29844fba1fc05ad9d0311310c75d%26sysparm_catalog%3D09a851b34faadbc05ad9d0311310c7e7%26sysparm_catalog_view%3Dsm_cat_categories%26sysparm_view%3Dtext_search")

            reprint_page.wait_for_load_state("domcontentloaded")

            try:
                frame = reprint_page.frame_locator("iframe[name=\"gsft_main\"]").first
                select_locator = frame.locator("select[name=\"IO\\:1352c389dbe080502d38cae43a96194c\"]")
                expect(select_locator).to_be_visible(timeout=1000)
            except Exception as e:
                QMessageBox.warning(self, "Login Required", "You should be logged into SNow first. You may be brought to the SSO screen by the Counter Stuff button in the Main Menu")
                return

            unikey_input = frame.locator("#sys_display\\.IO\\:35028389dbe080502d38cae43a961977")
            unikey_input.fill(unikey)
            frame.locator("select[name=\"IO\\:1352c389dbe080502d38cae43a96194c\"]").select_option(
                "Printing")
            frame.locator("select[name=\"IO\\:9eae32dedb5dc8502d38cae43a961914\"]").select_option(
                "Assistance")

            additional_details = frame.get_by_role("textbox", name="Additional details")
            additional_details.fill(
                f"Assisted with guiding through process of inputting correct login details to resolve hold for authentication error that occurs on mac device for {unikey}.")
            print("Successfully interacted with all form elements")
            QMessageBox.information(self, "Process Complete",
                                    "The reprint form has been completed. The pages will remain open for your review.")

        except Exception as e:
            QMessageBox.warning(self, "Error",
                                f"An error occurred: {str(e)}")
        finally:
            context.close()

    print("Reprint hold for auth process completed")

def filter_print_refund_tickets(self):
    if self.driver is None or not self.is_driver_alive():
        reply = QMessageBox.question(self, "Reopen Task List",
                                     "The task list is not open. Would you like to reopen it?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            if not self.open_task_list():
                return
        else:
            return
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
                                        click_xpath = '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span'
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
        QMessageBox.warning(self, "Error", f"An error occurred while filtering refund tickets: {str(e)}")