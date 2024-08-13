from playwright.sync_api import sync_playwright, expect
from PyQt5.QtWidgets import QInputDialog, QMessageBox


def run_okta_resets_for_new_phone(parent):
    unikey, ok = QInputDialog.getText(parent, "Enter Unikey", "Please enter the Unikey:")
    if not ok or not unikey:
        return

    extro_uid_dict = {}

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else browser.new_context()

            # IGA part
            iga_page = context.new_page()
            iga_page.goto("https://iga.sydney.edu.au/ui/a/admin/identities/all-identities")

            # Wait for the search input to be available and use it
            iga_page.wait_for_selector("input[placeholder='Search Identities']")
            iga_page.fill("input[placeholder='Search Identities']", unikey)
            iga_page.press("input[placeholder='Search Identities']", "Enter")

            # Check for the "no results" message
            no_results_selector = '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/app-identities-list-empty-state-container/section/span'

            try:
                iga_page.wait_for_selector(no_results_selector, timeout=1000)
                no_results_message = iga_page.inner_text(no_results_selector)
                if "We couldn't find anything that matches your query. Please try again." in no_results_message:
                    QMessageBox.critical(parent, "Error", f"The Unikey '{unikey}' doesn't exist in IGA.")
                    return
            except Exception:
                # If the selector is not found, it means results were found, so we continue
                pass

            # Wait for and click the search result
            iga_page.wait_for_selector(
                '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span')
            iga_page.click(
                '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span')

            # Extract user information
            extro_uid_dict = extract_user_info(iga_page)

            iga_page.close()

            # Okta part
            snow_page = context.new_page()
            snow_page.goto(
                "https://sydneyuni.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D3c714f09dbe080502d38cae43a9619cd%26sysparm_link_parent%3D5fbc29844fba1fc05ad9d0311310c75d%26sysparm_catalog%3D09a851b34faadbc05ad9d0311310c7e7%26sysparm_catalog_view%3Dsm_cat_categories%26sysparm_view%3Dtext_search")

            snow_page.wait_for_load_state("domcontentloaded")

            try:
                frame = snow_page.frame_locator("iframe[name=\"gsft_main\"]").first
                select_locator = frame.locator("select[name=\"IO\\:1352c389dbe080502d38cae43a96194c\"]")
                expect(select_locator).to_be_visible(timeout=1000)
            except Exception as e:
                QMessageBox.warning(parent, "Login Required",
                                    "You should be logged into SNow first. You may be brought to the SSO screen by the Counter Stuff button in the Main Menu")
                return

            fill_okta_form(frame, unikey, "Factor Reset (new phone)")

            create_floating_info_window(snow_page, unikey, extro_uid_dict)

            print("Successfully interacted with all form elements")
            QMessageBox.information(parent, "Process Complete",
                                    "The Okta reset process has been completed. The pages will remain open for your review.")

        except Exception as e:
            QMessageBox.warning(parent, "Error", f"An error occurred: {str(e)}")
        finally:
            context.close()

    print("Okta reset process completed")


def run_okta_resets_for_deleted_app(parent):
    unikey, ok = QInputDialog.getText(parent, "Enter Unikey", "Please enter the Unikey:")
    if not ok or not unikey:
        return

    extro_uid_dict = {}

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else browser.new_context()

            # IGA part
            iga_page = context.new_page()
            iga_page.goto("https://iga.sydney.edu.au/ui/a/admin/identities/all-identities")

            # Wait for the search input to be available and use it
            iga_page.wait_for_selector("input[placeholder='Search Identities']")
            iga_page.fill("input[placeholder='Search Identities']", unikey)
            iga_page.press("input[placeholder='Search Identities']", "Enter")

            # Check for the "no results" message
            no_results_selector = '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/app-identities-list-empty-state-container/section/span'

            try:
                iga_page.wait_for_selector(no_results_selector, timeout=1000)
                no_results_message = iga_page.inner_text(no_results_selector)
                if "We couldn't find anything that matches your query. Please try again." in no_results_message:
                    QMessageBox.critical(parent, "Error", f"The Unikey '{unikey}' doesn't exist in IGA.")
                    return
            except Exception:
                # If the selector is not found, it means results were found, so we continue
                pass

            # Wait for and click the search result
            iga_page.wait_for_selector(
                '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span')
            iga_page.click(
                '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span')

            # Extract user information
            extro_uid_dict = extract_user_info(iga_page)

            iga_page.close()

            # Okta part
            snow_page = context.new_page()
            snow_page.goto(
                "https://sydneyuni.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D3c714f09dbe080502d38cae43a9619cd%26sysparm_link_parent%3D5fbc29844fba1fc05ad9d0311310c75d%26sysparm_catalog%3D09a851b34faadbc05ad9d0311310c7e7%26sysparm_catalog_view%3Dsm_cat_categories%26sysparm_view%3Dtext_search")

            snow_page.wait_for_load_state("domcontentloaded")

            try:
                frame = snow_page.frame_locator("iframe[name=\"gsft_main\"]").first
                select_locator = frame.locator("select[name=\"IO\\:1352c389dbe080502d38cae43a96194c\"]")
                expect(select_locator).to_be_visible(timeout=1000)
            except Exception as e:
                QMessageBox.warning(parent, "Login Required",
                                    "You should be logged into SNow first. You may be brought to the SSO screen by the Counter Stuff button in the Main Menu")
                return

            fill_okta_form(frame, unikey, "Factor Reset (deleted the app)")

            create_floating_info_window(snow_page, unikey, extro_uid_dict)

            print("Successfully interacted with all form elements")
            QMessageBox.information(parent, "Process Complete",
                                    "The Okta reset process has been completed. The pages will remain open for your review.")

        except Exception as e:
            QMessageBox.warning(parent, "Error", f"An error occurred: {str(e)}")
        finally:
            context.close()

    print("Okta reset process completed")

def run_okta_resets_for_deleted_app_with_card(parent):
    card_number, ok = QInputDialog.getText(parent, "Scan number", "Please scan the card:")
    if not ok or not card_number:
        return

    extro_uid_dict = {}

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else browser.new_context()

            # PaperCut part
            ppc_page = context.new_page()
            ppc_page.goto("https://followme-print.sydney.edu.au:9192/app?service=page/UserList")

            #if the element with id loginBody appears, suggest the user to login to papercut first before pressing button
            try:
                ppc_page.wait_for_selector("#loginBody", timeout=1000)
                QMessageBox.warning(parent, "Login Required",
                                    "You should be logged into PaperCut first. You may be brought to the SSO screen by the Counter Stuff button in the Main Menu")
                return
            except Exception:
                # If the selector is not found, it means results were found, so we continue
                pass

            # Wait for the search input to be available and use it
            ppc_page.wait_for_selector("input[placeholder='Search e.g. user/card/email']")
            ppc_page.fill("input[placeholder='Search e.g. user/card/email']", card_number)
            ppc_page.press("input[placeholder='Search e.g. user/card/email']", "Enter")

            #the text that is at id="username" is the unikey
            ppc_page.wait_for_selector('//*[@id="content"]/div[2]/div[2]/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td/p[1]/span[2]')
            unikey = ppc_page.inner_text('//*[@id="content"]/div[2]/div[2]/table/tbody/tr[2]/td[2]/form/table/tbody/tr[1]/td/p[1]/span[2]')

            ppc_page.close()

            # IGA part
            iga_page = context.new_page()
            iga_page.goto("https://iga.sydney.edu.au/ui/a/admin/identities/all-identities")

            # Wait for the search input to be available and use it
            iga_page.wait_for_selector("input[placeholder='Search Identities']")
            iga_page.fill("input[placeholder='Search Identities']", unikey)
            iga_page.press("input[placeholder='Search Identities']", "Enter")

            # Check for the "no results" message
            no_results_selector = '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/app-identities-list-empty-state-container/section/span'

            try:
                iga_page.wait_for_selector(no_results_selector, timeout=1000)
                no_results_message = iga_page.inner_text(no_results_selector)
                if "We couldn't find anything that matches your query. Please try again." in no_results_message:
                    QMessageBox.critical(parent, "Error", f"The Unikey '{unikey}' doesn't exist in IGA.")
                    return
            except Exception:
                # If the selector is not found, it means results were found, so we continue
                pass

            # Wait for and click the search result
            iga_page.wait_for_selector(
                '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span')
            iga_page.click(
                '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span')

            # Extract user information
            extro_uid_dict = extract_user_info(iga_page)

            iga_page.close()

            # SNow part
            snow_page = context.new_page()
            snow_page.goto(
                "https://sydneyuni.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D3c714f09dbe080502d38cae43a9619cd%26sysparm_link_parent%3D5fbc29844fba1fc05ad9d0311310c75d%26sysparm_catalog%3D09a851b34faadbc05ad9d0311310c7e7%26sysparm_catalog_view%3Dsm_cat_categories%26sysparm_view%3Dtext_search")

            snow_page.wait_for_load_state("domcontentloaded")

            try:
                frame = snow_page.frame_locator("iframe[name=\"gsft_main\"]").first
                select_locator = frame.locator("select[name=\"IO\\:1352c389dbe080502d38cae43a96194c\"]")
                expect(select_locator).to_be_visible(timeout=1000)
            except Exception as e:
                QMessageBox.warning(parent, "Login Required",
                                    "You should be logged into SNow first. You may be brought to the SSO screen by the Counter Stuff button in the Main Menu")
                return

            fill_okta_form(frame, unikey, "Factor Reset (deleted the app)")

            # Okta part
            okta_page = context.new_page()
            okta_page.goto("https://sydneyuni-admin.okta.com/admin/dashboard")

        # if the element with id login28 appears, suggest the user to first login to okta
            try:
                okta_page.wait_for_selector("#login28", timeout=1000)
                QMessageBox.warning(parent, "Login Required",
                                    "You should be logged into Okta first. You may be brought to the SSO screen by the Counter Stuff button in the Main Menu")
                return
            except Exception:
                # If the selector is not found, it means results were found, so we continue
                pass
            create_floating_info_window(okta_page, unikey, extro_uid_dict)
            okta_page.wait_for_selector("#spotlight-search-bar")
            okta_page.fill("#spotlight-search-bar", unikey)
            okta_page.press("#spotlight-search-bar", "Enter")
            okta_page.wait_for_selector("spotlight-user-result-item-TOP_RESULT-00ublc0y5g9mrhuGg3l6")
            okta_page.click("spotlight-user-result-item-TOP_RESULT-00ublc0y5g9mrhuGg3l6")

            QMessageBox.information(parent, "Process Complete",
                                    "The Okta reset process has been completed. The pages will remain open for your review.")

        except Exception as e:
                QMessageBox.warning(parent, "Error", f"An error occurred: {str(e)}")
        finally:
                context.close()

    print("Okta reset process completed")

def extract_user_info(page):
    extro_uid_dict = {}

    # Wait for the page to load and extract ExtroUID
    page.wait_for_selector("//slpt-attribute[contains(., 'extroUID')]//span")
    extro_uid_element = page.query_selector("//slpt-attribute[contains(., 'extroUID')]//span")
    extro_uid = extro_uid_element.inner_text().strip() if extro_uid_element else "Not found"
    extro_uid_dict['extrouid'] = extro_uid

    # Extract other user information
    info_fields = {
        'studentid': "Student ID",
        'dob': "DOB",
        'studentdegreecode': "Student Degree Code",
        'uos': "UOS",
        'personalemail': "Personal Email"
    }

    for key, label in info_fields.items():
        page.wait_for_selector(f"//slpt-attribute[contains(., '{label}')]//span")
        element = page.query_selector(f"//slpt-attribute[contains(., '{label}')]//span")
        extro_uid_dict[key] = element.inner_text().strip() if element else "Not found"

    return extro_uid_dict


def fill_okta_form(frame, unikey, reset_type):
    unikey_input = frame.locator("#sys_display\\.IO\\:35028389dbe080502d38cae43a961977")
    unikey_input.fill(unikey)
    frame.locator("select[name=\"IO\\:1352c389dbe080502d38cae43a96194c\"]").select_option("Unikey/Okta Assistance")
    frame.locator("select[name=\"IO\\:d68099e6db29c4509909abf34a961949\"]").select_option(reset_type)

    additional_details = frame.get_by_role("textbox", name="Additional details")
    additional_details.fill(f"Assisted with reassigning Okta MFA onto phone after {reset_type.lower()}.")


def create_floating_info_window(page, unikey, extro_uid_dict):
    iframe_script = f"""
    // Create a wrapper div for the iframe
    var wrapper = document.createElement('div');
    wrapper.style.position = 'fixed';
    wrapper.style.top = '300px';
    wrapper.style.right = '10px';
    wrapper.style.width = '300px';
    wrapper.style.height = '320px';  // Increased to accommodate drag handle
    wrapper.style.zIndex = '9999';

    // Create drag handle
    var dragHandle = document.createElement('div');
    dragHandle.textContent = '{unikey} Student Information (Drag to move)';
    dragHandle.style.backgroundColor = '#007bff';
    dragHandle.style.color = 'white';
    dragHandle.style.padding = '5px';
    dragHandle.style.cursor = 'move';
    dragHandle.style.userSelect = 'none';

    // Create iframe
    var iframe = document.createElement('iframe');
    iframe.style.width = '100%';
    iframe.style.height = 'calc(100% - 30px)';  // Subtract height of drag handle
    iframe.style.border = '1px solid #ccc';
    iframe.style.borderRadius = '0 0 5px 5px';
    iframe.style.backgroundColor = '#f9f9f9';
    iframe.style.boxShadow = '0 0 10px rgba(0,0,0,0.1)';

    var iframeContent = `
        <html>
            <head>
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        padding: 10px;
                        margin: 0;
                    }}
                    h3 {{
                        margin: 10px 0 5px 0;
                        color: #333;
                        font-size: 14px;
                    }}
                    p {{
                        margin: 0 0 10px 0;
                        font-size: 14px;
                        font-weight: bold;
                        color: #007bff;
                    }}
                </style>
            </head>
            <body>
                <h3>ExtroUID</h3>
                <p>{extro_uid_dict['extrouid']}</p>
                <h3>Student ID</h3>
                <p>{extro_uid_dict['studentid']}</p>
                <h3>DOB</h3>
                <p>{extro_uid_dict['dob']}</p>
                <h3>Personal Email</h3>
                <p>{extro_uid_dict['personalemail']}</p>
                <h3>Student Degree Code</h3>
                <p>{extro_uid_dict['studentdegreecode']}</p>
                <h3>UOS</h3>
                <p>{extro_uid_dict['uos']}</p>
            </body>
        </html>
    `;

    // Append elements
    wrapper.appendChild(dragHandle);
    wrapper.appendChild(iframe);
    document.body.appendChild(wrapper);

    iframe.contentWindow.document.open();
    iframe.contentWindow.document.write(iframeContent);
    iframe.contentWindow.document.close();

    // Make wrapper draggable
    var isDragging = false;
    var currentX;
    var currentY;
    var initialX;
    var initialY;
    var xOffset = 0;
    var yOffset = 0;

    function dragStart(e) {{
        if (e.target === dragHandle) {{
            initialX = e.clientX - xOffset;
            initialY = e.clientY - yOffset;
            isDragging = true;
        }}
    }}

    function dragEnd(e) {{
        initialX = currentX;
        initialY = currentY;
        isDragging = false;
    }}

    function drag(e) {{
        if (isDragging) {{
            e.preventDefault();
            currentX = e.clientX - initialX;
            currentY = e.clientY - initialY;
            xOffset = currentX;
            yOffset = currentY;
            setTranslate(currentX, currentY, wrapper);
        }}
    }}

    function setTranslate(xPos, yPos, el) {{
        el.style.transform = "translate3d(" + xPos + "px, " + yPos + "px, 0)";
    }}

    document.addEventListener("mousedown", dragStart, false);
    document.addEventListener("mouseup", dragEnd, false);
    document.addEventListener("mousemove", drag, false);
    """

    page.evaluate(iframe_script)

