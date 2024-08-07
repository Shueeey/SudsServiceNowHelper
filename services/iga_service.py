from playwright.sync_api import sync_playwright, expect
from PyQt5.QtWidgets import QMessageBox

def search_user_in_iga(unikey):
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else browser.new_context()

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
                    return None, f"The Unikey '{unikey}' doesn't exist in IGA."
            except Exception:
                # If the selector is not found, it means results were found, so we continue
                pass

            # Wait for and click the search result
            result_selector = '//*[@id="single-spa-application:cloud-ui-admiral"]/app-cloud-ui-admiral-root/app-identities-list-page/div/div/app-identities-list/div/div/div/slpt-composite-card-grid/div/slpt-composite-data-grid/div/div[1]/div/slpt-data-grid/ag-grid-angular/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div/div[1]/slpt-data-grid-link-cell/slpt-link/a/div/span'
            iga_page.wait_for_selector(result_selector)
            iga_page.click(result_selector)

            # Extract user information
            user_info = extract_user_info(iga_page)

            iga_page.close()
            context.close()

            return user_info, None

        except Exception as e:
            return None, f"An error occurred while searching for the user: {str(e)}"


def extract_user_info(page):
    user_info = {}

    # Define the fields to extract
    fields = {
        'extrouid': "extroUID",
        'studentid': "Student ID",
        'dob': "DOB",
        'studentdegreecode': "Student Degree Code",
        'uos': "UOS",
        'personalemail': "Personal Email"
    }

    for key, label in fields.items():
        selector = f"//slpt-attribute[contains(., '{label}')]//span"
        try:
            page.wait_for_selector(selector, timeout=5000)
            element = page.query_selector(selector)
            user_info[key] = element.inner_text().strip() if element else "Not found"
        except Exception:
            user_info[key] = "Not found"

    return user_info


