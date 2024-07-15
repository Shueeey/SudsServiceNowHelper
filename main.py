import platform
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QComboBox, QLineEdit, QListWidget
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QTimer, Qt
import openpyxl
from datetime import datetime
import asyncio
from playwright.async_api import async_playwright
from threading import Event

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

        # Initialize the timer variables
        self.elapsed_time = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_timer)

        # Initialize the Excel workbook and sheet
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "Assistance Data"
        self.sheet.append(["Date", "Unikey", "Counter Location", "Assistance Category", "Time Spent (minutes)", "Notes"])

        self.browser = None
        self.page = None

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

    def open_task_list(self):
        asyncio.get_event_loop().run_until_complete(self._open_task_list_async())

    async def _open_task_list_async(self):
        async with async_playwright() as p:
            browser = await p.chromium.connect_over_cdp("http://localhost:9222")
            page = await browser.new_page()

            try:
                # Navigate to the task list URL
                await page.goto(task_list_url)

                # Wait for 2 seconds
                await page.wait_for_timeout(700)

                # Enter username
                await page.fill('#input28', 'spie2381')

                # Wait for 2 seconds
                await page.wait_for_timeout(700)

                # Enter password
                await page.fill('#input36', 'Ninsasin2002')

                # Wait for 2 seconds
                await page.wait_for_timeout(700)

                # Click the sign-in button
                await page.click('.o-form-button-bar')

                # Wait for 2 seconds
                await page.wait_for_timeout(700)

                # Click the push notification option
                await page.click('//*[@id="form61"]/div[2]/div/div[2]/div[2]/div[2]/a')

                print("Task list opened successfully!")

                self.browser = browser
                self.page = page

                # Perform the filter_print_refund_tickets function after the page has finished logging in and loading
                await page.wait_for_timeout(5000)

                await self.filter_print_refund_tickets()

                await page.pause()

            except Exception as e:
                print("An error occurred while opening the task list:", str(e))


    async def filter_print_refund_tickets(self):
        #open the task list

        await self._open_task_list_async()
        #get the list of all the children of the snIAVisualList
        children = await self.page.query_selector_all('//*[@id="snIAVisualList"]/div')
        for child in children:
            print(child)
        #get the text of all the children
        for child in children:
            print(await child.text_content())






    # def closeEvent(self, event):
    #     # Close the browser when the application is closed
    #     if self.browser:
    #         asyncio.get_event_loop().run_until_complete(self.browser.close())
    #     event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SnowSoftwareWindow()
    window.show()
    
    sys.exit(app.exec_())
