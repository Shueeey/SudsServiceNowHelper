import platform
import sys
import subprocess
from datetime import datetime
import time
import requests

from PyQt5.QtWidgets import QWidget, QPushButton, QLabel, QVBoxLayout, QComboBox, QLineEdit, QListWidget, \
    QDialog, QMessageBox, QInputDialog, QFrame, QHBoxLayout, QStackedWidget
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QTimer, Qt

from playwright.sync_api import Playwright, sync_playwright, expect
from playwright.sync_api import expect, TimeoutError as PlaywrightTimeoutError

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from ui.login_dialog import LoginDialog
from utils.browser_utils import open_url, is_chrome_debugger_running, start_chrome_debugging
from utils.excel_utils import initialize_workbook, save_to_excel
from services.okta_service import run_okta_resets_for_new_phone, run_okta_resets_for_deleted_app
from services.print_service import reprint, reprint_hold_for_auth, filter_print_refund_tickets

# URL constants
url = "https://sydneytest.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D3c714f09dbe080502d38cae43a9619cd%26sysparm_link_parent%3D5fbc29844fba1fc05ad9d0311310c75d%26sysparm_catalog%3D09a851b34faadbc05ad9d0311310c7e7%26sysparm_catalog_view%3Dsm_cat_categories%26sysparm_view%3Dtext_search"
task_list_url = "https://sydneytest.service-now.com/nav_to.do?uri=%2F$interactive_analysis.do%3Fsysparm_field%3Dopened_at%26sysparm_table%3Dtask%26sysparm_from_list%3Dtrue%26sysparm_query%3Dactive%3Dtrue%5Estate!%3D6%5Eassignment_group%3Djavascript:getMyGroups()%5Eassigned_to%3D%26sysparm_list_view%3Dcatalog"

class SnowSoftwareWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sud's Snow Software")
        self.setGeometry(100, 100, 500, 700)
        self.setStyleSheet("""
            QWidget {
                background-color: #1c2432;
                color: #ffffff;
                font-family: Arial, sans-serif;
            }
            QPushButton {
                background-color: #2f3c4e;
                border: none;
                color: #a7b3c6;
                padding: 10px 20px;
                border-radius: 3px;
                font-size: 14px;
                text-align: left;
            }
            QPushButton:hover {
                background-color: #3b4b61;
            }
            QPushButton:pressed {
                background-color: #4a5d78;
            }
            QLabel {
                font-size: 14px;
                color: #a7b3c6;
            }
            QLineEdit, QComboBox {
                background-color: #2f3c4e;
                border: 1px solid #4a5d78;
                padding: 5px;
                border-radius: 3px;
                color: #ffffff;
            }
            QListWidget {
                background-color: #2f3c4e;
                border: 1px solid #4a5d78;
                border-radius: 3px;
            }
            QFrame#sidebar {
                background-color: #151c28;
                max-width: 50px;
                padding-top: 20px;
            }
            QFrame#content {
                background-color: #1c2432;
                padding: 20px;
            }
        """)

        self.main_layout = QVBoxLayout(self)
        self.stacked_widget = QStackedWidget()
        self.main_layout.addWidget(self.stacked_widget)

        self.main_menu_widget = QWidget()
        self.counter_stuff_widget = QWidget()
        self.lab_menu_widget = QWidget()

        self.stacked_widget.addWidget(self.main_menu_widget)
        self.stacked_widget.addWidget(self.counter_stuff_widget)
        self.stacked_widget.addWidget(self.lab_menu_widget)

        self.init_main_menu()
        self.init_counter_stuff_menu()
        self.init_lab_menu()

        self.driver = None
        self.labs_managed = False
        self.login_successful = False

        # Initialize timer and other attributes
        self.elapsed_time = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_timer)
        self.initialize_attributes()

    def init_main_menu(self):
        layout = QVBoxLayout(self.main_menu_widget)

        # Timer section
        timer_layout = QHBoxLayout()
        self.timer_label = QLabel("00:00:00")
        self.timer_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        timer_layout.addWidget(self.timer_label)

        self.start_button = QPushButton("Start")
        self.stop_resume_button = QPushButton("Stop")
        self.reset_button = QPushButton("Reset")

        for button in [self.start_button, self.stop_resume_button, self.reset_button]:
            button.setFixedWidth(80)
            timer_layout.addWidget(button)

        layout.addLayout(timer_layout)

        # Input fields
        self.unikey_input = QLineEdit()
        self.unikey_input.setPlaceholderText("Enter Unikey")
        self.counter_location_dropdown = QComboBox()
        self.counter_location_dropdown.addItems(["Fisher", "Other Location"])
        self.assistance_category_dropdown = QComboBox()
        self.assistance_category_dropdown.addItems([
            "Card Encoding", "General Assist", "Lab Computer", "Loaned Device",
            "Lost & Found", "Navigation", "Personal Device", "Printing",
            "Software Installation", "Unikey Assistance", "University Application Assistance",
            "WiFi", "Other"
        ])

        for widget in [self.unikey_input, self.counter_location_dropdown, self.assistance_category_dropdown]:
            layout.addWidget(widget)

        # Action buttons
        self.counter_stuff_button = QPushButton("Counter Stuff")
        layout.addWidget(self.counter_stuff_button)

        self.manage_labs_button = QPushButton("Manage Labs")
        layout.addWidget(self.manage_labs_button)

        # Ticket list
        self.ticket_list = QListWidget()
        layout.addWidget(self.ticket_list)

        # Connect button signals
        self.start_button.clicked.connect(self.start_timer)
        self.stop_resume_button.clicked.connect(self.stop_resume_timer)
        self.reset_button.clicked.connect(self.reset_timer)
        self.counter_stuff_button.clicked.connect(self.show_counter_stuff_menu)
        self.manage_labs_button.clicked.connect(self.manage_labs)

    def init_counter_stuff_menu(self):
        layout = QVBoxLayout()

        back_button = QPushButton("Back")
        back_button.clicked.connect(self.show_main_menu)
        layout.addWidget(back_button)

        run_okta_resets_button = QPushButton("Run Okta Resets for New Phone")
        run_okta_resets_button.clicked.connect(self.run_okta_resets_for_new_phone)
        layout.addWidget(run_okta_resets_button)

        run_okta_resets_for_deleted_app_button = QPushButton("Run Okta Resets for Deleted App")
        run_okta_resets_for_deleted_app_button.clicked.connect(self.run_okta_resets_for_deleted_app)
        layout.addWidget(run_okta_resets_for_deleted_app_button)

        reprint_button = QPushButton("Reprint")
        reprint_button.clicked.connect(self.reprint)
        layout.addWidget(reprint_button)

        reprint_hold_for_auth_button = QPushButton("Printing Hold for Authentication Error")
        reprint_hold_for_auth_button.clicked.connect(self.reprint_hold_for_auth)
        layout.addWidget(reprint_hold_for_auth_button)

        self.counter_stuff_widget.setLayout(layout)

    def init_lab_menu(self):
        layout = QVBoxLayout(self.lab_menu_widget)

        back_button = QPushButton("Back")
        back_button.clicked.connect(self.show_main_menu)
        layout.addWidget(back_button)

        list_refund_tickets_button = QPushButton("List Refund Tickets")
        list_refund_tickets_button.clicked.connect(self.filter_print_refund_tickets)
        layout.addWidget(list_refund_tickets_button)

    def show_main_menu(self):
        self.stacked_widget.setCurrentWidget(self.main_menu_widget)

    def show_counter_stuff_menu(self):
        if not is_chrome_debugger_running():
            start_chrome_debugging()
        self.stacked_widget.setCurrentWidget(self.counter_stuff_widget)

    def show_lab_management_menu(self):
        if not self.login_successful:
            QMessageBox.warning(self, "Error", "Please log in successfully before accessing the lab management menu.")
            return
        self.stacked_widget.setCurrentWidget(self.lab_menu_widget)

    def manage_labs(self):
        if not self.labs_managed:
            if self.open_task_list():
                self.labs_managed = True
                self.show_lab_management_menu()
        else:
            if self.driver is None or not self.is_driver_alive():
                if self.open_task_list():
                    self.show_lab_management_menu()
            else:
                self.show_lab_management_menu()

    def initialize_attributes(self):
        self.workbook, self.sheet = initialize_workbook()

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
        unikey = self.unikey_input.text()
        counter_location = self.counter_location_dropdown.currentText()
        assistance_category = self.assistance_category_dropdown.currentText()
        save_to_excel(self.sheet, unikey, counter_location, assistance_category, self.elapsed_time)
        self.workbook.save("assistance_data.xlsx")

    def get_credentials(self):
        username, ok1 = QInputDialog.getText(self, "Login", "Enter your username:")
        if not ok1:
            return None, None
        password, ok2 = QInputDialog.getText(self, "Login", "Enter your password:", QLineEdit.Password)
        if not ok2:
            return None, None
        return username, password

    def run_okta_resets_for_new_phone(self):
        run_okta_resets_for_new_phone(self)

    def run_okta_resets_for_deleted_app(self):
        run_okta_resets_for_deleted_app(self)

    def reprint(self):
        reprint(self)

    def reprint_hold_for_auth(self):
        reprint_hold_for_auth(self)
    def is_driver_alive(self):
        try:
            self.driver.title
            return True
        except:
            return False

    def closeEvent(self, event):
        if self.driver:
            self.driver.quit()
        event.accept()
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
        filter_print_refund_tickets(self)
    def open_task_list(self):
        login_dialog = LoginDialog(self)
        if login_dialog.exec_() == QDialog.Accepted:
            username_input, password_input = login_dialog.get_credentials()

            try:
                self.driver = webdriver.Chrome()
                # Navigate to the task list URL
                self.driver.get(task_list_url)

                self.driver.implicitly_wait(2)

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

                choose_push_nofif = self.driver.find_element(By.XPATH,
                                                             '//*[@id="form61"]/div[2]/div/div[2]/div[2]/div[2]')
                choose_push_nofif.click()

                self.login_successful = True

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

                return True

            except Exception as e:
                print("An error occurred while opening the task list:", str(e))
                QMessageBox.warning(self, "Error", f"Failed to open task list: {str(e)}")
                if self.driver:
                    self.driver.quit()
                self.driver = None
                self.login_successful = False
                return False

            else:
                print("Login cancelled")
                self.login_successful = False
                return False

            #will need this later
            def is_driver_alive(self):
                try:
                    self.driver.title
                    return True
                except:
                    return False

            #trust bro it will be used later....i think. Nah its good to have tho
            def closeEvent(self, event):
                if self.driver:
                    self.driver.quit()
                event.accept()