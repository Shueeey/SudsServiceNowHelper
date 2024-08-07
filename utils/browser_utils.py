import platform
import subprocess
import requests
import webbrowser
from PyQt5.QtWidgets import QMessageBox

def open_url(url):
    if platform.system() == "Windows": #this bit is still hella scuffed needs work
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
            # Open the URL in a new tab of Google Chrome using AppleScript
            script = f'tell application "Google Chrome" to open location "{url}"'
            subprocess.call(["osascript", "-e", script])

            print("URL opened successfully in Chrome on macOS!")
        except Exception as e:
            print("An error occurred on macOS:", str(e))

    else:
        print("Unsupported operating system.")

def is_chrome_debugger_running():
    try:
        response = requests.get("http://localhost:9222/json/version")
        return response.status_code == 200
    except requests.exceptions.ConnectionError:
        return False

def start_chrome_debugging():
    url = "https://sydneyuni.service-now.com/nav_to.do?uri=%2Fhome_splash.do%3Fsysparm_direct%3Dtrue"

    system = platform.system()
    if system == "Windows":
        cmd = ['start', 'chrome.exe', '--remote-debugging-port=9222', url]
        subprocess.Popen(cmd, shell=True)
    elif system == "Darwin":  # macOS
        cmd = ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
               '--remote-debugging-port=9222',
               '--user-data-dir=/tmp/chrome-debug',
               url]
        subprocess.Popen(cmd)
    else:
        raise OSError("Unsupported operating system")

    print("Started new Chrome debugging instance with the specified URL")

def open_url_in_new_tab(url):
    try:
        webbrowser.open_new_tab(url)
        print(f"Opened {url} in a new tab")
    except Exception as e:
        print(f"Error opening {url}: {str(e)}")

def show_error_message(parent, title, message):
    QMessageBox.critical(parent, title, message)

def show_info_message(parent, title, message):
    QMessageBox.information(parent, title, message)

def show_warning_message(parent, title, message):
    QMessageBox.warning(parent, title, message)