import sys
from PyQt5.QtWidgets import QApplication
from ui.snow_software_window import SnowSoftwareWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SnowSoftwareWindow()
    window.show()
    sys.exit(app.exec_())