import sys
import win32com.client
import win32gui
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QMessageBox
from PyQt5.QtCore import Qt, QTimer
import pygetwindow as gw

class ControlPanel(QWidget):
    def __init__(self, ie):
        super().__init__()
        self.ie = ie
        self.initUI()

        # Timer to check the status of IE window
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_ie_status)
        self.timer.start(1000)  # Check every 1000 milliseconds (1 second)

    def initUI(self):
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        layout = QVBoxLayout()

        btn_back = QPushButton('Back', self)
        btn_back.clicked.connect(self.goBack)
        layout.addWidget(btn_back)

        btn_forward = QPushButton('Forward', self)
        btn_forward.clicked.connect(self.goForward)
        layout.addWidget(btn_forward)

        btn_refresh = QPushButton('Refresh', self)
        btn_refresh.clicked.connect(self.refreshPage)
        layout.addWidget(btn_refresh)

        self.setLayout(layout)

    def goBack(self):
        try:
            self.ie.GoBack()
        except Exception as e:
            print(f"Error going back: {e}")

    def goForward(self):
        try:
            self.ie.GoForward()
        except Exception as e:
            print(f"Error going forward: {e}")

    def refreshPage(self):
        try:
            self.ie.Refresh()
        except Exception as e:
            print(f"Error refreshing page: {e}")

    def position_next_to_ie(self):
        ie_windows = gw.getWindowsWithTitle('Internet Explorer')
        if ie_windows:
            ie_window = ie_windows[0]
            self.move(ie_window.right, ie_window.top)
            self.activateWindow()  # Bring the ControlPanel window to the foreground
            self.bring_ie_to_foreground(ie_window._hWnd)  # Use the handle to bring IE to the foreground

    def bring_ie_to_foreground(self, hwnd):
        win32gui.SetForegroundWindow(hwnd)

    def check_ie_status(self):
        ie_windows = gw.getWindowsWithTitle('Internet Explorer')
        if not ie_windows:
            self.close()

if __name__ == '__main__':
    app = QApplication([])

    # Check if a URL argument is provided
    if len(sys.argv) < 2:
        # Show a message box if no URL is provided
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("<URL> not given!")
        msg.setInformativeText("Please provide a URL as a command line argument.")
        msg.setWindowTitle("Argument Missing")
        msg.exec_()
        sys.exit(1)  # Exit the application

    ie = win32com.client.Dispatch("InternetExplorer.Application")
    ie.Visible = True
    ie.ToolBar = 0
    ie.Navigate(sys.argv[1])

    # Store the ControlPanel instance in a variable to prevent premature garbage collection
    ctrl_panel = ControlPanel(ie)
    ctrl_panel.show()
    ctrl_panel.position_next_to_ie()

    # Store the reference to ctrl_panel in the QApplication instance to ensure it's not garbage collected
    app.ctrl_panel = ctrl_panel

    app.exec_()

