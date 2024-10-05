from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QLabel, QSystemTrayIcon, QMenu, QAction
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer, Qt
import sys

class App(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        # Set up the system tray icon
        self.tray_icon = QSystemTrayIcon(QIcon("icon.png"), self)
        self.tray_icon.setToolTip("My Application")
        
        # Create context menu for tray icon
        tray_menu = QMenu()
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(QApplication.instance().quit)
        tray_menu.addAction(exit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

        # Set up the main window
        self.setWindowTitle("Sticky Input Window")
        self.setGeometry(300, 300, 400, 150)

        # Set up layout
        layout = QVBoxLayout()
        self.setLayout(layout)

        # Input field
        self.input_field = QLineEdit(self)
        self.input_field.setPlaceholderText("Enter command here...")
        layout.addWidget(self.input_field)

        # Execute button
        self.execute_button = QPushButton("Execute", self)
        self.execute_button.clicked.connect(self.process_input)
        layout.addWidget(self.execute_button)

        # Result label
        self.result_label = QLabel("", self)
        self.result_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.result_label)

        # Set up a timer for hiding notifications
        self.timer = QTimer()
        self.timer.timeout.connect(self.hide_message)

        self.show()

    def process_input(self):
        command = self.input_field.text()
        self.result_label.setText(f"Executing: {command}")
        # Simulate command execution
        QTimer.singleShot(1000, lambda: self.result_label.setText("Execution Complete"))
        self.timer.start(2000)  # Hide message after 2 seconds

    def hide_message(self):
        self.result_label.setText("")
        self.timer.stop()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
