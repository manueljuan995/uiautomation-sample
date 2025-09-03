from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QListWidget, QMessageBox
import sys
import psutil
import subprocess
import uiautomation as auto

class ProcessKillerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Uiautomation Test")
        self.setGeometry(300, 300, 400, 300)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        layout = QVBoxLayout(self.central_widget)
        
        # List to display running processes
        self.process_list = QListWidget()
        layout.addWidget(self.process_list)
        
        # Button to refresh the process list
        refresh_button = QPushButton("Refresh Process List")
        refresh_button.clicked.connect(self.populate_process_list)
        layout.addWidget(refresh_button)
        
        # Button to kill selected process
        kill_button = QPushButton("Kill Selected Process")
        kill_button.clicked.connect(self.kill_selected_process)
        layout.addWidget(kill_button)

        # Button to kill selected process
        test_button = QPushButton("Uiautomation Test")
        test_button.clicked.connect(self.uiautomation)
        layout.addWidget(test_button)
        
        # Populate the initial list
        self.populate_process_list()

    def populate_process_list(self):
        """Fetches running processes and displays them in the list widget."""
        self.process_list.clear()
        self.processes = []
        
        for process in psutil.process_iter(['pid', 'name']):
            try:
                process_info = f"{process.info['name']} (PID: {process.info['pid']})"
                self.process_list.addItem(process_info)
                self.processes.append(process)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

    def kill_selected_process(self):
        """Kills the selected process from the list."""
        selected_item = self.process_list.currentItem()
        
        if selected_item:
            process_text = selected_item.text()
            pid = int(process_text.split("(PID: ")[1].split(")")[0])

            # Confirm action
            reply = QMessageBox.question(self, 'Confirm Termination', 
                                         f"Are you sure you want to kill process with PID {pid}?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                try:
                    psutil.Process(pid).terminate()
                    QMessageBox.information(self, "Success", f"Process {pid} terminated successfully.")
                    self.populate_process_list()  # Refresh list after termination
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    QMessageBox.warning(self, "Error", f"Failed to terminate process {pid}.")
        else:
            QMessageBox.warning(self, "No Selection", "Please select a process to kill.")

    def uiautomation(self):
        subprocess.Popen('notepad.exe', shell=True)
        # you should find the top level window first, then find children from the top level window
        notepadWindow = auto.WindowControl(searchDepth=1, ClassName='Notepad')
        if not notepadWindow.Exists(3, 1):
            print('Can not find Notepad window')
            exit(0)
        notepadWindow.SetTopmost(True)
        # find the first EditControl in notepadWindow
        edit = notepadWindow.EditControl()
        try:
        # use value pattern to get or set value
            edit.GetValuePattern().SetValue('Hello Notepad')# or edit.GetPattern(auto.PatternId.ValuePattern)
        except auto.comtypes.COMError as ex:
            # maybe you don't run python as administrator 
            # or the control doesn't have a implementation for the pattern method(I have no solution for this)
            pass
        edit.SendKeys('{Ctrl}{End}{Enter}uiautomation')

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = ProcessKillerApp()
    main_window.show()
    sys.exit(app.exec_())