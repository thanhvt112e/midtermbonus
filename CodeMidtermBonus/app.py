from PyQt6.QtWidgets import QApplication, QMainWindow
from MainWindowExt import MainWindowExt

app = QApplication([])
main_window = QMainWindow()
ui = MainWindowExt(main_window)
main_window.show()
app.exec()
