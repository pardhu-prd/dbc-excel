'''Main file'''
from PyQt5.QtWidgets import QApplication
from dbcexcelwindow import DbcWindow

def run_app():
    '''Runs the main'''
    app = QApplication([])
    window = DbcWindow()
    window.show()
    app.exec_()

if __name__ == "__main__":
    run_app()
