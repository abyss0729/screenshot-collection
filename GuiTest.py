from PyQt5.QtGui import QPixmap
import Ui_main
from PyQt5.QtWidgets import QWidget, QApplication, QSplashScreen
import sys

if __name__ == '__main__':
    file = '青年大学习第八季第十期（收集结果）(1).xlsx'
    dir = 'ss'
    app = QApplication(sys.argv)
    splash = QSplashScreen(QPixmap('logo.ico'))
    splash.show()
    mainWindow = QWidget()
    ui = Ui_main.QmyWidget(mainWindow)
    ui.ui.SaveTextEdit.setText(dir)
    ui.ui.FileTextEdit.setText(file)
    mainWindow.show()
    splash.close()
    sys.exit(app.exec_())
