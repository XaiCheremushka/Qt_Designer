from PyQt5 import QtCore, QtGui, QtWidgets
import sys, Guest


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    GuestWin = QtWidgets.QDialog()
    ui = Guest.Ui_GuestWin()
    ui.setupUi(GuestWin)
    GuestWin.show()
    sys.exit(app.exec_())