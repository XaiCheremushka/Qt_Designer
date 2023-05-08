from PyQt5 import QtCore, QtGui, QtWidgets
import pandas as pd
from openpyxl import load_workbook
import time

data = {
    "Фамилия": "",
    "Имя": "",
    "Отчество": "",
    "Гражданство": "",
    "Когда выдан": "",
    "Кем выдан": "",
    "Серия": "",
    "Номер": "",
    "Пол": "",
    "Дата рождения": "",
    "Номер карты": "",
    "Срок действия": "",
    "Число на обратной стороне карты": "",
    "Дата бронирования": ""
}


class Ui_GuestWin(object):
    def setupUi(self, GuestWin):
        GuestWin.setObjectName("GuestWin")
        GuestWin.resize(619, 837)
        GuestWin.setStyleSheet("background-color: rgb(144, 144, 144);")
        self.label = QtWidgets.QLabel(GuestWin)
        self.label.setGeometry(QtCore.QRect(220, 10, 171, 41))
        self.label.setStyleSheet("font: 63 italic 20pt \"Lucida Fax\";")
        self.label.setObjectName("label")
        self.secondName = QtWidgets.QLineEdit(GuestWin)
        self.secondName.setGeometry(QtCore.QRect(30, 80, 131, 20))
        self.secondName.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.secondName.setText("")
        self.secondName.setObjectName("secondName")
        self.label_2 = QtWidgets.QLabel(GuestWin)
        self.label_2.setGeometry(QtCore.QRect(70, 60, 71, 21))
        self.label_2.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(GuestWin)
        self.label_3.setGeometry(QtCore.QRect(240, 60, 71, 21))
        self.label_3.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_3.setObjectName("label_3")
        self.name = QtWidgets.QLineEdit(GuestWin)
        self.name.setGeometry(QtCore.QRect(200, 80, 113, 20))
        self.name.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.name.setText("")
        self.name.setObjectName("name")
        self.label_4 = QtWidgets.QLabel(GuestWin)
        self.label_4.setGeometry(QtCore.QRect(380, 60, 71, 21))
        self.label_4.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_4.setObjectName("label_4")
        self.middleName = QtWidgets.QLineEdit(GuestWin)
        self.middleName.setGeometry(QtCore.QRect(340, 80, 151, 20))
        self.middleName.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.middleName.setText("")
        self.middleName.setObjectName("middleName")
        self.label_5 = QtWidgets.QLabel(GuestWin)
        self.label_5.setGeometry(QtCore.QRect(260, 110, 91, 41))
        self.label_5.setStyleSheet("font: 63 italic 16pt \"Lucida Fax\";")
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(GuestWin)
        self.label_6.setGeometry(QtCore.QRect(60, 220, 71, 21))
        self.label_6.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_6.setObjectName("label_6")
        self.pasSeries = QtWidgets.QLineEdit(GuestWin)
        self.pasSeries.setGeometry(QtCore.QRect(30, 240, 101, 20))
        self.pasSeries.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.pasSeries.setText("")
        self.pasSeries.setObjectName("pasSeries")
        self.pasWhoGive = QtWidgets.QLineEdit(GuestWin)
        self.pasWhoGive.setGeometry(QtCore.QRect(360, 180, 241, 20))
        self.pasWhoGive.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.pasWhoGive.setText("")
        self.pasWhoGive.setObjectName("pasWhoGive")
        self.label_7 = QtWidgets.QLabel(GuestWin)
        self.label_7.setGeometry(QtCore.QRect(220, 220, 71, 21))
        self.label_7.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(GuestWin)
        self.label_8.setGeometry(QtCore.QRect(440, 160, 91, 21))
        self.label_8.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_8.setObjectName("label_8")
        self.pasNumber = QtWidgets.QLineEdit(GuestWin)
        self.pasNumber.setGeometry(QtCore.QRect(172, 240, 141, 20))
        self.pasNumber.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.pasNumber.setText("")
        self.pasNumber.setObjectName("pasNumber")
        self.pasState = QtWidgets.QComboBox(GuestWin)
        self.pasState.setGeometry(QtCore.QRect(30, 180, 141, 21))
        self.pasState.setStyleSheet("background-color: rgb(176, 176, 176);\n"
                                    "font: italic 8pt \"Lucida Fax\";")
        self.pasState.setObjectName("pasState")
        self.pasState.addItem("")
        self.pasState.addItem("")
        self.pasState.addItem("")
        self.pasState.addItem("")
        self.pasState.addItem("")
        self.label_9 = QtWidgets.QLabel(GuestWin)
        self.label_9.setGeometry(QtCore.QRect(50, 160, 101, 21))
        self.label_9.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_9.setObjectName("label_9")
        self.pasWhenGive = QtWidgets.QLineEdit(GuestWin)
        self.pasWhenGive.setGeometry(QtCore.QRect(202, 180, 141, 20))
        self.pasWhenGive.setStyleSheet("background-color: rgb(176, 176, 176);\n"
                                       "")
        self.pasWhenGive.setText("")
        self.pasWhenGive.setObjectName("pasWhereGive")
        self.label_10 = QtWidgets.QLabel(GuestWin)
        self.label_10.setGeometry(QtCore.QRect(230, 160, 101, 21))
        self.label_10.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_10.setObjectName("label_10")
        self.label_11 = QtWidgets.QLabel(GuestWin)
        self.label_11.setGeometry(QtCore.QRect(450, 220, 71, 21))
        self.label_11.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_11.setObjectName("label_11")
        self.male = QtWidgets.QRadioButton(GuestWin)
        self.male.setGeometry(QtCore.QRect(380, 240, 82, 17))
        self.male.setStyleSheet("font: italic 8pt \"Lucida Fax\";")
        self.male.setObjectName("male")
        self.female = QtWidgets.QRadioButton(GuestWin)
        self.female.setGeometry(QtCore.QRect(480, 240, 82, 17))
        self.female.setStyleSheet("font: italic 8pt \"Lucida Fax\";")
        self.female.setObjectName("female")
        self.birthday = QtWidgets.QLineEdit(GuestWin)
        self.birthday.setGeometry(QtCore.QRect(32, 300, 141, 20))
        self.birthday.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.birthday.setText("")
        self.birthday.setObjectName("birthday")
        self.label_12 = QtWidgets.QLabel(GuestWin)
        self.label_12.setGeometry(QtCore.QRect(50, 280, 121, 21))
        self.label_12.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_12.setObjectName("label_12")
        self.label_13 = QtWidgets.QLabel(GuestWin)
        self.label_13.setGeometry(QtCore.QRect(180, 320, 271, 41))
        self.label_13.setStyleSheet("font: 63 italic 16pt \"Lucida Fax\";")
        self.label_13.setObjectName("label_13")
        self.cardNumber = QtWidgets.QLineEdit(GuestWin)
        self.cardNumber.setGeometry(QtCore.QRect(40, 390, 241, 20))
        self.cardNumber.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.cardNumber.setText("")
        self.cardNumber.setObjectName("cardNumber")
        self.label_14 = QtWidgets.QLabel(GuestWin)
        self.label_14.setGeometry(QtCore.QRect(110, 370, 91, 21))
        self.label_14.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_14.setObjectName("label_14")
        self.cardDate = QtWidgets.QLineEdit(GuestWin)
        self.cardDate.setGeometry(QtCore.QRect(330, 390, 241, 20))
        self.cardDate.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.cardDate.setText("")
        self.cardDate.setObjectName("cardDate")
        self.label_15 = QtWidgets.QLabel(GuestWin)
        self.label_15.setGeometry(QtCore.QRect(400, 370, 111, 21))
        self.label_15.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_15.setObjectName("label_15")
        self.cardBackNum = QtWidgets.QLineEdit(GuestWin)
        self.cardBackNum.setGeometry(QtCore.QRect(240, 450, 141, 20))
        self.cardBackNum.setStyleSheet("background-color: rgb(176, 176, 176);")
        self.cardBackNum.setText("")
        self.cardBackNum.setObjectName("cardBackNum")
        self.label_16 = QtWidgets.QLabel(GuestWin)
        self.label_16.setGeometry(QtCore.QRect(190, 430, 251, 21))
        self.label_16.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.label_16.setObjectName("label_16")
        self.broneButton = QtWidgets.QPushButton(GuestWin)
        self.broneButton.setGeometry(QtCore.QRect(210, 730, 221, 61))
        self.broneButton.setStyleSheet("background-color: rgb(26, 114, 255);\n"
                                       "font: italic 14pt \"Lucida Fax\";")
        self.broneButton.setObjectName("broneButton")
        self.calendarWidget = QtWidgets.QCalendarWidget(GuestWin)
        self.calendarWidget.setGeometry(QtCore.QRect(150, 520, 321, 191))
        self.calendarWidget.setStyleSheet("background-color: rgb(2, 190, 87);")
        self.calendarWidget.setObjectName("calendarWidget")
        self.label_17 = QtWidgets.QLabel(GuestWin)
        self.label_17.setGeometry(QtCore.QRect(220, 470, 191, 41))
        self.label_17.setStyleSheet("font: 63 italic 20pt \"Lucida Fax\";")
        self.label_17.setObjectName("label_17")
        self.labelResult = QtWidgets.QLabel(GuestWin)
        self.labelResult.setGeometry(QtCore.QRect(260, 800, 131, 21))
        self.labelResult.setStyleSheet("font: italic 11pt \"Lucida Fax\";")
        self.labelResult.setObjectName("labelResult")

        self.retranslateUi(GuestWin)
        QtCore.QMetaObject.connectSlotsByName(GuestWin)

        self.broneButton.clicked.connect(lambda: self.writeDatabase())

    def writeDatabase(self):
        # l = [secondName, name, middleName,  pasWhenGive, pasWhoGive,
        #     pasSeries, pasNumber, birthday, cardNumber, cardDate, cardBackNum]
        # for i in l:
        #     for j in data.keys():
        #         if str(i) == j:
        #             data[j] = self.i.text

        data["Фамилия"] = self.secondName.text()
        data["Имя"] = self.name.text()
        data["Отчество"] = self.middleName.text()
        data["Гражданство"] = self.pasState.currentText()
        data["Когда выдан"] = self.pasWhenGive.text()
        data["Номер"] = self.pasNumber.text()
        data["Серия"] = self.pasSeries.text()
        data["Кем выдан"] = self.pasWhoGive.text()
        data["Дата рождения"] = self.birthday.text()
        data["Номер карты"] = self.cardNumber.text()
        data["Срок действия карты"] = self.cardDate.text()
        data["Число на обратной стороне карты"] = self.cardBackNum.text()
        if self.male.isChecked() == True and self.female.isChecked() == False:
            data["Пол"] = "Мужской"
        else:
            data["Пол"] = "Женский"
        data["Дата бронирования"] = str(self.calendarWidget.selectedDate())[-12:-1].replace(", ", ".")


        # df = pd.DataFrame(data)
        # df.to_excel('DataBase.xlsx', index=False)


        # book = load_workbook('DataBase.xlsx')
        # writer = pandas.ExcelWriter('DataBase.xlsx', engine='openpyxl')
        # writer.book = book
        #
        # data_filtered.to_excel(writer,  ro=['Diff1', 'Diff2'])

        # writer.save()

        with open('DataBase.txt', 'a') as f:
            f.write('\nНовые данные\n')
            for i in data.keys():
                f.write(i + ": " + data[i] + "\n")


        print(data)
        print(data.keys())

        self.labelResult.setText(_translate("GuestWin", "Данные внесены"))
        time.sleep(8)
        GuestWin.close()
        # print(str(name))



    def retranslateUi(self, GuestWin):
        _translate = QtCore.QCoreApplication.translate
        GuestWin.setWindowTitle(_translate("GuestWin", "Жилец"))
        self.label.setText(_translate("GuestWin", "Ввод данных"))
        self.label_2.setText(_translate("GuestWin", "Фамилия"))
        self.label_3.setText(_translate("GuestWin", "Имя"))
        self.label_4.setText(_translate("GuestWin", "Отчество"))
        self.label_5.setText(_translate("GuestWin", "Паспорт"))
        self.label_6.setText(_translate("GuestWin", "Серия"))
        self.label_7.setText(_translate("GuestWin", "Номер"))
        self.label_8.setText(_translate("GuestWin", "Кем выдан"))
        self.pasState.setItemText(0, _translate("GuestWin", "Российская Федерация"))
        self.pasState.setItemText(1, _translate("GuestWin", "Республика Казахстан"))
        self.pasState.setItemText(2, _translate("GuestWin", "Республика Беларусь"))
        self.pasState.setItemText(3, _translate("GuestWin", "Республика Армения"))
        self.pasState.setItemText(4, _translate("GuestWin", "Разъединенные Штаты Пендостана"))
        self.label_9.setText(_translate("GuestWin", "Гражданство"))
        self.label_10.setText(_translate("GuestWin", "Когда выдан"))
        self.label_11.setText(_translate("GuestWin", "Пол"))
        self.male.setText(_translate("GuestWin", "Мужской"))
        self.female.setText(_translate("GuestWin", "Женский"))
        self.label_12.setText(_translate("GuestWin", "Дата рождения"))
        self.label_13.setText(_translate("GuestWin", "Данные банковской карты"))
        self.label_14.setText(_translate("GuestWin", "Номер карты"))
        self.label_15.setText(_translate("GuestWin", "Срок действия"))
        self.label_16.setText(_translate("GuestWin", "Число на обратной стороне карты"))
        self.broneButton.setText(_translate("GuestWin", "Забронировать"))
        self.label_17.setText(_translate("GuestWin", "Бронирование"))



if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    GuestWin = QtWidgets.QDialog()
    ui = Ui_GuestWin()
    ui.setupUi(GuestWin)
    GuestWin.show()
    sys.exit(app.exec_())
