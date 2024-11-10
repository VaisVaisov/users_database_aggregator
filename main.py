from os import system
system('pip install -r requirements.txt')
try:
    from random import randint
    from PyQt5 import QtCore, QtGui, QtWidgets
    from openpyxl import Workbook
except ImportError:
    system('pip install -r requirements.txt')

users_xlsx = Workbook()
users_xlsx_sheet = users_xlsx.active
users_xlsx_sheet.title = 'Users'
users_xlsx_sheet['A1'] = 'Username'
users_xlsx_sheet['B1'] = 'Password'


def generator_password(letter_flag, number_flag, special_character_flag, length_password):
    letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
               'u', 'v', 'w', 'x', 'y', 'z']
    numbers = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']
    special_characters = ['#', '№', '$', '%', '&', '!']
    try:
        length = int(length_password)
    except ValueError:
        return 'Generation error'
    symbols_in_password = []
    password = []
    if letter_flag:
        for symbol in letters:
            symbols_in_password.append(symbol)
    if number_flag:
        for symbol in numbers:
            symbols_in_password.append(symbol)
    if special_character_flag:
        for symbol in special_characters:
            symbols_in_password.append(symbol)
    try:
        for i in range(length):
            password.append(symbols_in_password[randint(0, len(symbols_in_password))])
    except IndexError:
        return 'Generation error'
    return ''.join(password)


class Ui_MainWindow(object):
    counter_row = 2

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowIcon(QtGui.QIcon("icon.png"))
        MainWindow.resize(491, 231)
        MainWindow.setMinimumSize(QtCore.QSize(491, 226))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.apply_button = QtWidgets.QPushButton(self.centralwidget)
        self.apply_button.setGeometry(QtCore.QRect(420, 180, 61, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.apply_button.setFont(font)
        self.apply_button.setObjectName("apply_button")
        self.login_field = QtWidgets.QLineEdit(self.centralwidget)
        self.login_field.setGeometry(QtCore.QRect(80, 30, 271, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.login_field.setFont(font)
        self.login_field.setObjectName("login_field")
        self.password_field = QtWidgets.QLineEdit(self.centralwidget)
        self.password_field.setGeometry(QtCore.QRect(80, 60, 271, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.password_field.setFont(font)
        self.password_field.setText("")
        self.password_field.setObjectName("password_field")
        self.password = QtWidgets.QLabel(self.centralwidget)
        self.password.setGeometry(QtCore.QRect(10, 60, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.password.setFont(font)
        self.password.setAlignment(QtCore.Qt.AlignCenter)
        self.password.setObjectName("password")
        self.login = QtWidgets.QLabel(self.centralwidget)
        self.login.setGeometry(QtCore.QRect(10, 30, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.login.setFont(font)
        self.login.setAlignment(QtCore.Qt.AlignCenter)
        self.login.setObjectName("login")
        self.generate_password = QtWidgets.QPushButton(self.centralwidget)
        self.generate_password.setGeometry(QtCore.QRect(360, 60, 71, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.generate_password.setFont(font)
        self.generate_password.setObjectName("generate_password")
        self.save_button = QtWidgets.QPushButton(self.centralwidget)
        self.save_button.setGeometry(QtCore.QRect(350, 180, 61, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.save_button.setFont(font)
        self.save_button.setObjectName("save_button")
        self.cancel_button = QtWidgets.QPushButton(self.centralwidget)
        self.cancel_button.setGeometry(QtCore.QRect(280, 180, 61, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.cancel_button.setFont(font)
        self.cancel_button.setObjectName("cancel_button")
        self.letter_flag = QtWidgets.QCheckBox(self.centralwidget)
        self.letter_flag.setGeometry(QtCore.QRect(10, 190, 70, 17))
        self.letter_flag.setObjectName("letter_flag")
        self.number_flag = QtWidgets.QCheckBox(self.centralwidget)
        self.number_flag.setGeometry(QtCore.QRect(90, 190, 70, 17))
        self.number_flag.setObjectName("number_flag")
        self.spec_symbols_flag = QtWidgets.QCheckBox(self.centralwidget)
        self.spec_symbols_flag.setGeometry(QtCore.QRect(170, 190, 91, 17))
        self.spec_symbols_flag.setObjectName("spec_symbols_flag")
        self.length_field = QtWidgets.QLineEdit(self.centralwidget)
        self.length_field.setGeometry(QtCore.QRect(80, 90, 271, 20))
        self.length_field.setObjectName("length_field")
        self.length = QtWidgets.QLabel(self.centralwidget)
        self.length.setGeometry(QtCore.QRect(10, 90, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.length.setFont(font)
        self.length.setAlignment(QtCore.Qt.AlignCenter)
        self.length.setObjectName("length")
        self.filename = QtWidgets.QLabel(self.centralwidget)
        self.filename.setGeometry(QtCore.QRect(10, 120, 61, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.filename.setFont(font)
        self.filename.setAlignment(QtCore.Qt.AlignCenter)
        self.filename.setObjectName("filename")
        self.filename_field = QtWidgets.QLineEdit(self.centralwidget)
        self.filename_field.setGeometry(QtCore.QRect(80, 120, 271, 20))
        self.filename_field.setObjectName("filename_field")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Users Database Creator by Vais Vaisov"))
        self.apply_button.setText(_translate("MainWindow", "Apply"))
        self.password.setText(_translate("MainWindow", "Password"))
        self.login.setText(_translate("MainWindow", "Login"))
        self.generate_password.setText(_translate("MainWindow", "Generate"))
        self.save_button.setText(_translate("MainWindow", "Save"))
        self.cancel_button.setText(_translate("MainWindow", "Cancel"))
        self.letter_flag.setText(_translate("MainWindow", "Letters"))
        self.number_flag.setText(_translate("MainWindow", "Numbers"))
        self.spec_symbols_flag.setText(_translate("MainWindow", "Spec. Symbols"))
        self.length.setText(_translate("MainWindow", "Length"))
        self.filename.setText(_translate("MainWindow", "Filename"))

    def logicUI(self):
        self.generate_password.clicked.connect(lambda: self.enter_password())
        self.save_button.clicked.connect(lambda: self.save_data())
        self.cancel_button.clicked.connect(lambda: self.cancel_data())
        self.apply_button.clicked.connect(lambda: self.apply_data())

    def save_data(self):
        users_xlsx_sheet[f'A{self.counter_row}'] = str(self.login_field.text())
        users_xlsx_sheet[f'B{self.counter_row}'] = str(self.password_field.text())
        self.counter_row += 1
        if len(self.filename_field.text()) > 0:
            users_xlsx.save(f'{self.filename_field.text()}.xlsx')
            exit()
        else:
            pass

    def apply_data(self):
        users_xlsx_sheet[f'A{self.counter_row}'] = str(self.login_field.text())
        users_xlsx_sheet[f'B{self.counter_row}'] = str(self.password_field.text())
        self.counter_row += 1

    def cancel_data(self):
        self.login_field.setText(''), self.password_field.setText('')

    def enter_password(self):
        if self.check_password():
            self.password_field.setText('')
            self.password_field.setText(generator_password(bool(self.letter_flag.isChecked()),
                                                           bool(self.number_flag.isChecked()),
                                                           bool(self.spec_symbols_flag.isChecked()),
                                                           self.length_field.text()))
        else:
            self.password_field.setText(generator_password(bool(self.letter_flag.isChecked()),
                                                           bool(self.number_flag.isChecked()),
                                                           bool(self.spec_symbols_flag.isChecked()),
                                                           self.length_field.text()))

    def check_password(self):
        if self.password_field.text():
            if self.password_field.text() == 'Generation error':
                return False
            else:
                return True
        else:
            return False


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    MainWindow.show()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.logicUI()
    sys.exit(app.exec_())
