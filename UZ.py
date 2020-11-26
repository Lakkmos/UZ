from PyQt5 import QtCore, QtGui, QtWidgets
from pytils import translit
import UZ_g, sys, string, xlwt, random, os, xlrd, xlwings, sip
boxobox = []
result = []
infile = []
c=''
imp=''
path2 = ''
flg = False
dict_set = dict()
k = 0

class UZ_form(QtWidgets.QMainWindow, UZ_g.Ui_Dialog):
    def __init__(self):
        global boxobox, path2, flg, dict_set

        super().__init__()
        self.setupUi(self)
        self.pushButton_3.clicked.connect(self.forming)
        self.pushButton_2.clicked.connect(self.export)
        self.pushButton.clicked.connect(self.forcopy)
        self.pushButton.setEnabled(False)
        self.pushButton_2.setEnabled(False)
        boxobox=[]
        try:
            file = open('skills.txt')
            f = file.read().split('\n')
            for i in range(len(f) - 1):
                c = QtWidgets.QCheckBox(self.groupBox)
                c.setGeometry(QtCore.QRect(10, 20 + i * 15, 220, 17))
                self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(430, 10, 220, 50 + i * 15))
                self.groupBox.setGeometry(QtCore.QRect(0, 0, 220, (26 + i * 16)))
                c.setObjectName("CheckBox_" + str(i))
                c.setText(f[i])
                boxobox.append(c)
            file.close()
        except:
            QtWidgets.QMessageBox.about(self, "Ошибка", "Проблема с файлом навыков. Его нет рядом.  ")

        try:
            file = open('settings.txt')
            set = file.read().split('\n')
            for item in set:
                key = item.split(": ")[0]
                item = item[item.find(': ')+2:]
                value = item
                dict_set[key] = str(value)
            file.close()
            flg = True
        except:
            flg = False

        self.comboBox_2.addItems(["Оператор"])
        self.comboBox_3.addItems(["Квалифицированный оператор"])
        self.comboBox.addItems(["1"])
        path_t = r"~"
        path2 = os.path.expanduser(path_t)




    def forming(self):
        global boxobox, result, infile, k, flg, dict_set
        self.pushButton.setEnabled(True)
        self.pushButton_2.setEnabled(True)
        list_skills = ''
        result = []
        infile = []
        fio = self.plainTextEdit.toPlainText()
        fio.strip()
        m = fio.split('\n')
        dr = self.plainTextEdit_2.toPlainText()
        dr.strip()
        d = dr.split('\n')
        tel = self.plainTextEdit_3.toPlainText()
        tel.strip()
        t = tel.split('\n')
        for i in range(len(boxobox)-1):
            if boxobox[i].isChecked():
                level = self.comboBox.currentText()
                if list_skills != '':
                   list_skills = list_skills+','
                list_skills = list_skills+(boxobox[i].text())+'/'+level

        result.append([])
        infile.append([])
        for i in range(len(m)-1):
            result.append([])
            infile.append([])
            if i ==0:
               result[i].append('Фамилия')
               result[i].append('Имя')
               result[i].append('Отчество')
            infile[i].append(m[i])
            login = m[i].split(' ')
            result[i + 1].append(login[0])
            result[i + 1].append(login[1])
            result[i + 1].append(login[2])
            login = str(login[0] + '_' + login[1][0] + login[2][0])
            login = login.lower()
            login = translit.translify(login)
            login = login.replace("'", "")
            if i == 0:
                result[i].append("Логин")
            result[i+1].append(login)
            infile[i].append(login)

            passw = random.choice(string.ascii_lowercase) + random.choice(string.ascii_lowercase) + str(random.choice(string.digits)) + str(random.choice(string.digits)) + str(random.choice(string.digits))
            if i == 0:
                result[i].append("Пароль")
            result[i+1].append(passw)
            infile[i].append(passw)

            try:
                d[i].strip()
                z = d[i].split('.')
                d[i] = (z[0]) + '.' + (z[1]) + '.' + (z[2][-2:])
                result[i + 1].append(d[i])
                if i == 0:
                    result[i].append('Дата рождения')
            except:
                k = 0

            try:
                if t[0] != '':
                    result[i + 1].append(t[i])
                    if i == 0:
                        result[i].append("Номер мобильного телефона")


            except:
                k = 0

            if i == 0:
                result[i].append("Навыки")
            result[i+1].append(list_skills)

            if i == 0:
                result[i].append("Категория оператора")
            result[i + 1].append(self.comboBox_3.currentText())
            if i == 0:
                result[i].append("Роли")
            result[i + 1].append(self.comboBox_2.currentText())
        if flg == False:
            self.pushButton_2.setToolTip("C:/Users/" + os.path.basename(path2) + "/Desktop/export.xls")
        else:
            self.pushButton_2.setToolTip(str(dict_set["path_save"]))

    def export(self):
        global result, path, path2, flg, dict_set
        try:
            wbk = xlwt.Workbook('utf - 8')
            sheet = wbk.add_sheet('sheet 1')

            for i in range(len(result)):
                for j in range(len(result[i])):
                    sheet.write(i, j, result[i][j])

            if flg == True:

                wbk.save(str(dict_set["path_save"]))
                QtWidgets.QMessageBox.about(self, "Готово", "Данные экспортированы в папку из настроек   ")
            else:
                wbk.save("C:/Users/" + os.path.basename(path2) + "/Desktop/export.xls")
                QtWidgets.QMessageBox.about(self, "Готово", "Данные экспортированы на рабочий стол   ")
        except:
            QtWidgets.QMessageBox.about(self, "Ошибка", "Сохранить не удалось   ")





    def forcopy(self):
        global c, infile, imp, path2
        imp = ''
        c = QtWidgets.QWidget(self)
        c.setStyleSheet('''
                        QWidget {
                            background : #d5f0d0;
                            border: 1px solid black; }
                        ''')
        self.pushButton.setEnabled(False)
        t = QtWidgets.QTableWidget(c)
        t.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        t.horizontalHeader().setVisible(False)
        t.verticalHeader().setVisible(False)
        t.setColumnCount(len(infile[0]))
        t.setRowCount(len(infile))

        for i in range(len(infile)-1):
            for j in range(len(infile[0])):
                t.setItem(i, j, QtWidgets.QTableWidgetItem(infile[i][j]))
                imp = imp+str(infile[i][j])+'\t'
            imp = imp + '\n'
        t.resizeColumnsToContents()
        b1= QtWidgets.QPushButton(c)
        b2 = QtWidgets.QPushButton(c)
        b3 = QtWidgets.QPushButton(c)
        b4 = QtWidgets.QPushButton(c)
        b5 = QtWidgets.QPushButton(c)
        b1.clicked.connect(self.close_forcopy)
        b2.clicked.connect(self.copytoclip)
        b3.clicked.connect(self.to_yar)
        b4.clicked.connect(self.to_che)
        b5.clicked.connect(self.to_br)
        c.setGeometry(QtCore.QRect(10, 10, 500, 200))
        t.setGeometry(QtCore.QRect(0, 0, 400, 200))
        b1.setGeometry(QtCore.QRect(410, 10, 80, 20))
        b2.setGeometry(QtCore.QRect(410, 40, 80, 20))
        b3.setGeometry(QtCore.QRect(410, 70, 80, 20))
        b4.setGeometry(QtCore.QRect(410, 100, 80, 20))
        b5.setGeometry(QtCore.QRect(410, 130, 80, 20))

        b3.setEnabled(False)
        b4.setEnabled(False)
        b5.setEnabled(False)
        t.setObjectName("TableWidget")
        c.setObjectName("widget")
        b1.setObjectName("button")
        b2.setObjectName("button2")
        b1.setText('Закрыть')
        b2.setText('Копировать')
        b3.setText('В Ярославль')
        if flg==True:
            if os.path.exists(dict_set["path_save_yar"]):
                b3.setEnabled(True)
                b3.setToolTip(dict_set["path_save_yar"])
        else:
            if os.path.exists(path2 + r'\YandexDisk\operator_account\operator_account.xlsx'):
                b3.setEnabled(True)
                b3.setToolTip(path2 + r'\YandexDisk\operator_account\operator_account.xlsx')
        b4.setText('В Череповец')
        if flg==True:
            if os.path.exists(dict_set["path_save_che"]):
                b4.setEnabled(True)
                b4.setToolTip(dict_set["path_save_che"])
        else:
            if os.path.exists(path2 + r'\YandexDisk\operator_account_2CHE\operator_account_2CHE.xlsx'):
                b4.setEnabled(True)
                b4.setToolTip(path2 + r'\YandexDisk\operator_account_2CHE\operator_account_2CHE.xlsx')

        b5.setText('В Брянск')
        if flg==True:
            if os.path.exists(dict_set["path_save_br"]):
                b5.setEnabled(True)
                b5.setToolTip(dict_set["path_save_br"])

        else:
            if os.path.exists(path2 + r'\YandexDisk\operator_account_BR\operator_account.xlsx'):
                b5.setEnabled(True)
                b5.setToolTip(path2 + r'\YandexDisk\operator_account_BR\operator_account.xlsx')


        c.show()
        t.show()

    def close_forcopy(self):
        global c
        self.pushButton.setEnabled(True)
        c.close()

    def copytoclip(self):
        global imp

        import pyperclip
        pyperclip.copy(imp)

    def to_yar(self):
        global infile, path2, dict_set
        x = dict_set.setdefault("path_save_yar", path2+r'\YandexDisk\operator_account\operator_account.xlsx')
        rb = xlrd.open_workbook(str(x))
        sheet = rb.sheet_by_index(0)
        num_rows = sheet.nrows
        wb = xlwings.Book(str(x))
        Sheet = wb.sheets[0]
        for i in range(len(infile)-1):
            for j in range(3):
                Sheet.range(num_rows+i+1, j+1).value = infile[i][j]
        wb.save()
        wb.close()

    def to_che(self):
        global infile,dict_set
        x = dict_set.setdefault("path_save_che", path2+r'\YandexDisk\operator_account_2CHE\operator_account_2CHE.xlsx')
        rb = xlrd.open_workbook(str(x))
        sheet = rb.sheet_by_index(0)
        num_rows = sheet.nrows
        wb = xlwings.Book(str(x))
        Sheet = wb.sheets[0]
        for i in range(len(infile) - 1):
            for j in range(3):
                Sheet.range(num_rows + i+1, j + 1).value = infile[i][j]
        wb.save()
        wb.close()
    def to_br(self):
        global infile,dict_set
        x = dict_set.setdefault("path_save_br", path2+r'\YandexDisk\operator_account_BR\operator_account.xlsx')
        rb = xlrd.open_workbook(str(x))
        sheet = rb.sheet_by_index(0)
        num_rows = sheet.nrows
        wb = xlwings.Book(str(x))
        Sheet = wb.sheets[0]
        for i in range(len(infile) - 1):
            for j in range(3):
                Sheet.range(num_rows + i+1, j + 1).value = infile[i][j]
        wb.save()
        wb.close()








def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = UZ_form()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение

if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()