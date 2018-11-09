import xlrd
import re
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys


class Application(QWidget):
    def __init__(self):
        super(Application, self).__init__()
        self.setWindowTitle('BOM check')
        self.setGeometry(400, 400, 1200, 480)

        layout = QVBoxLayout()

        self.Button = QPushButton(self)
        self.Button.setObjectName('Load_btn')
        self.Button.setText('Load')

        self.Button.clicked.connect(self.load)
        layout.addWidget(self.Button)

        self.table = QTableWidget(100, 5)
        # self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setColumnWidth(1, 500)
        self.table.setColumnWidth(2, 300)

        self.table.setHorizontalHeaderLabels(['Part number', 'Description', 'Location', 'Qty', 'Checked'])
        layout.addWidget(self.table)

        self.setLayout(layout)

        self.items = []

    def load(self):
        d, type = QFileDialog.getOpenFileName(self, 'Open')
        # d = QFileDialog.getExistingDirectory(self, 'Open')
        print(d)
        self.filter(d)

    def filter(self, d):
        book = xlrd.open_workbook(d)
        sheets = book.sheet_by_index(0)
        rows = sheets.nrows

        # table = []

        for i in range(7, rows):

            if re.findall('(.*)\nRef', sheets.row_values(i)[7]):
                item = {'PN': sheets.row_values(i)[5], 'Desc': re.findall('(.*)\nRef', sheets.row_values(i)[7]),
                        'Location': re.findall('RefDes:(.*)', sheets.row_values(i)[7]),
                        'Qty': sheets.row_values(i)[9]}
                self.items.append(item)

        self.fillTable(self.items)

    def fillTable(self, items):
        lo_num = 0
        i = 0
        for item in items:
            newitem = QTableWidgetItem(item['PN'])
            self.table.setItem(i, 0, newitem)

            newitem = QTableWidgetItem(item['Desc'][0])
            self.table.setItem(i, 1, newitem)

            if item['Location']:
                locations = item['Location'][0]
                lo_num = len(locations.split(','))
                # print(lo_num)
                newitem = QTableWidgetItem(locations)
                self.table.setItem(i, 2, newitem)

            num = re.findall('(.*).000', item['Qty'])
            newitem = QTableWidgetItem(num[0])
            self.table.setItem(i, 3, newitem)

            if lo_num == int(num[0]):
                # print('Y')
                newitem = QTableWidgetItem('√')
                newitem.setBackground(QBrush(QColor(0, 255, 0)))
                self.table.setItem(i, 4, newitem)
            else:
                newitem = QTableWidgetItem('×')
                newitem.setBackground(QBrush(QColor(255, 0, 0)))
                self.table.setItem(i, 4, newitem)
                # self.table.item(1, 4).setForeground(QBrush(QColor(0,255,0)))

            i = i + 1


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Application()
    ex.show()
    app.exit(app.exec())
    # sys.exit(app.exec())
