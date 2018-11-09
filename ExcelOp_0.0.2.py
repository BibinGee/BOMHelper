import xlrd
import re
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys
import xlwt

class Application(QWidget):
    def __init__(self):
        super(Application, self).__init__()
        self.setWindowTitle('BOM Helper')
        self.setGeometry(400, 400, 1200, 480)

        self.viewer = BOMViewer()

        hlayout = QHBoxLayout()

        self.load_Btn = QPushButton(self)
        self.load_Btn.setObjectName('load_btn')
        self.load_Btn.setText('Load')
        self.load_Btn.clicked.connect(self.load)
        hlayout.addWidget(self.load_Btn)

        self.generateBOM_Btn = QPushButton(self)
        self.setObjectName('gntBOM_Btn')
        self.generateBOM_Btn.setText('Generate Location BOM')
        self.generateBOM_Btn.clicked.connect(self.generateBOM)
        hlayout.addWidget(self.generateBOM_Btn)
        # self.setLayout(hlayout)

        vlayout = QVBoxLayout()
        vlayout.addLayout(hlayout)

        self.table = QTableWidget()
        self.table.setRowCount(100)
        self.table.setColumnCount(5)
        # self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setColumnWidth(1, 500)
        self.table.setColumnWidth(2, 300)
        self.table.setColumnWidth(3, 100)

        self.table.setHorizontalHeaderLabels(['Part number', 'Description', 'Location', 'Qty', 'Checked'])
        vlayout.addWidget(self.table)

        self.setLayout(vlayout)

        self.items = []
        self.generateBOM_Btn.setEnabled(False)

    @pyqtSlot()
    def load(self):
        d, t = QFileDialog.getOpenFileName(self, 'Open', './')
        print(d)
        if d:
            self.filter(d)
            self.generateBOM_Btn.setEnabled(True)

    @pyqtSlot()
    def generateBOM(self):
        # self.hide()
        self.viewer.show()
        self.viewer.createBOM(self.items)

    def filter(self, d):
        book = xlrd.open_workbook(d)
        sheets = book.sheet_by_index(0)
        rows = sheets.nrows
        self.table.setRowCount(rows - 6)

        for i in range(7, rows):
            try:
                item = {'PN': sheets.row_values(i)[5], 'Desc': sheets.row_values(i)[7],
                        'Qty': sheets.row_values(i)[9]}
                # print(item)
                self.items.append(item)
            except Exception as e:
                print(e)

            # if re.findall('(.*)\nRef', sheets.row_values(i)[7]):
            #     item = {'PN': sheets.row_values(i)[5], 'Desc': re.findall('(.*)\nRef', sheets.row_values(i)[7]),
            #             'Location': re.findall('RefDes:(.*)', sheets.row_values(i)[7]),
            #             'Qty': sheets.row_values(i)[9]}
            #     self.items.append(item)

        self.fillTable(self.items)

    def fillTable(self, items):
        i = 0

        for item in items:
            newitem = QTableWidgetItem(item['PN'])
            self.table.setItem(i, 0, newitem)

            # print(item['Desc'])
            des = re.findall('(.*)\nRef', item['Desc'])
            # print(des)
            if des:
                newitem = QTableWidgetItem(des[0])
                self.table.setItem(i, 1, newitem)
            else:
                newitem = QTableWidgetItem(item['Desc'])
                self.table.setItem(i, 1, newitem)

            num = re.findall('(.*).000', item['Qty'])
            if num:
                newitem = QTableWidgetItem(num[0])
                self.table.setItem(i, 3, newitem)

            string = re.findall('RefDes:(.*)', item['Desc'])
            if string:
                locations = string[0]
                lo_num = len(locations.split(','))
                # print(lo_num)
                newitem = QTableWidgetItem(locations)
                self.table.setItem(i, 2, newitem)

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


class BOMViewer(QWidget):
    def __init__(self):
        super(BOMViewer, self).__init__()
        self.setWindowTitle('BOM Viewer')
        self.setGeometry(410, 410, 1200, 480)

        v_layout = QVBoxLayout(self)
        self.save_btn = QPushButton(self)
        self.save_btn.setObjectName('save_btn')
        self.save_btn.setText('Save to Excel')
        self.save_btn.clicked.connect(self.createExcel)
        v_layout.addWidget(self.save_btn)

        self.table = QTableWidget(self)
        self.table.setColumnCount(4)
        self.table.setRowCount(100)
        self.table.setHorizontalHeaderLabels(['Part number', 'Description', 'Location', 'Qty'])
        self.table.setColumnWidth(1, 500)
        self.table.setColumnWidth(2, 300)
        v_layout.addWidget(self.table)
        #
        self.setLayout(v_layout)

        self.BOM = []

    def createBOM(self, items):
        i = 0
        for item in items:
            if re.findall('(.*)\nRef', item['Desc']):
                # n_item = QTableWidgetItem(item['PN'])
                self.table.setItem(i, 0, QTableWidgetItem(item['PN']))

                # n_item = QTableWidgetItem(re.findall('(.*)\nRef', item['Desc'])[0])
                self.table.setItem(i, 1, QTableWidgetItem(re.findall('(.*)\nRef', item['Desc'])[0]))

                self.table.setItem(i, 2, QTableWidgetItem(re.findall('RefDes:(.*)', item['Desc'])[0]))

                # num = re.findall('(.*).000', item['Qty'])
                # n_item = QTableWidgetItem(re.findall('(.*).000', item['Qty'])[0])
                qty = len(re.findall('RefDes:(.*)', item['Desc'])[0].split(','))
                self.table.setItem(i, 3, QTableWidgetItem(str(qty)))

                item = {'PN': item['PN'], 'Desc': re.findall('(.*)\nRef', item['Desc'])[0],
                        'Location': re.findall('RefDes:(.*)', item['Desc'])[0],
                        'Qty': qty}
                self.BOM.append(item)
                i = i + 1
        self.table.setRowCount(i)

        # self.createExcel(self.BOM)

    @pyqtSlot()
    def createExcel(self):
        i = 3
        if self.BOM:
            wb = xlwt.Workbook(encoding='utf-8')
            sheet = wb.add_sheet('Location BOM')
            sheet.write(2, 1, 'Part number')
            sheet.write(2, 2, 'Description')
            sheet.write(2, 3, 'Location')
            sheet.write(2, 4, 'Qty')
            for b in self.BOM:
                try:
                    sheet.write(i, 0, i + 1)
                    sheet.write(i, 1, b[i]['PN'])
                    sheet.write(i, 2, b[i]['Desc'])
                    sheet.write(i, 3, b[i]['Location'])
                    sheet.write(i, 4, b[i]['Qty'])
                except Exception as e:
                    print('Exception in writing to Excel: ', e)
                i = i + 1
            #
            f, t = QFileDialog.getSaveFileName(self, 'Save', '/')
            if f:
                wb.save(f)


if __name__ == '__main__':

    app = QApplication(sys.argv)

    ex = Application()

    ex.show()

    app.exit(app.exec())
    # sys.exit(app.exec())
