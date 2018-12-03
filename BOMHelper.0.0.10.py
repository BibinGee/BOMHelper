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
        self.setWindowTitle('BOM Helper Beta 1.0    Author: Daniel Gee')
        self.setGeometry(50, 50, 1200, 640)

        self.viewer = BOMViewer()
        self.reviewBoard = ReviewBoard()

        h_layout = QHBoxLayout()
        v_layout = QVBoxLayout()

        self.load_Btn = QPushButton(self)
        self.load_Btn.setObjectName('load_btn')
        self.load_Btn.setFont(QFont("Microsoft YaHei"))
        self.load_Btn.setText('Load')
        self.load_Btn.clicked.connect(self.load)
        h_layout.addWidget(self.load_Btn)

        self.find_btn = QPushButton(self)
        self.find_btn.setObjectName('find_btn')
        self.find_btn.setFont(QFont("Microsoft YaHei"))
        self.find_btn.setText('Find Difference')
        self.find_btn.clicked.connect(self.findDiff)
        h_layout.addWidget(self.find_btn)

        self.generateBOM_Btn = QPushButton(self)
        self.generateBOM_Btn.setObjectName('gntBOM_Btn')
        self.generateBOM_Btn.setFont(QFont("Microsoft YaHei"))
        self.generateBOM_Btn.setText('Generate Location BOM')
        self.generateBOM_Btn.clicked.connect(self.generateBOM)
        h_layout.addWidget(self.generateBOM_Btn)

        v_layout.addLayout(h_layout)

        self.table = QTableWidget()
        self.table.setRowCount(100)
        self.table.setColumnCount(5)
        self.table.resizeColumnToContents(2)
        # Description column
        self.table.setColumnWidth(1, 500)
        # Location column
        self.table.setColumnWidth(2, 300)
        # Quantity column
        self.table.setColumnWidth(3, 90)
        # Check column
        self.table.setColumnWidth(4, 90)
        self.table.setFont(QFont("Microsoft YaHei"))
        self.table.setHorizontalHeaderLabels(['Part number', 'Description', 'Location', 'Qty', 'Checked'])
        v_layout.addWidget(self.table)

        self.path_lb = QLabel(self)
        self.path_lb.setObjectName('path_lb')
        self.path_lb.setFont(QFont("Microsoft YaHei"))
        self.path_lb.setText('')
        v_layout.addWidget(self.path_lb)

        self.setLayout(v_layout)

        self.items = list()
        self.bom_name = str()

        self.find_btn.setEnabled(False)
        self.generateBOM_Btn.setEnabled(False)

    @pyqtSlot()
    def load(self):
        d, t = QFileDialog.getOpenFileName(self, 'Open', './', 'Excel(*.xls *.xlsx)')
        print(d)
        if d.find('Subassy') > 0:
            self.path_lb.setText(d)
            self.bom_name = re.findall('assy (.*) .', d)[0]
            # print(self.bom_name)
            self.table.clearContents()
            self.items.clear()
            self.filter(d)

            if not self.generateBOM_Btn.isEnabled():
                self.generateBOM_Btn.setEnabled(True)

            if not self.find_btn.isEnabled():
                self.find_btn.setEnabled(True)
        else:
            QMessageBox.warning(self, 'Warning', 'Please load a PDX BOM!')

    @pyqtSlot()
    def findDiff(self):
        d, t = QFileDialog.getOpenFileName(self, 'Open', './', 'Excel(*.xls *.xlsx)')
        print(d)
        if d.find('Subassy') > 0:
            if not self.reviewBoard.isVisible():
                self.reviewBoard.show()
                self.reviewBoard.findPDXDiff(d, self.items)
            else:
                self.reviewBoard.close()
                self.reviewBoard.show()
                self.reviewBoard.createBOM(d, self.items)
        else:
            QMessageBox.warning(self, 'Warning', 'Please load a PDX BOM!!')

    @pyqtSlot()
    def generateBOM(self):
        # self.hide()
        if not self.viewer.isVisible():
            self.viewer.show()
            self.viewer.createBOM(self.bom_name, self.items)
        else:
            self.viewer.close()
            # print('close viewer')
            self.viewer.show()
            self.viewer.createBOM(self.bom_name, self.items)

    def filter(self, d):
        book = xlrd.open_workbook(d)
        sheets = book.sheet_by_index(0)
        rows = sheets.nrows
        # print(rows)
        self.table.setRowCount(rows)

        header_pos = dict()
        st_row = 0
        for i in range(rows):
            string = sheets.row_values(i)
            k = 0
            for s in string:
                if isinstance(s, str):
                    if s == 'Number':
                        header_pos['Number'] = k
                        st_row = i
                    elif s == 'Name':
                        header_pos['Name'] = k
                    elif s == 'Quantity':
                        header_pos['Qty'] = k
                k = k + 1
            if st_row != 0:
                break
        print(header_pos)
        for i in range(st_row + 1, rows):
            string = sheets.row_values(i)
            if isinstance(string[header_pos['Number']], float):
                number = str(int(string[header_pos['Number']]))
            else:
                number = string[header_pos['Number']]
            if len(number):
                item = {'PN': number,
                        'Desc': string[header_pos['Name']],
                        'Qty': str(string[header_pos['Qty']])}
                self.items.append(item)

        book.release_resources()
        del book
        # print(self.items)
        self.table.clearContents()
        self.table.setRowCount(200)
        self.fillTable(self.items)

    def fillTable(self, items):
        i = 0
        err = 0

        for item in items:
            # print(item)
            # newitem = QTableWidgetItem(item['PN'])
            self.table.setItem(i, 0, QTableWidgetItem(item['PN']))

            # print(item['Desc'])
            # Retrieve description, exclude Location
            des = re.findall('(.*)\nRef', item['Desc'])
            # print(des)
            if des:
                # newitem = QTableWidgetItem(des[0])
                self.table.setItem(i, 1, QTableWidgetItem(des[0]))
            else:
                # newitem = QTableWidgetItem(item['Desc'])
                self.table.setItem(i, 1, QTableWidgetItem(item['Desc']))

            # num = re.findall('(\d).\d', item['Qty'])
            if item['Qty']:
                num = float(item['Qty'])
                # print(num)
                if num > 0:
                    # print('number')
                    newitem = QTableWidgetItem(str(num))
                    newitem.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(i, 3, newitem)

                    newitem = QTableWidgetItem('√')
                    newitem.setTextAlignment(Qt.AlignCenter)
                    newitem.setBackground(QBrush(QColor(0, 255, 0)))
                    self.table.setItem(i, 4, newitem)

                elif num == 0:
                    newitem = QTableWidgetItem(str(num))
                    newitem.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(i, 3, newitem)

                    newitem = QTableWidgetItem('!')
                    newitem.setTextAlignment(Qt.AlignCenter)
                    newitem.setBackground(QBrush(QColor(255, 255, 0)))
                    self.table.setItem(i, 4, newitem)
            else:
                newitem = QTableWidgetItem(item['Qty'])
                newitem.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(i, 3, newitem)

                newitem = QTableWidgetItem('!')
                newitem.setTextAlignment(Qt.AlignCenter)
                newitem.setBackground(QBrush(QColor(255, 255, 0)))
                self.table.setItem(i, 4, newitem)

            # Retrieve Locations and compare with quantity
            string = re.findall('RefDes:(.*)', item['Desc'])
            # print(string)
            if string:
                locations = string[0]
                lo_num = len(locations.split(','))
                # print(lo_num)
                # newitem = QTableWidgetItem(locations)
                self.table.setItem(i, 2, QTableWidgetItem(locations))

                num = float(item['Qty'])
                if lo_num == int(num):
                    # print('Y')
                    newitem = QTableWidgetItem('√')
                    newitem.setBackground(QBrush(QColor(0, 255, 0)))
                    newitem.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(i, 4, newitem)
                else:
                    newitem = QTableWidgetItem('×')
                    newitem.setTextAlignment(Qt.AlignCenter)
                    newitem.setBackground(QBrush(QColor(255, 0, 0)))
                    self.table.setItem(i, 4, newitem)
                    err = err + 1
                    # self.table.item(1, 4).setForeground(QBrush(QColor(0,255,0)))
            i = i + 1
        self.table.setRowCount(i)
        QMessageBox.information(self, "Checked Result", "Find " + str(err) + ' quantity error')


class BOMViewer(QWidget):
    def __init__(self):
        super(BOMViewer, self).__init__()
        self.setWindowTitle('BOM Viewer   Author: Daniel Gee')
        self.setGeometry(100, 100, 1200, 640)

        v_layout = QVBoxLayout(self)
        h_layout = QHBoxLayout(self)

        self.bom_lb = QLabel(self)
        self.bom_lb.setObjectName('bom_lb')
        self.bom_lb.setFont(QFont("Microsoft YaHei"))
        self.bom_lb.setAlignment(Qt.AlignCenter)
        self.bom_lb.setText('')
        v_layout.addWidget(self.bom_lb)

        self.save_btn = QPushButton(self)
        self.save_btn.setObjectName('save_btn')
        self.save_btn.setFont(QFont("Microsoft YaHei"))
        self.save_btn.setText('Save to Excel')
        self.save_btn.clicked.connect(self.createExcel)
        h_layout.addWidget(self.save_btn)

        self.findDiff_btn = QPushButton(self)
        self.findDiff_btn.setObjectName('findDiff_btn')
        self.findDiff_btn.setFont(QFont("Microsoft YaHei"))
        self.findDiff_btn.setText('Find Difference')
        self.findDiff_btn.clicked.connect(self.findBOMDiff)
        h_layout.addWidget(self.findDiff_btn)

        v_layout.addLayout(h_layout)

        self.table = QTableWidget(self)
        self.table.setColumnCount(4)
        self.table.setRowCount(100)
        self.table.setHorizontalHeaderLabels(['Part number', 'Description', 'Location', 'Qty'])
        # Description column
        self.table.setColumnWidth(1, 500)
        # Location column
        self.table.setColumnWidth(2, 300)
        # Qty column
        self.table.setColumnWidth(3, 90)
        self.table.setFont(QFont("Microsoft YaHei"))
        v_layout.addWidget(self.table)
        #
        self.setLayout(v_layout)

        self.reviewBoard = ReviewBoard()

        self.BOM = list()
        self.insertions = list()
        self.SMTs = list()
        self.b_name = str()

    @pyqtSlot()
    def findBOMDiff(self):
        d, t = QFileDialog.getOpenFileName(self, 'Open', './', 'Excel(*.xls *.xlsx)')
        d = d.upper()

        if d.find('LOCATION') > 0:
            # print(d.find('LOCATION'))
            if not self.reviewBoard.isVisible():
                self.reviewBoard.show()
                self.reviewBoard.findBOMDiff(d, self.BOM)
            else:
                self.reviewBoard.close()
                self.reviewBoard.show()
                self.reviewBoard.findBOMDiff(d, self.BOM)
        else:
            QMessageBox.warning(self, 'Warning', 'File name must include <Location>!')

    def createBOM(self, name, items):
        i = 0
        misc = list()

        if name is not None:
            self.b_name = name
            self.bom_lb.setText('Location BOM: ' + name)

        self.table.clearContents()

        # initial displaying rows, roughly
        self.table.setRowCount(len(items))
        for item in items:
            if re.findall('(.*)\nRef', item['Desc']):
                self.table.setItem(i, 0, QTableWidgetItem(item['PN']))
                # Retrieve Description
                self.table.setItem(i, 1, QTableWidgetItem(re.findall('(.*)\nRef', item['Desc'])[0]))
                # Retrieve Lccation
                self.table.setItem(i, 2, QTableWidgetItem(re.findall('RefDes:(.*)', item['Desc'])[0]))
                # Retrieve quantity
                qty = len(re.findall('RefDes:(.*)', item['Desc'])[0].split(','))
                t = QTableWidgetItem(str(qty))
                t.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(i, 3, t)
                # save to a list
                item = {'PN': item['PN'], 'Desc': re.findall('(.*)\nRef', item['Desc'])[0],
                        'Location': re.findall('RefDes:(.*)', item['Desc'])[0].strip(),
                        'Qty': qty}
                self.BOM.append(item)
                i = i + 1
            if re.findall('^PCB,|^FW(.*)', item['Desc']):
                # print(item)
                self.table.setItem(i, 0, QTableWidgetItem(item['PN']))
                self.table.setItem(i, 1, QTableWidgetItem(item['Desc']))
                t = QTableWidgetItem(re.findall('(.*).000', item['Qty'])[0])
                t.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(i, 3, t)
                item = {'PN': item['PN'], 'Desc': item['Desc'], 'Qty': re.findall('(.*).000', item['Qty'])[0]}
                i = i + 1
                misc.append(item)
        # Change to actual displaying rows.
        self.table.setRowCount(i)
        self.BOM = self.BOM + misc

        # Retrieve insertion parts and SMT parts
        for b in self.BOM:
            p = re.findall('S|s(.*)', b['PN'])
            if p:
                d = {'PN': b['PN'], 'Desc': b['Desc'], 'Location': b['Location'], 'Qty': b['Qty']}
                self.SMTs.append(d)
            elif re.match ('.*(SMD)|(SMT).*', b['Desc']) is not None:
                print(b)
                d = {'PN': b['PN'], 'Desc': b['Desc'], 'Location': b['Location'], 'Qty': b['Qty']}
                self.SMTs.append(d)
            else:
                # print(b)
                if not re.findall('^PCB,|^FW(.*)', b['Desc']):
                    d = {'PN': b['PN'], 'Desc': b['Desc'], 'Location': b['Location'], 'Qty': b['Qty']}
                    self.insertions.append(d)
                else:
                    # print(b)
                    if re.findall('^PCB(.*)', b['Desc']):
                        d = {'PN': b['PN'], 'Desc': b['Desc'], 'Qty': b['Qty']}
                        self.SMTs.append(d)
                    else:
                        d = {'PN': b['PN'], 'Desc': b['Desc'], 'Qty': b['Qty']}
                        self.insertions.append(d)


    @pyqtSlot()
    def createExcel(self):
        if self.BOM:
            # print(self.BOM)
            f, t = QFileDialog.getSaveFileName(self, 'Save', '/', 'Excel(*.xls)')
            if f:
                wb = xlwt.Workbook(encoding='utf-8')
                sheet = wb.add_sheet('Location BOM')

                style = xlwt.XFStyle()

                align1 = xlwt.Alignment()
                # Horizontal center
                align1.horz = xlwt.Alignment.HORZ_CENTER
                align1.wrap = xlwt.Alignment.WRAP_AT_RIGHT
                # Vertical center
                align1.vert = xlwt.Alignment.VERT_CENTER
                style.alignment = align1

                border = xlwt.Borders()
                border.left = xlwt.Borders.THIN
                border.right = xlwt.Borders.THIN
                border.top = xlwt.Borders.THIN
                border.bottom = xlwt.Borders.THIN
                style.borders = border

                font = xlwt.Font()
                font.name = 'Microsoft YaHei'
                font.bold = True
                style.font = font

                # Location BOM title
                sheet.write_merge(0, 1, 0, 6, 'Location BOM', style)
                sheet.write_merge(2, 2, 0, 1, 'PCBA part number:', style)
                sheet.write_merge(2, 2, 2, 4, self.b_name.split(' ')[0], style)
                sheet.write_merge(3, 3, 0, 1, 'PCBA Description:', style)
                sheet.write_merge(3, 3, 2, 4,  'PCBA ASSY. of ' + self.b_name.split(' ')[0], style)
                sheet.write_merge(2, 3, 5, 6, 'Rev' + self.b_name.split(' ')[-1], style)

                sheet.write_merge(4, 4, 0, 6, '', style)

                # Location BOM header
                sheet.write(5, 0, 'Index', style)

                sheet.col(1).width = 256 * 15
                sheet.write(5, 1, 'Part number', style)

                sheet.col(2).width = 256 * 60
                sheet.write(5, 2, 'Description', style)

                sheet.col(3).width = 256 * 6
                sheet.write(5, 3, 'Qty', style)

                sheet.col(4).width = 256 * 30
                sheet.write(5, 4, 'Loaction', style)

                sheet.col(5).width = 256 * 10
                sheet.write(5, 5, 'Position', style)

                sheet.col(6).width = 256 * 10
                sheet.write(5, 6, 'Remark', style)

                # setup cell style
                align2 = xlwt.Alignment()
                align2.horz = xlwt.Alignment.HORZ_LEFT
                align2.wrap = xlwt.Alignment.WRAP_AT_RIGHT
                # Vertical center
                align2.vert = xlwt.Alignment.VERT_CENTER
                style.alignment = align2

                c_font = xlwt.Font()
                c_font.name = 'Microsoft YaHei'
                c_font.bold = False
                style.font = c_font

                # fill with insertion parts
                i = 6
                sheet.write_merge(i, i, 0, 6, 'Insertion Parts', style)
                # i = i + 1
                k = 1
                for p in self.insertions:
                    try:
                        sheet.write(i + k, 0, k, style)
                        sheet.write(i + k, 1, p['PN'], style)
                        sheet.write(i + k, 2, p['Desc'], style)
                        sheet.write(i + k, 3, p['Qty'], style)
                        if 'Location' in p:
                            sheet.write(i + k, 4, p['Location'], style)
                        else:
                            sheet.write(i + k, 4, '', style)
                        
                        sheet.write(i + k, 5, 'MI', style)
                        sheet.write(i + k, 6, '', style)
                    except Exception as e:
                        print('Exception in writing to Excel: ', e)
                    k = k + 1

                # fill with SMT parts
                sheet.write_merge(i + k, i + k, 0, 6, 'SMT Parts', style)
                i = i + k
                k = 1
                for p in self.SMTs:
                    try:
                        sheet.write(i + k, 0, k, style)
                        sheet.write(i + k, 1, p['PN'], style)
                        sheet.write(i + k, 2, p['Desc'], style)
                        sheet.write(i + k, 3, p['Qty'], style)
                        if 'Location' in p:
                            sheet.write(i + k, 4, p['Location'], style)
                        else:
                            sheet.write(i + k, 4, '', style)
                        
                        sheet.write(i + k, 5, 'SMT', style)
                        sheet.write(i + k, 6, '', style)
                    except Exception as e:
                        print('Exception in writing to Excel: ', e)
                    k = k + 1
                wb.save(f)
                QMessageBox.information (self, "Complete", 'Location save complete!!')


class ReviewBoard(QTableWidget):
    def __init__(self):
        super(ReviewBoard, self).__init__()
        self.setWindowTitle('Review Board   Author: Daniel Gee')
        self.setGeometry(150, 150, 1000, 480)

        # self.setRowCount(10)
        self.setColumnCount(4)
        self.setFont(QFont("Microsoft YaHei"))
        self.setHorizontalHeaderLabels(['Part number', 'Current content', 'Referred content', 'Comments'])

        # Current content column
        self.setColumnWidth(1, 300)
        # Referred content column
        self.setColumnWidth(2, 300)
        # Comment column
        self.setColumnWidth(3, 200)

        # Disable edit
        # self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setWordWrap(True)
        self.setTextElideMode(Qt.ElideNone)
        self.resizeRowsToContents()
        self.BOM = list()

    # find difference in PDX BOM
    def findPDXDiff(self, path, cur_pdxbom):

        self.BOM = list()
        self.clearContents()

        if len(cur_pdxbom) and path:
            rb = xlrd.open_workbook(path)
            sheets = rb.sheet_by_index(0)
            rows = sheets.nrows

            # set a initial row count
            self.setRowCount(200)
            header_pos = dict()
            st_row = 0
            for i in range(rows):
                string = sheets.row_values(i)
                k = 0
                for s in string:
                    if isinstance(s, str):
                        if s == 'Number':
                            header_pos['Number'] = k
                            st_row = i
                        elif s == 'Name':
                            header_pos['Name'] = k
                        elif s == 'Quantity':
                            header_pos['Qty'] = k
                    k = k + 1
                if st_row != 0:
                    break

            for i in range(st_row + 1, rows):
                try:
                    string =  sheets.row_values(i)
                    if isinstance(string[header_pos['Number']], float):
                        number = str(int(string[header_pos['Number']]))
                    else:
                        number = string[header_pos['Number']]
                    if len(number):
                        item = {'PN': number,
                                'Desc': string[header_pos['Name']],
                                'Qty': str(string[header_pos['Qty']])}
                    # print(item)
                    self.BOM.append(item)
                except Exception as e:
                    print(e)
            rb.release_resources()
            del rb

            i = 0
            f = False
            p_err = list()
            d_err = list()
            q_err = list()
            err = 0
            for p in cur_pdxbom:
                for d in self.BOM:
                    if p['PN'] == d['PN']:
                        f = True
                        break
                if f:
                    f = False

                    if p['Desc'] == d['Desc']:
                        if float(p['Qty']) == float(d['Qty']):
                            i = i + 1
                        else:
                            print('cur pdx BOM: ' + p['Qty'], '\n', 'referred BOM: ' + d['Qty'])
                            item = {'PN': p['PN'], 'cur': p['Desc'] + ' --> ' + str(float(p['Qty'])),
                                    'ref': d['Desc'] + ' --> ' + str(float(d['Qty']))}
                            q_err.append(item)
                            err = err + 1
                    else:
                        print(p['Desc'], '\n', d['Desc'])
                        item = {'PN': p['PN'], 'cur': p['Desc'], 'ref': d['Desc']}
                        d_err.append(item)
                        err = err + 1
                        #
                        # if float(p['Qty']) != float(d['Qty']):
                        #     print(p['Qty'], '\n', d['Qty'])
                        #     item = {'PN': p['PN'], 'cur': p['Desc'] + ' --> ' + str(float(p['Qty'])),
                        #             'ref': d['Desc'] + ' --> ' + str(float(d['Qty']))}
                        #     q_err.append(item)
                        #     err = err + 1
                else:
                    print(p['PN'], '\n', d['PN'])
                    item = {'PN': p['PN'], 'cur': p['PN'], 'ref': 'No matched part number'}
                    p_err.append(item)
                    err = err + 1

            print('showing difference')

            i = 0
            if len(p_err):
                print('Part number Error:', len(p_err))
                for p in p_err:
                    m = QTableWidgetItem(p['PN'])
                    m.setBackground(QBrush(QColor(255, 250, 205)))
                    self.setItem(i, 0, m)
                    m = QTableWidgetItem(p['cur'])
                    m.setBackground(QBrush(QColor(255, 250, 205)))
                    self.setItem(i, 1, m)
                    m = QTableWidgetItem(p['ref'])
                    m.setBackground(QBrush(QColor(255, 250, 205)))
                    self.setItem(i, 2, m)
                    m = QTableWidgetItem('Part number difference')
                    m.setBackground(QBrush(QColor(255, 250, 205)))
                    self.setItem(i, 3, m)
                    i = i + 1

            if len(d_err):
                print('Description Error:', len(d_err))
                for p in d_err:
                    m = QTableWidgetItem(p['PN'])
                    m.setBackground(QBrush(QColor(0, 255, 255)))
                    self.setItem(i, 0, m)
                    m = QTableWidgetItem(p['cur'])
                    m.setBackground(QBrush(QColor(0, 255, 255)))
                    self.setItem(i, 1, m)
                    m = QTableWidgetItem(p['ref'])
                    m.setBackground(QBrush(QColor(0, 255, 255)))
                    self.setItem(i, 2, m)
                    m = QTableWidgetItem('Description difference')
                    m.setBackground(QBrush(QColor(0, 255, 255)))
                    self.setItem(i, 3, m)
                    i = i + 1

            if len(q_err):
                print('Quantity Error:', len(q_err))
                for p in q_err:
                    m = QTableWidgetItem(p['PN'])
                    m.setBackground(QBrush(QColor(192, 255, 62)))
                    self.setItem(i, 0, m)
                    m = QTableWidgetItem(p['cur'])
                    m.setBackground(QBrush(QColor(192, 255, 62)))
                    self.setItem(i, 1, m)
                    m = QTableWidgetItem(p['ref'])
                    m.setBackground(QBrush(QColor(192, 255, 62)))
                    self.setItem(i, 2, m)
                    m = QTableWidgetItem('Quantity difference')
                    m.setBackground(QBrush(QColor(192, 255, 62)))
                    self.setItem(i, 3, m)
                    i = i + 1
            self.setRowCount(err)
            self.setWordWrap(True)
            self.setTextElideMode(Qt.ElideNone)
            self.resizeRowsToContents()
            QMessageBox.warning(self, "Result", "Find " + str(err) + ' Differnce')

        else:
            QMessageBox.warning(self, "Warning", 'Please load a PDX BOM firstly')

    # find difference in Location BOM
    def findBOMDiff(self, path, cur_locationbom):
        self.BOM = list()

        if len(cur_locationbom):

            self.clearContents()
            rb = xlrd.open_workbook(path)
            sheet = rb.sheet_by_index(0)
            rows = sheet.nrows

            self.setRowCount(200)

            for i in range(5, rows):
                strings = sheet.row_values(i)
                k = 0
                for s in strings:

                        # m = re.match('[Ss0-9AB]{4,5}-[PGB0-9R]{4,5}', s)
                    # print(s)
                    if re.match('[Ss0-9AB]{4,5}-[PGA-D0-9R]{4,5}', str(s)) is not None:
                        if len(strings[k + 1]):
                            item = {'PN': strings[k], 'Desc': strings[k + 1], 'Location': strings[k + 3],
                                    'Qty': str(strings[k + 2])}
                            self.BOM.append(item)
                    elif re.match('[0-9]{7}', str(s)) is not None:
                        if len(strings[k + 1]):
                            item = {'PN': str(int(strings[k])), 'Desc': strings[k + 1], 'Location': strings[k + 3],
                                    'Qty': str(strings[k + 2])}
                            self.BOM.append(item)
                            print(str(int(s)))
                    k = k + 1
            # print(self.BOM)
            QMessageBox.information(self, "Result", 'Load Reference BOM with ' + str(len(self.BOM)) + 'items')
            p_err = list()
            d_err = list()
            q_err = list()
            err = 0

            f = False

            for p in cur_locationbom:
                for d in self.BOM:
                    if p['PN'] == d['PN']:
                        f = True
                        break
                if f:
                    f = False

                    if p['Location'].replace(' ', '') == d['Location'].replace(' ', ''):
                        if float(p['Qty']) == float(d['Qty']):
                            i = i + 1
                        else:
                            print(p['Qty'], '\n', d['Qty'])
                            item = {'PN': p['PN'], 'cur': p['Location'] + ' --> ' + str(float(p['Qty'])),
                                    'ref': d['Location'] + ' --> ' + str(float(d['Qty']))}
                            q_err.append(item)
                            err = err + 1
                    else:
                        print(p['Location'], '\n', d['Location'])
                        item = {'PN': p['PN'], 'cur': p['Location'], 'ref': d['Location']}
                        d_err.append(item)
                        err = err + 1

                else:
                    print(p['PN'], '\n', d['PN'])
                    item = {'PN': p['PN'], 'cur': p['PN'], 'ref': 'No matched part number'}
                    p_err.append(item)
                    err = err + 1

            i = 0
            if len(p_err):
                print('Part number Error:', len(p_err))
                for p in p_err:
                    m = QTableWidgetItem(p['PN'])
                    m.setBackground(QBrush(QColor(255, 250, 205)))
                    self.setItem(i, 0, m)
                    m = QTableWidgetItem(p['cur'])
                    m.setBackground(QBrush(QColor(255, 250, 205)))
                    self.setItem(i, 1, m)
                    m = QTableWidgetItem(p['ref'])
                    m.setBackground(QBrush(QColor(255, 250, 205)))
                    self.setItem(i, 2, m)
                    m = QTableWidgetItem('Part number difference')
                    m.setBackground(QBrush(QColor(255, 250, 205)))
                    self.setItem(i, 3, m)
                    i = i + 1

            if len(d_err):
                print('Description Error:', len(d_err))
                for p in d_err:
                    m = QTableWidgetItem(p['PN'])
                    m.setBackground(QBrush(QColor(0, 255, 255)))
                    self.setItem(i, 0, m)
                    m = QTableWidgetItem(p['cur'])
                    m.setBackground(QBrush(QColor(0, 255, 255)))
                    self.setItem(i, 1, m)
                    m = QTableWidgetItem(p['ref'])
                    m.setBackground(QBrush(QColor(0, 255, 255)))
                    self.setItem(i, 2, m)
                    m = QTableWidgetItem('Location difference')
                    m.setBackground(QBrush(QColor(0, 255, 255)))
                    self.setItem(i, 3, m)
                    i = i + 1

            if len(q_err):
                print('Quantity Error:', len(q_err))
                for p in q_err:
                    m = QTableWidgetItem(p['PN'])
                    m.setBackground(QBrush(QColor(192, 255, 62)))
                    self.setItem(i, 0, m)
                    m = QTableWidgetItem(p['cur'])
                    m.setBackground(QBrush(QColor(192, 255, 62)))
                    self.setItem(i, 1, m)
                    m = QTableWidgetItem(p['ref'])
                    m.setBackground(QBrush(QColor(192, 255, 62)))
                    self.setItem(i, 2, m)
                    m = QTableWidgetItem('Quantity difference')
                    m.setBackground(QBrush(QColor(192, 255, 62)))
                    self.setItem(i, 3, m)
                    i = i + 1
            self.setRowCount(err)
            self.setWordWrap(True)
            self.setTextElideMode(Qt.ElideNone)
            self.resizeRowsToContents()
            QMessageBox.warning(self, "Result", "Find " + str(err) + ' Differnce')

        else:
            QMessageBox.warning(self, "Warning", 'Please load a Location BOM firstly')


if __name__ == '__main__':
    app = QApplication(sys.argv)

    ex = Application()

    ex.show()

    app.exit(app.exec())
    # sys.exit(app.exec())
