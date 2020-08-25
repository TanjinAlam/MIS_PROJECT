import requests
import bs4
from datetime import datetime
import urllib.request as ur
import urllib.error
import os
from PyQt5.QtCore import pyqtSlot, QPoint, QObject, QFile, QDir
from PyQt5.QtGui import QTextDocument, QTextCursor, QTextTableFormat
from PyQt5.QtWidgets import QDialog, QApplication, QFileDialog, QGraphicsPixmapItem, QLabel, QGraphicsScene, QSlider, \
    QLineEdit, QMainWindow, QMessageBox, QTableWidgetItem, QAbstractItemView, QMenu, QTableWidget
from PyQt5.uic import loadUi
import sys
from PyQt5.QtWidgets import QApplication, QGraphicsView, QGraphicsScene, QGraphicsEllipseItem
import csv
import pandas as pd
from PyQt5 import QtPrintSupport
from PyQt5.QtCore import Qt
from PyQt5 import QtWidgets, QtCore, QtPrintSupport, QtGui




class MainForm(QMainWindow):
    def __init__(self):
        super(MainForm, self).__init__()
        loadUi('MIS.ui', self)
        self.known_values = ['SL NO', 'Date' , 'TRADING CODE', 'Update Status']
        self.fname = "Liste"
        self.tableView = TableWidgetDragRows()
        self.tableView.setGridStyle(1)
        self.tableView.setCornerButtonEnabled(False)
        self.actionLoad.triggered.connect(self.loadCsv)
        self.actionSave_FIle.triggered.connect(self.writeCsv)
        self.pushButton_1.clicked.connect(self.fetch_data)
        self.actionPrint_Preview.triggered.connect(self.handlePreview)

    def loadCsv(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "Open CSV",
                                                  (QDir.homePath() + "/Dokumente/CSV"), "CSV (*.csv *.tsv *.txt)")
        if fileName:
            self.loadCsvOnOpen(fileName)


    def fetch_data(self):
        update_time = self.lineEdit_1.text()
        update_time = update_time[0:3]
        print(update_time)
        if getattr(sys, 'frozen', False):
            path = os.path.join(sys._MEIPASS, "Company_Name.csv")
            df = pd.read_csv(path)
        else:
            df = pd.read_csv('Company_Name.csv')

        c_name = []
        for i in df[' TRADING CODE']:
            c_name.append(i)

        tm = datetime.today().strftime('%Y-%m-%d')
        row0 = ['SL NO', 'Date', 'TRADING CODE', 'July Update Status']
        sl = []
        date = []
        trading_code = []
        July_update = []
        idx = 0
        l = len(c_name)
        print("Please wait")
        print("Connecting to internet...")
        while l != 0:

            print(c_name[idx])
            url = ('https://www.dsebd.org/displayCompany.php?name=' + str(c_name[idx]))
            r = requests.get(url)
            soup = bs4.BeautifulSoup(r.text, "lxml")
            try:
                result = ur.urlopen(url)
            except urllib.error.HTTPError as e:
                print("ERROR")
            except requests.exceptions.InvalidURL:
                print("ERROR")

            table = soup.find_all('table', id='company')
            x = table[10:11]
            value = []
            for i in x:
                value.append((i.text))

            value = [w.replace("\r", "") for w in value]
            value = [w.replace("\t", "") for w in value]
            value = [w.replace("\n", "") for w in value]

            value = str(value)
            if update_time in value:
                sl.append(idx + 1)
                date.append(tm)
                trading_code.append(c_name[idx])
                July_update.append("YES")
            else:
                sl.append(idx + 1)
                date.append(tm)
                trading_code.append(c_name[idx])
                July_update.append("NO")
            if idx == 5:
                break
            idx += 1
            l -= 1

        rows = zip(sl, date, trading_code, July_update)
        with open('Latest_update.csv', 'w', encoding='utf-8',
                  newline='') as csvfile:
            lw = csv.writer(csvfile)
            lw.writerow(row0)
            for item in rows:
                lw.writerow(item)
        n_filename = 'Latest_update.csv'
        self.done(n_filename)

    def done(self, n_filename):
        if n_filename:
            print("RECEIVED")
            df = pd.read_csv(n_filename, header=None, delimiter=',', keep_default_na=False, error_bad_lines=False)
            print(len(df.index))
            print(len(df.columns))
            self.tableWidget.setColumnCount(len(df.columns))
            self.tableWidget.setRowCount(len(df.index))
            for i in range(len(df.index)):
                for j in range(len(df.columns)):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

            for j in range(self.tableWidget.columnCount()):
                m = QTableWidgetItem(str(j))
                self.tableWidget.setHorizontalHeaderItem(j, m)


    def loadCsvOnOpen(self, fileName):
        if fileName:
            print("RECEIVED")
            df = pd.read_csv(fileName, header=None, delimiter=',', keep_default_na=False, error_bad_lines=False)
            df = df[1:]
            df = df.sort_values(by=3,axis=0, ascending=False)
            self.tableWidget.setColumnCount(len(df.columns))
            self.tableWidget.setRowCount(len(df.index))
            for i in range(len(df.index)):
                for j in range(len(df.columns)):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

            self.tableView.setColumnCount(len(df.columns))
            self.tableView.setRowCount(len(df.index))

            for i in range(len(df.index)):
                for j in range(len(df.columns)):
                    self.tableView.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

    def writeCsv(self):
        path, _ = QFileDialog.getSaveFileName(self, 'Save File', QDir.homePath() + "/export.csv",
                                              "CSV Files(*.csv *.txt)")
        if path:
            with open(path, 'w', newline='') as stream:
                writer = csv.writer(stream)
                writer.writerow(self.known_values)
                for row in range(self.tableWidget.rowCount()):
                    rowdata = []
                    for column in range(self.tableWidget.columnCount()):
                        item = self.tableWidget.item(row, column)
                        if item is not None:
                            rowdata.append(item.text())
                    writer.writerow(rowdata)


    def handlePreview(self):
        dialog = QtPrintSupport.QPrintPreviewDialog()
        dialog.paintRequested.connect(self.handlePaintRequest)
        dialog.exec_()

    def handlePaintRequest(self, printer):
        printer.setDocName(self.fname)
        document = QTextDocument()
        cursor = QTextCursor(document)
        model = self.tableView.model()

        tableFormat = QTextTableFormat()
        tableFormat.setBorder(0.2)
        tableFormat.setBorderStyle(3)
        tableFormat.setCellSpacing(0);
        tableFormat.setTopMargin(0);
        tableFormat.setCellPadding(4)
        table = cursor.insertTable(model.rowCount() + 1, model.columnCount(), tableFormat)

        model = self.tableView.model()
        ### get headers
        myheaders = []
        for i in range(0, model.columnCount()):
            myheader = model.headerData(i, Qt.Horizontal)
            cursor.insertText(str(myheader))
            cursor.movePosition(QTextCursor.NextCell)
        ### get cells
        for row in range(0, model.rowCount()):
            for col in range(0, model.columnCount()):
                index = model.index(row, col)
                cursor.insertText(str(index.data()))
                cursor.movePosition(QTextCursor.NextCell)
        document.print_(printer)

class TableWidgetDragRows(QTableWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setDragDropOverwriteMode(False)
        self.setDropIndicatorShown(True)

        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setDragDropMode(QAbstractItemView.InternalMove)

app = QApplication(sys.argv)
window = MainForm()
window.show()
sys.exit(app.exec_())