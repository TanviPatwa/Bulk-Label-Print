from PyQt5 import QtCore, QtGui, QtWidgets
from barcode.codex import Code128
import pandas as pd
from datetime import date 
import xlsxwriter as xl 
from PyQt5.QtWidgets import QComboBox, QMessageBox
import re

import barcode 
from barcode.writer import ImageWriter
import labels
import shutil

import os.path
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFont, stringWidth
from reportlab.graphics import shapes
from reportlab.lib import colors

from PIL import Image
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.graphics.shapes import Image
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import os


data = pd.DataFrame()
today = date.today()



df=pd.read_excel('ModifiedDataset1.xlsx',engine="openpyxl")

number="020 2507823"
email="uvpune201@gmail.com"
address="B C Traders Village Urli Devachi NR Sonai Garden Pune-412308"

def reseting(self):
    self.code1.setText("")
    self.code2.setText("")
    self.itemName.setText("")
    self.vehicle.setText("")
    self.remark.setText("")
    self.quantity.setText("")
    self.mrp.setText("")

def msgpopup(self,text):
    msg = QMessageBox()
    msg.setWindowTitle("Message")
    msg.setText(text)
    msg.exec_()
    reseting(self)


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1000, 600)
        MainWindow.setMinimumSize(QtCore.QSize(1000, 600))
        MainWindow.setMaximumSize(QtCore.QSize(1000, 800))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.code1 = QtWidgets.QLineEdit(self.centralwidget)
        self.code1.setGeometry(QtCore.QRect(110, 110, 210, 31))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.code1.setFont(font)
        self.code1.setObjectName("code1")
        self.label_1 = QtWidgets.QLabel(self.centralwidget)
        self.label_1.setGeometry(QtCore.QRect(0, 120, 111, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_1.setFont(font)
        self.label_1.setAlignment(QtCore.Qt.AlignCenter)
        self.label_1.setObjectName("label_1")
        self.title = QtWidgets.QLabel(self.centralwidget)
        self.title.setGeometry(QtCore.QRect(0, 30, 1001, 51))
        font = QtGui.QFont()
        font.setPointSize(22)
        self.title.setFont(font)
        self.title.setFrameShadow(QtWidgets.QFrame.Plain)
        self.title.setAlignment(QtCore.Qt.AlignCenter)
        self.title.setTextInteractionFlags(QtCore.Qt.NoTextInteraction)
        self.title.setObjectName("title")
        self.code2 = QtWidgets.QLineEdit(self.centralwidget)
        self.code2.setGeometry(QtCore.QRect(510, 110, 210, 31))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.code2.setFont(font)
        self.code2.setObjectName("code2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(400, 120, 111, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.searchButton = QtWidgets.QPushButton(self.centralwidget)
        self.searchButton.setGeometry(QtCore.QRect(830, 100, 151, 51))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.searchButton.clicked.connect(self.Search)
        self.searchButton.setFont(font)
        self.searchButton.setObjectName("searchButton")
        self.searchButton.setAutoDefault(True)
        self.updateButton = QtWidgets.QPushButton(self.centralwidget)
        self.updateButton.setGeometry(QtCore.QRect(830, 250, 151, 51))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.updateButton.clicked.connect(self.updateItem)
        self.updateButton.setFont(font)
        self.updateButton.setObjectName("updateButton")
        self.updateButton.setAutoDefault(True)
        self.printButton = QtWidgets.QPushButton(self.centralwidget)
        self.printButton.setGeometry(QtCore.QRect(440, 520, 151, 51))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.printButton.clicked.connect(self.showLabel)
        self.printButton.setFont(font)
        self.printButton.setObjectName("printButton")
        self.printButton.setAutoDefault(True)
        self.deleteButton = QtWidgets.QPushButton(self.centralwidget)
        self.deleteButton.setGeometry(QtCore.QRect(830, 390, 151, 51))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.deleteButton.setFont(font)
        self.deleteButton.setObjectName("deleteButton")
        self.deleteButton.clicked.connect(self.deleteItem)
        self.deleteButton.setAutoDefault(True)
        self.addButton = QtWidgets.QPushButton(self.centralwidget)
        self.addButton.setGeometry(QtCore.QRect(830, 320, 151, 51))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.addButton.setFont(font)
        self.addButton.setObjectName("addButton")
        self.addButton.clicked.connect(self.AddItem)
        self.addButton.setAutoDefault(True)
        self.vehicle = QtWidgets.QLineEdit(self.centralwidget)
        self.vehicle.setGeometry(QtCore.QRect(530, 260, 250, 31))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.vehicle.setFont(font)
        self.vehicle.setText("")
        self.vehicle.setObjectName("vehicle")
        self.itemName = QtWidgets.QLineEdit(self.centralwidget)
        self.itemName.setGeometry(QtCore.QRect(130, 260, 250, 31))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.itemName.setFont(font)
        self.itemName.setText("")
        self.itemName.setObjectName("itemName")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(30, 270, 91, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(420, 270, 111, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_4.setFont(font)
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        self.label_4.setObjectName("label_4")
        self.details = QtWidgets.QLabel(self.centralwidget)
        self.details.setGeometry(QtCore.QRect(0, 170, 1001, 71))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.details.setFont(font)
        self.details.setAlignment(QtCore.Qt.AlignCenter)
        self.details.setTextInteractionFlags(QtCore.Qt.NoTextInteraction)
        self.details.setObjectName("details")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(0, 160, 1001, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(30, 320, 91, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_5.setFont(font)
        self.label_5.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_5.setObjectName("label_5")
        self.remark = QtWidgets.QLineEdit(self.centralwidget)
        self.remark.setGeometry(QtCore.QRect(130, 310, 250, 31))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.remark.setFont(font)
        self.remark.setText("")
        self.remark.setObjectName("remark")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(300, 370, 71, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_7.setFont(font)
        self.label_7.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_7.setObjectName("label_7")
        self.mrp = QtWidgets.QLineEdit(self.centralwidget)
        self.mrp.setGeometry(QtCore.QRect(380, 360, 91, 31))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.mrp.setFont(font)
        self.mrp.setText("")
        self.mrp.setObjectName("mrp")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(450, 320, 71, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_6.setFont(font)
        self.label_6.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_6.setObjectName("label_6")
        self.quantity = QtWidgets.QLineEdit(self.centralwidget)
        self.quantity.setGeometry(QtCore.QRect(530, 310, 110, 31))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.quantity.setFont(font)
        self.quantity.setText("")
        self.quantity.setObjectName("quantity")
        self.model = QtGui.QStandardItemModel()
        self.type = QtWidgets.QComboBox(self.centralwidget) 
        self.type.setGeometry(QtCore.QRect(660, 310, 120, 31))
        types = ['SET','KIT','NOS']
        for i in types:
            t = QtGui.QStandardItem(i)
            self.model.appendRow(t)
        # self.type.setFixedSize(130,31)
        # self.type.setFont('',12)
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.type.setFont(font)
        self.type.setModel(self.model)
        self.cc = QtWidgets.QLabel(self.centralwidget)
        self.cc.setGeometry(QtCore.QRect(0, 410, 201, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.cc.setFont(font)
        self.cc.setAlignment(QtCore.Qt.AlignCenter)
        self.cc.setTextInteractionFlags(QtCore.Qt.NoTextInteraction)
        self.cc.setObjectName("cc")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(220, 420, 131, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_8.setFont(font)
        self.label_8.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_8.setObjectName("label_8")
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setGeometry(QtCore.QRect(540, 420, 71, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_9.setFont(font)
        self.label_9.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_9.setObjectName("label_9")
        self.label_10 = QtWidgets.QLabel(self.centralwidget)
        self.label_10.setGeometry(QtCore.QRect(0, 460, 161, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_10.setFont(font)
        self.label_10.setAlignment(QtCore.Qt.AlignCenter)
        self.label_10.setTextInteractionFlags(QtCore.Qt.NoTextInteraction)
        self.label_10.setObjectName("label_10")
        self.number = QtWidgets.QLineEdit(self.centralwidget)
        self.number.setGeometry(QtCore.QRect(360, 420, 151, 21))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(9)
        font.setBold(False)
        font.setWeight(50)
        self.number.setFont(font)
        self.number.setText(number)
        self.number.setObjectName("number")
        self.email = QtWidgets.QLineEdit(self.centralwidget)
        self.email.setGeometry(QtCore.QRect(620, 420, 181, 21))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(9)
        font.setBold(False)
        font.setWeight(50)
        self.email.setFont(font)
        self.email.setText(email)
        self.email.setObjectName("email")
        self.address = QtWidgets.QLineEdit(self.centralwidget)
        self.address.setGeometry(QtCore.QRect(220, 460, 581, 21))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(9)
        font.setBold(False)
        font.setWeight(50)
        self.address.setFont(font)
        self.address.setText(address)
        self.address.setObjectName("address")
        self.label_11 = QtWidgets.QLabel(self.centralwidget)
        self.label_11.setGeometry(QtCore.QRect(840, 10, 55, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_11.setFont(font)
        self.label_11.setTextInteractionFlags(QtCore.Qt.NoTextInteraction)
        self.label_11.setObjectName("label_11")
        self.date = QtWidgets.QLabel(self.centralwidget)
        self.date.setGeometry(QtCore.QRect(890, 10, 100, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.date.setFont(font)
        self.date.setTextFormat(QtCore.Qt.MarkdownText)
        self.date.text()
        self.date.setObjectName("date")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.importOrderButton = QtWidgets.QPushButton(self.centralwidget)
        self.importOrderButton.setGeometry(QtCore.QRect(830, 450, 151, 51))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.importOrderButton.setFont(font)
        self.importOrderButton.setObjectName("importOrderButton")
        self.importOrderButton.setText("Print Order List")
        self.importOrderButton.clicked.connect(self.process_order_file)


        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.code1, self.code2)
        MainWindow.setTabOrder(self.code2, self.searchButton)
        MainWindow.setTabOrder(self.searchButton, self.itemName)
        MainWindow.setTabOrder(self.itemName, self.vehicle)
        MainWindow.setTabOrder(self.vehicle, self.remark)
        MainWindow.setTabOrder(self.remark, self.quantity)
        MainWindow.setTabOrder(self.quantity, self.mrp)
        MainWindow.setTabOrder(self.mrp, self.number)
        MainWindow.setTabOrder(self.number, self.email)
        MainWindow.setTabOrder(self.email, self.address)
        MainWindow.setTabOrder(self.address, self.printButton)
        MainWindow.setTabOrder(self.printButton, self.updateButton)
        MainWindow.setTabOrder(self.updateButton, self.addButton)
        MainWindow.setTabOrder(self.addButton, self.deleteButton)
        
        self.window = MainWindow

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Label Printer"))
        self.label_1.setText(_translate("MainWindow", "Code1"))
        self.title.setText(_translate("MainWindow", "Item Master"))
        self.label_2.setText(_translate("MainWindow", "Code2"))
        self.searchButton.setText(_translate("MainWindow", "Search"))
        self.updateButton.setText(_translate("MainWindow", "Apply Changes"))
        self.printButton.setText(_translate("MainWindow", "Print"))
        self.deleteButton.setText(_translate("MainWindow", "Delete Item"))
        self.addButton.setText(_translate("MainWindow", "Add Item"))
        self.label_3.setText(_translate("MainWindow", "Item Name"))
        self.label_4.setText(_translate("MainWindow", "Vehicle"))
        self.details.setText(_translate("MainWindow", "Details"))
        self.label_5.setText(_translate("MainWindow", "Remark"))
        self.label_7.setText(_translate("MainWindow", "M.R.P"))
        self.label_6.setText(_translate("MainWindow", "Quantity"))
        self.cc.setText(_translate("MainWindow", "Customer Care"))
        self.label_8.setText(_translate("MainWindow", "Phone Number"))
        self.label_9.setText(_translate("MainWindow", "Email id"))
        self.label_10.setText(_translate("MainWindow", "Packed By"))
        self.label_11.setText(_translate("MainWindow", "Date"))
        self.date.setText(_translate("MainWindow", "TextLabel"))
        self.importOrderButton.setText(_translate("MainWindow", "Print Order List"))


    def Search(self):
        global flg
        flg = 0
        global data
        if(self.code1.text()!=""):
            regex = re.compile(f'^{self.code1.text().strip()}.*', re.IGNORECASE)
            data = df[df['Code1'].str.match(regex)== True]
            # print(data)
            if(len(data)>1):
                self.searching()
            elif len(data)==1:
                self.code1.setText(data['Code1'].values[0])
                self.code2.setText(data['Code2'].values[0])
                self.itemName.setText(data['Item'].values[0])
                self.vehicle.setText(data['Vehicle'].values[0])
                self.quantity.setText(str(data['Quantity'].values[0]))
                self.mrp.setText(str(data['M.R.P'].values[0]))
                self.remark.setText(str(data['Remark'].values[0]))
            else: flg=1

        elif(self.code2.text()!=""):
            regex = re.compile(f"^{self.code2.text().strip()}.*", re.IGNORECASE)
            data = df[df['Code2'].str.match(regex)== True]
            # print(data)
            data.to_excel('cache.xlsx',index=False)
            if(len(data)>1):
                self.searching()
            elif len(data)==1:
                self.code1.setText(data['Code1'].values[0])
                self.code2.setText(data['Code2'].values[0])
                self.itemName.setText(data['Item'].values[0])
                self.vehicle.setText(data['Vehicle'].values[0])
                self.quantity.setText(str(data['Quantity'].values[0]))
                self.mrp.setText(str(data['M.R.P'].values[0]))
                self.remark.setText(str(data['Remark'].values[0]))
                
            else: flg=1
        else:
            msgpopup(self,"Nothing to search...")
            
        if(flg):
            msgpopup(self,"Results not Found")


    def updatetime(self):
        tdate = today.strftime("%d/%m/%Y")
        self.date.setText(tdate)

    def deleteItem(self):
        if self.code1.text()!="" and self.code2.text()!="" and self.itemName.text()!="" and self.vehicle.text()!="" and self.quantity.text()!="" and self.mrp.text()!="":
            df.drop(df[df.Code2==self.code2.text()].index,inplace=True)###working
            df.to_excel('ModifiedDataset1.xlsx',index=False)
            print("Deleted")
            msgpopup(self,"Item Deleted.")


    def updateItem(self):
        global df
        if self.code1.text()!="" and self.code2.text()!="" and self.itemName.text()!="" and self.vehicle.text()!="" and self.quantity.text()!="" and self.mrp.text()!="":
            df.loc[df['Code1'] == self.code1.text(),['Code1','Code2','Item','Vehicle','Quantity','M.R.P','Remark']] = [self.code1.text(),self.code2.text(),self.itemName.text(),self.vehicle.text(),int(self.quantity.text()),float(self.mrp.text()),self.remark.text()]
            df.to_excel('ModifiedDataset1.xlsx',index=False)
            print("Updated")
            msgpopup(self,"Updated Changes")
    
    def printLabel(self):
        pass

    def AddItem(self):
        global df
        
        # Load existing data if it exists
        try:
            df = pd.read_excel('ModifiedDataset1.xlsx')
        except FileNotFoundError:
            df = pd.DataFrame()  # Start with an empty DataFrame if the file doesn't exist

        if (self.code1.text() != "" and 
            self.code2.text() != "" and 
            self.itemName.text() != "" and 
            self.vehicle.text() != "" and 
            self.quantity.text() != "" and 
            self.mrp.text() != ""):
            
            new_item = {
                'Code1': self.code1.text(),
                'Code2': self.code2.text(),
                'Item': self.itemName.text(),
                'Vehicle': self.vehicle.text(),
                'Remark': self.remark.text(),
                'Quantity': self.quantity.text(),
                'M.R.P': self.mrp.text()
            }
            
            new_item_df = pd.DataFrame([new_item])
            df = pd.concat([df, new_item_df], ignore_index=True)
            print("df appended")
            
            # Save the updated DataFrame back to Excel
            df.to_excel('ModifiedDataset1.xlsx', index=False)
            print("Added")
            msgpopup(self, "Item Added")

    
    def generateQR(self, code1):
        """
        Generate a single QR code for the item.
        """
        number = code1
        import pyqrcode
        from pyqrcode import QRCode
        
        # Generate QR code and save it as an image
        url = pyqrcode.create(number)
        url.png('new_code1.png')  # Save the QR code as a PNG file
        print("QR Code generated: new_code1.png")

    def makeLabel(self, obj):
        """
        Modify the makeLabel method to handle additional arguments for drawing the label.
        """
        print("In make Lebel function",self.code1.text())
        base_path = os.path.dirname(__file__)
        specs = labels.Specification(55, 41, 1, 1, 50, 36, corner_radius=2)
        
        # Register fonts
        registerFont(TTFont('Sans Bold', os.path.join(base_path, 'OpenSans-Bold.ttf')))
        registerFont(TTFont('Sans Regular', os.path.join(base_path, 'OpenSans-Regular.ttf')))
        
        # Extract product details from the provided object (obj)
        itemName = obj['itemName']
        mrp = obj['mrp']
        qty = obj['quantity']
        type = obj['type']
        from datetime import date
        today = date.today()
        tdate = today.strftime("%b-%y")
        remark = obj['remark']
        vehicle = obj['vehicle']
        
        # Function to write the label's contents
        def write_name(label, width, height, name):
            label.add(shapes.Image(80, height - 40, width - 90, 40, os.path.join(base_path, "new_code1.png")))
            label.add(shapes.String(5, height - 30, self.code1.text().upper(), fontSize=9, fontName='Sans Regular'))
            label.add(shapes.String(5, height - 46, itemName, fontSize=10, fontName='Sans Bold'))
            label.add(shapes.String(5, height - 57, 'MRP - ' + mrp + '/-    Qty - ' + qty + '  ' + type, fontSize=10, fontName='Sans Bold'))
            label.add(shapes.String(5, height - 68, 'VH: ' + vehicle, fontSize=9, fontName='Sans Bold'))
            label.add(shapes.String(5, height - 76, 'CC - 0202507823 , uvpune201@gmail.com', fontSize=6, fontName='Sans Regular'))
            label.add(shapes.String(5, height - 84, 'Pkd by - B.C. Traders Village Urli Devachi', fontSize=6, fontName='Sans Regular'))
            label.add(shapes.String(5, height - 92, 'NR SonaiGarden Pune, 412308. Pkd Date:' + tdate, fontSize=6, fontName='Sans Regular'))
            label.add(shapes.String(5, height - 101.5, remark, fontSize=10, fontName='Sans Bold'))
        
        # Use the write_name function to generate the label
        # write_name(label, width, height, obj)
        sheet = labels.Sheet(specs, write_name, border=False)
        sheet.add_labels(i for i in range(1))
        sheet.save('nametags.pdf')

    def showLabel(self):
        obj = {
            'itemName': self.itemName.text(),
            'mrp': self.mrp.text(),
            'quantity': self.quantity.text(),
            'type': self.type.currentText(),
            'remark': self.remark.text(),
            'vehicle': self.vehicle.text(),
            'code1': self.code1.text(),
        }
        self.generateQR(obj['code1'])
        self.makeLabel(obj)
        os.system('nametags.pdf')

    def bulkprinting(self, orders):
        """
        Generate a single PDF with labels for all products in the orders.
        Each product's label is repeated based on its quantity.
        """
        specs = labels.Specification(55, 41, 1, 1, 50, 36, corner_radius=2)

        # Define the drawing_callable function
        def make_label(label, width, height, obj):
            """
            Draw the label content using the current product details.
            """
            base_path = os.path.dirname(__file__)

            # Register fonts
            registerFont(TTFont('Sans Bold', os.path.join(base_path, 'OpenSans-Bold.ttf')))
            registerFont(TTFont('Sans Regular', os.path.join(base_path, 'OpenSans-Regular.ttf')))

            # Add QR code image to the label
            label.add(shapes.Image(80, height - 50, 50, 50, obj['qr_code_path']))
            label.add(shapes.String(5, height - 45, obj['code1'].upper(), fontSize=9, fontName='Sans Regular'))
            label.add(shapes.String(5, height - 61, obj['itemName'], fontSize=10, fontName='Sans Bold'))
            label.add(shapes.String(5, height - 72, 'Qty: ' + obj['quantity'] + '  ' + obj['type'], fontSize=10, fontName='Sans Bold'))
            label.add(shapes.String(5, height - 83, 'VH: ' + obj['vehicle'], fontSize=10, fontName='Sans Bold'))

        # Initialize the labels sheet
        sheet = labels.Sheet(specs, make_label, border=False)

        # Iterate through the orders
        for order in orders:
            code1 = order["Code1"]
            quantity = int(order["Quantity"])

            # Skip products with quantity 0
            if quantity == 0:
                continue

            # Set the product details
            self.code1.setText(code1)
            self.Search()  # Updates relevant fields like itemName, mrp, etc.

            if (flg==0): 
                # Create label data
                label_data = {
                    'code1': self.code1.text(),
                    'itemName': self.itemName.text(),
                    'mrp': self.mrp.text(),
                    'quantity': self.quantity.text(),
                    'type': self.type.currentText(),
                    'remark': self.remark.text(),
                    'vehicle': self.vehicle.text(),
                }

                # Generate QR code for the current product
                self.generateQR(label_data['code1'])  # Overwrites "new_code1.png"
                qr_code_path = os.path.join(os.path.dirname(__file__), "new_code1.png")

                # Make a unique copy of the QR code image for this label
                unique_qr_code_path = os.path.join(os.path.dirname(__file__), f"{code1}_qr.png")
                shutil.copy(qr_code_path, unique_qr_code_path)
                label_data['qr_code_path'] = unique_qr_code_path

                # Add labels for the current product based on its quantity
                for _ in range(quantity):
                    sheet.add_label(label_data)

        # Save the PDF
        sheet.save("bulk_labels.pdf")
        print("Labels generated in bulk_labels.pdf.")
        os.system("bulk_labels.pdf")

        # Clean up temporary QR code files
        for order in orders:
            unique_qr_code_path = os.path.join(os.path.dirname(__file__), f"{order['Code1']}_qr.png")
            if os.path.exists(unique_qr_code_path):
                os.remove(unique_qr_code_path)

    def process_order_file(self):
        orders = []
        df=pd.read_excel('order.xlsx',engine="openpyxl")
        orders = df.to_dict(orient='records')
    
        # Ensure Quantity is an integer
        for order in orders:
            order["Quantity"] = int(order["Quantity"])

        self.bulkprinting(orders)

    def searching(self):
        from table import App
        from PyQt5.QtWidgets import QApplication, QTableView
        self.ui = App(data)

if __name__ == "__main__":
    import sys
    tdate = today.strftime("%d/%m/%Y")
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    ui.updatetime()
    sys.exit(app.exec_())

