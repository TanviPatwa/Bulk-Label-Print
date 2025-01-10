import sys
from PyQt5.QtWidgets import QApplication, QWidget, QTableWidget, QGridLayout,QTableWidgetItem
from PyQt5 import QtWidgets,Qt
from PyQt5 import QtCore,QtGui
import pandas as pd

data = pd.read_excel('ModifiedDataset1.xlsx',engine='openpyxl')
    
class App(QWidget):

    def __init__(self,data):
        super().__init__()
        self.colwidth = 130
        self.data = data
        self.rownums = len(self.data.index)
        self.colnums = len(self.data.columns)
        
        self.initUI()
           
    def initUI(self):
        self.setGeometry(640,280,800,600)
        self.createTable()
        self.layout = QGridLayout()
        self.layout.addWidget(self.tableWidget,0,0)
        self.setLayout(self.layout) 
        self.show()

    def setRowNCol(self):
        self.tableWidget.setRowCount(self.rownums)
        self.tableWidget.setColumnCount(self.colnums)

    def createTable(self):
        self.tableWidget = QTableWidget()
        self.tableWidget.setEditTriggers(QtWidgets.QTreeView.NoEditTriggers) 

        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(False)

        self.setRowNCol()

        for i in range(self.colnums):
            self.tableWidget.setColumnWidth(i, self.colwidth)
            
        self.setData()
                
        self.tableWidget.move(0,30)
        self.tableWidget.viewport().installEventFilter(self)

    def setData(self):
        for i in range(self.rownums):
            d = self.data.iloc[i].tolist()
            for j in range(len(d)):
                cell = QTableWidgetItem(str(d[j]))
                self.tableWidget.setItem(i, j, cell)

    def eventFilter(self, source, event):
        if self.tableWidget.selectedIndexes() != []:
            if event.type() == QtCore.QEvent.MouseButtonRelease:
                if event.button() == QtCore.Qt.LeftButton:
                    row = self.tableWidget.currentRow()
                    col = self.tableWidget.currentColumn()
                    if col==0 or col==1:
                        print(row,col)
                        print(self.data.iloc[row,1])
                        app = QtGui.QGuiApplication.instance()
                        app.closeAllWindows()
                        self.window=QtWidgets.QMainWindow()
                        from finalApp import Ui_MainWindow
                        self.ui=Ui_MainWindow()
                        self.ui.setupUi(self.window)
                        self.ui.code1.setText(str(self.data.iloc[row,0]))
                        self.ui.code2.setText(str(self.data.iloc[row,1]))
                        self.ui.itemName.setText(str(self.data.iloc[row,2]))
                        self.ui.vehicle.setText(str(self.data.iloc[row,3]))
                        # self.ui.company.setText(str(self.data.iloc[row,4]))
                        self.ui.quantity.setText(str(self.data.iloc[row,4]))
                        self.ui.mrp.setText(str(self.data.iloc[row,5]))
                        self.ui.remark.setText(str(self.data.iloc[row,6]))
                        self.ui.updatetime()
                        self.window.show()
                    else:
                        pass
       
        return QtCore.QObject.event(source, event)
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App(data)
    sys.exit(app.exec_())
