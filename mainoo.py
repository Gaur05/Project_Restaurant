import sys
import win32ui  #For printing the bill
import win32print
import win32con
from PyQt5.QtCore import pyqtSlot
from PyQt5.uic import loadUi 
from PyQt5.QtWidgets import QDialog ,QApplication, QMainWindow
# from PyQt5 import QtWidgets
from PyQt5 import QtGui
import xlrd #For excel file
import datetime


# class MyMainWindow(QMainWindow, Ui_MainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setupUi(self)

class theseencode(QDialog, QMainWindow):
    def __init__(self):
        super(theseencode,self).__init__()
        loadUi("C:\AkshatGaur\COLLEGE\Python projects\Restaurent Final\Program\BillingFinal.ui",self)
        self.Receipt=''
        self.TotalCost=0
        self.ReceiptNumber=0
        self.Reset_GUI()
        self.Print.clicked.connect(self.Print)
        self.Reset.clicked.connect(self.Reset)
        self.OK.clicked.connect(self.pushButton)
        self.TEXTRECEIPT.setText.connect(self.Receipt)

        self.Data=xlrd.open_workbook('ItemsDetails.xlsx')

        self.sheet1 = self.Data.sheet_by_index(0)

    @pyqtSlot()
    def Reset_GUI(self):

        #Write zero

        #SNACKS ITEMS
        self.SN1.setText('0')
        self.Samosa01.setText('0')
        self.Vburger01.setText('0')
        self.SN4.setText('0')
        self.Egg01.setText('0')
        self.SN4_2.setText('0')
        
        #SHAKES ITEMS
        self.SH11.setText('0')
        self.Mango01.setText('0')
        self.SH33.setText('0')
        self.SH44.setText('0')
        self.SH55.setText('0')
        self.Strawberry01.setText('0')

        #HOT_DRINKS ITEMS
        self.HT1.setText('0')
        self.HT2.setText('0')
        self.HT3.setText('0')
        self.HT4.setText('0')
        self.HT5.setText('0')

        #JUICE ITEMS
        self.J11.setText('0')
        self.J2.setText('0')
        self.J3.setText('0')
        self.J4.setText('0')
        self.J5.setText('0')

        self.textEdit.setText('')
        self.ReceiptNumber=self.ReceiptNumber+1
        DateTime = datetime.datetime.now()
        self.Receipt = str('Receipt@%s\t\t%s/%s/%s  %s:%s:%s' % (
            self.ReceiptNumber, DateTime.day, DateTime.month, DateTime.year, DateTime.hour, DateTime.minute, DateTime.second ))
        # self.TextReceipt.setText('%s'% self.Receipt)

    def getData(self):

        #Take data from text edit

        #SNACKS ITEM
        self.ROLLPARATHA = self.SN1.toPlainText()
        self.SAMOSA = self.Samosa01.toPlainText()
        self.VEGBURGER = self.Vburger01.toPlainText()
        self.CHICKENBURGER = self.SN4.toPlainText()
        self.EGGBURGER = self.Egg01.toPlainText()
        self.SANDWICH = self.SN4_2.toPlainText()

        #SHAKES ITEM
        self.BANANA = self.SH11.toPlainText()
        self.MANGO = self.Mango01.toPlainText()
        self.ALMOND = self.SH33.toPlainText()
        self.CHOCOLATE = self.SH44.toPlainText()
        self.OREO = self.SH55.toPlainText()
        self.STRAWBERRY = self.Strawberry01.toPlainText()

        #HOT_DRINKS ITEM
        self.GREENTEA = self.HT1.toPlainText()
        self.MILKTEA = self.HT2.toPlainText()
        self.COFFEE = self.HT3.toPlainText()
        self.HOTMILK = self.HT4.toPlainText()
        self.BLACKCOFFEE = self.HT5.toPlainText()

        #JUICE ITEM
        self.ORANGE = self.J11.toPlainText()
        self.MANGO = self.J2.toPlainText()
        self.APPLE = self.J3.toPlainText()
        self.GUAVA = self.J4.toPlainText()
        self.PINEAPPLE = self.J5.toPlainText()



    def Get_Price(self):

        self.getData()

        self.ROLLPARATHA=int(self.ROLLPARATHA)

        if(self.ROLLPARATHA>=1):

            Price=self.sheet1.cell_value(1,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.ROLLPARATHA*Price
            self.TotalCost = int(self.TotalCost)+Price
            # print(self.ROLLPARATHA)
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(1,0),self.sheet1.cell_value(1,1),int(self.sheet1.cell_value(1,3)),self.ROLLPARATHA 
            
            self.TextReceipt.setText(self.Receipt)


        self.SAMOSA=int(self.SAMOSA)

        if(self.SAMOSA >= 1):

            Price=self.sheet1.cell_value(2,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.SAMOSA * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(2,0),self.sheet1.cell_value(2,1),int(self.sheet1.cell_value(2,3)),self.SAMOSA 
            self.TextReceipt.setText(self.Receipt)

            self.VEGBURGER = int(self.VEGBURGER)

        if(self.VEGBURGER >= 1):

            Price=self.sheet1.cell_value(3,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.VEGBURGER * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(3,0),self.sheet1.cell_value(3,1),int(self.sheet1.cell_value(3,3)),self.VEGBURGER 
            self.TextReceipt.setText(self.Receipt)

            self.EEGBURGER = int(self.EEGBURGER)

        if(self.EEGBURGER >= 1):

            Price=self.sheet1.cell_value(4,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.EEGBURGER * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(4,0),self.sheet1.cell_value(4,1),int(self.sheet1.cell_value(4,3)),self.EGGBURGER 
            self.TextReceipt.setText(self.Receipt)


            self.CHICKENBURGER = int(self.CHICKENBURGER)

        if(self.CHICKENBURGER >= 1):

            Price=self.sheet1.cell_value(5,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.CHICKENBURGER * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(5,0),self.sheet1.cell_value(5,1),int(self.sheet1.cell_value(5,3)),self.CHICKENBURGER 
            self.TextReceipt.setText(self.Receipt)

            self.SANDWICH = int(self.SANDWICH)

        if(self.SANDWICH >= 1):

            Price=self.sheet1.cell_value(6,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.SANDWICH * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(6,0),self.sheet1.cell_value(6,1),int(self.sheet1.cell_value(6,3)),self.SANDWICH
            self.TextReceipt.setText(self.Receipt)

            self.GREENTEA = int(self.GREENTEA)

        if(self.GREENTEA >= 1):

            Price=self.sheet1.cell_value(7,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.GREENTEA * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(7,0),self.sheet1.cell_value(7,1),int(self.sheet1.cell_value(7,3)),self.GREENTEA
            self.TextReceipt.setText(self.Receipt)


            self.MILKTEA = int(self.MILKTEA)

        if(self.MILKTEA >= 1):

            Price=self.sheet1.cell_value(8,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.MILKTEA * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(8,0),self.sheet1.cell_value(8,1),int(self.sheet1.cell_value(8,3)),self.MILKTEA
            self.TextReceipt.setText(self.Receipt)

            self.COFFEE = int(self.COFFEE)

        if(self.COFFEE >= 1):

            Price=self.sheet1.cell_value(9,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.COFFEE * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(9,0),self.sheet1.cell_value(9,1),int(self.sheet1.cell_value(9,3)),self.COFFEE
            self.TextReceipt.setText(self.Receipt)


            self.HOTMILK = int(self.HOTMILK)

        if(self.HOTMILK >= 1):

            Price=self.sheet1.cell_value(10,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.HOTMILK * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(10,0),self.sheet1.cell_value(10,1),int(self.sheet1.cell_value(10,3)),self.HOTMILK
            self.TextReceipt.setText(self.Receipt)

            self.BLACKCOFFEE = int(self.BLACKCOFFEE)

        if(self.BLACKCOFFEE >= 1):

            Price=self.sheet1.cell_value(11,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.BLACKCOFFEE * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(11,0),self.sheet1.cell_value(11,1),int(self.sheet1.cell_value(11,3)),self.BLACKCOFFEE
            self.TextReceipt.setText(self.Receipt)

            self.ORANGEJ = int(self.ORANGEJ)

        if(self.ORANGEJ >= 1):

            Price=self.sheet1.cell_value(13,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.ORANGEJ * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(13,0),self.sheet1.cell_value(13,1),int(self.sheet1.cell_value(13,3)),self.ORANGEJ
            self.TextReceipt.setText(self.Receipt)

            self.MANGIJ = int(self.MANGOJ)

        if(self.MANGOJ >= 1):

            Price=self.sheet1.cell_value(14,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.MANGIJ * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(14,0),self.sheet1.cell_value(14,1),int(self.sheet1.cell_value(14,3)),self.MANGOJ
            self.TextReceipt.setText(self.Receipt)

            self.APPLEJ = int(self.APPLEJ)

        if(self.APPLEJ >= 1):

            Price=self.sheet1.cell_value(15,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.GUAVAJ * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(15,0),self.sheet1.cell_value(15,1),int(self.sheet1.cell_value(15,3)),self.APPLEJ
            self.TextReceipt.setText(self.Receipt)


            self.GUAVAJ = int(self.GUAVAJ)

        if(self.GUAVAJ >= 1):

            Price=self.sheet1.cell_value(16,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.GUAVAJ * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(16,0),self.sheet1.cell_value(16,1),int(self.sheet1.cell_value(16,3)),self.GUAVAJ
            self.TextReceipt.setText(self.Receipt)

            self.PINEAPPLEJ = int(self.PINEAPPLEJ)

        if(self.PINEAPPLEJ >= 1):

            Price=self.sheet1.cell_value(17,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.PINEAPPLEJ * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(17,0),self.sheet1.cell_value(17,1),int(self.sheet1.cell_value(17,3)),self.PINEAPPLEJ
            self.TextReceipt.setText(self.Receipt)

            self.BANANA = int(self.BANANA)

        if(self.BANANA >= 1):

            Price=self.sheet1.cell_value(18,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.BANANA * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(18,0),self.sheet1.cell_value(18,1),int(self.sheet1.cell_value(18,3)),self.BANANA
            self.TextReceipt.setText(self.Receipt)


            self.MANGO_S = int(self.MANGO_S)

        if(self.MANGO_S >= 1):

            Price=self.sheet1.cell_value(19,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.MANGO_S * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(19,0),self.sheet1.cell_value(19,1),int(self.sheet1.cell_value(19,3)),self.MANGO_S
            self.TextReceipt.setText(self.Receipt)

            self.ALMOND = int(self.ALMOND)

        if(self.ALMOND >= 1):

            Price=self.sheet1.cell_value(20,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.ALMOND * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(20,0),self.sheet1.cell_value(20,1),int(self.sheet1.cell_value(20,3)),self.ALMOND
            self.TextReceipt.setText(self.Receipt)

            self.CHOCOLATE = int(self.CHOCOLATE)

        if(self.CHOCOLATE >= 1):

            Price=self.sheet1.cell_value(21,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.CHOCOLATE * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(21,0),self.sheet1.cell_value(21,1),int(self.sheet1.cell_value(21,3)),self.CHOCOLATE
            self.TextReceipt.setText(self.Receipt)

            self.OREO = int(self.OREO)

        if(self.OREO >= 1):

            Price=self.sheet1.cell_value(22,3) #Value of the coloumn in the excel file
            Price=int(Price)
            Price=self.OREO * Price
            self.TotalCost = int(self.TotalCost) + Price
            self.Receipt='%s\n\n%s(%s)\t%sx%s\t %s'%(self.Receipt),self.sheet1.cell_value(22,0),self.sheet1.cell_value(22,1),int(self.sheet1.cell_value(22,3)),self.OREO
            self.TextReceipt.setText(self.Receipt)
        self.TextCost.setText('%s'% self.TextCost)
        self.TextReceipt.setText('%s\n\n-----------------------------\n\n\t\t\tTotal Cost=%s\n\n\n\n\tThankyou for shopping' % (self.Receipt,self.TotalCost))


def START_PRINT(self):
     
    INCH = 1440

    pDC = win32ui.CreateDC ()
    pDC.CreatePrinterDC (win32print.GetDefaulterPrinter ())
    pDC.StartDoc ("Test doc")
    pDC.StartPage ()
    pDC.SetMapMode (win32con.MM_TWIPS)

    # TEXT = open(self.Receipt)
    
    pDC.DrawText (self.Receipt, (0, INCH * -1, INCH * 8, INCH * -2), win32con.DT_CENTER)
    pDC.EndPage ()
    pDC.EndDoc ()


app = QApplication(sys.argv)
window = theseencode()
window.show()
# app = QtWidgets.QApplication(sys.argv)

try:
    sys.exit(app.exec_())
except:
    print("Exiting")

# 35:09 TIME VIDEO LINK -- https://youtu.be/xEs0s9d4mXw?si=ZMWshrXZ0sIxd7qO
#https://youtu.be/xEs0s9d4mXw?si=LEMmEws3gmdcvKQM
