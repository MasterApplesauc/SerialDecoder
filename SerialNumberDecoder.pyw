import sys, os, xlrd, datetime
from PyQt5 import QtWidgets, uic, QtGui
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QMenuBar
from PyQt5.QtGui import QIcon

class MainWindow(QtWidgets.QMainWindow):
	def __init__(self):
		self.createUI()
		self.createVars()

	def createVars(self):
		self.productCategoryDict = {}
		self.productCountryDict = {}
		self.productCodeDict = {}
		self.excelWorkbook = None

		try:
			prefFile = open(self.resourcePath('_.pref'), 'r')
			self.excelWorkbook = xlrd.open_workbook(str(prefFile.readline()))
			prefFile.close()
		except FileNotFoundError:
			try:
				self.excelWorkbook = xlrd.open_workbook(self.resourcePath('Serials.index'))
			except:
				try:
					self.excelWorkbook = xlrd.open_workbook(self.resourcePath('Serials.index.backup'))
				except:
					try:
						self.openFileDialog()
					except:
						pass

		try:
			self.sheetVerification = self.excelWorkbook.sheet_by_name('Verification')
			if int(self.sheetVerification.row_values(0)[0]) != 123:
				raise ValueError
		except ValueError:
			msg = QMessageBox()
			msg.setIcon(QMessageBox.Critical)
			msg.setText("Index file failed validation!")
			msg.setInformativeText('Index file currently loaded is invalid and appears to have been tampered with - please download a new one from the portal.')
			msg.setWindowTitle("INVALID FILE!")
			msg.exec_()
		except:
			msg = QMessageBox()
			msg.setIcon(QMessageBox.Critical)
			msg.setText("Invalid file/format")
			msg.setInformativeText('Program will not function without an index file. Please download one from the portal and try again.')
			msg.setWindowTitle("Error Loading Index File.")
			msg.exec_()

		try:
			self.sheetCountry = self.excelWorkbook.sheet_by_name('Product Country')
			self.sheetCategory = self.excelWorkbook.sheet_by_name('Product Category')
			self.sheetCode = self.excelWorkbook.sheet_by_name('Product Code')
			self.sheetVerification = self.excelWorkbook.sheet_by_name('Verification')

			for i in range(self.sheetCountry.nrows):
				self.productCountryDict[self.sheetCountry.row_values(i)[0]] = self.sheetCountry.row_values(i)[1]

			for i in range(self.sheetCategory.nrows):
				self.productCategoryDict[self.sheetCategory.row_values(i)[0]] = self.sheetCategory.row_values(i)[1]

			for i in range(self.sheetCode.nrows):
				self.productCodeDict[self.sheetCode.row_values(i)[0]] = self.sheetCode.row_values(i)[1]

			self.versionNumber = self.sheetVerification.row_values(1)[0]
			self.versionDate = (datetime.datetime(*xlrd.xldate_as_tuple(self.sheetVerification.row_values(2)[0], self.excelWorkbook.datemode)).date()).strftime('%d%b%Y')

		except:
			msg = QMessageBox()
			msg.setIcon(QMessageBox.Critical)
			msg.setText("Couldn't load index file required to decode serial.")
			msg.setInformativeText('Program will not function without an index file. Please download one from the portal and try again.')
			msg.setWindowTitle("Error Loading Index File.")
			msg.exec_()

	def openFileDialog(self):
		fileLocation = QFileDialog.getOpenFileName(self, 'Open Serial Index File')[0]
		
		if not fileLocation == "":
			prefFile = open('_.pref', 'w')
			prefFile.write(str(fileLocation))
			prefFile.close()

			self.excelWorkbook = xlrd.open_workbook(fileLocation)

			self.createVars()
			msg = QMessageBox()
			msg.setIcon(QMessageBox.Information)
			msg.setText("Loaded Index File.")
			msg.setInformativeText('Successfully loaded index file! Your results should update accordingly.')
			msg.setWindowTitle("File Loaded")
			msg.exec_()
			self.updateUI()

	def aboutBox(self):
			msg = QMessageBox()
			msg.setWindowIcon(QIcon(self.resourcePath('30x30ShimadzuLogo.png')))
			msg.setWindowTitle(f"SN Decoder - Version {self.versionNumber}")
			msg.setIcon(QMessageBox.Information)
			msg.setText("Shimadzu Serial Number Decoder")
			msg.setInformativeText(
f'''
Version {self.versionNumber} - {self.versionDate}
© 2020 Shimadzu Scientific Instruments, LLC. 
All Rights Reserved.
Created for Shimadzu (and ease of use) by
Joe Vincent, NCA Service Technician
''')

			msg.exec_()

	def createUI(self):
		super().__init__()
		uic.loadUi(self.resourcePath("mainwindow.ui"), self)
		self.setFixedSize(296, 307)
		self.setWindowIcon(QIcon(self.resourcePath('193x193ShimadzuLogo.png')))
		self.show()

		#Find and connect the Menu buttons
		self.menu = (self.findChild(QMenuBar))

		self.actionExit.triggered.connect(lambda: sys.exit())
		self.actionImport.triggered.connect(self.openFileDialog)
		self.actionAbout.triggered.connect(self.aboutBox)

		#Find all the visual UI Elements
		self.LE_SerialNumber = (self.findChild(QtWidgets.QLineEdit, 'lineEdit_12'))

		self.LE_SerialCategory = self.findChild(QtWidgets.QLineEdit, 'lineEdit_7')
		self.LE_ProductCategory = self.findChild(QtWidgets.QLineEdit, 'lineEdit')

		self.LE_SerialCode = self.findChild(QtWidgets.QLineEdit, 'lineEdit_8')
		self.LE_ProductCode = self.findChild(QtWidgets.QLineEdit, 'lineEdit_2')

		self.LE_SerialYear = self.findChild(QtWidgets.QLineEdit, 'lineEdit_9')
		self.LE_ProductYear1 = self.findChild(QtWidgets.QLineEdit, 'lineEdit_3')
		self.LE_ProductYear2 = self.findChild(QtWidgets.QLineEdit, 'lineEdit_4')

		self.LE_SerialSerial = self.findChild(QtWidgets.QLineEdit, 'lineEdit_10')
		self.LE_ProductSerial = self.findChild(QtWidgets.QLineEdit, 'lineEdit_6')

		self.LE_SerialCountry = self.findChild(QtWidgets.QLineEdit, 'lineEdit_11')
		self.LE_ProductCountry = self.findChild(QtWidgets.QLineEdit, 'lineEdit_5')

		self.LE_SerialNumber.textChanged.connect(self.updateUI)


	def updateUI(self):
		self.validateEntry()
		self.serialNumber = self.LE_SerialNumber.text()
		self.decodeSerialNumber()
		self.analyzeSerialNumber()

	def validateEntry(self):
		#CapitalizeEntry:
		unwantedChars = ['-',';',':','\'','"''\t']
		entry = self.LE_SerialNumber.text()
		entry = entry.upper()
		entry2 = []

		#Remove Unwanted Chars:
		for i in range(len(self.LE_SerialNumber.text())):
			if entry[i] in unwantedChars:
				entry = entry[0:i]

		self.LE_SerialNumber.setText(entry)

	def decodeSerialNumber(self):
		self.serialNumber = self.serialNumber.replace(' ','')
		#Pattern: [(Ab) cde] fg higjkl mn
		self.productCategory = self.serialNumber[0:2] #Ab
		self.LE_SerialCategory.setText(self.productCategory)

		self.productCode = self.serialNumber[0:5] #[(ab)cde]
		self.LE_SerialCode.setText(self.productCode)

		self.productYear = self.serialNumber[5:7] #fg
		self.LE_SerialYear.setText(self.productYear)

		self.productSerial = self.serialNumber[7:12] #hijkl
		self.LE_SerialSerial.setText(self.productSerial)

		if(len(self.serialNumber)) == 14: # mn
			self.productCountry = self.serialNumber[12:14]
			self.LE_SerialCountry.setText(self.productCountry)
		else:
			self.productCountry = "ZZ"
			self.LE_SerialCountry.setText(self.productCountry)

	def analyzeSerialNumber(self):
		if len(self.serialNumber) >= 2:
			self.findProductCategory()
		if len(self.serialNumber) >= 5:
			self.findProductCode()
		if len(self.serialNumber) >= 7:
			self.findProductYear()
		if len(self.serialNumber) >= 12:
			self.findProductSerial()
		if len(self.serialNumber) > 12:
			self.findProductCountry()

	def findProductCategory(self):
		if self.productCategory in self.productCategoryDict:
			self.LE_ProductCategory.setText(self.productCategoryDict[self.productCategory])
		else:
			self.LE_ProductCategory.setText('Unknown Instrument')

	def findProductCode(self):
		if self.productCode in self.productCodeDict:
			self.LE_ProductCode.setText(self.productCodeDict[self.productCode]) #[(Ab) cde]
		else:
			self.LE_ProductCode.setText('Unknown Model')

	def findProductYear(self):
		try:
			self.Y1 = int(self.productYear) + 1962
			self.Y2 = int(self.productYear) + 1963

			self.LE_ProductYear1.setText(f'Apr {self.Y1}')
			self.LE_ProductYear2.setText(f'Mar {self.Y2}')
		except:
			self.LE_ProductYear1.setText('Invalid')
			self.LE_ProductYear2.setText('Invalid')

	def findProductSerial(self):
		self.LE_ProductSerial.setText(f'Instrument #{self.productSerial}')


	def findProductCountry(self):
		if(self.productCountry in self.productCountryDict):
			self.LE_ProductCountry.setText(self.productCountryDict[self.productCountry])
		else:
			self.LE_ProductCountry.setText('Unkown Country')

	def resourcePath(self, relativePath):
		""" Get absolute path to resource, works for dev and for PyInstaller """
		try:
		# PyInstaller creates a temp folder and stores path in _MEIPASS
			basePath = sys._MEIPASS
		except Exception:
			basePath = os.path.abspath(".")

		return os.path.join(basePath, relativePath)


app = QtWidgets.QApplication(sys.argv)
window = MainWindow()
app.exec()