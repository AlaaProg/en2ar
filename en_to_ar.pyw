# A-Z {65 to 90}
# a-z {97 to  122}
# ar {1568 , 1610}

from xlrd import open_workbook
from PySide.QtGui import (QApplication,QAction,QMainWindow,QWidget,QLabel,
		QLineEdit,QPushButton,QVBoxLayout,QFileDialog,
		QTableWidget,QTableWidgetItem,QFont,QDialog)


style= open("style.css","r").read()

class Table(QTableWidget):
	def __init__(self):
		super(Table,self).__init__()


class Window(QMainWindow):
	def __init__(self):
		super(Window,self).__init__()
		open("en_to_ar.txt",'w').write("")
		self.table  = Table()
		self.dialog = QDialog(self) 
		self.action()
		self.widget()

	def action(self):
		bar  = self.menuBar()
		main = bar.addMenu("#")
		

 
	def widget(self):
		win = QWidget()
		self.setCentralWidget(win)
		self.setStyleSheet(style)
		self.setFixedSize(450,250)
		win.setWindowTitle("en To ar")
		vbox = QVBoxLayout(win)
		self.lab  = QLabel()
		self.line  = QLineEdit()
		self.line.setPlaceholderText("الكلمة")
		but  = QPushButton("ترجم")
		vbox.addWidget(self.line)
		vbox.addWidget(self.lab)
		vbox.addWidget(but)
		win.setTabOrder(but,self.line)
		self.lab.setText("<p align='center'>.&nbsp;&nbsp;.&nbsp;&nbsp;.</p>")
		but.clicked.connect(self.search)
		self.line.returnPressed.connect(self.search)
	def search(self):
		lab = self.lab;
		sew = self.line.text()
		index_1 = None;index_2 = None
		en = list(range(65,90))
		ar = list(range(1568,1610))
		if sew != "" :
			if ord(sew[0].upper().strip()) in en:
				index_1 = 0 ;index_2 = 1
			elif ord(sew[0].strip()) in ar:
				index_1 = 1 ;index_2 = 0


		if index_1 != None and index_2 != None :
			wb = open_workbook('c',logfile={"encoding":"u8"})
			se = sew.strip().lower()
			s = wb.sheets()[0]
			for i in range(s.nrows):
				word = str(s.cell(i,index_1).value).strip().lower()
				
				if word == se:
					lab.setText("<p align='center'>%s &nbsp;&nbsp;:&nbsp;&nbsp; %s</p>"%(str(s.cell(i,index_1).value).strip(),s.cell(i,index_2).value))
					break;
				else:
					lab.setText("<p align='center'>NotFund</p>")
		else:
			lab.setText("<p align='center'>.&nbsp;&nbsp;.&nbsp;&nbsp;.</p>")

def run():
	app = QApplication([])
	win = Window()
	win.show()
	app.exec_()

if __name__ == '__main__':
	run()
