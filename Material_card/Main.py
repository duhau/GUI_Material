# -*- coding: utf-8 -*-
#
# Created by: DuHua 2018-7-5
#
# The main function of this GUI is generate MCNP material cards
"""
Module implementing vol_fra_input_Dialog.
"""
from PyQt5 import QtCore, QtGui, QtWidgets 
from PyQt5.QtGui import QFont,QColor,QBrush,QPalette
from PyQt5.QtCore import pyqtSlot,Qt
from PyQt5.QtWidgets import QMessageBox,QApplication,QTableWidgetItem,QTableWidget,QHeaderView
from Ui_database import Ui_Main_Foram_Dialog
from Ui_input import Ui_vol_fra_input_Dialog
import os
import os.path
import sqlite3
import openpyxl
import re

class Main_Foram_Dialog(QtWidgets.QWidget):
	"""
	Class documentation goes here.
	"""
	def __init__(self, parent=None):
		"""
		Constructor
		@param parent reference to the parent widget
		@type QWidget
		"""
		super(Main_Foram_Dialog, self).__init__(parent)
		self.UI=Ui_Main_Foram_Dialog()
		self.UI.setupUi(self)
		#设置背景
		self.main_p = QPalette()
		self.main_p.setColor(QPalette.Background,QColor('#87CEFA'))
		self.setAutoFillBackground(True)
		self.setPalette(self.main_p);
		# if items in mat_type_listWidget are selected, get_item is activated
		self.UI.mat_type_listWidget.itemClicked.connect(lambda: self.get_item())
		self.UI.selected_listWidget.itemClicked.connect(lambda: self.set_content_tableWidget())
		self.UI.pushButton_ok.clicked.connect(lambda: self.write_result())
	
	# pop sub dialog, and complete Interaction with main window
	def get_item(self):
		#str1=self.UI.mat_type_listWidget.currentItem().text()
		self.input_dialg=Ui_vol_fra_input_Dialog()
		self.dilag_main = QtWidgets.QWidget()
		self.input_dialg.setupUi(self.dilag_main)
		self.dilag_main.show()
		self.dilag_main.setFixedSize(400,180)
		self.input_dialg.ok_volume_pushButton.clicked.connect(lambda: self.get_LineEdit_content(self.input_dialg.volume_input_lineEdit.text()))
		self.input_dialg.cacel_volume_pushButton_2.clicked.connect(self.dilag_main.close)
	
	#get contents in input_dialg, and check if it is effective
	def get_LineEdit_content(self,item):
		global material_content	
		is_digit_mark=re.match(r'\d{1,}(\.\d{1,})?',item)
		match_item=is_digit_mark.group(0)
		if len(match_item) == len(item):
			material_content=float(item)  #get contents ininput_dialg
			#print('material is {:<6.4f}'.format(float(item)))
			#self.ok_volume_pushButton.clicked.connect(lambda: self.set_material_content(material_content))
			self.dilag_main.close()       #hide sub window
			self.UI.selected_listWidget.addItem(self.UI.mat_type_listWidget.currentItem().text()) #set selected_listWidget item
			
		else:
			msgBox = QMessageBox()
			msgBox.setWindowTitle('警告')
			msgBox.setIcon(QMessageBox.Warning)
			msgBox.setText("输入数据类型错误，请重新输入一个数")
			msgBox.setStandardButtons(QMessageBox.Ok)
			reply = msgBox.exec() 
			self.input_dialg.volume_input_lineEdit.setText('')  #if fail, set null 
			
	def set_content_tableWidget(self):
		table_name=self.UI.selected_listWidget.currentItem().text()
		ZAID=[]
		Nuclide_Name=[]
		Ion_Density=[]
		#连接数据库，创建游标
		conn = sqlite3.connect('data.db')
		cur = conn.cursor()
		sql='SELECT * FROM %s'%table_name
		cursor =conn.execute(sql)
		for row in cursor:
			ZAID.append(row[0][:-4])
			Nuclide_Name.append(row[1])
			density='{:.5e}'.format(row[2]*material_content/100)  #乘以体积份额计算考虑占空比后的核子密度
			Ion_Density.append(density)
		insert_data=[ZAID,Nuclide_Name,Ion_Density]                             #每个表的元素
		conn.close()
		#在content_tableWidget中写数据
		horizontalHeader = ['ZAID','Nuclide_Name','Ion_Density']
		self.UI.content_tableWidget.setColumnCount(len(horizontalHeader))       #设置列数
		self.UI.content_tableWidget.setRowCount(len(ZAID))                      #设置行数
		self.UI.content_tableWidget.setHorizontalHeaderLabels(horizontalHeader)	#添加表头
		for i in range(self.UI.content_tableWidget.rowCount()):
			for j in range(self.UI.content_tableWidget.columnCount()):
				self.UI.content_tableWidget.setItem(i,j, QTableWidgetItem(insert_data[j][i]))
		#设置格式		
		for index in range(self.UI.content_tableWidget.columnCount()):
			headItem = self.UI.content_tableWidget.horizontalHeaderItem(index)
			headItem.setFont(QFont("song", 12, QFont.Bold))
			headItem.setForeground(QColor(60,60,60))
			headItem.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
			self.UI.content_tableWidget.setEditTriggers(QTableWidget.NoEditTriggers) #内容无法编辑
			self.UI.content_tableWidget.resizeColumnsToContents()  
			self.UI.content_tableWidget.resizeRowsToContents()
			self.UI.content_tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch) #自适应窗口
	
	def write_result(self):
		library_mark=self.UI.lib_comboBox.currentText()          #获取lib_comboBox的内容
		b=self.UI.selected_listWidget.count()
		total_text=''
		for i in range(0, self.UI.selected_listWidget.count()):  #获取selected_listWidget中的所有内容
			table_name=self.UI.selected_listWidget.item(i).text()
			ZAID,Nuclide_Name,Ion_Density=self.read_database(table_name) #从数据库读取每个表中的数据	
			comment=['c ********************************************c','c ***      %s'%table_name,
			         'c ***  volume fraction [%] - {0}'.format(str(material_content)),'c ***  t.a.d. =  %s' % tad,
			         'c ***  eff.density = %s'%eff_density,'c ********************************************c\n']
			text='\n'.join(comment)
			#获取每种成分的数据
			for j in range(len(ZAID)):
				item=['        ', ZAID[j]+library_mark, Ion_Density[j],'$', Nuclide_Name[j]]
				line=' '.join(item)+'\n' 
				text=text+line
			#获取总的数据
			total_text=total_text+text
		with open("output.txt", 'w') as fou:
			fou.write(total_text)
			fou.close()		
	
	#完成从数据库读取每个表中的数据	
	def read_database(self,table_name):
		global tad
		global eff_density		
		ZAID=[]
		Nuclide_Name=[]
		Ion_Density=[]
		tad=0
		eff_density=0		
		#连接数据库，创建游标
		conn = sqlite3.connect('data.db')
		cur = conn.cursor()
		sql='SELECT * FROM %s'%table_name
		cursor =conn.execute(sql)
		for row in cursor:
			ZAID.append(row[0][:-4])
			Nuclide_Name.append(row[1])
			density='{:.5e}'.format(row[2]*material_content/100)  #乘以体积份额计算考虑占空比后的核子密度
			Ion_Density.append(density)
			tad=row[2]+tad                                        #计算总核子密度
			eff_density=row[2]*material_content/100+eff_density   #计算材料总的有效核子密度
		data=[ZAID,Nuclide_Name,Ion_Density]                          #每个表中的元素
		tad='{:.5e}'.format(tad)
		eff_density='{:.5e}'.format(eff_density)		
		conn.close()
		return data
		
if __name__=="__main__":
	import sys
	app=QApplication(sys.argv)
	main_Fram=Main_Foram_Dialog()
	main_Fram.show()
	sys.exit(app.exec_())