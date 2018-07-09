# -*- coding: utf-8 -*-
#
# Created by: DuHua 2018-7-5
#
# The main function of this GUI is generate MCNP material cards


from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_vol_fra_input_Dialog(object):
        def setupUi(self, vol_fra_input_Dialog):
                vol_fra_input_Dialog.setObjectName("vol_fra_input_Dialog")
                vol_fra_input_Dialog.resize(376, 172)
                #vol_fra_input_Dialog.setSizeGripEnabled(True)
                self.layoutWidget = QtWidgets.QWidget(vol_fra_input_Dialog)
                self.layoutWidget.setGeometry(QtCore.QRect(50, 40, 295, 101))
                self.layoutWidget.setObjectName("layoutWidget")
                self.verticalLayout = QtWidgets.QVBoxLayout(self.layoutWidget)
                self.verticalLayout.setContentsMargins(0, 0, 0, 0)
                self.verticalLayout.setObjectName("verticalLayout")
                self.horizontalLayout = QtWidgets.QHBoxLayout()
                self.horizontalLayout.setObjectName("horizontalLayout")
                self.show_volume_label = QtWidgets.QLabel(self.layoutWidget)
                font = QtGui.QFont()
                font.setFamily("微软雅黑")
                font.setPointSize(11)
                font.setUnderline(False)
                font.setStrikeOut(False)
                self.show_volume_label.setFont(font)
                self.show_volume_label.setObjectName("show_volume_label")
                self.horizontalLayout.addWidget(self.show_volume_label)
                spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
                self.horizontalLayout.addItem(spacerItem)
                spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
                self.horizontalLayout.addItem(spacerItem1)
                self.volume_input_lineEdit = QtWidgets.QLineEdit(self.layoutWidget)
                self.volume_input_lineEdit.setObjectName("volume_input_lineEdit")
                self.horizontalLayout.addWidget(self.volume_input_lineEdit)
                self.label = QtWidgets.QLabel(self.layoutWidget)
                font = QtGui.QFont()
                font.setFamily("Consolas")
                font.setPointSize(16)
                self.label.setFont(font)
                self.label.setObjectName("label")
                self.horizontalLayout.addWidget(self.label)
                self.verticalLayout.addLayout(self.horizontalLayout)
                spacerItem2 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
                self.verticalLayout.addItem(spacerItem2)
                self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
                self.horizontalLayout_2.setObjectName("horizontalLayout_2")
                self.ok_volume_pushButton = QtWidgets.QPushButton(self.layoutWidget)
                font = QtGui.QFont()
                font.setFamily("微软雅黑")
                font.setPointSize(11)
                self.ok_volume_pushButton.setFont(font)
                self.ok_volume_pushButton.setObjectName("ok_volume_pushButton")
                self.horizontalLayout_2.addWidget(self.ok_volume_pushButton)
                spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
                self.horizontalLayout_2.addItem(spacerItem3)
                spacerItem4 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
                self.horizontalLayout_2.addItem(spacerItem4)
                self.cacel_volume_pushButton_2 = QtWidgets.QPushButton(self.layoutWidget)
                font = QtGui.QFont()
                font.setFamily("微软雅黑")
                font.setPointSize(11)
                self.cacel_volume_pushButton_2.setFont(font)
                self.cacel_volume_pushButton_2.setObjectName("cacel_volume_pushButton_2")
                self.horizontalLayout_2.addWidget(self.cacel_volume_pushButton_2)
                self.verticalLayout.addLayout(self.horizontalLayout_2)

                self.retranslateUi(vol_fra_input_Dialog)
                #self.ok_volume_pushButton.clicked.connect(lambda: self.get_LineEdit_content(self.volume_input_lineEdit.text()))
                QtCore.QMetaObject.connectSlotsByName(vol_fra_input_Dialog)

        def retranslateUi(self, vol_fra_input_Dialog):
                _translate = QtCore.QCoreApplication.translate
                vol_fra_input_Dialog.setWindowTitle(_translate("vol_fra_input_Dialog", "体积分数对话框"))
                self.show_volume_label.setText(_translate("vol_fra_input_Dialog", "体积份额"))
                self.label.setText(_translate("vol_fra_input_Dialog", "%"))
                self.ok_volume_pushButton.setText(_translate("vol_fra_input_Dialog", "确定"))
                self.cacel_volume_pushButton_2.setText(_translate("vol_fra_input_Dialog", "取消"))


if __name__ == "__main__":
        import sys
        app = QtWidgets.QApplication(sys.argv)
        vol_fra_input_Dialog = QtWidgets.QWidget()
        ui = Ui_vol_fra_input_Dialog()
        ui.setupUi(vol_fra_input_Dialog)
        vol_fra_input_Dialog.show()
        sys.exit(app.exec_())

