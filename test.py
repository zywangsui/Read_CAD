import random
import subprocess
import sys
import pandas as pd

import ezdxf
import win32api
import win32com.client
import win32con
import xlrd
from PyQt5.QtWidgets import QApplication, QMainWindow, QSpinBox
from PyQt5.QtWidgets import QApplication, QPushButton, QFileDialog
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QTreeWidgetItem, QDialog, QInputDialog, QMessageBox, QTableWidgetItem
from pyautocad import Autocad, APoint
from PyQt5.QtWidgets import QDialogButtonBox, QDateTimeEdit, QDialog, QComboBox, QTableView, QAbstractItemView, \
    QHeaderView, QTableWidget, QTableWidgetItem, QMessageBox, QListWidget, QListWidgetItem, QStatusBar, QMenuBar, QMenu, \
    QAction, QLineEdit, QStyle, QFormLayout, QVBoxLayout, QWidget, QApplication, QHBoxLayout, QPushButton, QMainWindow, \
    QGridLayout, QLabel
from PyQt5.QtGui import QIcon, QPixmap, QStandardItem, QStandardItemModel, QCursor, QFont, QBrush, QColor, QPainter, \
    QMouseEvent, QImage, QTransform
from PyQt5.QtCore import QStringListModel, QAbstractListModel, QModelIndex, QSize, Qt, QObject, pyqtSignal, QTimer, \
    QEvent, QDateTime, QDate
# 主窗口
from mainUI import Ui_MainWindow


class Area:
    def __init__(self):
        self.name = ""
        self.topLeft_x = 0.0
        self.topLeft_y = 0.0
        self.bottomRight_x = 0.0
        self.bottomRight_y = 0.0
        self.width = 0
        self.height = 0


class Led:
    def __init__(self):
        self.number = ""
        self.x = 0.0
        self.y = 0.0


table_RowCount = 1
AreaList = []
LedList = []
FloorName = ""
dxfName = ""


class Main(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(Main, self).__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        self.pushButton_2.setVisible(False)  # 取消excel表格的读取
        self.tableWidget.setVisible(False)
        self.pushButton_6.setEnabled(False)
        self.pushButton_7.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.move(1200, 200)
        if not os.path.exists("Data"):
            os.mkdir("Data")

        dialog2 = Dialog2(self)
        dialog2.dialogSignel2.connect(self.slot_NewData)
        dialog2.show()
        dialog2.move(800, 400)

    def slot_NewData(self, flag, str):
        global dxfName
        dxfName = str
        project_path = "Data/{}".format(str)
        if not os.path.exists(project_path):
            os.mkdir(project_path)
        self.label_2.setText(str)
        self.pushButton_6.setEnabled(True)

    def slot_ReadCad(self):
        global LedList
        global dxfName
        LedList.clear()

        try:
            filename = 'D:\{}11'.format(dxfName)
            acad_com = win32com.client.GetActiveObject('AutoCAD.Application')
            acad_com.ActiveDocument.SaveAs(filename, 25)
            doc = acad_com.ActiveDocument
            drawing = ezdxf.readfile(doc.FullName)
            msp = drawing.modelspace()
            # for cv in msp.query('INSERT'):
            #     print(cv.dxf.name)
            for led in msp.query('TEXT'):
                led_text = led.dxf.text
                isNumber = 1
                for i in range(len(led_text)):
                    # print(led_text[i])
                    if 47 < ord(led_text[i]) < 58:
                        pass
                    else:
                        isNumber = 0
                        break
                if isNumber == 1:
                    print("/*-------------*/")
                    print(led.dxf.text)
                    print(led.dxf.insert.x)
                    print(led.dxf.insert.y)
                    print("")

                    b = Led()
                    b.number = led_text
                    b.x = led.dxf.insert.x
                    b.y = led.dxf.insert.y
                    LedList.append(b)
            win32api.MessageBox(0, "完成,请执行第一步", "提醒", win32con.MB_OK | win32con.MB_TOPMOST)
            self.label_3.setText("图纸已确认")
            self.pushButton_7.setEnabled(True)


        except:
            win32api.MessageBox(0, "无法生成缓存图纸，请用AUDIT命令修复图纸", "提醒",
                                win32con.MB_OK | win32con.MB_TOPMOST)
            self.label_3.setText("图纸识别失败")
            self.pushButton_7.setEnabled(False)

    def slot_Step1(self):
        dialog = Dialog(self)
        # dialog.datetime.dateChanged.connect(self.slot_inner)
        dialog.dialogSignel.connect(self.slot_emit)
        dialog.show()
        dialog.move(800, 400)

    def slot_emit(self, flag, str):
        if flag == 0:
            self.label.setText(str)
            self.pushButton_8.setEnabled(True)
        else:
            self.pushButton_8.setEnabled(False)

        print(str)

    def slot_Step2(self):
        global table_RowCount
        global dxfName
        pos = []
        pos.clear()
        result_list = []
        result_list.clear()

        try:
            acad = Autocad(create_if_not_exists=True)
            autocad_ai = acad.doc.Utility

            point = autocad_ai.Getpoint(APoint(1, 1), "请选取平面图左上角")
            pos.append(list(point))
            point = autocad_ai.Getpoint(APoint(1, 1), "请选取平面图右下角")
            pos.append(list(point))

            a = Area()
            a.name = "test"
            a.topLeft_x = pos[0][0]
            a.topLeft_y = pos[0][1]
            a.bottomRight_x = pos[1][0]
            a.bottomRight_y = pos[1][1]
            a.width = a.bottomRight_x - a.topLeft_x
            a.height = a.topLeft_y - a.bottomRight_y
            AreaList.clear()
            AreaList.append(a)
            strlist = self.label.text().split(',')
            str_pro_name = self.label_2.text()
            # self.tableWidget.setRowCount(table_RowCount)
            #
            # self.tableWidget.setItem(table_RowCount - 1, 0, QTableWidgetItem(strlist[0]))
            # self.tableWidget.setItem(table_RowCount - 1, 1, QTableWidgetItem(str(a.topLeft_x * 18)))
            # self.tableWidget.setItem(table_RowCount - 1, 2, QTableWidgetItem(str(a.bottomRight_x * 12)))
            # self.tableWidget.setItem(table_RowCount - 1, 3, QTableWidgetItem(str(a.width * 51)))
            # self.tableWidget.setItem(table_RowCount - 1, 4, QTableWidgetItem("99." + ))

            for i in range(len(AreaList)):
                width = AreaList[i].width
                height = AreaList[i].height
                multiple_x = width / float(strlist[1])
                multiple_y = height / float(strlist[2])
                print(multiple_x, multiple_y, "放大倍数")

                topLeft_x = AreaList[i].topLeft_x
                topLeft_y = AreaList[i].topLeft_y
                bottomRight_x = AreaList[i].bottomRight_x
                bottomRight_y = AreaList[i].bottomRight_y

                for j in range(len(LedList)):
                    led_x = LedList[j].x
                    led_y = LedList[j].y
                    led_number = LedList[j].number
                    if bottomRight_x > led_x > topLeft_x and bottomRight_y < led_y < topLeft_y:
                        key_x = led_x - topLeft_x
                        key_y = topLeft_y - led_y
                        mkey_x = key_x / multiple_x
                        mkey_y = key_y / multiple_y

                        # QT绘制背景图时 长宽为原先的1/2
                        # mkey_x = mkey_x - float(strlist[1]) / 2
                        # mkey_y = mkey_y - float(strlist[2]) / 2

                        print("原始坐标", key_x, key_y, "转换后的坐标", mkey_x, mkey_y)
                        str1 = str_pro_name + "," + str(strlist[0]) + "," + str(led_number) + "," + str(mkey_x) + "," + str(mkey_y)
                        result_list.append(str1)
                pass
            path = 'Data/{}/{}.txt'.format(dxfName, strlist[0])
            mode = 'w'
            with open(path, mode) as f:
                for item in result_list:
                    f.write(item + '\n')

        except:
            print("错误")
        table_RowCount = table_RowCount + 1

    def slot_SaveAll(self):
        AllLed = []
        root = "Data\{}".format(dxfName)
        for dirpath, dirnames, filenames in os.walk(root):
            for filepath in filenames:
                o_path = os.path.join(dirpath, filepath)
                print(o_path)
                f = open(o_path, encoding="utf-8")
                # print(f.read())
                AllLed.append(str(f.read()))
        path = 'Data/{}/{}.txt'.format(dxfName, "All")
        mode = 'w'
        with open(path, mode) as f:
            for item in AllLed:
                f.write(item)
        win32api.MessageBox(0, "汇总完成,请导入至其它软件", "提醒", win32con.MB_OK | win32con.MB_TOPMOST)

    def slot_OpenCV(self):
        a = []
        # 读取excel文件
        excel_file = pd.ExcelFile('hh.xlsx')
        # 获取全部sheet名称
        sheet_names = excel_file.sheet_names
        # 读取全部sheet
        for sheet_name in sheet_names:
            sheet_data = pd.read_excel(excel_file, sheet_name)
            # 处理sheet_data的代码
            # print(sheet_data.values[:,[1,4]])
            x = sheet_data.values[:, [1, 4]]
            b = []
            for x1 in x:
                #print(x1[0],"+",x1[1])
                l = str(x1[1]).split('F')
                row = [x1[0],int(l[0])]
                #print(row)

                b.append(row)
            # a.append(sheet_data.values[:, [1, 4]])
            # print(a[0][0][0])
            b.sort(key=lambda x: x[1])
            a.append(b)
        for u in a:
            for s in u:
                print(s)


# 弹出框对象
class Dialog(QDialog):
    # 自定义消息
    dialogSignel = pyqtSignal(int, str)

    def __init__(self, parent=None):
        super(Dialog, self).__init__(parent)
        layout = QVBoxLayout(self)
        self.label = QLabel(self)
        self.label.setText("请输入楼层分辨率:")

        self.label_1 = QLabel(self)
        self.label_1.setText("长:")

        self.spinBox = QSpinBox(self)
        self.spinBox.setMaximum(20000)
        self.spinBox.setValue(3200)

        self.label_2 = QLabel(self)
        self.label_2.setText("宽:")

        self.spinBox_1 = QSpinBox(self)
        self.spinBox_1.setMaximum(20000)
        self.spinBox_1.setValue(2400)

        self.label_3 = QLabel(self)
        self.label_3.setText("楼层名称:")

        self.lineEdit = QLineEdit(self)

        layout.addWidget(self.label)
        layout.addWidget(self.label_1)
        layout.addWidget(self.spinBox)
        layout.addWidget(self.label_2)
        layout.addWidget(self.spinBox_1)
        layout.addWidget(self.label_3)
        layout.addWidget(self.lineEdit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)  # 点击ok
        buttons.rejected.connect(self.reject)  # 点击cancel
        layout.addWidget(buttons)

    def accept(self):  # 点击ok是发送内置信号
        print("accept")

        str = self.lineEdit.text() + "," + self.spinBox.text() + "," + self.spinBox_1.text()
        if str[0] == "," or str[1] == '0' or str[3] == '0':
            win32api.MessageBox(0, "数据不完整，请重新输入", "提醒", win32con.MB_OK | win32con.MB_TOPMOST)
        else:
            self.dialogSignel.emit(0, str)
            self.destroy()
            win32api.MessageBox(0, "完成,请执行第二步", "提醒", win32con.MB_OK | win32con.MB_TOPMOST)

    def reject(self):  # 点击cancel时，发送自定义信号
        print('reject')
        self.dialogSignel.emit(1, "无输入")
        self.destroy()


# 弹出框对象
class Dialog2(QDialog):
    # 自定义消息
    dialogSignel2 = pyqtSignal(int, str)

    def __init__(self, parent=None):
        super(Dialog2, self).__init__(parent)
        layout = QVBoxLayout(self)
        self.label = QLabel(self)
        self.label.setText("请输入项目名称：")

        self.lineEdit = QLineEdit(self)

        layout.addWidget(self.label)
        layout.addWidget(self.lineEdit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok, Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)  # 点击ok
        layout.addWidget(buttons)

    def accept(self):  # 点击ok是发送内置信号
        print("accept")

        str = self.lineEdit.text()
        if str == "":
            win32api.MessageBox(0, "数据不完整，请重新输入", "提醒", win32con.MB_OK | win32con.MB_TOPMOST)

        else:
            self.dialogSignel2.emit(0, str)
            self.destroy()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = Main()

    main.show()

    sys.exit(app.exec_())
