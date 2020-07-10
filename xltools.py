from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QWidget,QApplication,QFileDialog,QMessageBox,QDialog,QTextEdit,QVBoxLayout,QHBoxLayout,QLabel,QSpinBox,QPushButton
from PyQt5.QtGui import QIcon
import resource
import sys
import os
import pandas as pd

'''-------------------------------- 核心方法  --------------------------------'''
def merge_excel(dir):
    print('执行合并 ------')
    filename_excel = [] # 存表名
    frames = [] # 存表内容
    d = dir.replace('/','\\\\') # 因pandsa读取路径为双斜杠，需转换
    if d.endswith('\\\\') == False: # 若为磁盘根目录则路径结尾自带\\，若为文件夹则无，需添加\\
        d = d + '\\\\'
    print("路径是：",d,"\n有以下文件：")
    for files in os.listdir(path=dir):
        print(files)
        if 'xlsx' in files or 'xls' in files : # 搜索xlsx/xls后缀文件
            filename_excel.append(files)
            df = pd.read_excel(d+files)
            frames.append(df)
    if len(frames)!= 0: # 若存在EXCEL表则合并保存
        result = pd.concat(frames)
        result.to_excel(d+"合并结果表.xlsx")
    return filename_excel

def split_excel(path,num):
    # print("--- 执行拆分 ---")
    p = path.replace('/', '\\\\') # 传入pd库方法的路径
    dir = p[ : p.rfind('\\') + 1 ] # 输出被拆分表的目录
    sheetname = path[ path.rfind('/') + 1 :].strip('.xlsx').strip('.xlx') # 无后缀的文件名
    data = pd.read_excel(p) # 数据
    nrows = data.shape[0]  # 获取行数
    print("行数：",nrows)
    split_rows = num # 拆分的条数
    print("拆分的条数：", split_rows)
    count = int(nrows/split_rows) + 1  # 拆分的份数
    print("应当拆分成%d份"%count)
    begin = 0
    end = 0
    for i in range(1,count+1):
        sheetname_temp = sheetname+str(i)+'.xlsx'
        if i == 1:
            end = split_rows
        elif i == count:
            begin = end
            end = nrows
        else:
            begin = end
            end = begin + split_rows
        print(sheetname_temp)
        data_temp = data.iloc[ begin:end , : ] # [ 行范围 , 列范围 ]
        data_temp.to_excel(dir + sheetname_temp)
    # print('拆分完成')
    return count

'''-------------------------------- 界面窗口类  --------------------------------'''
class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(300, 190)
        Form.setMinimumSize(QtCore.QSize(300, 190))
        Form.setMaximumSize(QtCore.QSize(300, 190))
        Form.setAutoFillBackground(False)
        self.pB_Merge = QtWidgets.QPushButton(Form)
        self.pB_Merge.setGeometry(QtCore.QRect(31, 30, 110, 110))
        self.pB_Merge.setText("")
        self.pB_Merge.setObjectName("pB_Merge")
        self.pB_Split = QtWidgets.QPushButton(Form)
        self.pB_Split.setGeometry(QtCore.QRect(160, 30, 110, 110))
        self.pB_Split.setText("")
        self.pB_Split.setObjectName("pB_Split")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(58, 150, 54, 12))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(185, 150, 54, 12))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(5, 5, 80, 12))
        self.label_3.setObjectName("label_3")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "合并EXCEL"))
        self.label_2.setText(_translate("Form", "拆分EXCEL"))
        self.label_3.setText(_translate("Form", "@pan&yang"))

'''-------------------------------- 界面功能类  --------------------------------'''
class XltoolForm(QWidget,Ui_Form):
    def __init__(self):
        super(XltoolForm, self).__init__()
        self.setupUi(self)
        self.pB_Merge.setStyleSheet("QPushButton{border-image: url(:resource/merge.ico);};")
        self.pB_Split.setStyleSheet("QPushButton{border-image: url(:resource/split.ico);};")
        self.pB_Merge.clicked.connect(self.merge_excel)  # 合并按钮
        self.pB_Split.clicked.connect(self.open_dialog)  # 拆分按钮
        self.path = '' # 拆分文件路径

    def load_path(self):
        dir = QFileDialog.getExistingDirectory(self,"选择EXCEL文件所在文件夹","./")
        return dir

    def merge_excel(self):
        dir = self.load_path()
        if dir != '':
            try:
                filename_excel = merge_excel(dir)
                # 将合并结果输出到弹窗
                dialog = QDialog()
                dialog.resize(150,180)
                dialog.setWindowTitle("合并结果")
                txt = QTextEdit('已合并%s个表' % len(filename_excel),dialog) # 添加多行文本到弹窗
                txt.append('----------')
                txt.resize(150,180)
                txt.append('\n'.join(t for t in filename_excel)) # 逐行添加被合并的表名
                dialog.setWindowModality(Qt.ApplicationModal) # 模态：等待弹窗结束才可操作主窗口
                dialog.exec_()
            except:
                QMessageBox.about(self,'异常','请检查...')
        else:
            QMessageBox.about(self,"提示","未选择文件目录！")

    def load_file(self):
        filename,filetype = QFileDialog.getOpenFileName(self,'选择需要拆分的EXCEL文件','./','(*.xlx);(*.xlsx)')
        return filename

    def open_dialog(self):
        self.path = self.load_file()
        d = SplitDialog()
        d.IntSignal.connect(self.split_excel)
        d.show()
        d.exec()

    def split_excel(self,num):
        path = self.path
        # print('拆分文件路径:', path)
        # print('拆分条数:', num)
        if path != '':
            try:
                c = split_excel(path,num)
                QMessageBox.about(self, '拆分成功', '已拆分成%s份，请打开文件夹查看！'%c)
            except:
                QMessageBox.about(self, '异常', '请检查...')
        else:
            QMessageBox.about(self, "提示", "未选择文件！")


class SplitDialog(QDialog):  # 参数设置弹窗
    IntSignal = pyqtSignal(int)
    def __init__(self):
        super(SplitDialog, self).__init__()
        self.resize(180, 50)
        self.setWindowTitle("拆分设置")
        self.vbox = QVBoxLayout()  # 垂直布局
        self.hbox = QHBoxLayout()  # 水平布局

        self.lb = QLabel('请输入拆分条数，即拆分后多少行每个表')
        self.spin = QSpinBox()  # 计数器
        self.spin.resize(80, 20)
        self.spin.setRange(0, 2000)  # 设置计数器范围
        self.spin.setSingleStep(10)  # 设置计数器步长
        self.bt = QPushButton('OK')  # 按钮

        self.vbox.addWidget(self.lb)  # 将标签到vbox
        self.hbox.addWidget(self.spin)  # 将计数器到hbox
        self.hbox.addWidget(self.bt)  # 将按钮到hbox
        self.vbox.addLayout(self.hbox)  # 将hbox中的内容添加到将标签到vbox
        self.setLayout(self.vbox)  # 将布局内容添加到弹窗

        self.bt.clicked.connect(self.Num_emit)

    def Num_emit(self):
        num = self.spin.value()
        self.IntSignal.emit(num)
        self.close()




if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = XltoolForm()
    widget.setWindowIcon(QIcon(":resource/icon.ico"))
    title = "XLTools"
    datenum = "_20.5.26"
    strTitle = title+datenum
    widget.setWindowTitle(strTitle)
    widget.show()
    sys.exit(app.exec())