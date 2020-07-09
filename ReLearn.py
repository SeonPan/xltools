import pandas as pd
import os
import sys
from QtFile import mainWidget as mw
from QtFile import readWidget as rw
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QWidget, QApplication, QHeaderView, QTableWidget, QTableWidgetItem,  QMessageBox
from PyQt5.QtCore import Qt


class ReadForm(QWidget,rw.Ui_Form):
    def __init__(self):
        super(ReadForm, self).__init__()
        self.setupUi(self)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch) # 自适应列宽
        self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.Stretch) # 自适应行高
        self.tableWidget.verticalHeader().setVisible(False)  # 隐藏行表头
        # self.tableWidget.horizontalHeader().setVisible(False)  # 隐藏列表头
        self.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.pB_close.clicked.connect(self.closethis)
    def closethis(self):
        self.close()


class ReLearnForm(QWidget,mw.Ui_Form):
    def __init__(self):
        super(ReLearnForm, self).__init__()
        self.setupUi(self)
        self.cols = self.tableWidget.columnCount()
        self.rows = 0 # 任务数
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch) # 自适应列宽
        self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive) # 仅首列可手动调整
        self.row_flag = -1 # 当前被选中的行索引
        self.result = []
        self.tableWidget.itemSelectionChanged.connect(self.chioce) # 单元格选择改变
        self.pB_Add.clicked.connect(self.add_row) # 增
        self.pB_Del.clicked.connect(self.del_row) # 删
        self.pB_AllClean.clicked.connect(self.clean_row) # 清空
        self.pB_return.clicked.connect(self.return_result)  # 撤销
        self.pB_setPara.clicked.connect(self.prior_value) # 计算
        self.pB_save.clicked.connect(self.save) # 保存
        self.pB_read.clicked.connect(self.read) # 关于

    def read(self):
        self.d = ReadForm()
        self.d.show()
        self.d.setWindowTitle("参数参考")

    def add_row(self):
        print('添加一行任务记录')
        self.rows += 1
        self.tableWidget.setRowCount(self.rows)

    def del_row(self):
        print('删除一行任务记录')
        if self.row_flag == -1:
            QMessageBox.about(self,'提醒','未选择要删除的行！')
        else:
            self.tableWidget.removeRow(self.row_flag) # 删除指定行
            self.rows -= 1
            self.row_flag = -1

    def chioce(self): # 修改被选中的行索引
        self.row_flag = self.tableWidget.currentRow()
        print(f'选中第{self.row_flag + 1}行')

    def clean_row(self):
        print('清空任务记录')
        if self.rows != 0:
            reply = QMessageBox.question(self, '提示', '请确认是否要清空!', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.rows = 0
                self.tableWidget.setRowCount(self.rows)

    def return_result(self):
        print('撤销操作')
        print(self.result)
        if self.result:
            self.rows = len(self.result) # 恢复任务数
            self.tableWidget.setRowCount(self.rows) # 恢复行
            self.sort_move(self.result) # 恢复内容

    def prior_value(self):
        print('计算优先度')
        # print(f"共{self.row_nums}行{self.cols}列数据")
        list_para = []
        try:
            for r in range(self.rows):
                my_item = self.tableWidget.item(r, 0).text() # 任务名称
                urgent = int(self.tableWidget.item(r, 1).text()) # 紧急程度
                important = int(self.tableWidget.item(r, 2).text()) # 重要程度
                t = int(self.tableWidget.item(r, 3).text()) # 时间系数
                p = round(urgent*0.6 + important*0.4 + t*0.1,2) # 优先度
                item_p = QTableWidgetItem(str(p)) # 将优先度值添加为单元格
                item_p.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled) # 设置为可选择、不可编辑
                self.tableWidget.setItem(r, 4, item_p)
                list_para.append([my_item, urgent, important, t, p])
            self.result = sorted(list_para, key=lambda x: x[4], reverse=True) # 根据优先度将行元素列表降序
            self.sort_move(self.result)
        except:
            pass

    def sort_move(self, list_para_sorted): # 根据排序结果重写单元格
        print('执行了sort')
        for now in range(len(list_para_sorted)):
            item_it = QTableWidgetItem(str(list_para_sorted[now][0]))
            item_u = QTableWidgetItem(str(list_para_sorted[now][1]))
            item_im = QTableWidgetItem(str(list_para_sorted[now][2]))
            item_t = QTableWidgetItem(str(list_para_sorted[now][3]))
            item_p = QTableWidgetItem(str(list_para_sorted[now][4]))
            item_p.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)  # 设置为可选择、不可编辑
            print('得到了items')
            self.tableWidget.setItem(now, 0, item_it)
            self.tableWidget.setItem(now, 1, item_u)
            self.tableWidget.setItem(now, 2, item_im)
            self.tableWidget.setItem(now, 3, item_t)
            self.tableWidget.setItem(now, 4, item_p)
            print('设置了items')
        for i in range(int(self.rows)):  # 设置所有单元格文本居中
            for j in range(int(self.cols)):
                self.tableWidget.item(i, j).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)

    def save(self):
        print('保存结果')
        if self.result:
            # [my_item, urgent, important, t, p]
            df = pd.DataFrame(columns=['任务项', '紧急程度', '重要程度', '时间系数', '优先度'])
            for i in range(len(self.result)):
                df.loc[i] = self.result[i]
            df.to_excel('优先度结果表.xlsx', index = False)
            os.startfile('优先度结果表.xlsx')
        else:
            QMessageBox.about(self,'提醒','未产生计算结果！')



if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = ReLearnForm()
    title = "ReLearn"
    datenum = "_20.6.12"
    strTitle = title+datenum
    widget.setWindowTitle(strTitle)
    widget.show()
    sys.exit(app.exec())