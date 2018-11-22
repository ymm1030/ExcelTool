import CustomWidgets
from PyQt5.QtWidgets import *
import BaseMethods

class MatchAndInsertPanel(QWidget):
    def __init__(self, **kw):
        super(MatchAndInsertPanel, self).__init__(**kw)
        self.f1 = CustomWidgets.ExcelSelecter('文件1', parent = self)
        self.f2 = CustomWidgets.ExcelSelecter('文件2', parent = self)
        self.f1.setGeometry(50, 20, 500, 75)
        self.f2.setGeometry(50, 120, 500, 75)
        self.s1 = CustomWidgets.RowColumnSelector('文件1条码', parent = self)
        self.s2 = CustomWidgets.RowColumnSelector('文件1数量', parent = self)
        self.s3 = CustomWidgets.RowColumnSelector('文件2条码', parent = self)
        self.s4 = CustomWidgets.RowColumnSelector('文件2数量', parent = self)
        self.s1.setGeometry(50, 220, 250, 75)
        self.s2.setGeometry(350, 220, 250, 75)
        self.s3.setGeometry(50, 320, 250, 75)
        self.s4.setGeometry(350, 320, 250, 75)
        btn = QPushButton(self)
        btn.setText('执行')
        btn.setGeometry(650, 250, 50, 120)
        btn.clicked.connect(self.exec)
        btn2 = QPushButton(self)
        btn2.setText('对调')
        btn2.setGeometry(650, 50, 50, 120)
        btn2.clicked.connect(self.exchange)

    def exchange(self):
        fp1 = self.f1.filePath
        fp2 = self.f2.filePath
        self.f1.setFilePath(fp2)
        self.f2.setFilePath(fp1)

    def exec(self):
        try:
            BaseMethods.FindAndInsert(self.f1.filePath, self.s1.isColumn, self.s1.value, self.s2.value,
                                      self.f2.filePath, self.s3.isColumn, self.s3.value, self.s4.value)
            QMessageBox.warning(self, '成功', '处理已完成！请查看result.xls文件！', QMessageBox.Ok)
        except Exception as e:
            QMessageBox.warning(self, '错误', str(e), QMessageBox.Ok)