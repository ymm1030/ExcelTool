import CustomWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import BaseMethods

class MatchAndInsertPanel(QWidget):
    def __init__(self, **kw):
        super(MatchAndInsertPanel, self).__init__(**kw)
        desc = QLabel(self)
        desc.setWordWrap(True)
        desc.setText('指定源和目标的条码、数量，自动将源文件中对应条码的数量插入到目标文件中的指定行或者列')
        desc.setGeometry(10, 0, 710, 50)
        desc.setAlignment(Qt.AlignCenter)
        ft = desc.font()
        ft.setBold(True)
        desc.setFont(ft)
        self.f1 = CustomWidgets.ExcelSelecter('源', parent = self)
        self.f2 = CustomWidgets.ExcelSelecter('目标', parent = self)
        self.f1.setGeometry(50, 50, 500, 50)
        self.f2.setGeometry(50, 100, 500, 50)
        self.s1 = CustomWidgets.RowColumnSelector('源设置', parent = self)
        self.s2 = CustomWidgets.RowColumnSelector('目标设置', parent = self)
        self.s1.setGeometry(50, 215, 500, 50)
        self.s2.setGeometry(50, 265, 500, 50)
        btn = QPushButton(self)
        btn.setText('执行')
        btn.setGeometry(600, 220, 100, 100)
        btn.clicked.connect(self.exec)
        btn2 = QPushButton(self)
        btn2.setText('对调')
        btn2.setGeometry(600, 70, 100, 100)
        btn2.clicked.connect(self.exchange)
        self.output = CustomWidgets.OutputFile('输出文件', 'result.xls', parent = self)
        self.output.setGeometry(50, 330, 500, 50)
        self.f2.pathChanged.connect(self.output.dirChanged)
        sheetLabel = QLabel(self)
        sheetLabel.setText('选择sheet')
        sheetLabel.setGeometry(50, 155, 100, 30)
        self.sheetList_1 = QComboBox(self)
        self.sheetList_1.setGeometry(150, 160, 150, 20)
        self.sheetList_2 = QComboBox(self)
        self.sheetList_2.setGeometry(350, 160, 150, 20)
        self.f1.pathChanged.connect(self.sourcePathChanged)
        self.f2.pathChanged.connect(self.targetPathChanged)

    def exchange(self):
        fp1 = self.f1.filePath
        fp2 = self.f2.filePath
        self.f1.setFilePath(fp2)
        self.f2.setFilePath(fp1)

    def sourcePathChanged(self, text):
        try:
            sheet_list = BaseMethods.GetSheets(text)
        except Exception as e:
            print(e)
            sheet_list = []
        self.sheetList_1.clear()
        self.sheetList_1.addItems(sheet_list)

    def targetPathChanged(self, text):
        try:
            sheet_list = BaseMethods.GetSheets(text)
        except Exception as e:
            print(e)
            sheet_list = []
        self.sheetList_2.clear()
        self.sheetList_2.addItems(sheet_list)

    def exec(self):
        try:
            if not len(self.f1.filePath) or not len(self.f2.filePath):
                raise Exception('请先指定好文件路径！')
            if not QFile(self.f1.filePath).exists() or not QFile(self.f2.filePath).exists():
                raise Exception('指定的文件可能不存在！')
            if self.s1.codeIndex == self.s1.numIndex or self.s2.codeIndex == self.s2.numIndex:
                raise Exception('条码和数量的列好像没指定！')
            if not self.sheetList_1.count() or not self.sheetList_2.count():
                raise Exception('文件类型好像不对！')
            if QFile(self.output.resultName).exists():
                buttonReply = QMessageBox.question(self, '询问', '{}好像已经存在，要覆盖吗？'.format(self.output.resultName), 
                                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if buttonReply == QMessageBox.No:
                    return
            BaseMethods.FindAndInsert(self.f1.filePath, self.sheetList_1.currentIndex(), self.s1.isColumn, self.s1.codeIndex, self.s1.numIndex,
                                      self.f2.filePath, self.sheetList_2.currentIndex(), self.s2.isColumn, self.s2.codeIndex, self.s2.numIndex,
                                      self.output.resultName)
            QMessageBox.warning(self, '成功', '处理已完成！请查看{}文件！'.format(self.output.name), QMessageBox.Ok)
        except Exception as e:
            QMessageBox.warning(self, '错误', str(e), QMessageBox.Ok)