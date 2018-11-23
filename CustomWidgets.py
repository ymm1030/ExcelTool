from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import BaseMethods

class ExcelSelecter(QWidget):
    pathChanged = pyqtSignal(str)
    def __init__(self, name, **kw):
        super(ExcelSelecter, self).__init__(**kw)
        myLabel = QLabel(self)
        myLabel.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        myLabel.setText(name)
        self.myLineEdit = QLineEdit(self)
        self.myLineEdit.setReadOnly(True)
        myBtn = QPushButton(self)
        myBtn.setText('...')
        myLayout = QHBoxLayout(self)
        myLayout.addWidget(myLabel)
        myLayout.addWidget(self.myLineEdit)
        myLayout.addWidget(myBtn)
        self.setLayout(myLayout)
        myBtn.clicked.connect(self.btn_clicked)
        self.filePath = ''
    
    def btn_clicked(self):
        fp = QFileDialog.getOpenFileName(self, 'Select excel files', './', 'Excel Files (*.xls *.xlsx)')
        # print('--------- fp is:{}'.format(fp))
        self.myLineEdit.setText(fp[0])
        if self.filePath != fp[0]:
            self.pathChanged.emit(fp[0])
        self.filePath = fp[0]

    def setFilePath(self, fp):
        self.myLineEdit.setText(fp)
        self.filePath = fp
        self.pathChanged.emit(fp)

class RowColumnSelector(QWidget):
    def __init__(self, name, **kw):
        super(RowColumnSelector, self).__init__(**kw)
        self.isColumn = True
        self.codeIndex = 0
        self.numIndex = 0
        myLabel = QLabel(self)
        myLabel.setAlignment(Qt.AlignLeft|Qt.AlignVCenter)
        myLabel.setText(name)
        myList_1 = QComboBox(self)
        myList_1.addItem('列')
        myList_1.addItem('行')
        myLabel1 = QLabel(self)
        myLabel1.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        myLabel1.setText('条码')
        myList_1.currentTextChanged.connect(self.rowColumnTransition)
        myList_2 = QComboBox(self)
        myList_2.currentTextChanged.connect(self.codeIndexChanged)
        myList_2.addItems(['A', 'B', 'C', 'D', 'E', 'F', 'G',
                           'H', 'I', 'J', 'K', 'L', 'M', 'N',
                           'O', 'P', 'Q', 'R', 'S', 'T',
                           'U', 'V', 'W', 'X', 'Y', 'Z'])
        myLabel2 = QLabel(self)
        myLabel2.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        myLabel2.setText('数量')
        myList_3 = QComboBox(self)
        myList_3.currentTextChanged.connect(self.numIndexChanged)
        myList_3.addItems(['A', 'B', 'C', 'D', 'E', 'F', 'G',
                           'H', 'I', 'J', 'K', 'L', 'M', 'N',
                           'O', 'P', 'Q', 'R', 'S', 'T',
                           'U', 'V', 'W', 'X', 'Y', 'Z'])
        myLayout = QHBoxLayout(self)
        myLayout.addWidget(myLabel)
        myLayout.addWidget(myList_1)
        myLayout.addWidget(myLabel1)
        myLayout.addWidget(myList_2)
        myLayout.addWidget(myLabel2)
        myLayout.addWidget(myList_3)
        self.setLayout(myLayout)

    def rowColumnTransition(self, text):
        if text == '列':
            self.isColumn = True
        else:
            self.isColumn = False
        
    def codeIndexChanged(self, text):
        self.codeIndex = BaseMethods.GetValue(text)

    def numIndexChanged(self, text):
        self.numIndex = BaseMethods.GetValue(text)

class OutputFile(QWidget):
    def __init__(self, name, defaultName, **kw):
        super(OutputFile, self).__init__(**kw)
        self.dir = 'C:'
        self.name = defaultName
        self.resultName = self.dir + '/' + self.name 
        label = QLabel(self)
        label.setText(name)
        edit = QLineEdit(self)
        edit.setText(defaultName)
        self.result = QLabel(self)
        self.result.setText(self.resultName)
        layout = QHBoxLayout(self)
        layout.addWidget(label)
        layout.addWidget(edit)
        layout.addWidget(self.result)
        self.setLayout(layout)
        edit.textChanged.connect(self.textChanged)

    def textChanged(self, text):
        self.name = text
        if len(self.dir):
            self.result.setText(self.dir + '/' + self.name)
            self.resultName = self.result.text()
        else:
            self.resultName = ''

    def dirChanged(self, text):
        self.dir = QFileInfo(text).absoluteDir().path()
        if len(self.dir):
            self.result.setText(self.dir + '/' + self.name)
            self.resultName = self.result.text()
        else:
            self.resultName = ''