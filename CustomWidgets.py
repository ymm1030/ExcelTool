from PyQt5.QtWidgets import *
import BaseMethods

class ExcelSelecter(QWidget):
    def __init__(self, name, **kw):
        super(ExcelSelecter, self).__init__(**kw)
        myLabel = QLabel(self)
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
        print('--------- fp is:{}'.format(fp))
        self.myLineEdit.setText(fp[0])
        self.filePath = fp[0]

    def setFilePath(self, fp):
        self.myLineEdit.setText(fp)
        self.filePath = fp

class RowColumnSelector(QWidget):
    def __init__(self, name, **kw):
        super(RowColumnSelector, self).__init__(**kw)
        self.isColumn = True
        self.value = 0
        myLabel = QLabel(self)
        myLabel.setText(name)
        myList_1 = QComboBox(self)
        myList_1.addItem('列')
        myList_1.addItem('行')
        myList_1.currentTextChanged.connect(self.rowColumnTransition)
        myList_2 = QComboBox(self)
        myList_2.currentTextChanged.connect(self.rowColumnChanged)
        myList_2.addItems(['A', 'B', 'C', 'D', 'E', 'F', 'G',
                           'H', 'I', 'J', 'K', 'L', 'M', 'N',
                           'O', 'P', 'Q', 'R', 'S', 'T',
                           'U', 'V', 'W', 'X', 'Y', 'Z'])
        myLayout = QHBoxLayout(self)
        myLayout.addWidget(myLabel)
        myLayout.addWidget(myList_1)
        myLayout.addWidget(myList_2)
        self.setLayout(myLayout)

    def rowColumnTransition(self, text):
        if text == '列':
            self.isColumn = True
        else:
            self.isColumn = False
        
    def rowColumnChanged(self, text):
        self.value = BaseMethods.GetValue(text)