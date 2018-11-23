from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import MatchAndInsert

app = QApplication([])
win = QWidget()
f = win.font()
f.setPointSize(12)
win.setFont(f)

panel_1 = MatchAndInsert.MatchAndInsertPanel(parent = win)
panel_1.resize(730, 400)

win.setFixedSize(730, 400)
win.setWindowTitle('Excel工具')
win.show()
app.exec_()
