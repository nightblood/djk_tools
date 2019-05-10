from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

class MaskWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.setAttribute(Qt.WA_StyledBackground)
        self.setStyleSheet('background:rgba(0,0,0,102);')
        self.setAttribute(Qt.WA_DeleteOnClose)

    def show(self):
        """重写show，设置遮罩大小与parent一致
        """
        if self.parent() is None:
            return
        layout = QVBoxLayout()
        layout.setSizeConstraint(QLayout.SetMinimumSize)

        gif_label = QLabel()
        gif_label.setStyleSheet('background-color:transparent; color:white')
        gif = QMovie('./res/loading.gif')
        gif_label.setMovie(gif)
        gif.start()

        layout.addWidget(gif_label, 0, Qt.AlignHCenter)

        self.setLayout(layout)
        parent_rect = self.parent().geometry()
        self.setGeometry(0, 0, parent_rect.width(), parent_rect.height())
        super().show()
