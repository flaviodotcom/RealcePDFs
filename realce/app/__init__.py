from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence, QShortcut


def atalhos(self):
    fechar_janela = QKeySequence(Qt.CTRL | Qt.Key_W)
    fechar_janela = QShortcut(fechar_janela, self)
    fechar_janela.activated.connect(self.close)
