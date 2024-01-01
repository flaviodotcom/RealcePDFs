import os
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence, QShortcut


def resource_path(relative_path):
    current_dir = os.path.abspath(__file__)
    images_path = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))

    base_path = getattr(sys, "_MEIPASS", images_path)
    return os.path.join(base_path, relative_path)


def atalhos(self):
    fechar_janela = QKeySequence(Qt.CTRL | Qt.Key_W)
    fechar_janela = QShortcut(fechar_janela, self)
    fechar_janela.activated.connect(self.close)
