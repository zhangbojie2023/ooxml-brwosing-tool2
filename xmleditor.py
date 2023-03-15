
from PyQt6.QtWidgets import QTextEdit

class XmlEdit(QTextEdit):
    def __init__(self):
        super().__init__()
        self.setReadOnly(True)