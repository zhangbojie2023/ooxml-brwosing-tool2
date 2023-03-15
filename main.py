
import sys
import zipfile
import xml.dom.minidom
from PyQt6.QtWidgets import QWidget, QApplication, QMainWindow, QPushButton, QComboBox, QVBoxLayout, QHBoxLayout, QFileDialog

import xmleditor

# 自定义XML编辑器

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.result = dict()

        self.qcb1 = QComboBox()
        self.qcb1.currentTextChanged.connect(self.qcb1_currentTextChanged)

        self.qpb1 = QPushButton("开始解析")
        self.qpb1.setMaximumWidth(80)
        self.qpb1.clicked.connect(self.qpb1_clicked)

        # self.qpb2 = QPushButton("打包保存")
        # self.qpb2.setMaximumWidth(80)
        # self.qpb2.clicked.connect(self.qpb1_clicked)

        self.qbl1 = QHBoxLayout()
        self.qbl1.addWidget(self.qcb1)
        self.qbl1.addWidget(self.qpb1)
        # self.qbl1.addWidget(self.qpb2)

        self.qte1 = xmleditor.XmlEdit()

        self.qbl3 = QHBoxLayout()
        self.qbl3.addWidget(self.qte1)

        self.qbl2 = QVBoxLayout()
        self.qbl2.addLayout(self.qbl1)
        self.qbl2.addLayout(self.qbl3)

        self.container = QWidget()
        self.container.setLayout(self.qbl2)

        self.setCentralWidget(self.container)

    # 开始解析
    def qpb1_clicked(self):
        path = QFileDialog.getOpenFileName(self, "", "~", "*.docx")[0]
        if zipfile.is_zipfile(path):
            file = zipfile.ZipFile(path)
            for filename in sorted(file.namelist()):
                self.result[filename] = file.read(filename).decode()
                self.qcb1.addItem(filename)

    # 选择文件
    def qcb1_currentTextChanged(self, text):
        xm = xml.dom.minidom.parseString(self.result.get(text))
        self.qte1.setText(xm.toprettyxml())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle("OOXML Browsing Tool")
    window.resize(1100, 500)
    window.show()
    app.exec()
