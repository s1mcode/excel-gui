import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog
from PyQt5.QtCore import pyqtSlot

class ExcelProcessor(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel处理器')
        self.setGeometry(600, 300, 300, 200)

        layout = QVBoxLayout()
        widget = QWidget(self)
        self.setCentralWidget(widget)
        widget.setLayout(layout)

        # 上传按钮
        self.uploadBtn = QPushButton('上传Excel文件')
        self.uploadBtn.clicked.connect(self.uploadFile)
        layout.addWidget(self.uploadBtn)

        # 下载按钮
        self.downloadBtn = QPushButton('下载处理后的文件')
        self.downloadBtn.clicked.connect(self.downloadFile)
        layout.addWidget(self.downloadBtn)

        # 禁用下载按钮直到文件处理完成
        self.downloadBtn.setEnabled(False)

    @pyqtSlot()
    def uploadFile(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel Files (*.xls *.xlsx)", options=options)
        if fileName:
            self.processFile(fileName)
            self.downloadBtn.setEnabled(True)

    def processFile(self, path):
        # 这里实现你的Excel文件处理逻辑
        self.df = pd.read_excel(path)
        # 示例：创建一个新列
        self.df['新列'] = 'test'
        print("文件处理完成")

    @pyqtSlot()
    def downloadFile(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self, "保存文件", "", "Excel Files (*.xls *.xlsx)", options=options)
        if fileName:
            # 保存处理后的文件
            self.df.to_excel(fileName, index=False)
            print("文件保存成功")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelProcessor()
    ex.show()
    sys.exit(app.exec_())
