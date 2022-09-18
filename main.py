import os
import re
import sys
import glob
from bs4 import BeautifulSoup
from openpyxl import Workbook

from PySide6 import QtCore, QtWidgets


ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')


# os.makedirs('input_data', exist_ok=True)
# INPUT_PATH = glob.glob('input_data/*.html')
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), 'output')

os.makedirs(OUTPUT_PATH, exist_ok=True)


def html2excel_handle(filepath):
    html_doc = open(filepath, encoding='utf8')
    soup = BeautifulSoup(html_doc, 'html.parser')
    all_data = soup.find_all('span', class_='dspname left')

    wb = Workbook()
    ws = wb.active
    for item in all_data:
        # 名称
        # print(f'{item.text} {item.next_sibling}')
        content = item.parent.next_sibling.next_sibling.find('span')
        text = ''
        if content:
            text = content.text
            text = ILLEGAL_CHARACTERS_RE.sub(r'', text)
        ws.append([item.next_sibling, item.text, text])

    dirname = os.path.dirname(filepath)
    name = os.path.basename(filepath)
    output_filepath = os.path.join(dirname, name+'.xlsx')
    wb.save(output_filepath)
    return output_filepath


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.filepath = None

    def initUI(self):
        self.openFileButton = QtWidgets.QPushButton("打开文件")
        self.openFileButton.clicked.connect(self.openFile)
        self.filePathText = QtWidgets.QTextEdit()
        # self.filePathText.setMaximumHeight(10)

        self.html2excelButton = QtWidgets.QPushButton("转换")
        self.html2excelButton.clicked.connect(self.html2excel)

        self.hLayout = QtWidgets.QHBoxLayout(self)
        self.hLayout.addWidget(self.filePathText)
        self.hLayout.addWidget(self.openFileButton)
        self.hLayout.addWidget(self.html2excelButton)

        # self.layout = QtWidgets.QVBoxLayout(self)
        # self.layout.addLayout(self.hLayout)
        # self.layout.addWidget(self.html2excelButton)

        self.widget = QtWidgets.QWidget()
        self.widget.setLayout(self.hLayout)
        # self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)

    @QtCore.Slot()
    def html2excel(self):
        if not self.filepath:
            self.msgBox = QtWidgets.QMessageBox(self)
            self.msgBox.setText('文件路径不能为空')
            self.msgBox.exec()
            return
        filepath = html2excel_handle(self.filepath)
        self.msgBox = QtWidgets.QMessageBox(self)
        self.msgBox.setText(f'转换完成。输出路径为：{filepath}')
        self.msgBox.exec()

    @QtCore.Slot()
    def openFile(self):
        self.filepath, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,  # 父窗口对象
            "选择文件",  # 标题
            "",  # 起始目录
            "文件类型 (*.html)"  # 选择类型过滤项，过滤内容在括号中
        )
        self.filePathText.setText(self.filepath)


if __name__ == "__main__":
    app = QtWidgets.QApplication([])

    main_window = MainWindow()
    main_window.setWindowTitle('WechatExporter-html2excel')
    main_window.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
    main_window.resize(500, 100)
    main_window.show()

    sys.exit(app.exec())
