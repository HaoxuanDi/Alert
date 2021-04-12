"""
EVE 本地敌对检测脚本
"""

import time
import sys

import cv2
import numpy as np
import win32con
import win32gui
import win32ui
from PIL import Image

from PyQt5.QtWidgets import QWidget, QPushButton, QTextEdit, QLineEdit, QLabel, QApplication, QGridLayout, QComboBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QThread, pyqtSignal

import winsound

scale = 1

Name = ""
width = 1920
height = 1080

white = './target/white.png'
red = './target/red.png'
orange = './target/orange.png'
red_star = './target/red_star.png'


def get_image():
    hWnd = win32gui.FindWindow(None, f"星战前夜：晨曦 - {Name}")

    hWndDC = win32gui.GetWindowDC(hWnd)

    mfcDC = win32ui.CreateDCFromHandle(hWndDC)
    # 创建内存设备描述表
    saveDC = mfcDC.CreateCompatibleDC()
    # 创建位图对象准备保存图片
    saveBitMap = win32ui.CreateBitmap()

    # 为bitmap开辟存储空间
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    # 将截图保存到saveBitMap中
    saveDC.SelectObject(saveBitMap)
    # 保存bitmap到内存设备描述表
    saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()

    image = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)
    image.save("./target/image.png")


def find_image(color):
    img = cv2.imread('./target/image.png')
    img = cv2.resize(img, (0, 0), fx=scale, fy=scale)
    template = cv2.imread(color)
    result = cv2.matchTemplate(img, template, cv2.TM_CCOEFF_NORMED)
    threshold = 0.99
    loc = np.where(result >= threshold)

    return len(loc[0].tolist())


class MainWidget(QWidget):
    Name = "派森速成"
    width = 1920
    height = 1080

    def __init__(self):
        super(MainWidget, self).__init__()

        self.setFixedSize(400, 800)
        self.setWindowTitle("EVE本地敌对警报工具")
        self.setWindowIcon(QIcon("./target/icon.ico"))

        self.main_layout = QGridLayout()

        self.label_1 = QLabel("角色名称:")
        self.input_1 = QLineEdit()

        self.label_2 = QLabel("屏幕分辨率:")
        self.input_2 = QComboBox()
        self.input_2.addItems(["1920x1080", "2560x1440", "3840x2160"])
        self.input_2.setCurrentIndex(0)
        self.input_2.currentIndexChanged.connect(self._screen)
        self.input_2.setPlaceholderText("长像素x宽像素")

        self.text_1 = QTextEdit()
        self.text_1.setReadOnly(True)

        self.btu_1 = QPushButton("启动")
        self.btu_1.clicked.connect(self.start)

        self.btu_2 = QPushButton("停止")
        self.btu_2.clicked.connect(self.end)

        self.btu_3 = QPushButton("帮助")

        self.main_layout.addWidget(self.label_1, 0, 0, 1, 1)
        self.main_layout.addWidget(self.input_1, 0, 1, 1, 3)
        self.main_layout.addWidget(self.btu_1, 0, 4, 1, 1)
        self.main_layout.addWidget(self.label_2, 1, 0, 1, 1)
        self.main_layout.addWidget(self.input_2, 1, 1, 1, 3)
        self.main_layout.addWidget(self.btu_2, 1, 4, 1, 1)
        self.main_layout.addWidget(self.text_1, 3, 0, 4, 5)
        self.main_layout.addWidget(self.btu_3, 7, 0, 1, 5)

        self.setLayout(self.main_layout)

        self.thread = Thread()
        self.thread.signal_1.connect(self.callback_1)
        self.thread.signal_2.connect(self.callback_2)

        self.thread_ = Alert()

    def callback_1(self, string):
        self.text_1.append(string)

    def callback_2(self):
        self.thread_.start()

    def start(self):
        self.Name = self.input_1.text()
        self.thread.start()
        self.text_1.append("-----------------检测启动-----------------")

    def end(self):
        self.thread.quit()
        self.thread.terminate()
        self.text_1.append("-----------------检测停止-----------------")

    def _screen(self, s):
        self.input_2.itemText(s)
        self.width = self.input_2.itemText(s).split("x")[0]
        self.height = self.input_2.itemText(s).split("x")[1]


class Thread(QThread):
    signal_1 = pyqtSignal(str)
    signal_2 = pyqtSignal()

    def __init__(self):
        super(Thread, self).__init__()

    def run(self):
        n_white_ = 0
        n_red_ = 0
        n_orange_ = 0
        n_red_star_ = 0
        while True:
            get_image()
            n_white = find_image(white)
            n_red = find_image(red)
            n_orange = find_image(orange)
            n_red_star = find_image(red_star)

            now = int(time.time())
            timeArray = time.localtime(now)
            otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", timeArray)

            info = f"[{otherStyleTime}] 发现{n_white}个白\n[{otherStyleTime}] 发现{n_red}个红\n[{otherStyleTime}] 发现{n_orange}个橙\n[{otherStyleTime}] 发现{n_red_star}个红星\n------------------------------------\n "

            if n_white_ == n_white and n_red_ == n_red and n_orange_ == n_orange and n_red_star_ == n_red_star:
                continue

            n_white_ = n_white
            n_red_ = n_red
            n_orange_ = n_orange
            n_red_star_ = n_red_star

            self.signal_1.emit(info)

            if n_white != 0 or n_red != 0 or n_orange != 0 or n_red_star != 0:
                self.signal_2.emit()


class Alert(QThread):

    def __init__(self):
        super(Alert, self).__init__()

    def run(self):
        winsound.Beep(2000, 500)
        winsound.Beep(1000, 500)
        winsound.Beep(2000, 500)
        winsound.Beep(1000, 500)
        winsound.Beep(2000, 500)
        winsound.Beep(1000, 500)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainWidget()
    win.show()
    sys.exit(app.exec())
