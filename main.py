# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os
import win32ui
import sys
from PyQt5.QtCore import QObject, pyqtSignal, QThread
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QHBoxLayout, QGroupBox, \
    QInputDialog
import matplotlib.pyplot as plt
import matplotlib
import copy
from KeySight import KeySight

matplotlib.use("Qt5Agg")

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import pandas as pd
import numpy as np


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        # 此处开始写总体布局 左侧为按钮层 右侧为图表层
        # 各个按钮垂直分布，按钮分两个组，单独布局， 各个图表垂直分布
        self.layout0 = QHBoxLayout()

        self.layout_figure = QVBoxLayout()

        self.layout_button = QVBoxLayout()

        self.layout_experiment = QVBoxLayout()
        self.layout_show = QHBoxLayout()

        # 总体初始化结束（设定在最后）

        # 此处开始写左边的按钮界面
        self.groupbox1 = QGroupBox("实验记录区")
        self.groupbox2 = QGroupBox("效果展示区")
        self.btn0 = QPushButton("\n快速设置阻抗分析仪\n", self)
        self.btn1 = QPushButton("\n新建记录文本\n", self)
        self.btn1.resize(200, 50)
        self.btn1_1 = QPushButton("\n加载上次文本\n", self)
        self.btn1_1.resize(200, 50)
        self.btn2 = QPushButton("\n追加测量一次\n", self)
        self.btn2.resize(200, 50)
        self.btn3 = QPushButton("\n结束测量并保存文件\n", self)
        self.btn3.resize(200, 50)

        self.layout_experiment.addWidget(self.btn0)
        self.layout_experiment.addWidget(self.btn1)
        self.layout_experiment.addWidget(self.btn1_1)
        self.layout_experiment.addWidget(self.btn2)
        self.layout_experiment.addWidget(self.btn3)

        self.tip = QLabel("1.本次实验中：只有创建文本或者加载上次创建的\n文本才能成功记录实验结果。\n"
                          "2.本控制台能够自动测量各频率的阻抗值。\n3.首先，快速设置keysight仪器，然后新建或者加载文本。\n"
                          "4.追加测量时要填入备注信息。\n5.若仅仅观看扫描结果，则不用加载文本\n"
                          "6.文件名会自动以‘.csv’结尾", self)
        self.layout_experiment.addWidget(self.tip)
        self.layout_experiment.setStretch(0, 2)
        self.layout_experiment.setStretch(1, 2)
        self.layout_experiment.setStretch(2, 2)
        self.layout_experiment.setStretch(3, 2)
        self.layout_experiment.setStretch(4, 2)
        self.layout_experiment.setStretch(5, 1)
        self.groupbox1.setLayout(self.layout_experiment)

        self.btn4 = QPushButton("\n实时检测\n（需要几秒钟才能刷新）", self)
        self.btn4.resize(200, 50)
        self.btn5 = QPushButton("\n检测结束\n", self)
        self.btn5.resize(200, 50)

        self.layout_show.addWidget(self.btn4)
        self.layout_show.addWidget(self.btn5)

        self.groupbox2.setLayout(self.layout_show)

        self.layout_button.addWidget(self.groupbox1)
        self.layout_button.addWidget(self.groupbox2)
        self.layout_button.setStretch(0, 3)
        self.layout_button.setStretch(1, 1)

        self.btn0.clicked.connect(lambda: self.connect_task())
        self.btn1.clicked.connect(lambda: self.new_file_task())
        self.btn2.clicked.connect(lambda: self.add_one_record_task())
        self.btn1_1.clicked.connect(lambda: self.load_last_file_task())
        # self.btn3.clicked.connect(lambda: self.w3.show())
        # 按钮界面结束

        # 此处开始写图表界面，使用pyqt5内嵌matplotlib

        # 图表界面结束

        plt.title('Bode Diagram')
        self.figure_mag = plt.figure()  # 模值图
        self.figure_ang = plt.figure()  # 角度图

        self.canvas_mag = FigureCanvas(self.figure_mag)  # 将模值图放入内嵌体
        self.canvas_ang = FigureCanvas(self.figure_ang)  # 将角度图放入内嵌体

        self.layout_figure.addWidget(self.canvas_mag)
        self.layout_figure.addWidget(self.canvas_ang)

        self.groupbox3 = QGroupBox("Bode图,模值单位是分贝，相位单位是度，横坐标为10的幂")
        self.groupbox3.setLayout(self.layout_figure)
        # 此处开始总布局的设定
        self.groupbox4 = QGroupBox("操作面板")
        self.groupbox4.setLayout(self.layout_button)
        self.groupbox4.resize(300, 800)
        self.groupbox3.resize(1000, 800)
        self.layout0.addWidget(self.groupbox4)  # 左边画按键组
        self.layout0.setStretch(0, 1)
        self.layout0.addWidget(self.groupbox3)  # 右边画图表组
        self.layout0.setStretch(1, 5)

        self.setLayout(self.layout0)
        self.resize(2000, 1500)
        self.move(300, 300)
        self.trainDict = {}
        self.setWindowTitle('keysight 简易上位机')


        # 总布局设定结束，展示主窗口

        #加载keysight类型
        self.ks = KeySight()
        self.zmag = []
        self.zang = []
        # 默认文件名
        self.file_name = ' '
        # 用于自动将频率点转为列表形式的字符串
        self.freqEvalStr = ""
        self.freq_list = []  # freq_list 和freqEvalStr的关系在connect_task中
        self.show()

    def connect_task(self):
        self.ks.keysight_connect()
        self.ks.keysight_set_measurement()
        # self.freqEvalStr = "["
        # realfreq = copy.deepcopy(self.ks.keysight_return_freq())
        # self.freqEvalStr.append(realfreq)
        # self.freqEvalStr.append("]")
        # self.freq_list = list(self.freqEvalStr)
        return None

    def new_file_task(self):  # 创建新文件
        self.connect_task()
        self.file_name = self.show_dialog()

        with open(self.file_name, "w", encoding="utf-8") as f:
            f.write("density,1kHz_mag,2kHz_mag,5kHz_mag,7kHz_mag,10kHz_mag,20kHz_mag,50kHz_mag,70kHz_mag,"
                    "100kHz_mag,200kHz_mag,500kHz_mag,700kHz_mag,1MHz_mag,2MHz_mag,1kHz_ang,2kHz_ang,5kHz_ang,"
                    "7kHz_ang,10kHz_ang,20kHz_ang,50kHz_ang,70kHz_ang,100kHz_ang,200kHz_ang,500kHz_ang,"
                    "700kHz_ang,1MHz_ang,2MHz_ang\n")
            f.close()
        return None

    def show_dialog(self):
        sender = self.sender()
        if sender == self.btn1:
            text, ok = QInputDialog.getText(self, '新建文件', '请输入文件名：')
            if ok:
                return text
        if sender == self.btn2:
            text, ok = QInputDialog.getText(self, '新建记录', '请输入属性（浓度）：')
            if ok:
                return text
        return None

    def load_last_file_task(self):
        self.connect_task()
        # 当前文件夹路径
        dirpath = os.path.dirname(__file__)
        dlg = win32ui.CreateFileDialog(True)  # True表示打开文件对话框
        # 设置打开文件对话框中的初始显示目录
        dlg.SetOFNInitialDir(dirpath)
        dlg.DoModal()
        # 等待获取用户选择的文件
        self.file_name = dlg.GetPathName()  # 获取选择的文件名称
        # 如果没选择文件则filename是空的,即=""
        print(self.file_name)

        return None

    def add_one_record_task(self):  # 增加一条阻抗测量记录 以细菌浓度作为备注
        # 测量前询问物体属性
        density = self.show_dialog()
        # 首先进行一次测量
        self.ks.keysight_start_measurement()
        # 测量后得到数据zMag和zAng
        # zmag = [12000,10000,8000,6000,5000,4000,3500,3000,2500,2000,1500,1000,500,250]
        # zang = [-70,-65,-60,-55,-50,-45,-40,-35,-30,-25,-20,-15,-10,-5]
        self.zmag = None
        self.zang = None
        self.zmag = copy.deepcopy(self.ks.zMag)
        self.zang = copy.deepcopy(self.ks.zAng)
        #
        print(self.zmag)
        print(self.zang)
        # 将数据存入文件内
        with open(self.file_name, "a") as f:
            f.write(density)

            for i in range(14):
                f.write(",")
                f.write(str(self.zmag[i]))
            for i in range(14):
                f.write(",")
                f.write(str(self.zang[i]))
            f.write("\n")
            f.close()
        # 绘图
        self.draw_mag_N_ang()

        # collectdict = {"density": [], "1kHz_mag": [], "2kHz_mag": [], "5kHz_mag": [], "7kHz_mag": [], "10kHz_mag": [],
        #                "20kHz_mag": [], "50kHz_mag": [], "70kHz_mag": [], "100kHz_mag": [],
        #                "200kHz_mag": [], "500kHz_mag": [], "700kHz_mag": [], "1MHz_mag": [], "2MHz_mag": [],
        #                "1kHz_ang": [], "2kHz_ang": [], "5kHz_ang": [], "7kHz_ang": [], "10kHz_ang": [],
        #                "20kHz_ang": [], "50kHz_ang": [], "70kHz_ang": [], "100kHz_ang": [],
        #                "200kHz_ang": [], "500kHz_ang": [], "700kHz_ang": [], "1MHz_ang": [], "2MHz_ang": []
        #                }
        # df = pd.DataFrame(collectdict)
        # df.to_csv(file_name, mode="a", encoding="utf-8", header=False, index=False)
        return None

    def save_file_task(self):  # 保存文件，结束测量
        return None

    def draw_mag_N_ang(self):  # 绘制阻抗分析仪返回的图形

        # axis_x=[1000,2000,5000,7000,10000,20000,50000,70000,100000,200000,500000,700000,1000000,2000000]
        axis_x = list(eval(self.ks.pre_setting_frequency_str))
        # zmag = [12000,10000,8000,6000,5000,4000,3500,3000,2500,2000,1500,1000,500,250]
        # zang = [-70,-65,-60,-55,-50,-45,-40,-35,-30,-25,-20,-15,-10,-5]

        self.figure_mag.clear()
        self.figure_ang.clear()

        bodex = np.array(axis_x)

        bodex = np.log10(bodex)  # bode图中 横坐标是10的幂

        magDB = np.array(self.zmag)
        magDB = 20*np.log10(magDB)  # bode图中 模值的纵坐标以分贝为单位

        angDG = np.array(self.zang)


        ax = self.figure_mag.add_subplot(111)
        bx = self.figure_ang.add_subplot(111)
        ax.plot(bodex, magDB, "H-")
        bx.plot(bodex, angDG, "H-")
        self.canvas_mag.draw()
        # self.canvas_mag.flush_events()
        self.canvas_ang.draw()
        # self.canvas_ang.flush_events()

        return None


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    # plt.ion()
    # 创建QApplication实例
    app = QApplication(sys.argv)  # 获取命令行参数
    # 创建一个窗口
    w = MainWindow()
    '''
    进入程序主循环，循环扫描响应在窗口上的事件，让整个程序不会退出
    通过exit函数确保主循环安全结束
    '''
    sys.exit(app.exec_())
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
