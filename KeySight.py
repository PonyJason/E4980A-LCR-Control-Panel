import pyvisa
import time

class KeySight:
    def __init__(self):
        self.myInst = None

        self.pre_setting_frequency_str = "1000,1200,1400,1600,1800,2000,2500,3000,3500,4000,5000,6000,7000,8000,9000," \
                                         "10000,15000,20000,25000,30000,35000,40000,50000,60000,70000,80000,90000," \
                                         "100000,200000,300000,400000,600000,800000," \
                                         "1000000,1500000,2000000"
        self.freqNumber = self.pre_setting_frequency_str.count(',')+1  # 频率点的数量是逗号的数量加一，这样缓冲器只要刚好达到长度即可
        self.dirTextList = []
        self.forceList = []
        self.zMag = []
        self.zAng = []
        self.rm = pyvisa.ResourceManager('')
        self.rm.list_resources()

    def keysight_return_freq(self):
        return self.pre_setting_frequency_str

    def keysight_connect(self):
        # self.myInst = self.rm.open_resource('USB0::0x0957::0x0909::MY46205006::0::INSTR')  # 该指令用于设置连接
        self.myInst = self.rm.open_resource('TCPIP0::192.168.31.212::inst0::INSTR')  # 设置tcpip连接
        self.myInst.write("*RST;*CLS")
        self.myInst.write("*IDN?")
        print(self.myInst.read())  # 显示仪器信息
        return None

    def keysight_set_measurement(self):
        self.myInst.write("FORM ASC")
        self.myInst.write("TRIG:SOUR BUS ")
        self.myInst.write("APER MEDium,5")
        self.myInst.write(":CURRent:LEVel 0.001")
        self.myInst.write("COMP ON")
        self.myInst.write("LIST:MODE SEQ")
        self.myInst.write("LIST:FREQ "+self.pre_setting_frequency_str)  # 设置频率列表
        self.myInst.write("DISP:PAGE LIST")  # 用列表方式显示和记录

        # 启用数据缓冲器
        self.myInst.write("MEM:DIM DBUF, " + str(self.freqNumber))
        self.myInst.write("MEM:CLE DBUF")
        self.myInst.write("MEM:FILL DBUF")

        self.myInst.write("INIT:CONT ON")  # 启用连续测量
        self.myInst.write(":FUNC:IMP:TYPE ZTD")
        return None

    def keysight_reset_measurement(self):
        # 截取set_measurement的一部分
        self.myInst.write("TRIG:SOUR BUS ")
        # self.myInst.write("LIST:MODE STEP")
        self.myInst.write("LIST:MODE SEQ")  # 本次实验采用一次触发连续测量
        self.myInst.write("LIST:FREQ "+self.pre_setting_frequency_str)  # 设置频率列表
        self.myInst.write("DISP:PAGE LIST")

        # 启用数据缓冲器
        self.myInst.write("MEM:DIM DBUF, " + str(self.freqNumber))
        self.myInst.write("MEM:CLE DBUF")  # 清空缓冲器
        self.myInst.write("MEM:FILL DBUF")

        self.myInst.write("INIT:CONT ON")  # 启用连续测量
        self.myInst.write(":FUNC:IMP:TYPE ZTD")  # 设置测量格式
        time.sleep(1)
        return None

    def keysight_start_measurement(self):
        self.keysight_set_measurement()
        # time.sleep(3)
        self.myInst.write(":TRIG:IMM")  # 启用了连续测量触发一次即可连续测量直到填满缓存区
        # for i in range(self.freqNumber):
        #     self.myInst.write(":INIT:IMM;:TRIG:IMM")  # 没启用连续测量每发出一个触发信号则进行一次测量
        #     time.sleep(0.2)
        time.sleep(35)
        # res = self.myInst.query_ascii_values("MEM:READ? DBUF")  # 单点触发模式下立刻读取结果
        res = self.myInst.query_ascii_values(":FETC?")  # 连续测量模式下自动返回结果

        self.zMag.clear()
        self.zAng.clear()
        for i in range(self.freqNumber * 4):
            if i % 4 == 0:
                self.zMag.append(res[i])
            elif i % 4 == 1:
                self.zAng.append(res[i])
        print(res)
        time.sleep(2)
        return None  # 测量程序结束后，zMag和zAng中保存到了测量结果


