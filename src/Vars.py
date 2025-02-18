#  Vars
from PyQt5.QtCore import pyqtSignal, QObject


class gVar(QObject):
    # 定义 Prd2Rw 改变时的信号
    Prd2Rw_changed = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        # 初始化 Prd2Rw
        self._Prd2Rw = None
        # 其他不需要监控的全局变量，可直接定义
        self.attNames = ["cm", "iBodys", "iMaterial",
                         "iDensity", "iMass", "iThickness"]
        self.read_only_cols = [0, 2, 4, 6, 8, 10,
                               12, 13, 9, 11]  # 9和11目前也屏蔽，自定义参数可以运行后修改

    @property
    def Prd2Rw(self):
        return self._Prd2Rw

    @Prd2Rw.setter
    def Prd2Rw(self, new_Prd2Rw):
        if self._Prd2Rw != new_Prd2Rw:
            self._Prd2Rw = new_Prd2Rw
            # 当 Prd2Rw 改变时发出信号
            self.Prd2Rw_changed.emit(new_Prd2Rw)


# 创建全局变量实例
gVar = gVar()
