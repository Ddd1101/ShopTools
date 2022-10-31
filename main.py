from PySide2 import QtCore
from PySide2.QtWidgets import QApplication
from  PrepareGoods import PrepareGoods

if __name__ == '__main__':
    # 高分辨率适配
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication([])
    w = PrepareGoods()
    w.ui.show()

    app.exec_()