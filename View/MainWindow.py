# -*- coding: UTF-8 -*-
"""纯代码布局的主窗口与两个参数 Tab。

保持所有控件的 objectName 与旧 Main.ui 完全一致，
使 main.py 中 `self.ui.XXX` 的访问无需改动即可继续工作。
"""
from PySide2.QtWidgets import (
    QMainWindow,
    QWidget,
    QTabWidget,
    QHBoxLayout,
    QVBoxLayout,
    QGridLayout,
    QGroupBox,
    QLabel,
    QPushButton,
    QComboBox,
    QLineEdit,
    QDateTimeEdit,
    QRadioButton,
    QCheckBox,
    QTextBrowser,
    QSplitter,
)
from PySide2.QtCore import Qt


class _NamedMixin:
    """统一命名注册：setObjectName + setattr + 记录到 _widgets。"""

    _widgets = None

    def _make(self, name, cls, **kwargs):
        w = cls(**kwargs)
        w.setObjectName(name)
        setattr(self, name, w)
        self._widgets[name] = w
        return w


def _grid_row(grid, row, text, widget):
    label = QLabel(text)
    grid.addWidget(label, row, 0)
    grid.addWidget(widget, row, 1)
    return label


class ChildClothTab(QWidget, _NamedMixin):
    """1688童装 参数页（objectName=tab）。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("tab")
        self._widgets = {}

        root = QVBoxLayout(self)

        # ---- 时间 ----
        time_box = QGroupBox("时间")
        time_grid = QGridLayout(time_box)

        self.startTime = self._make("startTime", QDateTimeEdit)
        _grid_row(time_grid, 0, "开始日期", self.startTime)

        self.endTime = self._make("endTime", QDateTimeEdit)
        _grid_row(time_grid, 1, "结束日期", self.endTime)

        self.deleveredStartTime = self._make("deleveredStartTime", QDateTimeEdit)
        _grid_row(time_grid, 2, "发货开始时间", self.deleveredStartTime)

        self.deleveredEndTime = self._make("deleveredEndTime", QDateTimeEdit)
        _grid_row(time_grid, 3, "发货截止时间", self.deleveredEndTime)

        self.isLimitDeleveredTime = self._make("isLimitDeleveredTime", QRadioButton)
        self.isLimitDeleveredTime.setText("是否限制发货时间")
        self.isLimitDeleveredTime.setAutoExclusive(False)
        time_grid.addWidget(self.isLimitDeleveredTime, 4, 1)

        root.addWidget(time_box)

        # ---- 订单参数 ----
        param_box = QGroupBox("参数")
        param_grid = QGridLayout(param_box)

        self.shopName = self._make("shopName", QComboBox)
        _grid_row(param_grid, 0, "店    铺", self.shopName)

        self.Tag = self._make("Tag", QComboBox)
        _grid_row(param_grid, 1, "色    标", self.Tag)

        self.filter = self._make("filter", QComboBox)
        _grid_row(param_grid, 2, "过    滤", self.filter)

        self.orderStatus = self._make("orderStatus", QComboBox)
        _grid_row(param_grid, 3, "订单状态", self.orderStatus)

        self.orderId = self._make("orderId", QLineEdit)
        _grid_row(param_grid, 4, "订 单 号", self.orderId)

        # 价格表（输入 + 浏览按钮）
        self.priceTablePath = self._make("priceTablePath", QLineEdit)
        self.priceTablePathButton = self._make("priceTablePathButton", QPushButton)
        self.priceTablePathButton.setText("...")
        price_row = QHBoxLayout()
        price_row.addWidget(self.priceTablePath)
        price_row.addWidget(self.priceTablePathButton)
        param_grid.addWidget(QLabel("价 格 表"), 5, 0)
        param_grid.addLayout(price_row, 5, 1)

        # 保存目录（输入 + 浏览按钮）
        self.saveFilePath = self._make("saveFilePath", QLineEdit)
        self.saveFilePathButton = self._make("saveFilePathButton", QPushButton)
        self.saveFilePathButton.setText("...")
        save_row = QHBoxLayout()
        save_row.addWidget(self.saveFilePath)
        save_row.addWidget(self.saveFilePathButton)
        param_grid.addWidget(QLabel("表格保存目录"), 6, 0)
        param_grid.addLayout(save_row, 6, 1)

        root.addWidget(param_box)

        # ---- 选项 ----
        opt_box = QGroupBox("选项")
        opt_layout = QHBoxLayout(opt_box)

        self.IsPrintOwn = self._make("IsPrintOwn", QRadioButton)
        self.IsPrintOwn.setText("打印本厂订单")
        self.IsPrintOwn.setAutoExclusive(False)
        self.IsPrintOwn.setChecked(True)
        opt_layout.addWidget(self.IsPrintOwn)

        self.IsPrintUnitPrice = self._make("IsPrintUnitPrice", QRadioButton)
        self.IsPrintUnitPrice.setText("是否打印单价")
        self.IsPrintUnitPrice.setAutoExclusive(False)
        self.IsPrintUnitPrice.setChecked(True)
        opt_layout.addWidget(self.IsPrintUnitPrice)

        self.DebugLogOutput = self._make("DebugLogOutput", QCheckBox)
        self.DebugLogOutput.setText("debug日志输出到输出框")
        self.DebugLogOutput.setChecked(False)
        opt_layout.addWidget(self.DebugLogOutput)

        root.addWidget(opt_box)

        # ---- 操作 ----
        btn_box = QGroupBox("操作")
        btn_layout = QHBoxLayout(btn_box)

        self.commit = self._make("commit", QPushButton)
        self.commit.setText("出单")
        btn_layout.addWidget(self.commit)

        self.commit_2 = self._make("commit_2", QPushButton)
        self.commit_2.setText("核算成本")
        btn_layout.addWidget(self.commit_2)

        self.checkDelivery = self._make("checkDelivery", QPushButton)
        self.checkDelivery.setText("异常物流检测")
        btn_layout.addWidget(self.checkDelivery)

        root.addWidget(btn_box)
        root.addStretch(1)


class AccessorTab(QWidget, _NamedMixin):
    """1688饰品 参数页（objectName=tab_2）。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("tab_2")
        self._widgets = {}

        root = QVBoxLayout(self)

        time_box = QGroupBox("时间")
        time_grid = QGridLayout(time_box)

        self.AliAccessorStartTime = self._make("AliAccessorStartTime", QDateTimeEdit)
        _grid_row(time_grid, 0, "开始日期", self.AliAccessorStartTime)

        self.AliAccessorEndTime = self._make("AliAccessorEndTime", QDateTimeEdit)
        _grid_row(time_grid, 1, "结束日期", self.AliAccessorEndTime)

        self.AliAccessorDeleveredStartTime = self._make(
            "AliAccessorDeleveredStartTime", QDateTimeEdit
        )
        _grid_row(time_grid, 2, "发货开始时间", self.AliAccessorDeleveredStartTime)

        self.AliAccessorDeleveredEndTime = self._make(
            "AliAccessorDeleveredEndTime", QDateTimeEdit
        )
        _grid_row(time_grid, 3, "发货截止时间", self.AliAccessorDeleveredEndTime)

        self.AliAccessorIsLimitDeleveredTime = self._make(
            "AliAccessorIsLimitDeleveredTime", QRadioButton
        )
        self.AliAccessorIsLimitDeleveredTime.setText("是否限制发货时间")
        self.AliAccessorIsLimitDeleveredTime.setAutoExclusive(False)
        time_grid.addWidget(self.AliAccessorIsLimitDeleveredTime, 4, 1)

        root.addWidget(time_box)

        param_box = QGroupBox("参数")
        param_grid = QGridLayout(param_box)

        self.AliAccessorShopName = self._make("AliAccessorShopName", QComboBox)
        _grid_row(param_grid, 0, "店铺", self.AliAccessorShopName)

        self.AliAccessorTag = self._make("AliAccessorTag", QComboBox)
        _grid_row(param_grid, 1, "色标", self.AliAccessorTag)

        self.AliAccessorOrderStatus = self._make("AliAccessorOrderStatus", QComboBox)
        _grid_row(param_grid, 2, "订单状态", self.AliAccessorOrderStatus)

        self.AliAccessorOrderId = self._make("AliAccessorOrderId", QLineEdit)
        _grid_row(param_grid, 3, "订单号", self.AliAccessorOrderId)

        self.AliAccessorPriceTablePath = self._make(
            "AliAccessorPriceTablePath", QLineEdit
        )
        self.AliAccessorPriceTablePathButton = self._make(
            "AliAccessorPriceTablePathButton", QPushButton
        )
        self.AliAccessorPriceTablePathButton.setText("...")
        price_row = QHBoxLayout()
        price_row.addWidget(self.AliAccessorPriceTablePath)
        price_row.addWidget(self.AliAccessorPriceTablePathButton)
        param_grid.addWidget(QLabel("价格表"), 4, 0)
        param_grid.addLayout(price_row, 4, 1)

        self.AliAccessorSaveFilePath = self._make(
            "AliAccessorSaveFilePath", QLineEdit
        )
        self.AliAccessorSaveFilePathButton = self._make(
            "AliAccessorSaveFilePathButton", QPushButton
        )
        self.AliAccessorSaveFilePathButton.setText("...")
        save_row = QHBoxLayout()
        save_row.addWidget(self.AliAccessorSaveFilePath)
        save_row.addWidget(self.AliAccessorSaveFilePathButton)
        param_grid.addWidget(QLabel("表格保存目录"), 5, 0)
        param_grid.addLayout(save_row, 5, 1)

        root.addWidget(param_box)

        opt_box = QGroupBox("选项")
        opt_layout = QHBoxLayout(opt_box)

        self.AliAccessorIsPrintOwn = self._make(
            "AliAccessorIsPrintOwn", QRadioButton
        )
        self.AliAccessorIsPrintOwn.setText("打印本厂订单")
        self.AliAccessorIsPrintOwn.setAutoExclusive(False)
        opt_layout.addWidget(self.AliAccessorIsPrintOwn)

        self.AliAccessorIsPrintUnitPrice = self._make(
            "AliAccessorIsPrintUnitPrice", QRadioButton
        )
        self.AliAccessorIsPrintUnitPrice.setText("是否打印单价")
        self.AliAccessorIsPrintUnitPrice.setAutoExclusive(False)
        opt_layout.addWidget(self.AliAccessorIsPrintUnitPrice)

        root.addWidget(opt_box)

        btn_box = QGroupBox("操作")
        btn_layout = QHBoxLayout(btn_box)

        self.AliAccessorCommit = self._make("AliAccessorCommit", QPushButton)
        self.AliAccessorCommit.setText("出单")
        btn_layout.addWidget(self.AliAccessorCommit)

        self.AliAccessorCost = self._make("AliAccessorCost", QPushButton)
        self.AliAccessorCost.setText("核算成本")
        btn_layout.addWidget(self.AliAccessorCost)

        self.AliAccessorCheckDelivery = self._make(
            "AliAccessorCheckDelivery", QPushButton
        )
        self.AliAccessorCheckDelivery.setText("异常物流检测")
        btn_layout.addWidget(self.AliAccessorCheckDelivery)

        root.addWidget(btn_box)
        root.addStretch(1)


class MainWindow(QMainWindow):
    """主窗口：左侧三个 Tab，右侧日志框。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("MainWindow")
        self.resize(926, 530)

        central = QWidget()
        central.setObjectName("centralwidget")
        self.setCentralWidget(central)

        splitter = QSplitter(Qt.Horizontal)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(splitter)
        central.setLayout(layout)

        # 左侧 Tab
        self.tabWidget = QTabWidget()
        self.tabWidget.setObjectName("tabWidget")
        self._cloth_tab = ChildClothTab()
        self._accessor_tab = AccessorTab()
        pdd_tab = QWidget()
        pdd_tab.setObjectName("tab_3")
        self.tabWidget.addTab(self._cloth_tab, "1688童装")
        self.tabWidget.addTab(self._accessor_tab, "1688饰品")
        self.tabWidget.addTab(pdd_tab, "PDD童装")
        splitter.addWidget(self.tabWidget)

        # 右侧日志框
        self.costomLogging = QTextBrowser()
        self.costomLogging.setObjectName("costomLogging")
        self.costomLogging.setReadOnly(True)
        splitter.addWidget(self.costomLogging)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        # 把两个 Tab 内已命名的控件提升到 MainWindow 属性，
        # 使 self.ui.XXX 的访问语义与旧 QUiLoader 结果一致。
        for tab in (self._cloth_tab, self._accessor_tab):
            for name, widget in tab._widgets.items():
                setattr(self, name, widget)