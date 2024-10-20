#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project ：ShopTools
@Author  ：Ddd
@Date    ：2023/8/13 19:05 
"""
from PySide2.QtCore import QFile, Qt
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QDialog,
    QMessageBox,
)
import yaml


class MenuSettingView(QDialog):
    def __init__(self):
        super(MenuSettingView, self).__init__()

        # # 从文件中加载UI定义
        # qfile = QFile("QtUi.ui")
        # qfile.open(QFile.ReadOnly)
        # qfile.close()
        # self.ui = QUiLoader().load(qfile)

        # 加载 YAML 文件
        with open("config.yaml", "r", encoding="utf-8") as file:
            config = yaml.safe_load(file)

        layout = QVBoxLayout()

        for shop in config["shops"]:
            label = QLabel(shop["name"], self)
            line_edit = QLineEdit(self)

            layout.addWidget(label)
            layout.addWidget(line_edit)

        self.setLayout(layout)
        self.setAttribute(Qt.WA_DeleteOnClose)

        print(config["shops"])
        print(len(config["shops"]))

    def closeEvent(self, event):
        # 这是窗口关闭时调用的方法

        # 示例: 询问用户是否确定关闭
        reply = QMessageBox.question(
            self,
            "Message",
            "Are you sure you want to close?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    # def __init__(self):
    #     super().__init__()
    #     layout = QVBoxLayout()
    #
    #     self.label = QLabel("", self)
    #     self.line_edit = QLineEdit(self)
    #     self.save_button = QPushButton("Save", self)
    #     self.save_button.clicked.connect(self.on_save)
    #
    #     layout.addWidget(self.label)
    #     layout.addWidget(self.line_edit)
    #     layout.addWidget(self.save_button)
    #
    #     self.setLayout(layout)
    #     self.setAttribute(Qt.WA_DeleteOnClose)
    #
    # def on_save(self):
    #     # You can save the setting here or do whatever you need with it.
    #     setting_value = self.line_edit.text()
    #     print("Setting saved:", setting_value)
    #
    #
    #
    # def show_settings_dialog(self):
    #     settings_dialog = MenuSettingView(self)
    #     settings_dialog.exec_()
