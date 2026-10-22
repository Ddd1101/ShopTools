# -*- coding: utf-8 -*-
"""PyInstaller 运行时钩子：在 exe 所在目录创建运行所需的工作目录。

备货单 / Log / images 三个目录在打包产物中可能不存在（空目录不会被 PyInstaller
默认收集），这里在程序启动时统一创建，保证日志写文件、备货单导出、图片存取可用。
"""
import os
import sys


def _create_work_dirs():
    # 打包后 sys.executable 指向 exe，取其所在目录作为基准
    base_dir = os.path.dirname(os.path.abspath(sys.executable))
    for name in ("备货单", "Log", "images"):
        os.makedirs(os.path.join(base_dir, name), exist_ok=True)


_create_work_dirs()