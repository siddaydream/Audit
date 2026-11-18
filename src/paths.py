# -*- coding: utf-8 -*-
"""项目路径解析：兼容源码运行与 PyInstaller 打包（frozen 时以可执行文件所在目录为根）。"""
import os
import sys


def base_dir() -> str:
    if getattr(sys, "frozen", False):          # PyInstaller 打包后
        # 可执行文件需放在项目文件夹内（与 config/ data/ 同级或位于 .app 包内）；
        # Mac .app 包内路径随 PyInstaller 版本可能是 Contents/MacOS 或 Contents/Frameworks，
        # 故从可执行文件逐级向上查找包含 config/ 与 data/ 的目录（最多 6 级）。
        d = os.path.dirname(os.path.abspath(sys.executable))
        for _ in range(6):
            if os.path.isdir(os.path.join(d, "config")) and os.path.isdir(os.path.join(d, "data")):
                return d
            parent = os.path.dirname(d)
            if parent == d:
                break
            d = parent
        raise RuntimeError(
            f"未能在可执行文件附近找到 config/ 与 data/ 目录。"
            f"请把程序放在项目文件夹内（与 config、data 文件夹同级）。起点: {sys.executable}")
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


DATA_DIR = os.path.join(base_dir(), "data")
CONFIG_DIR = os.path.join(base_dir(), "config")
OUT_DIR = os.path.join(base_dir(), "output")
