"""销售预测助手"""

import os
import sys


def _get_base_dir():
    """获取项目根目录（兼容 PyInstaller 打包）"""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.dirname(__file__))


_VERSION_FILE = os.path.join(_get_base_dir(), 'VERSION')


def get_version():
    try:
        with open(_VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return '3.0.0'


__version__ = get_version()
