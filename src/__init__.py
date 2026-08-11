"""销售预测助手"""

import os

_VERSION_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'VERSION')

def get_version():
    try:
        with open(_VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return '0.0.0'

__version__ = get_version()
