from PyQt6.QtGui import QAction, QIcon
import sys
import os
sys.path.append(os.getcwd())


class ActionCreator:
    @staticmethod
    def create_action(parent, text: str, status_tip: str, shortcut: str = None, icon_path: str = None):
        # 创建一个QAction对象
        action = QAction(text, parent)
        if icon_path:
            action.setIcon(QIcon(icon_path))
        action.setStatusTip(status_tip)
        if shortcut:
            action.setShortcut(shortcut)
        return action