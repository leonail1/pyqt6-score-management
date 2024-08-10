import os
import json
from PyQt6.QtWidgets import QMenu, QMessageBox
from PyQt6.QtGui import QAction
from .table_file_dealer import FileDealer
from .action_creator import ActionCreator

class MenuManager:
    def __init__(self, parent):
        self.parent = parent
        self.file_dealer = FileDealer(self.parent)
        self._actions = {}
        self._action_connections = {
            'import': self.file_dealer.import_file
        }

        # 获取当前脚本的目录
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # 构建配置文件的绝对路径
        actions_config_path = os.path.join(current_dir, '..', 'config', 'actions_config.json')
        menu_config_path = os.path.join(current_dir, '..', 'config', 'menu_config.json')

        try:
            self._action_config = self._load_config(actions_config_path)
            self._menu_config = self._load_config(menu_config_path)
        except FileNotFoundError as e:
            QMessageBox.critical(self.parent, "Error", f"配置文件未找到: {str(e)}")
            raise

        self._create_actions()

    def _load_config(self, filename):
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            raise FileNotFoundError(f"无法找到配置文件: {filename}")
        except json.JSONDecodeError:
            raise ValueError(f"配置文件格式错误: {filename}")

    def _create_actions(self):
        for action_config in self._action_config['actions']:
            action = ActionCreator.create_action(
                self.parent,
                action_config['text'],
                action_config['status_tip'],
                action_config.get('shortcut'),
                action_config.get('icon_path')
            )
            self._actions[action_config['action_id']] = action

            # 连接动作
            self._connect_action(action_config['action_id'], action)

    def _connect_action(self, action_id, action):
        if action_id in self._action_connections:
            action.triggered.connect(self._action_connections[action_id])
        else:
            print(f"Warning: No connection defined for action '{action_id}'")

    def _add_menu_items(self, menu, items):
        for item in items:
            if item['action_id']:
                if item['action_id'] in self._actions:
                    action = self._actions[item['action_id']]
                    action.setText(item['name'])  # 设置显示的名称
                    menu.addAction(action)
                else:
                    raise KeyError(f"Action '{item['action_id']}' not found in self._actions")

            if item.get('separator', False):
                menu.addSeparator()

            if item['submenu']:
                submenu = QMenu(item['name'], self.parent)
                menu.addMenu(submenu)
                self._add_menu_items(submenu, item['submenu'])
            elif not item['action_id']:
                # 如果没有 action_id 且没有子菜单，添加一个禁用的 action
                disabled_action = QAction(item['name'], self.parent)
                disabled_action.setEnabled(False)
                menu.addAction(disabled_action)

    def setup_menu(self):
        try:
            menubar = self.parent.menuBar()
            for menu_item in self._menu_config['menu_structure']:
                menu = menubar.addMenu(menu_item['name'])
                self._add_menu_items(menu, menu_item['items'])
        except KeyError as e:
            QMessageBox.critical(self.parent, "Error", f"Missing key in menu configuration: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"An error occurred while setting up the menu: {str(e)}")