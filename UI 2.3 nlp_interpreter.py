
import subprocess
import time
from pywinauto import Application
import win32gui
import win32com.client
import os

class AppController:
    def __init__(self):
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.active_app = None

    def execute(self, parsed_command):
        action = parsed_command['action']
        app = parsed_command.get('app', self.active_app)
        params = parsed_command['params']

        if action == 'start':
            return self.start_app(app)
        elif action == 'create_folder':
            return self.create_folder(params['name'])
        elif action == 'create_file':
            return self.create_file(params['name'])
        elif action == 'delete':
            return self.delete_item(params['name'])
        elif action == 'type':
            return self.type_text(app, params['text'])
        elif action == 'click':
            return self.click_element(app, params['element'])
        elif action == 'select_menu':
            return self.select_menu(app, params['menu_path'])
        else:
            raise ValueError(f"Unknown action: {action}")

    def start_app(self, app_name):
        try:
            subprocess.Popen(app_name)
            time.sleep(2)  # Wait for the app to start
            self.active_app = app_name
            return f"Started {app_name}"
        except Exception as e:
            return f"Failed to start {app_name}: {str(e)}"

    def create_folder(self, folder_name):
        try:
            os.makedirs(folder_name, exist_ok=True)
            return f"Created folder: {folder_name}"
        except Exception as e:
            return f"Failed to create folder: {str(e)}"

    def create_file(self, file_name):
        try:
            with open(file_name, 'w') as f:
                pass
            return f"Created file: {file_name}"
        except Exception as e:
            return f"Failed to create file: {str(e)}"

    def delete_item(self, item_name):
        try:
            if os.path.isfile(item_name):
                os.remove(item_name)
            elif os.path.isdir(item_name):
                os.rmdir(item_name)
            return f"Deleted: {item_name}"
        except Exception as e:
            return f"Failed to delete {item_name}: {str(e)}"

    def type_text(self, app_name, text):
        try:
            app = Application().connect(title_re=f".*{app_name}.*")
            main_window = app.top_window()
            main_window.set_focus()
            self.shell.SendKeys(text)
            return f"Typed '{text}' in {app_name}"
        except Exception as e:
            return f"Failed to type text in {app_name}: {str(e)}"

    def click_element(self, app_name, element_name):
        try:
            app = Application().connect(title_re=f".*{app_name}.*")
            main_window = app.top_window()
            element = main_window.child_window(title=element_name, control_type="Button")
            element.click_input()
            return f"Clicked {element_name} in {app_name}"
        except Exception as e:
            return f"Failed to click {element_name} in {app_name}: {str(e)}"

    def select_menu(self, app_name, menu_path):
        try:
            app = Application().connect(title_re=f".*{app_name}.*")
            main_window = app.top_window()
            main_window.menu_select(menu_path)
            return f"Selected menu {menu_path} in {app_name}"
        except Exception as e:
            return f"Failed to select menu in {app_name}: {str(e)}"