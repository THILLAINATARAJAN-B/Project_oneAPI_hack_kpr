import os
import subprocess
import shutil
from PIL import ImageGrab
import time
import win32com.client
import openai
import ctypes
from datetime import datetime
import screen_brightness_control as sbc
import pygetwindow as gw
import win32gui
import win32con
import win32api
import pyautogui
import pyperclip
import json
from groq import Groq
from generator import content_generator
from pywinauto import Application
from CLICK_CV import CLICK_ELEMENT
from VERIFY_CV import verify_screen_status

class PCController_MS_Store:
    def __init__(self, root):
        self.root = root
        self.ms_store_hwnd = None

    def _execute_with_root_management(self, func, *args, **kwargs):
        self.root.withdraw()
        try:
            return func(*args, **kwargs)
        finally:
            self.root.deiconify()

    def _ensure_ms_store_open(self):
        # Check if MS Store is already open
        self.ms_store_hwnd = win32gui.FindWindow(None, "Microsoft Store")
        
        if not self.ms_store_hwnd:
            # If not open, launch MS Store
            subprocess.Popen("start ms-windows-store:", shell=True)
            time.sleep(5)  # Wait for the store to open
            self.ms_store_hwnd = win32gui.FindWindow(None, "Microsoft Store")

        # Bring MS Store window to front
        #win32gui.ShowWindow(self.ms_store_hwnd, win32con.SW_RESTORE)
        #win32gui.SetForegroundWindow(self.ms_store_hwnd)

    def screen(self):
        return self._execute_with_root_management(verify_screen_status)

    """
        def search_result_page(self):
            def _search_result_page():
                if self.screen() == 'ms_search_app':
                    pyautogui.press('tab', 1)
                if self.screen() == 'ms_search_serach_result':
                    pyautogui.press('tab', 2)
                    pyautogui.press('enter')
                    return self.search_result_page()
                else:
                    return 0
            return self._execute_with_root_management(_search_result_page)
"""

    def open_microsoft_store(self):
        return self._execute_with_root_management(self._ensure_ms_store_open)

    def search_for_app(self, app_name):
        def _search_for_app():
            self._ensure_ms_store_open()
            time.sleep(3)
            pyautogui.hotkey('ctrl', 'E') 
            time.sleep(3) # Focus search bar
            pyautogui.write(app_name)
            time.sleep(3)
            pyautogui.press('enter')
        return self._execute_with_root_management(_search_for_app)

    def install_app(self, app_name):
        def _install_app():
            self.search_for_app(app_name)
            time.sleep(2)  # Wait for search results
            pyautogui.press('tab', 2)
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.press('enter')  # Click "Get" or "Install"
        return self._execute_with_root_management(_install_app)

    def update_installed_apps(self):
        def _update_installed_apps():
            self._ensure_ms_store_open()
            CLICK_ELEMENT('ms_store_downloads')  # Switch to Library
            time.sleep(1)
            CLICK_ELEMENT('ms_store_update_all')
        return self._execute_with_root_management(_update_installed_apps)

    def check_for_app_updates(self):
        return self._execute_with_root_management(self.update_installed_apps)

    def enable_auto_update_apps(self):
        def _enable_auto_update_apps():
            self._ensure_ms_store_open()
            time.sleep(1)
            pyautogui.press('tab', presses=2)  # Navigate to settings
            pyautogui.press('down', presses=5)
            pyautogui.press('enter')
            pyautogui.press('tab', presses=1)  # Navigate to settings
            pyautogui.press('space')
            time.sleep(1)
            pyautogui.press('tab', presses=2)  # Navigate to "App updates"
            pyautogui.press('space')  # Toggle auto-update
        return self._execute_with_root_management(_enable_auto_update_apps)

    def disable_auto_update_apps(self):
        return self._execute_with_root_management(self.enable_auto_update_apps)

    def filter_apps_by_category_store(self, category):
        def _filter_apps_by_category_store():
            self._ensure_ms_store_open()
            pyautogui.hotkey('ctrl', 'e')  # Focus search bar
            pyautogui.write(category)
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.press('enter')
        return self._execute_with_root_management(_filter_apps_by_category_store)

    def search_for_free_apps(self):
        def _search_for_free_apps():
            self._ensure_ms_store_open()
            pyautogui.hotkey('ctrl', 'e')  # Focus search bar
            pyautogui.write("free")
            pyautogui.press('enter')
        return self._execute_with_root_management(_search_for_free_apps)

    def search_for_paid_apps(self):
        def _search_for_paid_apps():
            self._ensure_ms_store_open()
            pyautogui.hotkey('ctrl', 'e')  # Focus search bar
            pyautogui.write("paid")
            pyautogui.press('enter')
        return self._execute_with_root_management(_search_for_paid_apps)