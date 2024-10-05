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

class PCController_files:
    def __init__(self,root):   
        self.root = root
        self.current_directory = os.getcwd()
        self.pc_controller_basic = PCController_Basic(root)
        print("ready to execute tasks")
    def get_window_elements(self,window_title):
        try:
            # Connect to the File Explorer window
            app = Application(backend="uia").connect(title=window_title)
            window = app.window(title=window_title)

            # Get all elements (children)
            elements = window.children()

            # Print details of each element
            for elem in elements:
                try:
                    # Get element details
                    class_name = elem.class_name()
                    rect = elem.rectangle()  # Get the bounding rectangle
                    position = (rect.left, rect.top, rect.width(), rect.height())

                    print(f"Element: {elem.window_text()}, Class: {class_name}, Position: {position}")
                except Exception as e:
                    print(f"Error getting details for element: {e}")

        except Exception as e:
            print(f"Error connecting to window '{window_title}': {e}")

    def get_current_directory(self):
        """
        Updates the current directory based on the active File Explorer window.
        """
        try:
            self.root.withdraw()
            time.sleep(0.5)
            
            file_explorer_windows = [win for win in gw.getWindowsWithTitle('File Explorer')]

            if not file_explorer_windows:
                raise RuntimeError("No File Explorer window found.")

            file_explorer_window = file_explorer_windows[0]
            file_explorer_window.activate()
            time.sleep(0.5)

            pyautogui.hotkey('ctrl', 'l')
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)

            clipboard_content = pyperclip.paste().strip()
            if not clipboard_content:
                raise RuntimeError("Clipboard is empty after copying path.")
            print(f"Clipboard content after copying: '{clipboard_content}'")

            path = clipboard_content
            if os.path.isdir(path):
                self.current_directory = path
                print(f"Updated current directory to: {path}")
            else:
                print(f"Invalid path copied: {path}")

        except Exception as e:
            print(f"Failed to get current directory: {e}")
        finally:
            self.root.deiconify()

        return self.current_directory


    def open(self, target):
        print("Opening file or directory...")
        """
        Opens the specified target in File Explorer and brings the window to the foreground.
        """
        if not isinstance(target, str):
            print(f"Invalid target type: {type(target)}")
            raise ValueError("Target should be a string.")

        try:
            # Check if the target is a drive
            if target.lower() in ['c:', 'd:', 'e:', 'f:', 'a:', 'b:']:
                formatted_target = target.rstrip('\\') + '\\'
                self.open_in_file_explorer(formatted_target)
                self.set_current_directory(formatted_target)
                print(f"Opened drive: {target}")
                self.focus_on_explorer_window(formatted_target)
                return

            # Get the current directory
            current_directory = self.get_current_directory()
            if not isinstance(current_directory, str):
                print(f"Invalid current directory type: {type(current_directory)}")
                raise ValueError("Current directory should be a string.")

            # Check if the target is a directory in the current directory
            full_path = os.path.join(current_directory, target)
            if os.path.isdir(full_path):
                self.open_in_file_explorer(full_path)
                print(f"Opened folder: {full_path}")
                self.set_current_directory(full_path)
                #self.focus_on_explorer_window(full_path)
                return

            # Check if the target is an existing file
            if os.path.exists(target):
                self.open_in_file_explorer(target)
                print(f"Opened file: {target}")
                #self.focus_on_explorer_window(target)
                return

            raise FileNotFoundError(f"The path '{target}' does not exist.")
        except Exception as e:
            print(f"Error opening target '{target}': {e}")


    def set_current_directory(self, path):
        """
        Sets the current directory to the specified path if it is valid.
        """
        if os.path.isdir(path):
            self.current_directory = path
            print("settted current_directory : ",self.current_directory)
        else:
            raise ValueError(f"Invalid directory path: {path}")

    def focus_on_explorer_window(self, path):
        """
        Brings the File Explorer window containing the specified path to the foreground.
        """
        try:
            time.sleep(0.5)

            file_explorer_windows = [win for win in gw.getWindowsWithTitle('File Explorer')]

            if not file_explorer_windows:
                raise RuntimeError("No File Explorer window found.")

            for win in file_explorer_windows:
                if path.lower() in win.title.lower():
                    win.activate()
                    print(f"Focused on File Explorer window: {win.title}")
                    return

            print(f"No File Explorer window found with path: {path}")

        except Exception as e:
            print(f"Failed to focus on File Explorer window: {e}")

    def open_in_file_explorer(self,path):
        try:
            if os.path.isdir(path) or os.path.isfile(path):
                os.startfile(path)
                print(f"Opened: {path}")
            else:
                print(f"Path does not exist: {path}")
        except Exception as e:
            print(f"Failed to open path '{path}': {e}")

    def find_folder_in_directory(self, directory, folder_name):
        """
        Searches for a folder with the given name in the specified directory.
        """
        if not isinstance(directory, str):
            print(f"Invalid directory type: {type(directory)}")
            return None
        
        try:
            for entry in os.listdir(directory):
                full_path = os.path.join(directory, entry)
                if os.path.isdir(full_path) and entry.lower() == folder_name.lower():
                    return full_path
            return None
        except Exception as e:
            print(f"Error searching for folder '{folder_name}': {e}")
            return None
    

    def create_folder(self, folder_name):
        print("Creating folder")
        """
        Creates a folder in the path of the currently active File Explorer window.
        """
        try:
            # Temporarily hide the Tkinter window
            if self.root:
                self.root.withdraw()  # Equivalent to iconify() but hides the window
            
            # Copy the path from File Explorer
            self.copy_path_from_file_explorer()

            # Retrieve the path from clipboard
            path = pyperclip.paste().strip()
            if not path:
                raise RuntimeError("No path found in clipboard.")
            print(f"Copied path from clipboard: '{path}'")

            # Restore the Tkinter window
            if self.root:
                self.root.deiconify()

            # Create the folder
            if not folder_name:
                folder_name = 'New folder'
            folder_path = os.path.join(path, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            print(f"Folder created: {folder_path}")

        except Exception as e:
            print(f"Failed to create folder: {e}")

    def copy_path_from_file_explorer(self):
        """
        Copies the path from the currently active File Explorer window.
        """
        try:
            # Find and focus the File Explorer window
            file_explorer_windows = [win for win in gw.getWindowsWithTitle('File Explorer')]

            if not file_explorer_windows:
                raise RuntimeError("No File Explorer window found.")

            # Focus on the first File Explorer window
            file_explorer_window = file_explorer_windows[0]
            file_explorer_window.activate()
            time.sleep(1)  # Wait for the window to activate

            # Simulate key presses to copy the path
            pyautogui.hotkey('ctrl', 'l')  # Focus address bar
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'c')  # Copy path to clipboard
            time.sleep(1)  # Wait for clipboard to be updated

            # Verify clipboard content
            clipboard_content = pyperclip.paste().strip()
            if not clipboard_content:
                raise RuntimeError("Clipboard is empty after copying path.")
            print(f"Clipboard content after copying: '{clipboard_content}'")

        except Exception as e:
            print(f"Failed to copy path from File Explorer: {e}")

    def rename_folder(self, old_name, new_name):
        old_path = os.path.join(self.current_directory, old_name)
        new_path = os.path.join(self.current_directory, new_name)
        os.rename(old_path, new_path)
        print(f'Folder renamed from {old_name} to {new_name}')

    def delete_file_or_folder(self, item_name):
        print("deleting file")
        print(self.get_current_directory())
        item_path = os.path.join(self.current_directory, item_name)
        if os.path.isfile(item_path):
            os.remove(item_path)
        elif os.path.isdir(item_path):
            shutil.rmtree(item_path)
        print(f'Item deleted: {item_path}')

    def find_search_bar(self,window):
        """Find the search bar in the given window."""
        try:
            # Iterate through all elements to find the search bar
            elements = window.children()
            for elem in elements:
                if "Search" in elem.window_text():  # Search for the search bar
                    return elem
        except Exception as e:
                print(f"Error finding search bar: {e}")                
        return None
    
    def search_file(self, file_name):
        print("Searching for file:", file_name)
        try:
            # Withdraw the main application window
            self.root.withdraw()
            
            # Open File Explorer to the specified directory
            directory_to_search = "C:\\" 
            command = f'powershell -command "Start-Process explorer.exe -ArgumentList \'{directory_to_search}\'"'
            subprocess.Popen(command, shell=True)

            time.sleep(5)  # Wait for the File Explorer to open

            # Use pyautogui to focus on the File Explorer window
            #pyautogui.hotkey('alt', 'tab')  # Switch to the File Explorer window
            time.sleep(0.5)  # Brief pause to ensure the switch

            # Simulate pressing Tab 8 times to focus the search bar
            for _ in range(9):
                pyautogui.press('tab')
                time.sleep(0.2)  # Small delay to ensure the tab presses are registered

            # Type the file name to search for
            pyautogui.write(file_name)
            time.sleep(0.5)  # Brief pause before pressing enter
            pyautogui.press('enter')  # Press Enter to initiate the search

            print(f"Search initiated for: {file_name}")

            # After initiating the search, deiconify the main application window
            time.sleep(1)  # Wait briefly before showing the window again
            self.root.deiconify()

        except Exception as e:
            print(f"Error initiating search for '{file_name}': {e}")
            # Ensure the root window is shown again if there's an error
            self.root.deiconify()


    """
    def open_my_pc(self):
        subprocess.Popen(['explorer', 'shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}'], shell=True)
        time.sleep(2)  # Wait for File Explorer to open
    """

    def open_file(self, path):
        import os
        import subprocess
        
        if os.path.exists(path):
            if os.path.isdir(path):
                # Open the directory in File Explorer
                subprocess.run(['explorer', path])
            else:
                # Open the file in its default application
                os.startfile(path)
        else:
            raise FileNotFoundError(f"The path '{path}' does not exist.")
        

    def open_my_pc(self):
        try:
            result = subprocess.run(["explorer", "shell:MyComputerFolder"], capture_output=True, text=True, check=True)
            print(f"Command output: {result.stdout}")
        except subprocess.CalledProcessError as e:
            print(f"Failed to open 'My Computer': {e}")
            print(f"Error output: {e.stderr}")


    def close_my_pc(self):
        """
        Close the 'My PC' (or 'This PC') window if it's open.
        """
        for window in gw.getWindowsWithTitle('This PC') or gw.getWindowsWithTitle('My Computer'):
            window.activate()
            time.sleep(1)
            pyautogui.hotkey('alt', 'f4')
            print("Closed My PC window.")

    def close(self, folder_name):
        """
        Closes the file explorer window for the specified folder name.
        """
        if not folder_name or folder_name.strip() == "":
            raise ValueError("Folder name cannot be empty.")

        # Temporarily hide the Tkinter window
        if self.root:
            self.root.withdraw()  # Equivalent to iconify() but hides the window
        
        # Copy the path from the currently active File Explorer
        self.copy_path_from_file_explorer()

        # Retrieve the path from the clipboard
        path = pyperclip.paste().strip()
        if not path:
            raise RuntimeError("No path found in clipboard.")
        print(f"Copied path from clipboard: '{path}'")

        # Restore the Tkinter window
        if self.root:
            self.root.deiconify()

        if not os.path.exists(path):
            print(f"The path '{path}' does not exist.")
            raise FileNotFoundError(f"The path '{path}' does not exist.")

        print("Attempting to close path:", path)

        # Proceed to close the window for the resolved path
        if os.path.isdir(path):
            # Look for windows that contain the folder name in their title
            windows = [w for w in gw.getWindowsWithTitle('File Explorer') if folder_name.lower() in w.title.lower()]

            # Debugging information
            print("Currently open windows:")
            for window in gw.getAllWindows():
                print(f"- {window.title}")

            if not windows:
                print(f"No windows found for folder: {folder_name}")
                return

            for window in windows:
                try:
                    window.activate()
                    time.sleep(1)  # Allow time for the window to activate
                    pyautogui.hotkey('alt', 'f4')  # Send Alt+F4 to close the window
                    print(f"Closed window with title: {window.title}")
                except Exception as e:
                    print(f"Failed to close window with title: {window.title} - Error: {e}")
        else:
            print(f"Cannot close the path '{path}' as it is not a directory.")

    def close_file_explorers(self):
        """
        Closes all open File Explorer windows.
        """
        try:
            # Find all File Explorer windows
            file_explorer_windows = [win for win in gw.getWindowsWithTitle('File Explorer')]

            if not file_explorer_windows:
                print("No File Explorer windows found.")
                return

            # Close each File Explorer window
            for win in file_explorer_windows:
                try:
                    print(f"Closing File Explorer window: {win.title}")
                    win.activate()  # Bring the window to the foreground
                    time.sleep(1)  # Wait for the window to activate
                    pyautogui.hotkey('alt', 'f4')  # Send Alt+F4 to close the window
                    time.sleep(1)  # Wait a moment for the window to close
                except Exception as e:
                    print(f"Failed to close window with title '{win.title}': {e}")

        except Exception as e:
            print(f"Failed to close File Explorer windows: {e}")
    
    def close_target(self, target):
        print("closeing taerger")
        """
        Closes the target, which could be 'current' or an application name.
        """
        if target == 'current':
            print("current close")
            self.pc_controller_basic.close_current_window()
        else:
            self.PCController_Application.close_application(target)

class PCController_Basic:
    def __init__(self,root):
        self.root = root
        self.current_directory = os.getcwd()

    def focus_on_window(self, title):
        """
        Brings the window with the specified title to the foreground.
        """
        try:
            time.sleep(0.5)
            windows = gw.getWindowsWithTitle(title)
            
            if not windows:
                raise RuntimeError(f"No window found with title containing '{title}'.")

            # Activate the first window that matches the title
            window = windows[0]
            window.activate()
            print(f"Focused on window: {window.title}")

        except Exception as e:
            print(f"Failed to focus on window with title '{title}': {e}")

    def close_current_window(self):
        """Closes the currently active window by sending an Alt+F4 keystroke."""
        if self.root:
            self.root.withdraw()
        
        time.sleep(2)  # Delay to ensure focus
        
        active_window = gw.getActiveWindow()
        print(f"Active window before closing: {active_window.title if active_window else 'None'}")

        if active_window:
            active_window.activate()
            time.sleep(0.5)  # Short delay before sending Alt+F4
            pyautogui.hotkey('alt', 'f4')
            time.sleep(1)  # Wait to ensure the window has time to close
            print("Attempted to close the window.")
        else:
            print("No active window to close.")
        
        if self.root:
            self.root.deiconify()


    def take_screenshot(self, file_name=None):
        print("Taking screenshot")
        """
        Takes a screenshot and saves it with the given file name.
        If no file name is provided, generates a default name based on the current timestamp.
        """
        if file_name is None or not file_name.lower().endswith(('.png', '.jpg', '.jpeg')):
            # Generate a default filename based on the current timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            file_name = f'screenshot_{timestamp}.png'
        
        try:
            # Temporarily hide the window
            self.root.withdraw()
            # Give it a moment to ensure the window is hidden
            time.sleep(1)
            # Capture the screenshot of the entire screen
            screenshot = ImageGrab.grab()
            screenshot.save(file_name)
            print(f'Screenshot saved as {file_name}')
        except Exception as e:
            print(f'Failed to save screenshot: {e}')
        finally:
            # Restore the window state
            self.root.deiconify()



    def set_brightness(self,level):
        try:
            # Ensure the level is between 0 and 100
            level = max(0, min(100, level))
            # Set brightness
            print("level:",level)
            sbc.set_brightness(level)
            print(f"Brightness set to {level}%")
        except Exception as e:
            print(f"Failed to set brightness: {e}")

    def close_settings(self):
        """
        Close the Settings app if it's open.
        """
        for window in gw.getWindowsWithTitle('Settings'):
            window.activate()
            time.sleep(1)
            pyautogui.hotkey('alt', 'f4')
            print("Closed Settings window.")

    def open_task_manager(self):
        subprocess.Popen(['taskmgr'], shell=True)
        time.sleep(2)  # Wait for Task Manager to open

    def open_start_menu(self):
        subprocess.Popen(['powershell', '-command', 'Start-Process "start"'], shell=True)
        time.sleep(1)  # Wait for the Start Menu to open

    def open_settings(self):
        subprocess.Popen(['start', 'ms-settings:'], shell=True)
        time.sleep(2)  # Wait for Settings to open
      

class PCController_MS_App:
    def __init__(self,root):
        self.ms_apps = {
            'Word': 'Word',
            'Excel': 'Excel',
            'PowerPoint': 'PowerPoint',
            'Outlook': 'Outlook',
            'OneNote': 'OneNote',
            'Teams': 'Microsoft Teams',
            'OneDrive': 'OneDrive',
            'Skype': 'Skype',
            'Edge': 'Microsoft Edge',
            'Visual Studio Code': 'Visual Studio Code'
        }
        self.root = root
        self.current_directory = os.getcwd() 

    def open_ms_app(self, app_name):
        """
        Opens a Microsoft application using the Start menu.

        Args:
            app_name (str): The name of the Microsoft application to open.
        """
        if app_name not in self.ms_apps:
            print(f"Unsupported application: {app_name}")
            return

        search_term = self.ms_apps[app_name]

        try:
            # Press the Windows key to open the Start menu
            pyautogui.press('win')
            time.sleep(1)  # Wait for the Start menu to open

            # Type the application name
            pyautogui.write(search_term)
            time.sleep(1)  # Wait for search results

            # Press Enter to open the application
            pyautogui.press('enter')
            time.sleep(3)  # Wait for the application to open

            print(f"Successfully opened {app_name}.")
        except Exception as e:
            print(f"Failed to open {app_name}: {e}")

    def open_blank_word_doc(self):
        """
        Opens a blank Microsoft Word document.

        Returns:
            tuple: (word_app, doc) if successful, (None, None) otherwise
        """
        try:
            word_app = win32com.client.Dispatch('Word.Application')
            word_app.Visible = True
            doc = word_app.Documents.Add()
            print("Successfully opened a blank Word document.")
            return word_app, doc
        except Exception as e:
            print(f'Failed to open Word document: {e}')
            return None, None

    def open_blank_excel_workbook(self):
        """
        Opens a blank Microsoft Excel workbook.

        Returns:
            tuple: (excel_app, workbook) if successful, (None, None) otherwise
        """
        try:
            excel_app = win32com.client.Dispatch('Excel.Application')
            excel_app.Visible = True
            workbook = excel_app.Workbooks.Add()
            print("Successfully opened a blank Excel workbook.")
            return excel_app, workbook
        except Exception as e:
            print(f'Failed to open Excel workbook: {e}')
            return None, None

    def open_blank_powerpoint_presentation(self):
        """
        Opens a blank Microsoft PowerPoint presentation.

        Returns:
            tuple: (powerpoint_app, presentation) if successful, (None, None) otherwise
        """
        try:
            powerpoint_app = win32com.client.Dispatch('PowerPoint.Application')
            powerpoint_app.Visible = True
            presentation = powerpoint_app.Presentations.Add()
            print("Successfully opened a blank PowerPoint presentation.")
            return powerpoint_app, presentation
        except Exception as e:
            print(f'Failed to open PowerPoint presentation: {e}')
            return None, None

    def open_outlook(self):
        """
        Opens Microsoft Outlook.

        Returns:
            outlook_app if successful, None otherwise
        """
        try:
            outlook_app = win32com.client.Dispatch('Outlook.Application')
            namespace = outlook_app.GetNamespace("MAPI")
            print("Successfully opened Outlook.")
            return outlook_app
        except Exception as e:
            print(f'Failed to open Outlook: {e}')
            return None

    def open_onenote(self):
        """Opens Microsoft OneNote."""
        self.open_ms_app('OneNote')

    def open_teams(self):
        """Opens Microsoft Teams."""
        self.open_ms_app('Teams')

    def open_onedrive(self):
        """Opens Microsoft OneDrive."""
        self.open_ms_app('OneDrive')

    def open_skype(self):
        """Opens Skype."""
        self.open_ms_app('Skype')

    def open_edge(self):
        """Opens Microsoft Edge."""
        self.open_ms_app('Edge')

    def open_vscode(self):
        """Opens Visual Studio Code."""
        self.open_ms_app('Visual Studio Code')

    def close_ms_app(self, app):
        """
        Closes a Microsoft application.

        Args:
            app: The application object to close
        """
        try:
            app.Quit()
            print(f"Successfully closed the application.")
        except Exception as e:
            print(f"Failed to close the application: {e}")

    def save_document(self, doc, file_path):
        """
        Saves a document to the specified file path.

        Args:
            doc: The document object to save
            file_path (str): The file path to save the document to
        """
        try:
            doc.SaveAs(file_path)
            print(f"Successfully saved the document to {file_path}")
        except Exception as e:
            print(f"Failed to save the document: {e}")

    def paste_content_to_word(self, doc, content):
        """
        Pastes the provided content into a Word document.

        Args:
            doc: The Word document object
            content (str): The content to paste into the document
        """
        try:
            doc.Content.Text = content
            print("Successfully pasted content into the Word document.")
        except Exception as e:
            print(f"Failed to paste content into the Word document: {e}")

    def create_excel_chart(self, workbook, sheet, data, chart_type='xlColumnClustered'):
        """
        Creates a chart in an Excel workbook.

        Args:
            workbook: The Excel workbook object
            sheet: The worksheet object
            data (list): A list of data points for the chart
            chart_type (str): The type of chart to create (default is column chart)
        """
        try:
            # Input data
            for i, value in enumerate(data, start=1):
                sheet.Cells(i, 1).Value = value

            # Create chart
            chart = sheet.Shapes.AddChart2(-1, chart_type).Chart
            chart.SetSourceData(sheet.Range(f"A1:A{len(data)}"))
            print("Successfully created chart in Excel.")
        except Exception as e:
            print(f"Failed to create chart in Excel: {e}")

    def add_slide_to_powerpoint(self, presentation, layout_index=1):
        """
        Adds a new slide to a PowerPoint presentation.

        Args:
            presentation: The PowerPoint presentation object
            layout_index (int): The index of the slide layout to use (default is 1, which is usually a blank layout)

        Returns:
            The newly created slide object
        """
        try:
            slide = presentation.Slides.Add(presentation.Slides.Count + 1, layout_index)
            print("Successfully added a new slide to the PowerPoint presentation.")
            return slide
        except Exception as e:
            print(f"Failed to add a new slide to the PowerPoint presentation: {e}")
            return None

    def send_email(self, outlook_app, to, subject, body, attachments=None):
        """
        Sends an email using Outlook.

        Args:
            outlook_app: The Outlook application object
            to (str): The recipient's email address
            subject (str): The email subject
            body (str): The email body
            attachments (list): Optional list of file paths to attach to the email
        """
        try:
            mail = outlook_app.CreateItem(0)  # 0 represents olMailItem
            mail.To = to
            mail.Subject = subject
            mail.Body = body

            if attachments:
                for attachment in attachments:
                    mail.Attachments.Add(attachment)

            mail.Send()
            print("Successfully sent the email.")
        except Exception as e:
            print(f"Failed to send the email: {e}")

class PCController_Win_security:
    def __init__(self, root):
        self.root = root
        self.current_directory = os.getcwd()
        self.security_actions = {
            'Open Windows Security': self.open_windows_security,
            'Run a quick scan': self.run_quick_scan_windows_security,
            'Run a full scan': self.run_full_scan_windows_security,
            'Run a custom scan': self.run_custom_scan_windows_security,
            'Check for security updates': self.check_for_security_updates_windows_security,
            'View security dashboard': self.view_security_dashboard_windows_security,
            'Update virus definitions': self.update_virus_definitions_windows_security,
            'Turn on real-time protection': self.turn_on_real_time_protection_windows_security,
            'Turn off real-time protection': self.turn_off_real_time_protection_windows_security,
            'View protection history': self.view_protection_history_windows_security,
            'Allow an app through firewall': self.allow_app_through_firewall_windows_security,
            'Block an app through firewall': self.block_app_through_firewall_windows_security,
            'Enable firewall': self.enable_firewall_windows_security,
            'Disable firewall': self.disable_firewall_windows_security,
            'Check device security': self.check_device_security_windows_security
        }

    def open_windows_security(self):
        """Opens Windows Security application using the Start menu."""
        print("opening winsecurity")
        try:
            # Press the Windows key to open the Start menu
            pyautogui.press('win')
            time.sleep(1)  # Wait for the Start menu to open

            # Type "Windows Security"
            pyautogui.write('Windows Security')
            time.sleep(1)  # Wait for search results

            # Press Enter to open Windows Security
            pyautogui.press('enter')
            time.sleep(3)  # Wait for Windows Security to open

            print("Successfully opened Windows Security.")
        except Exception as e:
            print(f"Failed to open Windows Security: {e}")

    def run_quick_scan_windows_security(self):
        """Runs a quick scan in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('quick_scan_button.png')
            print("Quick scan initiated.")
        except Exception as e:
            print(f"Failed to initiate quick scan: {e}")

    def run_full_scan_windows_security(self):
        """Runs a full scan in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('full_scan_button.png')
            print("Full scan initiated.")
        except Exception as e:
            print(f"Failed to initiate full scan: {e}")

    def run_custom_scan_windows_security(self):
        """Runs a custom scan in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('custom_scan_button.png')
            print("Custom scan initiated.")
        except Exception as e:
            print(f"Failed to initiate custom scan: {e}")

    def check_for_security_updates_windows_security(self):
        """Checks for security updates in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('check_for_updates_button.png')
            print("Checking for security updates.")
        except Exception as e:
            print(f"Failed to check for security updates: {e}")

    def view_security_dashboard_windows_security(self):
        """Views the security dashboard in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('dashboard_button.png')
            print("Viewing security dashboard.")
        except Exception as e:
            print(f"Failed to view security dashboard: {e}")

    def update_virus_definitions_windows_security(self):
        """Updates virus definitions in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('update_definitions_button.png')
            print("Updating virus definitions.")
        except Exception as e:
            print(f"Failed to update virus definitions: {e}")

    def turn_on_real_time_protection_windows_security(self):
        """Turns on real-time protection in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('real_time_protection_toggle.png')
            pyautogui.click('turn_on_button.png')
            print("Real-time protection turned on.")
        except Exception as e:
            print(f"Failed to turn on real-time protection: {e}")

    def turn_off_real_time_protection_windows_security(self):
        """Turns off real-time protection in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('real_time_protection_toggle.png')
            pyautogui.click('turn_off_button.png')
            print("Real-time protection turned off.")
        except Exception as e:
            print(f"Failed to turn off real-time protection: {e}")

    def view_protection_history_windows_security(self):
        """Views protection history in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('protection_history_button.png')
            print("Viewing protection history.")
        except Exception as e:
            print(f"Failed to view protection history: {e}")

    def allow_app_through_firewall_windows_security(self, app_path):
        """Allows an app through the firewall in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('firewall_settings.png')
            pyautogui.click('allow_app_button.png')
            pyautogui.write(app_path)
            pyautogui.press('enter')
            print(f"Allowed {app_path} through firewall.")
        except Exception as e:
            print(f"Failed to allow app through firewall: {e}")

    def block_app_through_firewall_windows_security(self, app_path):
        """Blocks an app through the firewall in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('firewall_settings.png')
            pyautogui.click('block_app_button.png')
            pyautogui.write(app_path)
            pyautogui.press('enter')
            print(f"Blocked {app_path} through firewall.")
        except Exception as e:
            print(f"Failed to block app through firewall: {e}")

    def enable_firewall_windows_security(self):
        """Enables firewall in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('firewall_settings.png')
            pyautogui.click('enable_firewall_button.png')
            print("Firewall enabled.")
        except Exception as e:
            print(f"Failed to enable firewall: {e}")

    def disable_firewall_windows_security(self):
        """Disables firewall in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('firewall_settings.png')
            pyautogui.click('disable_firewall_button.png')
            print("Firewall disabled.")
        except Exception as e:
            print(f"Failed to disable firewall: {e}")

    def check_device_security_windows_security(self):
        """Checks device security in Windows Security."""
        self.open_windows_security()
        try:
            pyautogui.click('device_security_button.png')
            print("Checking device security.")
        except Exception as e:
            print(f"Failed to check device security: {e}")

    def perform_security_action(self, action_name, *args):
        """
        Performs a specified security action.

        Args:
            action_name (str): The name of the security action to perform.
            *args: Additional arguments required for the action.
        """
        action = self.security_actions.get(action_name)
        if action:
            try:
                action(*args)
            except Exception as e:
                print(f"Failed to perform {action_name}: {e}")
        else:
            print(f"Action '{action_name}' not found in the security actions list.")

        
class PCController_Application:
    def __init__(self,root):
        self.applications = {
            'Notepad': 'notepad',  # Just use the name for built-in apps
            'Calculator': 'calc',  # Just use the name for built-in apps
            'WordPad': 'write',    # Just use the name for built-in apps
            'WhatsApp': 'WhatsApp'  # Use the name for apps with a known Windows Search shortcut
        }   
        self.root = root
        self.current_directory = os.getcwd() 



    def close_application(self, app_name=None):
        """
        Closes the specified application using the taskkill command.
        """
        if not app_name.lower().endswith('.exe'):
            app_name += '.exe'
        try:
            subprocess.run(['taskkill', '/IM', app_name, '/F'], check=True)
            print(f"Successfully closed {app_name}")
        except subprocess.CalledProcessError as e:
            print(f"Failed to close {app_name}: {e}")
        except Exception as e:
            print(f"An error occurred: {e}")
    def open_application(self, app_name):
        """
        Opens the specified application based on its name using Windows Search.

        Args:
            app_name (str): The name of the application to open.
        """
        # Retrieve the application path or command from the dictionary using the provided name
        app_command = None
        print(f"Application command: {app_command}")

        if app_name.lower()=='whatsapp':
            self.open_whatsapp()

        if app_command:
            try:
                # Open the application using the start command
                subprocess.Popen(['start', app_command], shell=True)
                print(f"Successfully opened {app_name}.")
            except Exception as e:
                # Handle exceptions (e.g., path issues, permission issues)
                print(f"Failed to open {app_name}: {e}")
        else:
            # Handle the case where the application name is not found
            print(f"Application '{app_name}' not found in the applications list.")

    def open_whatsapp(self):
        """Open WhatsApp and perform actions."""
        # Open WhatsApp from the Windows search bar
        pyautogui.hotkey('win', 's')
        time.sleep(1)  # Wait for search bar to open
        pyautogui.write('whatsapp')
        time.sleep(1)  # Wait for search results
        pyautogui.press('enter')
        time.sleep(5)  # Wait for WhatsApp to open
            
        # Maximize the WhatsApp window
        whatsapp_window = gw.getWindowsWithTitle('WhatsApp')[0]
        whatsapp_window.maximize()
        time.sleep(2)  # Wait for the window to maximize
    
    def search_whatsapp(self,contant):
        self.open_whatapp()

class PCController_Paint:
    def __init__(self, root):
        self.root = root
        self.paint_hwnd = None

    def _ensure_paint_open(self):
        self.paint_hwnd = win32gui.FindWindow(None, "Untitled - Paint")
        if not self.paint_hwnd:
            try:
                pyautogui.press('win')
                time.sleep(1)
                pyautogui.write('paint')
                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(2)
                self.paint_hwnd = win32gui.FindWindow(None, "Untitled - Paint")
            except Exception as e:
                print(f"Failed to open Paint: {e}")
                return False
        return True

    def _focus_paint(self):
        if self.paint_hwnd:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(self.paint_hwnd)

    def _perform_action(self, action):
        if self._ensure_paint_open():
            self.root.withdraw()
            self._focus_paint()
            try:
                action()
            except Exception as e:
                print(f"Action failed: {e}")
            finally:
                self.root.deiconify()

    def open_microsoft_paint(self):
        self._perform_action(lambda: None)  # Paint is already opened in _ensure_paint_open
        print("Successfully opened Microsoft Paint.")

    def create_new_canvas_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('ctrl', 'n'))
        print("Created a new blank canvas.")

    def open_existing_image_paint(self, file_path):
        def action():
            pyautogui.hotkey('ctrl', 'o')
            time.sleep(1)
            pyautogui.write(file_path)
            pyautogui.press('enter')
        self._perform_action(action)
        print(f"Opened existing image: {file_path}")

    def save_current_image_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('ctrl', 's'))
        print("Saved the current image.")

    def save_image_as_paint(self, file_path):
        def action():
            pyautogui.hotkey('ctrl', 'shift', 's')
            time.sleep(1)
            pyautogui.write(file_path)
            pyautogui.press('enter')
        self._perform_action(action)
        print(f"Saved image as: {file_path}")

    def close_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('alt', 'f4'))
        print("Closed Paint application.")

    def resize_canvas_paint(self, width, height):
        def action():
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(1)
            pyautogui.write(str(width))
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.write(str(height))
            pyautogui.press('enter')
        self._perform_action(action)
        print(f"Resized canvas to {width}x{height}.")

    def undo_last_action_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('ctrl', 'z'))
        print("Undid last action.")

    def redo_last_action_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('ctrl', 'y'))
        print("Redid last action.")

    def draw_rectangle_paint(self, start_x, start_y, end_x, end_y):
        def action():
            pyautogui.press('r')  # Select rectangle tool
            pyautogui.mouseDown(start_x, start_y)
            pyautogui.moveTo(end_x, end_y)
            pyautogui.mouseUp()
        self._perform_action(action)
        print("Drew a rectangle.")

    def draw_circle_paint(self, center_x, center_y, radius):
        def action():
            pyautogui.press('o')  # Select oval tool
            pyautogui.mouseDown(center_x - radius, center_y - radius)
            pyautogui.moveTo(center_x + radius, center_y + radius)
            pyautogui.mouseUp()
        self._perform_action(action)
        print("Drew a circle.")

    def draw_line_paint(self, start_x, start_y, end_x, end_y):
        def action():
            pyautogui.press('l')  # Select line tool
            pyautogui.mouseDown(start_x, start_y)
            pyautogui.moveTo(end_x, end_y)
            pyautogui.mouseUp()
        self._perform_action(action)
        print("Drew a straight line.")

    def draw_freeform_shape_paint(self, points):
        def action():
            pyautogui.press('p')  # Select pencil tool
            pyautogui.mouseDown(points[0][0], points[0][1])
            for point in points[1:]:
                pyautogui.moveTo(point[0], point[1])
            pyautogui.mouseUp()
        self._perform_action(action)
        print("Drew a freeform shape.")

    def change_brush_size_paint(self, size):
        def action():
            pyautogui.press('b')  # Select brush tool
            for _ in range(size):
                pyautogui.hotkey('ctrl', '+')
        self._perform_action(action)
        print(f"Changed brush size to {size}.")

    def change_brush_color_paint(self, color):
        def action():
            pyautogui.press('b')  # Select brush tool
            pyautogui.click(color)  # Click on the color in the color palette
        self._perform_action(action)
        print(f"Changed brush color to {color}.")

    def change_brush_type_paint(self, brush_type):
        def action():
            pyautogui.press('b')  # Select brush tool
            pyautogui.press('right', presses=brush_type)
        self._perform_action(action)
        print(f"Changed brush type to {brush_type}.")

    def fill_shape_with_color_paint(self, x, y, color):
        def action():
            pyautogui.press('f')  # Select fill tool
            pyautogui.click(color)  # Select color
            pyautogui.click(x, y)  # Click inside the shape to fill
        self._perform_action(action)
        print("Filled shape with color.")

    def use_eraser_tool_paint(self, start_x, start_y, end_x, end_y):
        def action():
            pyautogui.press('e')  # Select eraser tool
            pyautogui.mouseDown(start_x, start_y)
            pyautogui.moveTo(end_x, end_y)
            pyautogui.mouseUp()
        self._perform_action(action)
        print("Used eraser tool.")

    def select_area_paint(self, start_x, start_y, end_x, end_y):
        def action():
            pyautogui.press('s')  # Select selection tool
            pyautogui.mouseDown(start_x, start_y)
            pyautogui.moveTo(end_x, end_y)
            pyautogui.mouseUp()
        self._perform_action(action)
        print("Selected area.")

    def copy_selected_area_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('ctrl', 'c'))
        print("Copied selected area.")

    def cut_selected_area_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('ctrl', 'x'))
        print("Cut selected area.")

    def paste_copied_area_paint(self, x, y):
        def action():
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.moveTo(x, y)
            pyautogui.click()
        self._perform_action(action)
        print("Pasted copied area.")

    def zoom_in_image_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('ctrl', '+'))
        print("Zoomed in on image.")

    def zoom_out_image_paint(self):
        self._perform_action(lambda: pyautogui.hotkey('ctrl', '-'))
        print("Zoomed out of image.")

    def set_image_as_background_paint(self):
        def action():
            pyautogui.click('Set as desktop background')
        self._perform_action(action)
        print("Set image as desktop background.")

    def add_text_to_image_paint(self, text, x, y):
        def action():
            pyautogui.press('t')  # Select text tool
            pyautogui.click(x, y)
            pyautogui.write(text)
        self._perform_action(action)
        print(f"Added text '{text}' to image.")

    def change_font_size_paint(self, size):
        def action():
            pyautogui.press('t')  # Select text tool
            pyautogui.hotkey('ctrl', 'a')
            for _ in range(size):
                pyautogui.hotkey('ctrl', '.')
        self._perform_action(action)
        print(f"Changed font size to {size}.")

    def change_font_color_paint(self, color):
        def action():
            pyautogui.press('t')  # Select text tool
            pyautogui.click(color)  # Click on the color in the color palette
        self._perform_action(action)
        print(f"Changed font color to {color}.")

    def draw_polygon_paint(self, points):
        def action():
            pyautogui.press('g')  # Select polygon tool
            for point in points:
                pyautogui.click(point[0], point[1])
            pyautogui.doubleClick()  # Close the polygon
        self._perform_action(action)
        print("Drew a polygon.")

    def change_canvas_background_color_paint(self, color):
        def action():
            pyautogui.press('f')  # Select fill tool
            pyautogui.click(color)  # Select color
            pyautogui.rightClick(1, 1)  # Right-click on the top-left corner
            pyautogui.press('f')  # Choose 'Fill with color'
        self._perform_action(action)
        print("Changed canvas background color.")

    def delete_selected_area_paint(self):
        self._perform_action(lambda: pyautogui.press('delete'))
        print("Deleted selected area.")

    def move_selected_area_paint(self, dx, dy):
        def action():
            pyautogui.press('s')  # Select selection tool
            pyautogui.moveRel(dx, dy)
            pyautogui.click()
        self._perform_action(action)
        print(f"Moved selected area by ({dx}, {dy}).")

    def merge_two_images_paint(self, second_image_path):
        def action():
            self.open_existing_image_paint(second_image_path)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.hotkey('ctrl', 'c')
            self.undo_last_action_paint()  # To unselect
            pyautogui.hotkey('ctrl', 'v')
        self._perform_action(action)
        print("Merged two images.")

    def use_color_picker_paint(self, x, y):
        def action():
            pyautogui.press('i')  # Select color picker tool
            pyautogui.click(x, y)
        self._perform_action(action)
        print(f"Picked color at position ({x}, {y}).")

    def resize_selected_area_paint(self, width, height):
        def action():
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(1)
            pyautogui.write(str(width))
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.write(str(height))
            pyautogui.press('enter')
        self._perform_action(action)
        print(f"Resized selected area to {width}x{height}.")

    def export_image_png_paint(self, file_path):
        def action():
            pyautogui.hotkey('ctrl', 's')
            time.sleep(1)
            pyautogui.write(file_path)
            pyautogui.press('down')  # To select PNG from dropdown
            pyautogui.press('enter')
        self._perform_action(action)
        print(f"Exported image as PNG: {file_path}")

    def export_image_jpeg_paint(self, file_path):
        def action():
            pyautogui.hotkey('ctrl', 's')
            time.sleep(1)
            pyautogui.write(file_path)
            pyautogui.press('tab')
            pyautogui.press('up', presses=3)  # To select JPEG from dropdown
            pyautogui.press('enter')
        self._perform_action(action)
        print(f"Exported image as JPEG: {file_path}")

    def export_image_bmp_paint(self, file_path):
        def action():
            pyautogui.hotkey('ctrl', 's')
            time.sleep(1)
            pyautogui.write(file_path)
            pyautogui.press('tab')
            pyautogui.press('up', presses=4)  # To select BMP from dropdown
            pyautogui.press('enter')
        self._perform_action(action)
        print(f"Exported image as BMP: {file_path}")

    def print_image_paint(self):
        def action():
            pyautogui.hotkey('ctrl', 'p')
            time.sleep(1)
            pyautogui.press('enter')
        self._perform_action(action)
        print("Sent image to printer.")

    def set_print_options_paint(self):
        def action():
            pyautogui.hotkey('ctrl', 'p')
            time.sleep(1)
            pyautogui.press('tab', presses=3)
            pyautogui.press('enter')
        self._perform_action(action)
        print("Opened print options dialog.")

    def view_image_full_screen_paint(self):
        self._perform_action(lambda: pyautogui.press('f11'))
        print("Viewing image in full screen.")


class FULL_FUNC:

    def paste_content_to_word(self, word_app, doc, content, file_name):
        """
        Pastes the generated content into the Word document and saves it.
        """
        try:
            # Add content to the document
            doc.Content.Text = content
            # Save the document in the current directory
            file_path = os.path.join(self.current_directory, file_name)
            doc.SaveAs(file_path)
            doc.Close()
            word_app.Quit()
            print(f'Word document created and saved as {file_path}.')
        except Exception as e:
            print(f'Failed to paste content or save document: {e}')

    def create_word_doc_with_content(self, topic, file_name='output.docx'):
        """
        Open a blank Word document, generate content about the topic, paste it, and save the document.
        """
        # Update the current directory from the active File Explorer window
        self.get_current_directory()

        # Open a blank Word document
        word_app, doc = self.PCController_MS_APP.open_blank_word_doc()
        if word_app is None or doc is None:
            return

        # Generate content about the specified topic
        prompt = f"Write an informative article about {topic}."
        content = self.generate_content(prompt)  # Ensure generate_content returns a string
        if not content:
            print("No content generated.")
            return
        print(f"Generated Content:\n{content}\n")

        # Focus on the Word document
        PCController_Basic.focus_on_window("Document1 - Word")  # Ensure the title is correct

        # Wait for the document to come to the foreground
        time.sleep(1)  # Adjust the sleep duration as needed for stability

        # Paste content into the Word document and save it
        self.paste_content_to_word(word_app, doc, content, file_name)

        # Optionally, you can add a step to ensure the document is saved and closed properly
        word_app.Quit()  # Ensure Word application is closed
        print(f'Word document created and saved as {file_name}.')

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
        """Ensure the Microsoft Store is open and in the foreground."""
        # Check if MS Store is already open
        self.ms_store_hwnd = win32gui.FindWindow(None, "Microsoft Store")
        
        if not self.ms_store_hwnd:
            # If not open, launch MS Store
            try:
                subprocess.Popen("start ms-windows-store:", shell=True)
                time.sleep(5)  # Wait for the store to open
                self.ms_store_hwnd = win32gui.FindWindow(None, "Microsoft Store")
            except Exception as e:
                print(f"Error launching Microsoft Store: {e}")
                return

        # If the window handle is still not found, log the failure
        if not self.ms_store_hwnd:
            print("Failed to find Microsoft Store window after attempting to open it.")
            return

        # Bring MS Store window to front
        try:
            win32gui.ShowWindow(self.ms_store_hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(self.ms_store_hwnd)
            pyautogui.click('win', 'up')  # Optional: Simulate a click to ensure focus
        except Exception as e:
            print(f"Error bringing Microsoft Store to front: {e}")


    def screen(self):
        return self._execute_with_root_management(verify_screen_status)

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
            time.sleep(8)  # Wait for search results
            CLICK_ELEMENT('ms_store_filters','single')
            # Press 'tab' key with a specified gap duration
            for _ in range(4):
                pyautogui.press('tab')
                time.sleep(1)  # Duration between pressing 'tab'
            
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(5)
            pyautogui.press('enter')  # Click "Get" or "Install"
        
        return self._execute_with_root_management(_install_app)

    def update_installed_apps(self):
        def _update_installed_apps():
            self._ensure_ms_store_open()
            CLICK_ELEMENT('ms_store_downloads','double')  # Switch to Library
            time.sleep(1)
            CLICK_ELEMENT('ms_store_update_all','single')
        return self._execute_with_root_management(_update_installed_apps)

    def check_for_app_updates(self):
        return self._execute_with_root_management(self.update_installed_apps)


    def filter_apps_by_category_store(self, category):
        def _filter_apps_by_category_store():
            self._ensure_ms_store_open()
            time.sleep(3)
            pyautogui.hotkey('ctrl', 'E') 
            time.sleep(3) # Focus search bar
            pyautogui.write(category)
            time.sleep(3)
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

class PCController_file_explorer:
    def __init__(self, root):
        self.root = root
        self.file_explorer_hwnd = None

    def _execute_with_root_management(self, func, *args, **kwargs):
        self.root.withdraw()
        try:
            return func(*args, **kwargs)
        finally:
            self.root.deiconify()

    def _ensure_file_explorer_open(self):
        # Check if File Explorer is already open
        self.file_explorer_hwnd = win32gui.FindWindow(None, "File Explorer")
        
        if not self.file_explorer_hwnd:
            # If not open, launch File Explorer
            subprocess.Popen("explorer", shell=True)
            time.sleep(5)  # Wait for File Explorer to open
            self.file_explorer_hwnd = win32gui.FindWindow(None, "File Explorer")

        # Bring File Explorer window to front
        win32gui.ShowWindow(self.file_explorer_hwnd, 5)  # SW_RESTORE
        win32gui.SetForegroundWindow(self.file_explorer_hwnd)

    def open_this_pc(self):
        def _open_this_pc():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('win', 'e')  # Open "This PC"
        return self._execute_with_root_management(_open_this_pc)

    def search_document(self):
        def _search_document():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('ctrl', 'f')  # Open search box
            pyautogui.write('document')  # Search for document
            pyautogui.press('enter')
        return self._execute_with_root_management(_search_document)

    def move_file(self, source, destination):
        def _move_file():
            self._ensure_file_explorer_open()
            # Select the file and move it
            # Implement file navigation logic here
            # For example, navigating to the file using pyautogui
            pyautogui.write(source)  # Navigate to the file
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'x')  # Cut file
            pyautogui.write(destination)  # Navigate to destination
            pyautogui.hotkey('ctrl', 'v')  # Paste file
        return self._execute_with_root_management(_move_file)

    def cut_file(self, file_path):
        def _cut_file():
            self._ensure_file_explorer_open()
            # Navigate to the file and cut it
            pyautogui.write(file_path)  # Navigate to file
            pyautogui.hotkey('ctrl', 'x')  # Cut file
        return self._execute_with_root_management(_cut_file)

    def paste_file(self, destination_folder):
        def _paste_file():
            self._ensure_file_explorer_open()
            pyautogui.write(destination_folder)  # Navigate to destination folder
            pyautogui.hotkey('ctrl', 'v')  # Paste file
        return self._execute_with_root_management(_paste_file)

    def empty_recycle_bin(self):
        def _empty_recycle_bin():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('win', 'r')  # Open Run dialog
            pyautogui.write('shell:RecycleBinFolder')
            pyautogui.press('enter')
            time.sleep(2)  # Wait for the Recycle Bin to open
            pyautogui.hotkey('shift', 'delete')  # Empty Recycle Bin
            pyautogui.press('enter')  # Confirm deletion
        return self._execute_with_root_management(_empty_recycle_bin)

    def create_shortcut(self, file_path):
        def _create_shortcut():
            self._ensure_file_explorer_open()
            # Navigate to the file
            pyautogui.write(file_path)
            time.sleep(1)
            # Right-click to open context menu
            pyautogui.press('shift', 'f10')  # Right-click
            time.sleep(1)
            # Select "Create Shortcut"
            pyautogui.press('down', presses=3)  # Navigate to "Create Shortcut"
            pyautogui.press('enter')
        return self._execute_with_root_management(_create_shortcut)

    def delete_all_files_in_downloads(self):
        def _delete_files():
            self._ensure_file_explorer_open()
            pyautogui.write('C:\\Users\\YourUsername\\Downloads')  # Change to your username
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.hotkey('ctrl', 'a')  # Select all files
            pyautogui.press('delete')  # Delete selected files
            pyautogui.press('enter')  # Confirm deletion
        return self._execute_with_root_management(_delete_files)

    def open_documents_folder(self):
        def _open_documents():
            self._ensure_file_explorer_open()
            pyautogui.write('C:\\Users\\YourUsername\\Documents')  # Change to your username
            pyautogui.press('enter')
        return self._execute_with_root_management(_open_documents)

    def open_pictures_folder(self):
        def _open_pictures():
            self._ensure_file_explorer_open()
            pyautogui.write('C:\\Users\\YourUsername\\Pictures')  # Change to your username
            pyautogui.press('enter')
        return self._execute_with_root_management(_open_pictures)

    def open_music_folder(self):
        def _open_music():
            self._ensure_file_explorer_open()
            pyautogui.write('C:\\Users\\YourUsername\\Music')  # Change to your username
            pyautogui.press('enter')
        return self._execute_with_root_management(_open_music)

    def open_downloads_folder(self):
        def _open_downloads():
            self._ensure_file_explorer_open()
            pyautogui.write('C:\\Users\\YourUsername\\Downloads')  # Change to your username
            pyautogui.press('enter')
        return self._execute_with_root_management(_open_downloads)

    def create_text_file(self):
        def _create_text_file():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('ctrl', 'n')  # New file
            pyautogui.write('New Text Document.txt')  # Change file name if needed
            pyautogui.press('enter')  # Save file
        return self._execute_with_root_management(_create_text_file)

    def check_file_size(self, file_path):
        def _check_file_size():
            self._ensure_file_explorer_open()
            pyautogui.write(file_path)  # Navigate to file
            # Select the file and check properties
            pyautogui.press('shift', 'f10')  # Right-click
            pyautogui.press('down', presses=4)  # Navigate to Properties
            pyautogui.press('enter')
            time.sleep(1)  # Wait for properties to open
            # Here, additional logic can be added to extract file size
        return self._execute_with_root_management(_check_file_size)

    def restore_file_from_recycle_bin(self, file_name):
        def _restore_file():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('win', 'r')  # Open Run dialog
            pyautogui.write('shell:RecycleBinFolder')
            pyautogui.press('enter')
            time.sleep(2)  # Wait for the Recycle Bin to open
            pyautogui.write(file_name)  # Search for the file
            pyautogui.press('enter')  # Select file
            pyautogui.press('ctrl', 'r')  # Restore file
        return self._execute_with_root_management(_restore_file)

    def rename_folder_in_documents(self, old_name, new_name):
        def _rename_folder():
            self._ensure_file_explorer_open()
            self.open_documents_folder()
            time.sleep(2)
            pyautogui.write(old_name)  # Navigate to folder
            pyautogui.press('f2')  # Rename
            pyautogui.write(new_name)  # New name
            pyautogui.press('enter')  # Confirm rename
        return self._execute_with_root_management(_rename_folder)

    def show_hidden_files(self):
        def _show_hidden_files():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('alt', 'v')  # Open View menu
            pyautogui.press('h')  # Select Hidden items
        return self._execute_with_root_management(_show_hidden_files)

    def hide_file_extensions(self):
        def _hide_file_extensions():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('alt', 't')  # Open Tools menu
            pyautogui.press('o')  # Select Folder options
            time.sleep(1)
            pyautogui.press('tab', presses=3)  # Navigate to View tab
            pyautogui.press('down', presses=5)  # Navigate to Hide extensions
            pyautogui.press('space')  # Toggle option
            pyautogui.press('enter')  # Confirm
        return self._execute_with_root_management(_hide_file_extensions)

    def show_file_extensions(self):
        def _show_file_extensions():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('alt', 't')  # Open Tools menu
            pyautogui.press('o')  # Select Folder options
            time.sleep(1)
            pyautogui.press('tab', presses=3)  # Navigate to View tab
            pyautogui.press('down', presses=5)  # Navigate to Show extensions
            pyautogui.press('space')  # Toggle option
            pyautogui.press('enter')  # Confirm
        return self._execute_with_root_management(_show_file_extensions)

    def compress_folder(self, folder_path):
        def _compress_folder():
            self._ensure_file_explorer_open()
            pyautogui.write(folder_path)  # Navigate to folder
            pyautogui.press('shift', 'f10')  # Right-click
            pyautogui.press('down', presses=4)  # Select Send to
            pyautogui.press('down')  # Navigate to Compressed (zipped) folder
            pyautogui.press('enter')  # Create zip
        return self._execute_with_root_management(_compress_folder)

    def extract_zip_file(self, zip_file_path):
        def _extract_zip_file():
            self._ensure_file_explorer_open()
            pyautogui.write(zip_file_path)  # Navigate to zip file
            pyautogui.press('shift', 'f10')  # Right-click
            pyautogui.press('down', presses=2)  # Select Extract All
            pyautogui.press('enter')  # Confirm extraction
        return self._execute_with_root_management(_extract_zip_file)

    def sort_files_by_size(self):
        def _sort_files_by_size():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('alt', 'v')  # Open View menu
            pyautogui.press('s')  # Select Sort by
            pyautogui.press('z')  # Sort by size
        return self._execute_with_root_management(_sort_files_by_size)

    def sort_files_by_date_modified(self):
        def _sort_files_by_date_modified():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('alt', 'v')  # Open View menu
            pyautogui.press('s')  # Select Sort by
            pyautogui.press('d')  # Sort by date modified
        return self._execute_with_root_management(_sort_files_by_date_modified)

    def select_all_files_in_folder(self):
        def _select_all_files():
            self._ensure_file_explorer_open()
            pyautogui.hotkey('ctrl', 'a')  # Select all files
        return self._execute_with_root_management(_select_all_files)

    def delete_folder_permanently(self, folder_path):
        def _delete_folder():
            self._ensure_file_explorer_open()
            pyautogui.write(folder_path)  # Navigate to folder
            pyautogui.press('shift', 'f10')  # Right-click
            pyautogui.press('d')  # Select Delete
            pyautogui.press('enter')  # Confirm deletion
        return self._execute_with_root_management(_delete_folder)

    def pin_folder_to_quick_access(self, folder_path):
        def _pin_folder():
            self._ensure_file_explorer_open()
            pyautogui.write(folder_path)  # Navigate to folder
            pyautogui.press('shift', 'f10')  # Right-click
            pyautogui.press('p')  # Select Pin to Quick Access
        return self._execute_with_root_management(_pin_folder)

    def unpin_folder_from_quick_access(self, folder_name):
        def _unpin_folder():
            self._ensure_file_explorer_open()
            # Navigate to Quick Access
            pyautogui.hotkey('win', 'e')  # Open File Explorer
            time.sleep(1)
            pyautogui.write(folder_name)  # Navigate to folder
            pyautogui.press('shift', 'f10')  # Right-click
            pyautogui.press('u')  # Select Unpin from Quick Access
        return self._execute_with_root_management(_unpin_folder)

    def share_file_via_network(self, file_path):
        def _share_file():
            self._ensure_file_explorer_open()
            pyautogui.write(file_path)  # Navigate to file
            pyautogui.press('shift', 'f10')  # Right-click
            pyautogui.press('s')  # Select Share
            # Additional sharing logic may be needed here
        return self._execute_with_root_management(_share_file)

    def rename_multiple_files(self, old_name, new_name):
        def _rename_multiple_files():
            self._ensure_file_explorer_open()
            # Implement logic for renaming multiple files
            # This could involve navigating to a folder, selecting files, and renaming them
        return self._execute_with_root_management(_rename_multiple_files)


class PCController_gmail:
    def __init__(self, root):
        self.root = root
        self.gmail_hwnd = None

    def _execute_with_root_management(self, func, *args, **kwargs):
        self.root.withdraw()
        try:
            return func(*args, **kwargs)
        finally:
            self.root.deiconify()

    def _ensure_gmail_open(self):
        """Ensure Gmail is open and in the foreground."""
        # Check if Gmail is already open
        self.gmail_hwnd = win32gui.FindWindow(None, "Gmail")

        if not self.gmail_hwnd:
            # If not open, launch Gmail
            try:
                subprocess.Popen("start chrome https://mail.google.com", shell=True)
                time.sleep(5)  # Wait for Gmail to open
                self.gmail_hwnd = win32gui.FindWindow(None, "Gmail")
            except Exception as e:
                print(f"Error launching Gmail: {e}")
                return

        # If the window handle is still not found, log the failure
        if not self.gmail_hwnd:
            print("Failed to find Gmail window after attempting to open it.")
            return

        # Bring Gmail window to front
        try:
            win32gui.ShowWindow(self.gmail_hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(self.gmail_hwnd)
            pyautogui.click('win', 'up')  # Optional: Simulate a click to ensure focus
        except Exception as e:
            print(f"Error bringing Gmail to front: {e}")

    def open_gmail(self):
        """Open Gmail in the browser."""
        return self._execute_with_root_management(self._ensure_gmail_open)

    def compose_new_email(self):
        """Compose a new email in Gmail."""
        def _compose_new_email():
            self._ensure_gmail_open()
            time.sleep(5)  # Wait for Gmail to be ready
            CLICK_ELEMENT('gmail_compose','single')
        return self._execute_with_root_management(_compose_new_email)

    def send_email(self,to):
        """Send the currently composed email."""
        def _send_email():
            self._ensure_gmail_open()
            time.sleep(1)
            CLICK_ELEMENT('gmail_recipients','single')
            time.sleep(1)
            pyautogui.write(to)
            time.sleep(1)
            CLICK_ELEMENT('gmail_subject','single')
            time.sleep(1)
            pyautogui.write('hi bro')
            time.sleep(3)
            pyautogui.hotkey('ctrl', 'enter')  # Shortcut to send email
        return self._execute_with_root_management(_send_email)

    def reply_to_email(self):
        """Reply to the currently selected email."""
        def _reply_to_email():
            self._ensure_gmail_open()
            time.sleep(3)
            pyautogui.hotkey('r')  # Shortcut to reply to email
        return self._execute_with_root_management(_reply_to_email)

    def forward_email(self):
        """Forward the currently selected email."""
        def _forward_email():
            self._ensure_gmail_open()
            time.sleep(3)
            pyautogui.hotkey('f')  # Shortcut to forward email
        return self._execute_with_root_management(_forward_email)

    def search_email(self, search_query):
        """Search for an email using the search query."""
        def _search_email():
            self._ensure_gmail_open()
            time.sleep(3)
            pyautogui.hotkey('ctrl', 'e')  # Focus search bar
            pyautogui.write(search_query)
            pyautogui.press('enter')
        return self._execute_with_root_management(_search_email)

    def mark_email_read(self):
        """Mark the currently selected email as read."""
        def _mark_email_read():
            self._ensure_gmail_open()
            time.sleep(3)
            pyautogui.hotkey('shift', 'i')  # Mark as read
        return self._execute_with_root_management(_mark_email_read)

    def mark_email_unread(self):
        """Mark the currently selected email as unread."""
        def _mark_email_unread():
            self._ensure_gmail_open()
            time.sleep(3)
            pyautogui.hotkey('shift', 'u')  # Mark as unread
        return self._execute_with_root_management(_mark_email_unread)

    def delete_email(self):
        """Delete the currently selected email."""
        def _delete_email():
            self._ensure_gmail_open()
            time.sleep(3)
            pyautogui.hotkey('shift', '3')  # Delete email
        return self._execute_with_root_management(_delete_email)

    def open_spam_folder(self):
        """Open the Spam folder in Gmail."""
        def _open_spam_folder():
            self._ensure_gmail_open()
            time.sleep(3)
            pyautogui.click(x=200, y=200)  # Adjust coordinates to click on the Spam folder
        return self._execute_with_root_management(_open_spam_folder)

    def logout_gmail(self):
        """Logout of Gmail."""
        def _logout_gmail():
            self._ensure_gmail_open()
            time.sleep(3)
            pyautogui.click(x=200, y=200)  # Adjust coordinates for the logout button
        return self._execute_with_root_management(_logout_gmail)


