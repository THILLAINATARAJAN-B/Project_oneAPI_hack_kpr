import win32gui
import win32process
import psutil

def get_app_path(window_title):
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) == window_title:
            hwnds.append(hwnd)
        return True

    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    
    if not hwnds:
        return None

    _, pid = win32process.GetWindowThreadProcessId(hwnds[0])
    return psutil.Process(pid).exe()