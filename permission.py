import ctypes
import sys

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin(argv=None, debug=False):
    if argv is None:
        argv = sys.argv
    if is_admin():
        return True
    else:
        if sys.version_info[0] >= 3:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(argv), None, 1)
        else:
            ctypes.windll.shell32.ShellExecuteW(None, u"runas", unicode(sys.executable), unicode(" ".join(argv)), None, 1)
    return None
