from __future__ import print_function
import ctypes, sys
import win32gui
import win32api
import win32con
import time
from datashape import unicode

classname = "Win32Window0"
titlename = "阴阳师-网易游戏"
# 获取句柄
hwnd = win32gui.FindWindow(classname, titlename)



def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def useAdmin():
    if is_admin():
        pass
    else:
        if sys.version_info[0] == 3:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        else:  # in python2.x
            ctypes.windll.shell32.ShellExecuteW(None, u"runas", unicode(sys.executable), unicode(__file__), None, 1)


def click_position(hwd, x_position, y_position):
    """
    鼠标左键点击指定坐标
    :param hwd:
    :param x_position:
    :param y_position:
    :param sleep:
    :return:
    """
    # 将两个16位的值连接成一个32位的地址坐标
    long_position = win32api.MAKELONG(x_position, y_position)
    win32api.PostMessage(hwd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position)
    win32api.PostMessage(hwd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position)


def team():
    if is_admin():
        print("请使用CTRL+C退出脚本！")
        while True:
            right, top, left, bottom = win32gui.GetWindowRect(hwnd)
            print(win32gui.GetWindowRect(hwnd))
            click_position(hwnd, int((left - right) * 0.873), int((bottom - top) * 0.78))#组队开始
            # print(int((left - right) * 0.873), int((bottom - top) * 0.79))
            time.sleep(0.3)
            click_position(hwnd, int((left - right) * 0.873), int((bottom - top) * 0.78))  # 组队开始
            # click_position(hwnd, int((left - right) * 0.745), int((bottom - top) * 0.705))  # 业原火挑战
            # print(int((left - right) * 0.745), int((bottom - top) * 0.705))
            time.sleep(0.4)
            # click_position(hwnd, int((left - right) * 0.873), int((bottom - top) * 0.77))  # 业原火开始
            # click_position(hwnd, int((left - right) * 0.73), int((bottom - top) * 0.282))#点怪
            time.sleep(0.5)
    else:
        if sys.version_info[0] == 3:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        else:  # in python2.x
            ctypes.windll.shell32.ShellExecuteW(None, u"runas", unicode(sys.executable), unicode(__file__), None, 1)


team()
