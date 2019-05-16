import win32api
import win32con
import win32gui
import win32ui
import win32com.client
import re
from PIL import Image
from ctypes import windll


class WindowManager:

    def __init__(self, window_title):
        self.title = window_title
        self._handle = None
        self._width = 0
        self._height = 0
        self._xCoord = 0
        self._yCoord = 0

        self._find_window_wildcard(self.title)
        self.get_screen_pos()

    # ---------------------PRIVATE WINDOW FINDER METHODS-----------------------

    def _window_enum_callback(self, hwnd, wildcard):
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) != None:
            self._handle = hwnd

    def _find_window_wildcard(self, wildcard):
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)
        if (self._handle == None):
            print("Error in setting window: {0}".format(wildcard))

    # -------------------------------------------------------------------------

    # ----------------------------WINDOW METHODS-------------------------------

    def get_screen_pos(self):
        if (self._handle is not None):
            self._xCoord, self._yCoord, botX, botY = win32gui.GetWindowRect(self._handle)
            self._width = botX - self._xCoord
            self._height = botY - self._yCoord

    def set_foreground(self):
        win32gui.ShowWindow(self._handle, win32con.SW_RESTORE)
        win32gui.SetWindowPos(self._handle, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(self._handle, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(self._handle, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE + win32con.SWP_NOSIZE + win32con.SWP_SHOWWINDOW)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(self._handle)

    def set_foreground_maximize(self):
        win32gui.ShowWindow(self._handle, win32con.SW_MAXIMIZE)

    def move_window(self, w_NewX, w_NewY):
        win32gui.SetWindowPos(self._handle, win32con.HWND_TOP, w_NewX, w_NewY, self._width, self._height,
                              win32con.SWP_NOSIZE)
        self.get_screen_pos()

    def get_client_screen(self):
        self.get_screen_pos()

        hwndDC = win32gui.GetWindowDC(self._handle)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, self._width, self._height)

        saveDC.SelectObject(saveBitMap)

        result = windll.user32.PrintWindow(self._handle, saveDC.GetSafeHdc(), 0)
        print(result)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)

        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(self._handle, hwndDC)

        if result == 1:
            im.save(self.title + ".png")

    # -------------------------------------------------------------------------

    # ----------------------------CURSOR METHODS-------------------------------

    def set_cursor_pos_window(self, m_x, m_y):
        win32api.SetCursorPos((m_x + self._xCoord, m_y + self._yCoord))

    def set_cursor_absolute(self, m_x, m_y):
        win32api.SetCursorPos((m_x, m_y))

    def get_cursor_pos(self):
        m_x, m_y = win32api.GetCursorPos()
        print("Abs_MouseX: {0}, Abs_MouseY: {1}, Win_MouseX: {2}, Win_MouseY: {3}".format(m_x, m_y, m_x - self._xCoord,
                                                                                          m_y - self._yCoord), end='\r')

    def click_in_window(self, m_x, m_y):
        lParam = (m_y) << 16 | (m_x)
        win32gui.SendMessage(self._handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        win32gui.SendMessage(self._handle, win32con.WM_LBUTTONUP, 0, lParam)

    def right_click_in_window(self, m_x, m_y):
        lParam = (m_y) << 16 | (m_x)
        win32gui.SendMessage(self._handle, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lParam)
        win32gui.SendMessage(self._handle, win32con.WM_RBUTTONUP, 0, lParam)

    # -------------------------------------------------------------------------
