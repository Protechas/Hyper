"""Tint the native Windows frame while retaining standard window controls."""
import ctypes
from ctypes import wintypes
import sys

from PyQt5 import sip
from PyQt5.QtCore import QEvent, QObject, Qt, QTimer
from PyQt5.QtWidgets import QWidget


def colorref(hex_color):
    value = hex_color.lstrip('#')
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return red | green << 8 | blue << 16


class NativeFrameTheme(QObject):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.light = False
        self.set_attribute = None
        if sys.platform == 'win32':
            self.set_attribute = ctypes.windll.dwmapi.DwmSetWindowAttribute
            self.set_attribute.argtypes = (wintypes.HWND, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD)
            self.set_attribute.restype = ctypes.c_long
        app.installEventFilter(self)

    def set_light(self, light):
        self.light = bool(light)
        for widget in self.app.topLevelWidgets():
            if widget.isVisible():
                self.apply(widget)

    def apply(self, widget):
        if self.set_attribute is None or sip.isdeleted(widget) or not widget.isWindow():
            return {}
        if widget.windowFlags() & Qt.FramelessWindowHint:
            return {}
        palette = ('#dfcece', '#4b1c2a', '#b98997') if self.light else ('#34151f', '#fff0f4', '#6f2a3c')
        attributes = {20: int(not self.light), 35: colorref(palette[0]),
                      36: colorref(palette[1]), 34: colorref(palette[2])}
        results = {}
        for key, value in attributes.items():
            native_value = wintypes.DWORD(value)
            results[key] = self.set_attribute(int(widget.winId()), key, ctypes.byref(native_value), ctypes.sizeof(native_value))
        return results

    def eventFilter(self, obj, event):
        if isinstance(obj, QWidget) and event.type() in (QEvent.Show, QEvent.WinIdChange):
            if obj.isWindow() and not obj.windowFlags() & Qt.FramelessWindowHint:
                QTimer.singleShot(0, lambda widget=obj: self.apply(widget))
        return False
