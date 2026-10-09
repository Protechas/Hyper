"""Tint the native Windows frame while retaining standard window controls."""
import ctypes
from ctypes import wintypes
import logging
import sys
import weakref

from PyQt5 import sip
from PyQt5.QtCore import QEvent, QObject, Qt, QTimer
from PyQt5.QtWidgets import QWidget


def colorref(hex_color):
    value = hex_color.lstrip('#')
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return red | green << 8 | blue << 16


class NativeFrameTheme(QObject):
    MANAGED_PROPERTY = '_hyper_native_frame_theme'

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.light = False
        self.set_attribute = None
        if sys.platform == 'win32':
            try:
                self.set_attribute = ctypes.windll.dwmapi.DwmSetWindowAttribute
                self.set_attribute.argtypes = (wintypes.HWND, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD)
                self.set_attribute.restype = ctypes.c_long
            except (AttributeError, OSError):
                logging.exception('Native title-bar theming is unavailable; using the system title bar.')
                self.set_attribute = None
        app.installEventFilter(self)

    def watch(self, widget):
        """Opt a long-lived Hyper window into native frame theming."""
        if widget is None or sip.isdeleted(widget):
            return
        widget.setProperty(self.MANAGED_PROPERTY, True)
        if widget.isVisible():
            self._schedule(widget)

    def _schedule(self, widget):
        """Schedule through a timer owned by the window being themed."""
        try:
            if widget is None or sip.isdeleted(widget):
                return
            timer = getattr(widget, '_hyper_native_frame_timer', None)
            if timer is None:
                timer = QTimer(widget)
                timer.setSingleShot(True)
                widget_ref = weakref.ref(widget)

                def apply_if_alive():
                    candidate = widget_ref()
                    if candidate is not None:
                        self._apply_queued(candidate)

                timer.timeout.connect(apply_if_alive)
                widget._hyper_native_frame_timer = timer
            timer.start(0)
        except (RuntimeError, TypeError, AttributeError):
            logging.exception('Could not schedule native title-bar update.')

    def set_light(self, light):
        self.light = bool(light)
        for widget in self.app.topLevelWidgets():
            if widget.isVisible() and widget.property(self.MANAGED_PROPERTY):
                self.apply(widget)

    def apply(self, widget):
        """Apply native colors without allowing a DWM failure to escape into Qt."""
        try:
            if self.set_attribute is None or widget is None or sip.isdeleted(widget) or not widget.isWindow():
                return {}
            if widget.windowFlags() & Qt.FramelessWindowHint:
                return {}

            native_id = widget.winId()
            if native_id is None:
                # Qt can temporarily remove the native handle while a window
                # is hidden, minimized, or being recreated.
                return {}
            window_handle = int(native_id)
            palette = ('#dfcece', '#4b1c2a', '#b98997') if self.light else ('#34151f', '#fff0f4', '#6f2a3c')
            attributes = {20: int(not self.light), 35: colorref(palette[0]),
                          36: colorref(palette[1]), 34: colorref(palette[2])}
            results = {}
            for key, value in attributes.items():
                native_value = wintypes.DWORD(value)
                try:
                    results[key] = self.set_attribute(
                        window_handle, key, ctypes.byref(native_value), ctypes.sizeof(native_value)
                    )
                except (OSError, TypeError, ValueError):
                    logging.exception('Could not apply native title-bar attribute %s.', key)
                    results[key] = None
            return results
        except (RuntimeError, TypeError, ValueError, AttributeError):
            # A queued callback may run while Qt is replacing or destroying the
            # native window handle. The system title bar is a safe fallback.
            logging.exception('Could not update native title-bar colors; using the system title bar.')
            return {}

    def _apply_queued(self, widget):
        """Safety boundary for callbacks scheduled during native-handle changes."""
        try:
            self.apply(widget)
        except Exception:
            # apply() already handles expected native/Qt failures. Keep this
            # final boundary so an unexpected callback error cannot terminate
            # the frontend event loop.
            logging.exception('Unexpected native title-bar callback failure.')

    def eventFilter(self, obj, event):
        try:
            if (isinstance(obj, QWidget) and event.type() == QEvent.Show
                    and obj.property(self.MANAGED_PROPERTY)):
                if obj.isWindow() and not obj.windowFlags() & Qt.FramelessWindowHint:
                    self._schedule(obj)
        except (RuntimeError, TypeError, AttributeError):
            logging.exception('Could not schedule native title-bar update.')
        return False
