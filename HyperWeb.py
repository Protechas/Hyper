"""HTML presentation for Hyper. The original widgets remain the source of truth."""
import json
import logging
import os
from pathlib import Path
import sys
import tempfile

# The bundled interface is entirely local and does not benefit from Chromium's
# shared GPU/disk caches. Disabling those caches prevents a stale GPUCache lock
# from aborting startup before the first window is shown.
_webengine_flags = os.environ.get('QTWEBENGINE_CHROMIUM_FLAGS', '').split()
for _flag in ('--disable-gpu-shader-disk-cache', '--disable-gpu-program-cache', '--disk-cache-size=0'):
    if _flag not in _webengine_flags:
        _webengine_flags.append(_flag)
os.environ['QTWEBENGINE_CHROMIUM_FLAGS'] = ' '.join(_webengine_flags)

from PyQt5.QtCore import QEvent, QObject, QTimer, QUrl, Qt, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.QtWebChannel import QWebChannel
from PyQt5.QtWebEngineWidgets import QWebEnginePage, QWebEngineProfile, QWebEngineView

from Hyper import LoginDialog, SeleniumAutomationApp
from desktop_icon import DesktopIcon, link_icon
from window_theme import NativeFrameTheme

ROOT = Path(__file__).resolve().parent
_WEBENGINE_RUNTIME = Path(tempfile.mkdtemp(prefix='hyper-webengine-'))
_WEBENGINE_CONFIGURED = False


def configure_webengine_profile():
    """Give this Hyper process an isolated, disposable WebEngine profile."""
    global _WEBENGINE_CONFIGURED
    if _WEBENGINE_CONFIGURED:
        return QWebEngineProfile.defaultProfile()

    profile = QWebEngineProfile.defaultProfile()
    profile.setHttpCacheType(QWebEngineProfile.MemoryHttpCache)
    profile.setPersistentCookiesPolicy(QWebEngineProfile.NoPersistentCookies)
    profile.setCachePath(str(_WEBENGINE_RUNTIME / 'cache'))
    profile.setPersistentStoragePath(str(_WEBENGINE_RUNTIME / 'storage'))
    _WEBENGINE_CONFIGURED = True
    return profile


class LocalPage(QWebEnginePage):
    def javaScriptConsoleMessage(self, level, message, line, source):
        import logging
        logging.warning('HTML interface: %s (line %s)', message, line)

    def acceptNavigationRequest(self, url, navigation_type, is_main_frame):
        # Only the bundled interface can access the desktop bridge.
        return url.isLocalFile() and Path(url.toLocalFile()).resolve() == ROOT / 'ui' / 'index.html'


def embed(host, bridge, login=False):
    app = QApplication.instance()
    configure_webengine_profile()
    if not hasattr(app, 'native_frame_theme'):
        app.native_frame_theme = NativeFrameTheme(app)
    app.native_frame_theme.watch(host)
    view = QWebEngineView(host)
    view.setPage(LocalPage(view))
    channel = QWebChannel(view.page())
    channel.registerObject('hyper', bridge)
    view.page().setWebChannel(channel)
    view.setContextMenuPolicy(Qt.NoContextMenu)
    host.web_view, host.web_channel, host.web_bridge = view, channel, bridge
    url = QUrl.fromLocalFile(str(ROOT / 'ui' / 'index.html'))
    if login:
        url.setQuery('login')
    view.setUrl(url)
    host.layout().addWidget(view)
    return view


class LoginBridge(QObject):
    def __init__(self, host):
        super().__init__(host)
        self.host = host

    @pyqtSlot(str, str)
    def signIn(self, username, password):
        self.host.user_edit.setText(username)
        self.host.pass_edit.setText(password)
        self.host.try_login()

    @pyqtSlot()
    def cancel(self):
        self.host.reject()

    @pyqtSlot(bool)
    def theme(self, light):
        QApplication.instance().native_frame_theme.set_light(light)


class WebLogin(LoginDialog):
    def __init__(self):
        super().__init__(max_attempts=5)
        layout = self.layout()
        while layout.count():
            layout.takeAt(0)
        for widget in self.findChildren(QObject):
            if hasattr(widget, 'hide'):
                widget.hide()
        layout.setContentsMargins(0, 0, 0, 0)
        self.resize(860, 560)
        self.setWindowIcon(link_icon())
        self.setMinimumSize(480, 500)
        embed(self, LoginBridge(self), login=True)


class WorkspaceBridge(QObject):
    changed = pyqtSignal(str)
    GROUPS = {'manufacturers': 'manufacturer_checkboxes', 'years': 'year_checkboxes',
              'adas': 'adas_checkboxes', 'repair': 'repair_checkboxes'}
    CONTROLS = ('mode_switch', 'excel_mode_switch', 'cleanup_checkbox',
                'upload_mode_checkbox', 'always_on_top_checkbox', 'theme_toggle')
    BUTTONS = ('select_file_button', 'select_all_manufacturers_button',
               'select_all_years_button', 'select_all_adas_button',
               'select_all_repair_button', 'start_button', 'pause_button', 'view_log_button')

    def __init__(self, host):
        super().__init__(host)
        self.host = host
        self.previous = ''
        self.file_picker_scheduled = False
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.publish)
        self.timer.start(150)

    def snapshot(self):
        host = self.host
        def checkbox(widget):
            return {'text': widget.text(), 'checked': widget.isChecked(), 'enabled': widget.isEnabled()}
        state = {
            'groups': {key: [checkbox(w) for w in getattr(host, attr)] for key, attr in self.GROUPS.items()},
            'controls': {key: checkbox(getattr(host, key)) for key in self.CONTROLS},
            'buttons': {key: {'text': getattr(host, key).text(), 'enabled': getattr(host, key).isEnabled()}
                        for key in self.BUTTONS},
            'adasTitle': host.adas_label.text(),
            'files': [host.excel_list.item(i).text() for i in range(host.excel_list.count())],
            'compact': host._progress_only_mode,
            'running': host.automation_active(),
            'opacity': host.opacity_slider.value(),
            'log': host.activity_log_panel.terminal_output.toPlainText(),
            'progress': [],
        }
        for label, bar in ((host.current_manufacturer_label, host.current_manufacturer_progress),
                           (host.manufacturer_hyperlink_label, host.manufacturer_hyperlink_bar),
                           (host.overall_progress_label, host.overall_progress_bar)):
            maximum = bar.maximum()
            state['progress'].append({'text': label.text(), 'value': max(0, bar.value()),
                                      'maximum': maximum, 'stopped': bool(bar.property('stopped'))})
        return state

    @pyqtSlot(result=str)
    def getState(self):
        return json.dumps(self.snapshot())

    def publish(self):
        payload = self.getState()
        if payload != self.previous:
            self.previous = payload
            self.changed.emit(payload)

    @pyqtSlot(str, int, bool)
    def select(self, group, index, checked):
        attr = self.GROUPS.get(group)
        if attr:
            widgets = getattr(self.host, attr)
            if 0 <= index < len(widgets) and widgets[index].isEnabled():
                widgets[index].setChecked(checked)
        self.publish()

    @pyqtSlot(str, bool)
    def toggle(self, name, checked):
        if name in self.CONTROLS:
            widget = getattr(self.host, name)
            if widget.isEnabled():
                widget.setChecked(checked)
                if name == 'theme_toggle':
                    self.host.toggle_theme()
        self.publish()

    @pyqtSlot(str)
    def click(self, name):
        if name == 'select_file_button':
            # A native Windows file dialog runs its own modal message loop.
            # Launch it only after this WebChannel callback has returned;
            # otherwise dismissing the dialog can re-enter Qt WebEngine and
            # terminate the application process.
            if not self.file_picker_scheduled:
                self.file_picker_scheduled = True
                QTimer.singleShot(0, self._open_file_picker)
            return

        self._click_now(name)

    def _open_file_picker(self):
        self.file_picker_scheduled = False
        self._click_now('select_file_button')

    def _click_now(self, name):
        try:
            if name in self.BUTTONS:
                widget = getattr(self.host, name)
                if widget.isEnabled():
                    widget.click()
            self.publish()
        except Exception:
            # Never allow a native/Qt callback failure to escape through the
            # WebChannel and terminate the frontend process.
            logging.exception('Frontend button callback failed: %s', name)

    @pyqtSlot(int)
    def transparency(self, value):
        self.host.opacity_slider.setValue(value)
        self.publish()

    @pyqtSlot()
    def collapse(self):
        self.host.collapse_to_icon()


class WebWorkspace(SeleniumAutomationApp):
    STARTUP_WIDTH = 1280
    # The compact HTML workspace naturally fits at this height. Keeping the
    # window taller only exposes unused background below the footer.
    STARTUP_HEIGHT = 780
    STARTUP_MARGIN = 24

    def __init__(self):
        super().__init__()
        # Retain the entire original widget tree and all of its signal connections.
        self.workspace_scroll.hide()
        embed(self, WorkspaceBridge(self))
        self.setWindowIcon(link_icon())
        self.desktop_icon = DesktopIcon(self)
        self.apply_startup_geometry()

    def apply_startup_geometry(self, screen=None):
        """Center a roomy normal window without filling the entire display."""
        app = QApplication.instance()
        screen = screen or (app.primaryScreen() if app is not None else None)
        if screen is None:
            self.resize(self.STARTUP_WIDTH, self.STARTUP_HEIGHT)
            return

        available = screen.availableGeometry()
        horizontal_room = max(1, available.width() - (self.STARTUP_MARGIN * 2))
        vertical_room = max(1, available.height() - (self.STARTUP_MARGIN * 2))
        window_width = min(self.STARTUP_WIDTH, horizontal_room)
        window_height = min(self.STARTUP_HEIGHT, vertical_room)
        self.resize(window_width, window_height)
        self.move(
            available.x() + (available.width() - window_width) // 2,
            available.y() + (available.height() - window_height) // 2,
        )

    def toggle_theme(self):
        super().toggle_theme()
        app = QApplication.instance()
        if hasattr(app, 'native_frame_theme'):
            app.native_frame_theme.set_light(self.theme_toggle.isChecked())

    def automation_active(self):
        return bool(getattr(self, 'is_running', False) or getattr(self, '_progress_only_mode', False))

    def collapse_to_icon(self):
        if not self.automation_active():
            return
        self.desktop_icon.present()
        self.hide()

    def restore_progress(self):
        self.showNormal()
        self.desktop_icon.dismiss()
        self.raise_()
        self.activateWindow()

    def closeEvent(self, event):
        if self.automation_active():
            event.ignore()
            self.collapse_to_icon()
        else:
            if hasattr(self, 'desktop_icon'):
                self.desktop_icon.dismiss()
            super().closeEvent(event)

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QEvent.WindowStateChange and self.isMinimized() and self.automation_active():
            QTimer.singleShot(0, self.collapse_to_icon)


def main():
    os.chdir(ROOT)  # Existing worker commands use repository-relative paths.
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName('Hyper')
    app.setWindowIcon(link_icon())
    app.setStyle('Fusion')
    app.setFont(QFont('Segoe UI', 10))
    login = WebLogin()
    if login.exec_() != QDialog.Accepted:
        return 1
    window = WebWorkspace()
    # Be explicit so a previous maximized/fullscreen state can never leak into
    # the newly launched workspace.
    window.setWindowState(Qt.WindowNoState)
    window.showNormal()
    return app.exec_()


if __name__ == '__main__':
    sys.exit(main())
