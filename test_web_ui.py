"""Exercise the real Qt/HTML bridge without running SharePoint automation."""
import json
import os
from pathlib import Path
import sys
import time
import unittest
from unittest.mock import patch

os.environ.setdefault('QT_QPA_PLATFORM', 'windows' if sys.platform == 'win32' else 'offscreen')
from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
from HyperWeb import WebWorkspace, WebLogin


class WebInterfaceTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        cls.app = QApplication.instance() or QApplication(sys.argv)
        cls.app.setApplicationName('Hyper UI Tests')
        cls.window = WebWorkspace()
        loaded = []
        cls.window.web_view.loadFinished.connect(loaded.append)
        cls.window.resize(1280, 960)
        cls.window.show()
        cls.wait_for(lambda: bool(loaded))
        if not loaded[0]:
            raise AssertionError('HTML interface failed to load')
        cls.wait_for(lambda: cls.js("typeof current !== 'undefined' && current !== null"))

    @classmethod
    def wait_for(cls, predicate, timeout=15):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            cls.app.processEvents()
            if predicate():
                return
            time.sleep(.01)
        raise AssertionError('Timed out waiting for HTML bridge')

    @classmethod
    def js(cls, code):
        results = []
        cls.window.web_view.page().runJavaScript(code, results.append)
        deadline = time.monotonic() + 5
        while not results and time.monotonic() < deadline:
            cls.app.processEvents()
            time.sleep(.005)
        if not results:
            raise AssertionError('No JavaScript response')
        return results[0]

    def test_01_labels_and_defaults(self):
        state = self.window.web_bridge.snapshot()
        self.assertEqual(len(state['groups']['manufacturers']), 36)
        for name, items in state['groups'].items():
            labels = self.js(f"Array.from(document.getElementById('{name}').children).map(x=>x.textContent)")
            self.assertEqual(labels, [item['text'] for item in items])
        self.assertFalse(state['buttons']['pause_button']['enabled'])
        self.assertTrue(all(not item['enabled'] for item in state['groups']['repair']))

    def test_02_html_selection_and_mode_signals(self):
        self.js("document.querySelector('#manufacturers input').click()")
        self.wait_for(lambda: self.window.manufacturer_checkboxes[0].isChecked())
        self.js("document.querySelector('[data-control=excel_mode_switch] [data-value=true]').click()")
        self.wait_for(lambda: len(self.window.adas_checkboxes) == 15)
        self.wait_for(lambda: self.js("document.querySelectorAll('#adas input').length") == 15)
        self.js("document.querySelector('#adas input').click()")
        self.wait_for(lambda: self.window.adas_checkboxes[0].isChecked())
        self.js("document.querySelector('[data-control=mode_switch] [data-value=true]').click()")
        self.wait_for(lambda: self.window.mode_switch.isChecked())
        self.assertFalse(self.window.excel_mode_switch.isEnabled())
        self.assertTrue(all(not cb.isChecked() for cb in self.window.adas_checkboxes))
        self.assertTrue(all(cb.isEnabled() for cb in self.window.repair_checkboxes))
        self.js("document.querySelector('[data-control=mode_switch] [data-value=false]').click()")
        self.wait_for(lambda: not self.window.mode_switch.isChecked())
        self.window.web_bridge.toggle('excel_mode_switch', False)

    def test_03_original_buttons_and_progress(self):
        self.js("document.querySelector('[data-action=select_all_years_button]').click()")
        self.wait_for(lambda: all(cb.isChecked() for cb in self.window.year_checkboxes))
        self.window.web_bridge.click('select_all_years_button')
        self.assertFalse(any(cb.isChecked() for cb in self.window.year_checkboxes))
        self.window.web_bridge.toggle('upload_mode_checkbox', True)
        self.assertTrue(self.window.upload_type_container.isEnabled())
        self.window.web_bridge.toggle('upload_mode_checkbox', False)
        self.window.current_manufacturer_progress.setValue(37)
        self.window.current_manufacturer_label.setText('Current Manufacturer: Acura')
        self.wait_for(lambda: self.js("document.querySelector('.percent').textContent") == '37%')
        self.window._set_progress_only_mode(True)
        self.window.activity_log_panel.append_output('<script>never execute</script>')
        self.wait_for(lambda: self.js("document.getElementById('setup').hidden"))
        self.assertEqual(self.js("document.getElementById('log').textContent"), '<script>never execute</script>')
        self.window._set_progress_only_mode(False)
        self.window.current_manufacturer_progress.setValue(0)
        self.window.current_manufacturer_label.setText('Current Manufacturer: None')

    def test_04_themes_and_screenshots(self):
        output = Path(os.environ.get('HYPER_SCREENSHOTS', 'work/screenshots'))
        output.mkdir(parents=True, exist_ok=True)
        for theme in ('dark', 'light'):
            self.window.resize(1180, 860)
            self.window.web_bridge.toggle('theme_toggle', theme == 'light')
            self.wait_for(lambda: self.js('document.documentElement.dataset.theme') == theme)
            self.wait_for(lambda: not self.js("document.getElementById('setup').hidden"))
            QTest.qWait(1900)  # Capture after the one-time entrance and sheen finish.
            self.assertFalse(self.js('document.documentElement.scrollWidth > innerWidth'))
            self.assertTrue(self.js("document.getElementById('progress').getBoundingClientRect().bottom <= innerHeight"))
            self.window.web_view.grab().save(str(output / f'hyper-{theme}.png'))
        self.window.resize(760, 900)
        for _ in range(30):
            self.app.processEvents()
            time.sleep(.01)
        self.assertFalse(self.js('document.documentElement.scrollWidth > innerWidth'))
        self.window.web_view.grab().save(str(output / 'hyper-compact-width.png'))
        login = WebLogin()
        login.show()
        loaded = []
        login.web_view.loadFinished.connect(loaded.append)
        self.wait_for(lambda: bool(loaded))
        self.assertTrue(loaded[0])
        for _ in range(40):
            self.app.processEvents()
            time.sleep(.01)
        login.web_view.grab().save(str(output / 'hyper-login.png'))
        login.reject()

    def test_05_native_file_and_confirmation_actions(self):
        with patch.object(QFileDialog, 'getOpenFileNames', return_value=(['C:/test/Acura.xlsx'], '')) as picker:
            self.window.web_bridge.click('select_file_button')
            picker.assert_called_once()
            self.assertEqual(self.window.excel_paths, ['C:/test/Acura.xlsx'])
        self.window.adas_checkboxes[0].setChecked(True)
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.No) as confirm:
            self.window.web_bridge.click('start_button')
            confirm.assert_called_once()
            self.assertIn('Acura.xlsx', confirm.call_args.args[2])
            self.assertIn('ADAS SI (Old acronyms)', confirm.call_args.args[2])
            self.assertFalse(self.window.is_running)
        self.window.excel_paths = []
        self.window.excel_list.clear()
        self.window.excel_list.addItem('No files selected, please select files')
        self.window.adas_checkboxes[0].setChecked(False)

    def test_06_floating_icon_lifecycle(self):
        host = self.window
        icon = host.desktop_icon
        host.is_running = True  # Simulate UI state only; no worker is launched.
        host.current_manufacturer_label.setText('Current Manufacturer: Acura')
        host.current_manufacturer_progress.setValue(37)
        host.overall_progress_bar.setValue(25)
        host.web_bridge.publish()
        self.wait_for(lambda: not self.js("document.getElementById('collapse-icon').hidden"))
        self.assertFalse(host.close())  # X hides the UI instead of closing the application.
        self.assertFalse(host.isVisible())
        self.assertTrue(host.is_running)
        self.assertTrue(icon.isVisible())
        self.assertTrue(icon.bubble.isVisible())
        self.assertIn('Acura · 37%', icon.bubble.label.text())
        self.assertEqual(icon.bubble.hide_timer.interval(), 6000)
        now = icon.last_notice
        with patch('desktop_icon.time.monotonic', return_value=now + 2), patch.object(icon.bubble, 'show_status') as notice:
            icon.refresh()
            notice.assert_not_called()
        with patch('desktop_icon.time.monotonic', return_value=now + 31), patch.object(icon.bubble, 'show_status') as notice:
            icon.refresh()
            notice.assert_called_once()
        host.pause_requested = True
        icon.refresh()
        self.assertTrue(icon.bubble.label.text().startswith('Paused'))
        output = Path(os.environ.get('HYPER_SCREENSHOTS', 'work/screenshots'))
        output.mkdir(parents=True, exist_ok=True)
        icon.grab().save(str(output / 'hyper-desktop-icon.png'))
        icon.bubble.grab().save(str(output / 'hyper-status-bubble.png'))
        area = icon.screen_area()
        icon.move_within_screen(QPoint(area.right(), area.bottom()))
        self.assertTrue(area.contains(icon.geometry()))
        QTest.mouseClick(icon, Qt.LeftButton, pos=QPoint(32, 32))
        self.assertTrue(host.isVisible())
        self.assertTrue(host.is_running)
        self.assertFalse(icon.isVisible())
        self.assertFalse(icon.bubble.isVisible())
        host.pause_requested = False
        host.showMinimized()
        self.wait_for(icon.isVisible)
        self.assertFalse(host.isVisible())
        host.is_running = False
        host.current_manufacturer_label.setText('Current Manufacturer: Complete')
        host.overall_progress_label.setText('Overall Progress: Complete')
        icon.refresh()
        self.assertIn('Complete', icon.bubble.label.text())
        self.assertTrue(icon.isVisible())
        host.restore_progress()
        host.collapse_to_icon()  # Idle collapse cannot strand the application.
        self.assertTrue(host.isVisible())
        self.assertFalse(icon.isVisible())

    @unittest.skipUnless(sys.platform == 'win32' and sys.getwindowsversion().build >= 22000,
                         'Native caption colors require Windows 11')
    def test_07_native_titlebar_theme(self):
        # Caption/text colors are set-only DWM attributes. Check native HRESULTs.
        for light in (False, True):
            self.window.web_bridge.toggle('theme_toggle', light)
            results = self.app.native_frame_theme.apply(self.window)
            self.assertEqual(results, {20: 0, 35: 0, 36: 0, 34: 0})
            self.assertEqual(self.app.native_frame_theme.light, light)
        self.window.showMaximized()
        self.app.processEvents()
        self.assertTrue(self.window.isMaximized())
        self.window.showNormal()
        self.assertFalse(self.window.isMaximized())

    @classmethod
    def tearDownClass(cls):
        cls.window.close()


if __name__ == '__main__':
    unittest.main(verbosity=2)
