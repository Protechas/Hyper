"""Interactive sample of the desktop icon; never imports automation code."""
import sys

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QCheckBox, QLabel, QProgressBar, QPushButton, QVBoxLayout, QWidget

from desktop_icon import DesktopIcon, link_icon


class Preview(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Hyper — floating icon preview')
        self.setWindowIcon(link_icon())
        self.resize(440, 330)
        self.setStyleSheet('QWidget {background:#1e1b20; color:#faf2f4;} QPushButton {background:#b91f40; border:0; border-radius:7px; padding:10px;} QProgressBar {border:1px solid #4d2d3a; border-radius:5px; text-align:center;} QProgressBar::chunk {background:#dc5a78;}')
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(12)
        intro = QLabel('Floating icon preview\n\nSimulated progress only. Drag the icon to move it;\nhover for status; click to restore this window.')
        layout.addWidget(intro)
        self.pause_requested = False
        self.theme_toggle = QCheckBox('Light theme')
        layout.addWidget(self.theme_toggle)
        self.current_manufacturer_label = QLabel('Preview — Current Manufacturer: Honda')
        self.manufacturer_hyperlink_label = QLabel('Manufacturer Hyperlinks: 42 / 120')
        self.overall_progress_label = QLabel('Overall Progress: 1 / 3 manufacturers')
        self.current_manufacturer_progress = QProgressBar()
        self.current_manufacturer_progress.setValue(35)
        self.overall_progress_bar = QProgressBar()
        self.overall_progress_bar.setValue(33)
        for widget in (self.current_manufacturer_label, self.manufacturer_hyperlink_label,
                       self.current_manufacturer_progress, self.overall_progress_label):
            layout.addWidget(widget)
        self.pause_button = QPushButton('Pause preview')
        self.pause_button.clicked.connect(self.toggle_pause)
        layout.addWidget(self.pause_button)
        collapse = QPushButton('Collapse to icon')
        collapse.clicked.connect(self.collapse)
        layout.addWidget(collapse)
        self.start_button = QPushButton('Close preview')
        self.start_button.clicked.connect(QApplication.instance().quit)
        layout.addWidget(self.start_button)
        self.desktop_icon = DesktopIcon(self)
        self.desktop_icon.setWindowTitle('Hyper — sample floating icon')
        self.demo_timer = QTimer(self)
        self.demo_timer.timeout.connect(self.advance)
        self.demo_timer.start(2000)
        QApplication.instance().aboutToQuit.connect(self.desktop_icon.dismiss)

    def automation_active(self):
        return True

    def toggle_pause(self):
        self.pause_requested = not self.pause_requested
        self.pause_button.setText('Resume preview' if self.pause_requested else 'Pause preview')
        self.desktop_icon.refresh(force=True)

    def advance(self):
        if not self.pause_requested:
            value = (self.current_manufacturer_progress.value() + 3) % 101
            self.current_manufacturer_progress.setValue(value)
            self.manufacturer_hyperlink_label.setText(f'Manufacturer Hyperlinks: {value * 120 // 100} / 120')

    def collapse(self):
        self.desktop_icon.present()
        self.hide()

    def restore_progress(self):
        self.showNormal()
        self.desktop_icon.dismiss()
        self.raise_()
        self.activateWindow()

    def closeEvent(self, event):
        self.desktop_icon.dismiss()
        QApplication.instance().quit()
        event.accept()


if __name__ == '__main__':
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    app = QApplication(sys.argv)
    app.setApplicationName('Hyper Floating Preview')
    app.setQuitOnLastWindowClosed(False)
    app.setFont(QFont('Segoe UI', 10))
    preview = Preview()
    preview.collapse()
    sys.exit(app.exec_())
