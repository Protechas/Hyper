"""A small desktop companion for an existing Hyper automation window."""
import time

from PyQt5.QtCore import QPoint, QRectF, Qt, QTimer
from PyQt5.QtGui import QColor, QFont, QLinearGradient, QPainter, QPen
from PyQt5.QtWidgets import QApplication, QLabel, QMenu, QVBoxLayout, QWidget


def paint_link(painter, rect, color):
    """Draw the same interlocking link motif used by the HTML logo."""
    painter.save()
    painter.translate(rect.center())
    painter.rotate(-40)
    painter.setBrush(Qt.NoBrush)
    painter.setPen(QPen(QColor(color), rect.width() * .075, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
    unit = rect.width() / 32
    painter.drawRoundedRect(QRectF(-13 * unit, -6 * unit, 17 * unit, 12 * unit), 6 * unit, 6 * unit)
    painter.drawRoundedRect(QRectF(-4 * unit, -6 * unit, 17 * unit, 12 * unit), 6 * unit, 6 * unit)
    painter.restore()


class StatusBubble(QWidget):
    def __init__(self, owner):
        super().__init__(None, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.owner = owner
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_QuitOnClose, False)
        self.setFixedWidth(300)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 11, 14, 11)
        self.label = QLabel()
        self.label.setTextFormat(Qt.PlainText)
        self.label.setWordWrap(True)
        self.label.setFont(QFont('Segoe UI', 9))
        layout.addWidget(self.label)
        self.hide_timer = QTimer(self)
        self.hide_timer.setSingleShot(True)
        self.hide_timer.timeout.connect(self.hide)

    def show_status(self, text):
        light = self.owner.host.theme_toggle.isChecked()
        self.label.setStyleSheet('color: %s; background: transparent;' % ('#40242e' if light else '#fff0f4'))
        self.label.setText(text)
        self.adjustSize()
        self.owner.position_bubble()
        self.show()
        self.hide_timer.start(6000)

    def paintEvent(self, event):
        light = self.owner.host.theme_toggle.isChecked()
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QPen(QColor('#d9a0af' if light else '#9f5267'), 1))
        painter.setBrush(QColor('#e4d8d5' if light else '#321c26'))
        painter.drawRoundedRect(QRectF(.5, .5, self.width() - 1, self.height() - 1), 10, 10)


class DesktopIcon(QWidget):
    STATUS_INTERVAL = 30  # seconds between periodic updates
    PHASE_INTERVAL = 15

    def __init__(self, host):
        super().__init__(None, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.host = host
        self.setWindowTitle('Hyper — automation status')
        self.setWindowIcon(host.windowIcon())
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WA_QuitOnClose, False)
        self.setFixedSize(64, 64)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip('Hyper — click to show progress; drag to move; right-click for controls')
        self.bubble = StatusBubble(self)
        self.tick = QTimer(self)
        self.tick.setInterval(1000)
        self.tick.timeout.connect(self.refresh)
        self.last_notice = 0
        self.last_phase = None
        self.was_running = False
        self.was_paused = False
        self.initial_position = False
        self.press_point = None
        self.dragged = False

    def present(self):
        if not self.initial_position:
            screen = self.host.screen() or QApplication.primaryScreen()
            area = screen.availableGeometry()
            self.move(area.right() - self.width() - 24, area.bottom() - self.height() - 24)
            self.initial_position = True
        self.show()
        self.tick.start()
        self.refresh(force=True)

    def dismiss(self):
        self.tick.stop()
        self.bubble.hide_timer.stop()
        self.bubble.hide()
        self.hide()

    def status_text(self):
        prefix = 'Paused\n' if self.host.pause_requested else ''
        bar = self.host.current_manufacturer_progress
        percent = f' · {max(0, bar.value()) * 100 // bar.maximum()}%' if bar.maximum() > 0 else ''
        return prefix + '\n'.join((self.host.current_manufacturer_label.text() + percent,
                                   self.host.manufacturer_hyperlink_label.text(),
                                   self.host.overall_progress_label.text()))

    def refresh(self, force=False):
        running = self.host.automation_active()
        paused = bool(self.host.pause_requested)
        phase = self.host.current_manufacturer_label.text()
        now = time.monotonic()
        urgent = self.was_running != running or self.was_paused != paused
        phase_changed = phase != self.last_phase and now - self.last_notice >= self.PHASE_INTERVAL
        periodic = running and now - self.last_notice >= self.STATUS_INTERVAL
        if self.isVisible() and (force or urgent or phase_changed or periodic):
            self.bubble.show_status(self.status_text())
            self.last_notice, self.last_phase = now, phase
        self.was_running, self.was_paused = running, paused
        self.update()

    def screen_area(self):
        screen = QApplication.screenAt(self.geometry().center()) or QApplication.primaryScreen()
        return screen.availableGeometry()

    def move_within_screen(self, point):
        center = point + QPoint(self.width() // 2, self.height() // 2)
        # During a drag the proposed center can briefly fall outside every
        # screen even though the proposed top-left point is still on the
        # current screen. Try both before falling back to the primary display.
        screen = (QApplication.screenAt(point) or QApplication.screenAt(center)
                  or self.screen() or QApplication.primaryScreen())
        area = screen.availableGeometry()
        self.move(max(area.left(), min(point.x(), area.right() - self.width() + 1)),
                  max(area.top(), min(point.y(), area.bottom() - self.height() + 1)))
        self.position_bubble()

    def position_bubble(self):
        area = self.screen_area()
        x = max(area.left(), min(self.x() + self.width() // 2 - self.bubble.width() // 2,
                                area.right() - self.bubble.width() + 1))
        y = self.y() - self.bubble.height() - 10
        if y < area.top():
            y = self.y() + self.height() + 10
        y = max(area.top(), min(y, area.bottom() - self.bubble.height() + 1))
        self.bubble.move(x, y)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.press_point = event.globalPos()
            self.press_origin = self.pos()
            self.dragged = False

    def mouseMoveEvent(self, event):
        if self.press_point is not None and event.buttons() & Qt.LeftButton:
            delta = event.globalPos() - self.press_point
            if delta.manhattanLength() > 5:
                self.dragged = True
            if self.dragged:
                self.move_within_screen(self.press_origin + delta)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.press_point is not None:
            self.press_point = None
            if not self.dragged:
                self.host.restore_progress()

    def enterEvent(self, event):
        self.refresh(force=True)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        restore = menu.addAction('Show progress')
        pause = menu.addAction(self.host.pause_button.text())
        pause.setEnabled(self.host.pause_button.isEnabled())
        stop = menu.addAction(self.host.start_button.text())
        stop.setEnabled(self.host.automation_active())
        chosen = menu.exec_(event.globalPos())
        if chosen is not None:
            self.host.restore_progress()
            if chosen is pause:
                self.host.pause_button.click()
            elif chosen is stop:
                self.host.start_button.click()

    def closeEvent(self, event):
        event.ignore()
        self.host.restore_progress()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        gradient = QLinearGradient(7, 5, 56, 60)
        gradient.setColorAt(0, QColor('#df5472'))
        gradient.setColorAt(1, QColor('#9f1838'))
        painter.setPen(QPen(QColor('#ffacbd'), 1))
        painter.setBrush(gradient)
        painter.drawRoundedRect(QRectF(6, 6, 52, 52), 17, 17)
        painter.setBrush(Qt.NoBrush)
        painter.setPen(QPen(QColor(255, 255, 255, 48), 3))
        painter.drawArc(QRectF(2, 2, 60, 60), 90 * 16, -360 * 16)
        bar = self.host.overall_progress_bar
        percent = max(0, bar.value()) / bar.maximum() if bar.maximum() > 0 else 0
        painter.setPen(QPen(QColor('#ffc4d1'), 3, Qt.SolidLine, Qt.RoundCap))
        painter.drawArc(QRectF(2, 2, 60, 60), 90 * 16, -int(min(1, percent) * 360 * 16))
        paint_link(painter, QRectF(14, 14, 36, 36), '#ffffff')
        painter.setPen(QPen(QColor('#1d1218'), 2))
        painter.setBrush(QColor('#f2b969' if self.host.pause_requested else '#6ce1b6'))
        painter.drawEllipse(QRectF(48, 48, 10, 10))


def link_icon():
    from PyQt5.QtGui import QIcon, QPixmap
    image = QPixmap(64, 64)
    image.fill(Qt.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setPen(Qt.NoPen)
    painter.setBrush(QColor('#bf284b'))
    painter.drawRoundedRect(QRectF(2, 2, 60, 60), 18, 18)
    paint_link(painter, QRectF(10, 10, 44, 44), '#ffffff')
    painter.end()
    return QIcon(image)
