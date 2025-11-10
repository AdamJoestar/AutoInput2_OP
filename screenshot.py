import tempfile
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QPushButton
from PyQt5.QtCore import Qt, QPoint, QRect, QSize
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor, QImage
from PIL import ImageGrab, Image


class ScreenshotSelector(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setWindowModality(Qt.ApplicationModal)
        self.showFullScreen()
        self.setCursor(Qt.CrossCursor)

        # Grab screenshot with PIL
        self.screenshot_pil = ImageGrab.grab()
        # Convert to QPixmap
        self.screenshot_qimage = QImage(self.screenshot_pil.tobytes(), self.screenshot_pil.width, self.screenshot_pil.height, QImage.Format_RGB888)
        self.screenshot_pixmap = QPixmap.fromImage(self.screenshot_qimage)

        # Selection rect
        self.selection_rect = QRect()
        self.start_point = QPoint()
        self.selecting = False

        # Add OK and Cancel buttons at the top
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(10, 10, 10, 10)
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.addStretch()

        main_layout = QVBoxLayout()
        main_layout.addLayout(button_layout)
        main_layout.addStretch()
        self.setLayout(main_layout)

        self.raise_()
        self.activateWindow()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_point = event.pos()
            self.selection_rect = QRect(self.start_point, QSize())
            self.selecting = True
            self.update()

    def mouseMoveEvent(self, event):
        if self.selecting:
            self.selection_rect = QRect(self.start_point, event.pos()).normalized()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.selecting = False

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self.screenshot_pixmap)
        if self.selection_rect.isValid():
            pen = QPen(QColor(255, 0, 0), 2, Qt.SolidLine)
            painter.setPen(pen)
            painter.drawRect(self.selection_rect)

    def get_selected_image(self):
        if self.selection_rect.isValid():
            # Crop the PIL image
            left = self.selection_rect.left()
            top = self.selection_rect.top()
            width = self.selection_rect.width()
            height = self.selection_rect.height()
            cropped = self.screenshot_pil.crop((left, top, left + width, top + height))
            return cropped
        else:
            return self.screenshot_pil  # Return full screenshot if no selection
