from PyQt6.QtWidgets import QGraphicsRectItem
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QBrush, QColor, QFont


class StageItem(QGraphicsRectItem):
    """Класс для создания item'ов"""
    def __init__(self, x, y, width, height, voltage, number, params=None):
        super().__init__(x, y, width, height)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemIsSelectable, True)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)

        self.voltage = voltage
        self.number = number
        self.params = params or {}  # Словарь параметров
        # self.setBrush(QBrush(QColor(100, 62, 255)))  # Был синий
        self.setBrush(QBrush(QColor(255, 160, 47)))

        self.font = QFont("Arial", 12)

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)

        # Основной текст
        painter.setFont(self.font)
        main_text = f"{self.number}. {self.voltage}"
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignCenter, main_text)


    def itemChange(self, change, value):
        if change == QGraphicsRectItem.GraphicsItemChange.ItemPositionChange:
            # Обновляем размеры сцены при перемещении
            # if self.scene():
            #     self.scene().views()[0].updateSceneRect(self.scene().itemsBoundingRect())
            return QPointF(self.x(), value.y())
        return super().itemChange(change, value)