from PyQt6.QtWidgets import QGraphicsRectItem
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QBrush, QColor, QFont


class StageItem(QGraphicsRectItem):
    def __init__(self, x, y, width, height, stage_type, number, params=None):
        super().__init__(x, y, width, height)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemIsSelectable, True)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)

        self.stage_type = stage_type
        self.number = number
        self.params = params or {}  # Словарь параметров
        # self.setBrush(QBrush(QColor(100, 62, 255)))  # Был синий
        self.setBrush(QBrush(QColor(255, 160, 47)))

        self.font = QFont("Arial", 9)
        self.small_font = QFont("Arial", 8)

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)

        # Основной текст
        painter.setFont(self.font)
        main_text = f"{self.number}. {self.stage_type}"
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop, main_text)

        # Параметры
        if self.params:
            painter.setFont(self.small_font)
            params_text = ""
            if self.stage_type == "Relay":
                params_text = f"{self.params.get('relay_type', '')} {'ON' if self.params.get('relay_state', False) else 'OFF'}"
            elif self.stage_type == "Resistance":
                params_text = f"{self.params.get('resistance_value', '')}"

            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom, params_text)

    def itemChange(self, change, value):
        if change == QGraphicsRectItem.GraphicsItemChange.ItemPositionChange:
            # Обновляем размеры сцены при перемещении
            # if self.scene():
            #     self.scene().views()[0].updateSceneRect(self.scene().itemsBoundingRect())
            return QPointF(value.x(), self.y())
        return super().itemChange(change, value)