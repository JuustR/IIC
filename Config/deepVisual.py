import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QGraphicsView,
                             QGraphicsScene, QGraphicsRectItem, QVBoxLayout,
                             QWidget, QPushButton, QLabel, QHBoxLayout,
                             QInputDialog, QComboBox, QCheckBox, QFormLayout,
                             QDialogButtonBox, QStackedWidget, QDialog)
from PyQt6.QtCore import Qt, QPointF, QTimer
from PyQt6.QtGui import QBrush, QColor, QFont, QPainter


class StageItem(QGraphicsRectItem):
    def __init__(self, x, y, width, height, stage_type, number, params=None):
        super().__init__(x, y, width, height)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemIsSelectable, True)
        self.setFlag(QGraphicsRectItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)

        self.stage_type = stage_type
        self.number = number
        self.params = params or {}  # Словарь параметров
        self.setBrush(QBrush(QColor(100, 150, 255)))

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


class CustomGraphicsScene(QGraphicsScene):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent

    def mouseReleaseEvent(self, event):
        """Обработчик отпускания мыши"""
        super().mouseReleaseEvent(event)

        if event.button() == Qt.MouseButton.LeftButton:
            # Отложенный вызов авто-выравнивания
            if self.main_window:
                QTimer.singleShot(100, self.main_window.safe_auto_arrange)


class StageConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройка этапа")

        # Основные элементы
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Relay", "SeebeckUp", "SeebeckDown",
                                  "Resistance", "Heater", "Pause"])

        # Создаем stacked widget для разных конфигураций
        self.stacked_config = QStackedWidget()

        # Конфигурация для Relay
        self.relay_widget = QWidget()
        relay_layout = QFormLayout()
        self.relay_type_combo = QComboBox()
        self.relay_type_combo.addItems(["Heater", "Current", "Sample"])
        self.relay_state_check = QCheckBox("Включено")
        relay_layout.addRow("Тип реле:", self.relay_type_combo)
        relay_layout.addRow("Состояние:", self.relay_state_check)
        self.relay_widget.setLayout(relay_layout)

        # Конфигурация для Resistance (пример)
        self.resistance_widget = QWidget()
        resistance_layout = QFormLayout()
        self.resistance_value = QComboBox()
        self.resistance_value.addItems(["10k", "100k", "1M"])
        resistance_layout.addRow("Значение:", self.resistance_value)
        self.resistance_widget.setLayout(resistance_layout)

        # Пустой виджет для этапов без параметров
        self.empty_widget = QWidget()

        # Добавляем виджеты в stacked
        self.stacked_config.addWidget(self.relay_widget)  # index 0
        self.stacked_config.addWidget(self.resistance_widget)  # index 1
        self.stacked_config.addWidget(self.empty_widget)  # index 2

        # Кнопки диалога
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                                           QDialogButtonBox.StandardButton.Cancel)

        # Основной layout
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Тип этапа:"))
        layout.addWidget(self.type_combo)
        layout.addWidget(self.stacked_config)
        layout.addWidget(self.button_box)
        self.setLayout(layout)

        # Связываем сигналы
        self.type_combo.currentTextChanged.connect(self.update_config_view)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Начальное состояние
        self.update_config_view(self.type_combo.currentText())

    def update_config_view(self, stage_type):
        """Динамически меняем виджет конфигурации"""
        if stage_type == "Relay":
            self.stacked_config.setCurrentIndex(0)
        elif stage_type == "Resistance":
            self.stacked_config.setCurrentIndex(1)
        else:
            self.stacked_config.setCurrentIndex(2)

    def get_params(self):
        """Возвращает параметры в виде словаря"""
        params = {}
        current_type = self.type_combo.currentText()

        if current_type == "Relay":
            params = {
                'relay_type': self.relay_type_combo.currentText(),
                'relay_state': self.relay_state_check.isChecked()
            }
        elif current_type == "Resistance":
            params = {
                'resistance_value': "" + self.resistance_value.currentText()
            }

        return current_type, params


class StagesEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Редактор последовательности этапов")
        self.setGeometry(100, 100, 1200, 600)

        self.scene = CustomGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.TextAntialiasing)

        self.stage_items = []
        self.current_number = len(self.stage_items) + 1

        self.measurement_types = ["Relay", "SeebeckUp", "SeebeckDown", "Resistance", "Heater"]

        self.init_ui()

    def safe_auto_arrange(self):
        """Защищённый вызов авто-выравнивания"""
        try:
            if all(hasattr(self, attr) for attr in ['scene', 'stage_items']):
                self.auto_arrange()
        except Exception as e:
            print(f"Ошибка выравнивания: {str(e)}")

    def init_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

        # Назначаем objectName для дополнительной стилизации
        central_widget.setObjectName("mainWidget")

        # Панель управления
        control_panel = QWidget()
        control_panel.setObjectName("controlPanel")
        control_layout = QHBoxLayout()

        self.add_button = QPushButton("Добавить этап")
        self.add_button.setObjectName("addButton")

        self.remove_button = QPushButton("Удалить выбранный")
        self.remove_button.setObjectName("removeButton")

        self.edit_button = QPushButton("Редактировать выбранный")
        self.edit_button.setObjectName("editButton")

        self.auto_arrange_button = QPushButton("Авто-расположение")
        self.auto_arrange_button.setObjectName("autoArrangeButton")

        # Добавляем кнопки в layout
        for btn in [self.add_button, self.remove_button, self.edit_button, self.auto_arrange_button]:
            control_layout.addWidget(btn)

        control_panel.setLayout(control_layout)

        # Информационная метка
        self.info_label = QLabel("Перетаскивайте этапы для изменения порядка")
        self.info_label.setObjectName("infoLabel")

        # Настройка скроллбаров
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # Добавляем элементы в основной layout
        layout.addWidget(control_panel)
        layout.addWidget(self.info_label)
        layout.addWidget(self.view)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Подключаем сигналы
        self.add_button.clicked.connect(self.add_stage)
        self.remove_button.clicked.connect(self.remove_selected)
        self.edit_button.clicked.connect(self.edit_selected)
        self.auto_arrange_button.clicked.connect(self.auto_arrange)

    def add_stage(self):
        dialog = StageConfigDialog(self)

        if dialog.exec() == QInputDialog.DialogCode.Accepted:
            stage_type, params = dialog.get_params()

            # Рассчитываем позицию для нового этапа
            if self.stage_items:
                # Находим самый правый этап
                rightmost_item = max(self.stage_items, key=lambda item: item.x())
                new_x = rightmost_item.x() + rightmost_item.rect().width() + 30  # Ширина + отступ 30
            else:
                new_x = 50  # Начальная позиция если нет этапов

            # Определяем высоту в зависимости от наличия параметров
            height = 80 if params else 50

            stage = StageItem(0, 50, 150, height, stage_type, self.current_number, params)

            self.scene.addItem(stage)
            self.stage_items.append(stage)
            self.current_number += 1

            # Устанавливаем позицию после добавления на сцену
            stage.setPos(new_x, 50)
            self.renumber_stages()
            self.auto_arrange()  # Добавил, чтоб не было фантомного смещения по y

            # Прокручиваем к новому элементу
            self.view.ensureVisible(stage)

    def remove_selected(self):
        """Безопасное удаление выбранных элементов"""
        try:
            items_to_remove = []
            for item in self.scene.selectedItems():
                if isinstance(item, StageItem) and item in self.stage_items:
                    items_to_remove.append(item)

            for item in items_to_remove:
                try:
                    self.scene.removeItem(item)
                    self.stage_items.remove(item)
                except:
                    continue

            self.renumber_stages()
            QTimer.singleShot(100, self.safe_auto_arrange)
        except Exception as e:
            print(f"Ошибка при удалении: {str(e)}")

    def edit_selected(self):
        selected_items = self.scene.selectedItems()
        if not selected_items:
            return

        item = selected_items[0]
        if not isinstance(item, StageItem):
            return

        # Создаем диалог редактирования
        dialog = StageConfigDialog(self)
        dialog.setWindowTitle("Редактирование этапа")

        # Устанавливаем текущие значения
        dialog.type_combo.setCurrentText(item.stage_type)

        # Загружаем параметры в соответствующие виджеты
        if item.stage_type == "Relay":
            dialog.relay_type_combo.setCurrentText(item.params.get('relay_type', 'Heater'))
            dialog.relay_state_check.setChecked(item.params.get('relay_state', False))
        elif item.stage_type == "Resistance":
            dialog.resistance_value.setCurrentText(item.params.get('resistance_value', '10k'))

        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Получаем обновленные данные
            new_type, new_params = dialog.get_params()

            # Обновляем элемент
            item.stage_type = new_type
            item.params = new_params

            # Корректируем высоту в зависимости от наличия параметров
            new_height = 80 if new_params else 50
            item.setRect(item.rect().x(), item.rect().y(), item.rect().width(), new_height)

            # Обновляем отображение
            item.update()

            # Пересчитываем layout если нужно
            self.auto_arrange()

    def auto_arrange(self):
        """Безопасное автоматическое выравнивание этапов"""
        try:
            if not hasattr(self, 'stage_items') or not self.stage_items:
                return

            # Создаем временный список для безопасного доступа
            valid_items = [item for item in self.stage_items
                           if isinstance(item, StageItem) and item.scene() is not None]

            if not valid_items:
                return

            # Сортируем по текущей позиции X
            sorted_items = sorted(valid_items, key=lambda x: x.x())

            # Выравниваем с отступами
            x_pos = 50
            y_pos = 50
            spacing = 30

            for item in sorted_items:
                try:
                    if item and item.scene():  # Дополнительная проверка
                        item.setPos(x_pos, y_pos)
                        x_pos += item.rect().width() + spacing
                except:
                    continue  # Пропускаем проблемные элементы

            # Обновляем размер сцены
            try:
                self.scene.setSceneRect(0, 0, x_pos + 50, 300)

                # Управление скроллбаром
                show_scroll = len(valid_items) > 1
                self.view.setHorizontalScrollBarPolicy(
                    Qt.ScrollBarPolicy.ScrollBarAsNeeded if show_scroll
                    else Qt.ScrollBarPolicy.ScrollBarAlwaysOff
                )
            except:
                pass

            self.renumber_stages()

        except Exception as e:
            print(f"Ошибка в auto_arrange: {str(e)}")

    def renumber_stages(self):
        """Безопасная перенумерация этапов"""
        try:
            valid_items = [item for item in self.stage_items
                           if isinstance(item, StageItem) and item.scene() is not None]

            for i, item in enumerate(sorted(valid_items, key=lambda x: x.x()), 1):
                try:
                    item.number = i
                    item.update()
                except:
                    continue
        except:
            pass

    def closeEvent(self, event):
        # Выводим итоговый порядок этапов при закрытии
        sorted_items = sorted(self.stage_items, key=lambda x: x.x())
        print("\nИтоговый порядок этапов:")
        for item in sorted_items:
            print(f"{item.number}. {item.text}")

            # Переписать если нужно
            # if stage.stage_type == "Relay":
            #     relay_type = stage.params.get('relay_type')
            #     relay_state = stage.params.get('relay_state', False)
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Дополнительные глобальные стили
    app.setStyleSheet("""
            QScrollBar:vertical {
                background: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0.0 #121212, stop: 0.2 #282828, stop: 1 #484848);
                width: 12px;
                margin: 0px;
                border: none;
            }
            QScrollBar:horizontal {
                background: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0.0 #121212, stop: 0.2 #282828, stop: 1 #484848);
                height: 12px;
                margin: 0px;
                border: none;
            }
            QScrollBar::handle:vertical {
                background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0, stop: 0 #ffa02f, stop: 0.5 #d7801a, stop: 1 #ffa02f);
                min-height: 20px;
                border-radius: 6px;
                margin: 2px;
            }
            QScrollBar::handle:horizontal {
                background: QLinearGradient( x1: 0, y1: 0, x2: 1, y2: 0, stop: 0 #ffa02f, stop: 0.5 #d7801a, stop: 1 #ffa02f);
                min-width: 20px;
                border-radius: 6px;
                margin: 2px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                border: none;
                background: none;
                width: 0;
                height: 0;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
            QToolTip {
                border: 1px solid black;
                background-color: #616161;
                padding: 1px;
                border-radius: 3px;
                opacity: 100;
                color: #FFFFFF;
            }
            QWidget {
                color: #b1b1b1;
                background-color: #323232;
            }
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #565656, stop:0.1 #525252, stop:0.5 #4e4e4e,
                    stop:0.9 #4a4a4a, stop:1 #464646);
                border: 1px solid #1e1e1e;
                border-radius: 6px;
                padding: 3px;
                font-size: 12px;
                min-width: 40px;
                color: #b1b1b1;
            }
            QPushButton:hover {
                border: 2px solid qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ffa02f, stop:1 #d7801a);
            }
            QPushButton:pressed {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2d2d2d, stop:0.1 #2b2b2b, stop:0.5 #292929,
                    stop:0.9 #282828, stop:1 #252525);
            }
            QLabel {
                color: #b1b1b1;
                font-size: 12px;
            }
            QGraphicsView {
                background-color: #242424;
                border: 1px solid #444;
            }
            QInputDialog QLineEdit {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #4d4d4d, stop:0 #646464, stop:1 #5d5d5d);
                border: 1px solid #1e1e1e;
                border-radius: 5px;
                color: #b1b1b1;
                padding: 1px;
            }
            QMenu {
                background-color: #323232;
                border: 1px solid #444;
            }
            QMenu::item:selected {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ffa02f, stop:1 #d7801a);
                color: #000000;
            }
            QComboBox QAbstractItemView {
                background-color: #323232;
                color: #b1b1b1;
                border: 1px solid #444;
                selection-background-color: #616161;
                selection-color: #ffffff;
            }
            QComboBox {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #565656, stop:0.1 #525252, stop:0.5 #4e4e4e,
                    stop:0.9 #4a4a4a, stop:1 #464646);
                color: #b1b1b1;
                border: 1px solid #1e1e1e;
                border-radius: 3px;
                padding: 1px 18px 1px 3px;
                min-width: 6em;
            }
            
            QInputDialog {
                background-color: #323232;
            }
            QInputDialog QLabel {
                color: #b1b1b1;
                font-size: 12px;
            }
        """)

    window = StagesEditor()
    window.show()
    sys.exit(app.exec())
