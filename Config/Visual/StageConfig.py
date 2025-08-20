from PyQt6.QtWidgets import (QVBoxLayout, QWidget, QLabel, QComboBox, QCheckBox, QFormLayout,
                             QDialogButtonBox, QStackedWidget, QDialog)


class StageConfigDialog(QDialog):
    """Класс для работы со стадиями"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройка стадии")

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
        layout.addWidget(QLabel("Тип стадии:"))
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