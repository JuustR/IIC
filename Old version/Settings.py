import sys
import json
from PyQt6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QLabel,
    QLineEdit, QPushButton, QMessageBox, QWidget
)
from PyQt6.QtCore import QSettings


class SettingsDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Настройки")
        self.resize(500, 400)

        # Хранилище настроек
        self.settings = QSettings("lab425", "IIC")

        # Загрузка словаря категорий
        self.categories = self.load_categories()

        # Элементы интерфейса
        layout = QVBoxLayout(self)

        # Виджет вкладок
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        # Поле для имени категории
        name_layout = QHBoxLayout()
        self.label_category_name = QLabel("Название категории:")
        name_layout.addWidget(self.label_category_name)

        self.input_category_name = QLineEdit()
        name_layout.addWidget(self.input_category_name)

        self.update_name_button = QPushButton("Обновить имя")
        self.update_name_button.clicked.connect(self.update_category_name)
        name_layout.addWidget(self.update_name_button)

        layout.addLayout(name_layout)

        # Кнопки управления категориями
        button_layout = QHBoxLayout()
        self.add_category_button = QPushButton("Добавить категорию")
        self.add_category_button.clicked.connect(self.add_category)
        button_layout.addWidget(self.add_category_button)

        self.delete_category_button = QPushButton("Удалить категорию")
        self.delete_category_button.clicked.connect(self.delete_category)
        button_layout.addWidget(self.delete_category_button)

        layout.addLayout(button_layout)

        # Кнопка для сохранения параметров
        self.save_button = QPushButton("Сохранить параметры")
        self.save_button.clicked.connect(self.save_categories)
        layout.addWidget(self.save_button)

        self.setLayout(layout)

        # Загрузка категорий во вкладки
        self.load_tabs()

    def load_categories(self):
        """Загружает словарь категорий из QSettings."""
        categories_json = self.settings.value("categories", "{}")
        return json.loads(categories_json)

    def save_categories(self):
        """Сохраняет словарь категорий в QSettings."""
        categories_json = json.dumps(self.categories)
        self.settings.setValue("categories", categories_json)
        QMessageBox.information(self, "Настройки", "Все параметры сохранены!")

    def load_tabs(self):
        """Создаёт вкладки для всех категорий."""
        self.tab_widget.clear()
        for category, params in self.categories.items():
            self.add_tab(category, params)

    def add_tab(self, category, params):
        """Добавляет вкладку для категории."""
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)

        label_param1 = QLabel("Параметр 1:")
        tab_layout.addWidget(label_param1)

        input_param1 = QLineEdit(params.get("параметр1", ""))
        tab_layout.addWidget(input_param1)

        label_param2 = QLabel("Параметр 2:")
        tab_layout.addWidget(label_param2)

        input_param2 = QLineEdit(params.get("параметр2", ""))
        tab_layout.addWidget(input_param2)

        self.tab_widget.addTab(tab, category)
        tab.setProperty("params", {"параметр1": input_param1, "параметр2": input_param2})

    def add_category(self):
        """Добавляет новую категорию."""
        new_category = f"Новая категория {len(self.categories) + 1}"
        self.categories[new_category] = {"параметр1": "", "параметр2": ""}
        self.add_tab(new_category, self.categories[new_category])
        self.tab_widget.setCurrentIndex(self.tab_widget.count() - 1)
        self.input_category_name.setText(new_category)

    def delete_category(self):
        """Удаляет выбранную категорию."""
        current_index = self.tab_widget.currentIndex()
        if current_index == -1:
            QMessageBox.warning(self, "Ошибка", "Нет выбранной категории!")
            return

        category_name = self.tab_widget.tabText(current_index)
        reply = QMessageBox.question(
            self, "Подтверждение", f"Удалить категорию '{category_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.categories.pop(category_name, None)
            self.tab_widget.removeTab(current_index)

    def update_category_name(self):
        """Обновляет имя текущей категории."""
        current_index = self.tab_widget.currentIndex()
        if current_index == -1:
            QMessageBox.warning(self, "Ошибка", "Нет выбранной категории!")
            return

        new_name = self.input_category_name.text().strip()
        if not new_name:
            QMessageBox.warning(self, "Ошибка", "Имя категории не может быть пустым!")
            return

        old_name = self.tab_widget.tabText(current_index)
        if old_name != new_name:
            # Обновить название в словаре категорий
            self.categories[new_name] = self.categories.pop(old_name)
            self.tab_widget.setTabText(current_index, new_name)

    def save_tabs_to_categories(self):
        """Обновляет словарь категорий данными из вкладок."""
        for index in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(index)
            category_name = self.tab_widget.tabText(index)
            params_widgets = tab.property("params")
            self.categories[category_name] = {
                "параметр1": params_widgets["параметр1"].text(),
                "параметр2": params_widgets["параметр2"].text(),
            }

    def closeEvent(self, event):
        """Сохраняет данные перед закрытием окна."""
        self.save_tabs_to_categories()
        self.save_categories()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    dialog = SettingsDialog()
    dialog.exec()
