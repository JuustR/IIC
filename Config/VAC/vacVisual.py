import sys
import os
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QVBoxLayout, QWidget,
                             QPushButton, QLabel, QHBoxLayout, QFileDialog, QMessageBox,
                             QLineEdit, QFrame, QCheckBox)
from PyQt6.QtCore import Qt, QTimer, QSettings
from PyQt6.QtGui import QPainter, QBrush, QColor, QFont

from Config.VAC.vacItemCreator import StageItem

SETTINGS_ORG = "lab425"
SETTINGS_APP = "vac"  # Поменять можно
SETTINGS_LAST_FILE = "last_file_path"


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


class StagesEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Измерение вольт-амперных характеристик")
        self.setGeometry(100, 100, 550, 800)

        # Флаги и состояние файла
        self.is_saved_flag = True  # Флаг сохраненности
        self.current_file_path = None  # Текущий открытый файл

        self.scene = CustomGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.TextAntialiasing)

        self.voltage_items = []
        self.current_number = len(self.voltage_items) + 1

        # Настройки приложения
        self.settings = QSettings(SETTINGS_ORG, SETTINGS_APP)

        self.init_ui()

        # Добавляем кнопки для сохранения/загрузки в менюбар
        self.init_menu()

        # Автозагрузка последнего файла при старте
        self.auto_load_last_file()

        # Подключаем отслеживание изменений
        self.setup_change_tracking()

    # new_file | save_stages_in_file | load_stages_from_file | update_recent_files_menu |
    def init_menu(self):
        """Инициализация меню файлов"""
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Файл")

        new_action = file_menu.addAction("Новый файл")
        new_action.triggered.connect(self.new_file)

        file_menu.addSeparator()

        save_action = file_menu.addAction("Сохранить")
        save_action.triggered.connect(self.save_stages_in_file)

        load_action = file_menu.addAction("Загрузить")
        load_action.triggered.connect(self.load_stages_from_file)

        # Меню последних файлов
        self.recent_menu = file_menu.addMenu("Последние файлы")
        self.update_recent_files_menu()

        file_menu.addSeparator()

        exit_action = file_menu.addAction("Выход")
        exit_action.triggered.connect(self.close)

    def new_file(self):
        """Создает новый файл"""
        if not self.is_saved_flag:
            reply = QMessageBox.question(
                self, "Несохраненные изменения",
                "Есть несохраненные изменения. Сохранить перед созданием нового файла?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.save_stages()
            elif reply == QMessageBox.StandardButton.Cancel:
                return

        self.clear_scene()
        self.current_file_path = None
        self.mark_as_saved()  # Новый файл считается "сохраненным" (пустым)

    def save_stages_in_file(self):
        """Сохраняет стадии в JSON файл"""
        try:
            # Получаем путь для сохранения
            current_dir = os.getcwd()
            parent_dir = os.path.dirname(os.path.dirname(current_dir))
            base_dir = os.path.join(os.path.join(parent_dir, "Stages"), "VAC")
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Сохранить этап", base_dir, "JSON Files (*.json)"
            )

            if not file_path:
                return

            # Подготавливаем данные для сохранения
            stages_data = []
            sorted_items = sorted(self.voltage_items, key=lambda x: x.y())

            for item in sorted_items:
                stage_data = {
                    'voltage': item.voltage,
                    'number': item.number,
                    'y_position': item.y(),
                    'params': item.params
                }
                stages_data.append(stage_data)

            # Сохраняем в файл
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(stages_data, f, indent=4, ensure_ascii=False)

            # ОБНОВЛЯЕМ ФЛАГ СОХРАНЕННОСТИ
            self.mark_as_saved(file_path)

            # Добавляем в последние файлы
            self.add_to_recent_files(file_path)

            # Сохраняем путь в настройках
            self.settings.setValue(SETTINGS_LAST_FILE, file_path)

            QMessageBox.information(self, "Успех", "Этап успешно сохранен!")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении: {str(e)}")

    def load_stages_from_file(self):
        """Загружает стадии из ВЫБРАННОГО JSON файла"""
        try:
            # Получаем путь для загрузки
            current_dir = os.getcwd()
            parent_dir = os.path.dirname(os.path.dirname(current_dir))
            base_dir = os.path.join(os.path.join(parent_dir, "Stages"), "VAC")
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Загрузить этап", base_dir, "JSON Files (*.json)"
            )

            if not file_path:
                return

            self.load_from_file(file_path)

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке: {str(e)}")

    def load_stage(self, stage_data):
        """Загружает одну стадию из данных"""
        try:
            height = 50
            # Загружаем с сохраненной Y позицией
            stage = StageItem(
                50,  # Фиксированная X
                0,
                150, height,
                stage_data['voltage'],
                stage_data['number'],
                stage_data.get('params', {})
            )

            stage.setPos(50, stage_data['y_position'])

            self.scene.addItem(stage)
            self.voltage_items.append(stage)

            # Обновляем счетчик номеров
            self.current_number = max(self.current_number, stage_data['number'] + 1)

        except Exception as e:
            print(f"Ошибка при загрузке этапа: {str(e)}")

    def setup_change_tracking(self):
        """Настройка отслеживания изменений"""
        # Будем отслеживать изменения через таймер
        self.change_timer = QTimer()
        self.change_timer.setInterval(500)  # Проверка каждые 500ms
        self.change_timer.timeout.connect(self.check_for_changes)
        self.change_timer.start()

        # Сохраняем начальное состояние
        self.last_saved_state = self.get_current_state_hash()

    def get_current_state_hash(self):
        """Создает хэш текущего состояния для сравнения"""
        try:
            state_data = []
            # Сортируем по Y для вертикальной ориентации
            for item in sorted(self.voltage_items, key=lambda x: x.y()):
                state_data.append({
                    'voltage': item.voltage,
                    'number': item.number,
                    'y_position': item.y(),  # ← Используем Y
                    'params': item.params
                })
            return hash(str(state_data))
        except:
            return 0

    def check_for_changes(self):
        """Проверяет, были ли изменения с последнего сохранения"""
        if not self.is_saved_flag:
            return  # Уже помечено как несохраненное

        current_state = self.get_current_state_hash()
        if current_state != self.last_saved_state:
            self.mark_as_unsaved()

    def mark_as_unsaved(self):
        """Помечает файл как несохраненный"""
        if self.is_saved_flag:
            self.is_saved_flag = False
            self.update_window_title()

    def mark_as_saved(self, file_path=None):
        """Помечает файл как сохраненный"""
        self.is_saved_flag = True
        if file_path:
            self.current_file_path = file_path
        self.last_saved_state = self.get_current_state_hash()
        self.update_window_title()

    def update_window_title(self):
        """Обновляет заголовок окна с указанием статуса сохранения"""
        base_title = "Измерение вольт-амперных характеристик"

        if self.current_file_path:
            filename = os.path.basename(self.current_file_path)
            title = f"{base_title} - {filename}"
        else:
            title = base_title

        if not self.is_saved_flag:
            title = "* " + title  # Добавляем звездочку для несохраненных файлов

        self.setWindowTitle(title)

    def auto_load_last_file(self):
        """Автоматическая загрузка последнего файла при запуске"""
        last_file = self.settings.value(SETTINGS_LAST_FILE, "")

        if last_file and os.path.exists(last_file):
            try:
                # Загружаем файл
                self.load_from_file(last_file)

            except Exception as e:
                print(f"Ошибка при автозагрузке: {str(e)}")
                # Восстанавливаем стандартный заголовок при ошибке
                self.update_window_title()

    def save_stages(self):
        """Сохраняет стадии в JSON файл"""
        try:
            # Если файл уже был сохранен, используем тот же путь
            if self.current_file_path and self.is_saved_flag:
                file_path = self.current_file_path
            else:
                file_path, _ = QFileDialog.getSaveFileName(
                    self, "Сохранить этап", self.current_file_path or "", "JSON Files (*.json)"
                )

            if not file_path:
                return

            # Сохраняем данные
            stages_data = []
            sorted_items = sorted(self.voltage_items, key=lambda x: x.x())

            for item in sorted_items:
                stage_data = {
                    'voltage': item.voltage,
                    'number': item.number,
                    'y_position': item.y(),
                    'params': item.params
                }
                stages_data.append(stage_data)

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(stages_data, f, indent=4, ensure_ascii=False)

            # Обновляем флаги
            self.mark_as_saved(file_path)
            self.add_to_recent_files(file_path)
            self.settings.setValue(SETTINGS_LAST_FILE, file_path)

            QMessageBox.information(self, "Успех", "Этап успешно сохранены!")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении: {str(e)}")

    def load_from_file(self, file_path):
        """Загружает стадии из указанного файла"""
        try:
            # Проверяем, нужно ли сохранить текущие изменения
            if not self.is_saved_flag:
                reply = QMessageBox.question(
                    self, "Несохраненные изменения",
                    "Есть несохраненные изменения. Сохранить перед загрузкой?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
                )

                if reply == QMessageBox.StandardButton.Yes:
                    self.save_stages()
                elif reply == QMessageBox.StandardButton.Cancel:
                    return

            if not os.path.exists(file_path):
                QMessageBox.warning(self, "Ошибка", f"Файл не существует: {file_path}")
                return

            with open(file_path, 'r', encoding='utf-8') as f:
                stages_data = json.load(f)

            # Очищаем текущую сцену
            self.clear_scene()

            # Восстанавливаем этапы
            for stage_data in stages_data:
                self.load_stage(stage_data)

            # Обновляем нумерацию и layout
            self.renumber_stages()
            self.auto_arrange()

            # Обновляем флаги
            self.mark_as_saved(file_path)
            self.add_to_recent_files(file_path)
            self.settings.setValue(SETTINGS_LAST_FILE, file_path)
            self.update_window_title()

        except json.JSONDecodeError:
            QMessageBox.critical(self, "Ошибка", "Некорректный JSON файл")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке: {str(e)}")

    def clear_scene(self):
        """Очищает сцену и сбрасывает состояние"""
        for item in self.voltage_items[:]:
            self.scene.removeItem(item)
            self.voltage_items.remove(item)
        self.current_number = 1
        # self.setWindowTitle("Редактор последовательности этапов")
        self.mark_as_unsaved()

    def update_recent_files_menu(self):
        """Обновляет меню последних файлов"""
        self.recent_menu.clear()

        recent_files = self.settings.value("recent_files", [])

        if not recent_files:
            action = self.recent_menu.addAction("Нет последних файлов")
            action.setEnabled(False)
            return

        for file_path in recent_files[:5]:  # Последние 5 файлов
            if os.path.exists(file_path):
                action = self.recent_menu.addAction(os.path.basename(file_path))
                action.setData(file_path)
                action.triggered.connect(lambda checked, path=file_path: self.load_from_file(path))

    def add_to_recent_files(self, file_path):
        """Добавляет файл в список последних"""
        recent_files = self.settings.value("recent_files", [])

        # Удаляем дубликаты
        if file_path in recent_files:
            recent_files.remove(file_path)

        # Добавляем в начало
        recent_files.insert(0, file_path)

        # Ограничиваем количество
        recent_files = recent_files[:10]

        self.settings.setValue("recent_files", recent_files)
        self.update_recent_files_menu()

    def safe_auto_arrange(self):
        """Защищённый вызов авто-выравнивания"""
        try:
            if all(hasattr(self, attr) for attr in ['scene', 'voltage_items']):
                self.auto_arrange()
        except Exception as e:
            print(f"Ошибка выравнивания: {str(e)}")

    # init_ui
    def init_ui(self):
        """Основное рабочее окно"""
        central_widget = QWidget()
        layout = QVBoxLayout()

        # Назначаем objectName для дополнительной стилизации
        central_widget.setObjectName("mainWidget")

        # Панель управления
        control_panel = QWidget()
        control_panel.setObjectName("controlPanel")
        control_layout = QHBoxLayout()

        self.value_label = QLabel("Значение")
        self.value_label.setObjectName("valueLabel")

        self.voltage_le = QLineEdit("0.1")
        self.voltage_le.setObjectName("voltage_le")
        self.voltage_le.setFixedWidth(45)

        self.add_button = QPushButton("Добавить")
        self.add_button.setObjectName("addButton")

        self.vert_line = QFrame()
        self.vert_line.setFrameShape(QFrame.Shape.VLine)
        self.vert_line.setObjectName("vertLine")

        self.step_label = QLabel("Настройка шага")
        self.step_label.setObjectName("stepLabel")

        self.step_le = QLineEdit("0.1")
        self.step_le.setObjectName("step_le")
        self.step_le.setFixedWidth(45)

        self.plus_cb = QCheckBox("+/-")
        self.plus_cb.setChecked(True)
        self.plus_cb.setObjectName("plus_cb")

        self.vert_line2 = QFrame()
        self.vert_line2.setFrameShape(QFrame.Shape.VLine)
        self.vert_line2.setObjectName("vertLine2")

        self.remove_button = QPushButton("Удалить выбранное")
        self.remove_button.setObjectName("removeButton")

        # Добавляем элементы в layout
        for elm in [self.value_label, self.voltage_le, self.add_button, self.vert_line, self.step_label, self.step_le,
                    self.plus_cb, self.vert_line2, self.remove_button]:
            control_layout.addWidget(elm)

        control_panel.setLayout(control_layout)

        # Информационная метка
        self.info_label = QLabel("Перетаскивайте стадии для изменения порядка")
        self.info_label.setObjectName("infoLabel")

        # Настройка скроллбаров
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # Добавляем элементы в основной layout
        layout.addWidget(control_panel)
        layout.addWidget(self.info_label)
        layout.addWidget(self.view)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Подключаем сигналы
        self.add_button.clicked.connect(self.add_stage)
        self.remove_button.clicked.connect(self.remove_selected)

    def add_stage(self):
        current_voltage = self.voltage_le.text()

        # Рассчитываем позицию для нового этапа (вертикально)
        if self.voltage_items:
            # Находим самый нижний элемент по Y
            bottommost_item = max(self.voltage_items, key=lambda item: item.y())
            new_y = bottommost_item.y() + bottommost_item.rect().height() + 30
        else:
            new_y = 50

        height = 50
        # Создаем элемент с фиксированной X позицией
        voltage_item = StageItem(50, 0, 150, height, current_voltage, self.current_number)

        self.scene.addItem(voltage_item)
        self.voltage_items.append(voltage_item)
        self.current_number += 1
        voltage_item.setPos(50, new_y)

        if self.plus_cb.isChecked():
            text = round((float(self.voltage_le.text()) + float(self.step_le.text())), 4)
        else:
            text = round((float(self.voltage_le.text()) - float(self.step_le.text())), 4)
        self.voltage_le.setText(str(text))

        self.renumber_stages()
        self.auto_arrange()

        self.view.ensureVisible(voltage_item)
        self.mark_as_unsaved()


    def remove_selected(self):
        """Безопасное удаление выбранных элементов"""
        try:
            items_to_remove = []
            for item in self.scene.selectedItems():
                if isinstance(item, StageItem) and item in self.voltage_items:
                    items_to_remove.append(item)

            for item in items_to_remove:
                try:
                    self.scene.removeItem(item)
                    self.voltage_items.remove(item)
                except:
                    continue

            self.renumber_stages()
            QTimer.singleShot(100, self.safe_auto_arrange)

            self.mark_as_unsaved()
        except Exception as e:
            print(f"Ошибка при удалении: {str(e)}")

    def auto_arrange(self):
        """Автоматическое выравнивание стадий"""
        try:
            if not hasattr(self, 'voltage_items') or not self.voltage_items:
                return

            # Создаем временный список для безопасного доступа
            valid_items = [item for item in self.voltage_items
                           if isinstance(item, StageItem) and item.scene() is not None]

            if not valid_items:
                return

            # Сортируем по текущей позиции Y
            sorted_items = sorted(valid_items, key=lambda x: x.y())

            # Выравниваем с отступами
            x_pos = 50
            y_pos = 50
            spacing = 30

            for item in sorted_items:
                try:
                    if item and item.scene():  # Дополнительная проверка
                        item.setPos(x_pos, y_pos)
                        y_pos += item.rect().height() + spacing
                except:
                    continue  # Пропускаем проблемные элементы

            # Обновляем размер сцены
            try:
                self.scene.setSceneRect(0, 0, 250, y_pos + 50)

                # Управление скроллбаром
                if len(valid_items) > 1:
                    self.view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
                    self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
                else:
                    self.view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            except:
                pass

            self.renumber_stages()

        except Exception as e:
            print(f"Ошибка в auto_arrange: {str(e)}")

    def renumber_stages(self):
        """Перенумерация этапов"""
        try:
            valid_items = [item for item in self.voltage_items
                           if isinstance(item, StageItem) and item.scene() is not None]

            # Сортируем по Y позиции (сверху вниз)
            for i, item in enumerate(sorted(valid_items, key=lambda x: x.y()), 1):
                try:
                    item.number = i
                    item.update()
                except:
                    continue
        except:
            pass

    def closeEvent(self, event):
        """Ивент при закрытии программы, который проверяет сохраненность файлов"""
        if not self.is_saved_flag:
            reply = QMessageBox.question(
                self, "Несохраненные изменения",
                "Есть несохраненные изменения. Сохранить перед выходом?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.save_stages()
                event.accept()
            elif reply == QMessageBox.StandardButton.No:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


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
