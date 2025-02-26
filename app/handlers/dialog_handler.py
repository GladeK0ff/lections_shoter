# dialogs.py
from PyQt5 import QtWidgets, QtCore


class InfoDialog(QtWidgets.QDialog):
    def __init__(self, message, show_button=True, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Инструкция")
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)

        screen_geometry = QtWidgets.QApplication.primaryScreen().geometry()
        width = int(screen_geometry.width() * 0.3)
        height = int(screen_geometry.height() * 0.2)
        x_pos = int(screen_geometry.width() * 0.05)
        y_pos = int(screen_geometry.height() * 0.4)

        self.setGeometry(x_pos, y_pos, width, height)

        layout = QtWidgets.QVBoxLayout()

        self.label = QtWidgets.QLabel(message)
        layout.addWidget(self.label)

        if show_button:
            button = QtWidgets.QPushButton("Понятно", self)
            button.clicked.connect(self.accept)
            layout.addWidget(button)

        self.setLayout(layout)


class SlideCountDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Количество слайдов")
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)

        screen_geometry = QtWidgets.QApplication.primaryScreen().geometry()
        width = int(screen_geometry.width() * 0.3)
        height = int(screen_geometry.height() * 0.2)
        x_pos = int(screen_geometry.width() * 0.05)
        y_pos = int(screen_geometry.height() * 0.4)

        self.setGeometry(x_pos, y_pos, width, height)

        layout = QtWidgets.QVBoxLayout()

        label = QtWidgets.QLabel('''
                                 Убедитесь, что вы в начале презентации\n
                                 После нажатия кнопки "Готово" программа начнет свою работу\n
                                 Введите количество слайдов:
                                 ''')

        layout.addWidget(label)

        self.spin_box = QtWidgets.QSpinBox()
        self.spin_box.setMinimum(1)
        self.spin_box.setMaximum(9999)
        layout.addWidget(self.spin_box)

        submit_btn = QtWidgets.QPushButton("Готово", self)
        submit_btn.clicked.connect(self.accept)
        layout.addWidget(submit_btn)

        self.setLayout(layout)

    def get_slide_count(self):
        return self.spin_box.value()


class SaveFileDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Сохранить как")
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)

        layout = QtWidgets.QVBoxLayout()

        # Поле для ввода пути
        self.file_path_edit = QtWidgets.QLineEdit()
        layout.addWidget(self.file_path_edit)

        # Кнопка выбора файла
        browse_button = QtWidgets.QPushButton("Обзор")
        browse_button.clicked.connect(self.browse_file)
        layout.addWidget(browse_button)

        # Кнопка сохранения
        save_button = QtWidgets.QPushButton("Сохранить")
        save_button.clicked.connect(self.accept)  # Подтверждение сохранения
        layout.addWidget(save_button)

        # Кнопка отмены
        cancel_button = QtWidgets.QPushButton("Отмена")
        cancel_button.clicked.connect(self.reject)  # Отмена сохранения
        layout.addWidget(cancel_button)

        self.setLayout(layout)

    def browse_file(self):
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Выберите файл",
            "",
            "Microsoft Word Documents (*.docx);;PDF Documents (*.pdf)"
        )
        if file_path:
            self.file_path_edit.setText(file_path)

    def get_save_file_path(self):
        return self.file_path_edit.text()


class QuestionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Анимация в презентации?")
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)

        screen_geometry = QtWidgets.QApplication.primaryScreen().geometry()
        width = int(screen_geometry.width() * 0.3)
        height = int(screen_geometry.height() * 0.2)
        x_pos = int(screen_geometry.width() * 0.05)
        y_pos = int(screen_geometry.height() * 0.4)

        self.setGeometry(x_pos, y_pos, width, height)

        layout = QtWidgets.QVBoxLayout()

        label = QtWidgets.QLabel(
            "Есть ли анимация в презентации?\n"
            "- Если есть, нажмите 'Да'.\n"
            "- Если нет, нажмите 'Нет'."
        )
        layout.addWidget(label)

        yes_button = QtWidgets.QPushButton("Да", self)
        yes_button.clicked.connect(self.accept_with_animation)
        layout.addWidget(yes_button)

        no_button = QtWidgets.QPushButton("Нет", self)
        no_button.clicked.connect(self.accept_without_animation)
        layout.addWidget(no_button)

        self.setLayout(layout)

    def accept_with_animation(self):
        self.result = "with_animation"
        self.accept()

    def accept_without_animation(self):
        self.result = "without_animation"
        self.accept()