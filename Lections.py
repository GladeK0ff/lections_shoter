import docx
from PIL import ImageGrab
import pyautogui
from pynput import mouse
import os
import sys
import win32com.client
import time
from PyQt5 import QtWidgets, QtCore

doc = docx.Document()

class ClickPositions:
    def __init__(self):
        self.positions = []

    def on_click(self, x, y, _button, pressed):
        if pressed:
            # Save coordinates
            self.positions.append((x, y))
            # Stop listener after capturing two positions
            if len(self.positions) == 3:
                return False

class AnimationDialog(QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Слайды с анимацией")

        # Установка флага, чтобы окно всегда было сверху
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)

        # Устанавливаем расположение окна
        screen_geometry = QtWidgets.QDesktopWidget().screenGeometry()
        width_ratio = 0.3
        height_ratio = 0.2
        x_offset_ratio = 0.01
        y_offset_ratio = 0.3

        width = int(screen_geometry.width() * width_ratio)
        height = int(screen_geometry.height() * height_ratio)

        # Calculate position based on offset ratios
        x_pos = int(screen_geometry.width() * x_offset_ratio)
        y_pos = int(screen_geometry.height() * y_offset_ratio)


        self.setGeometry(x_pos, y_pos, width, height)

        self.layout = QtWidgets.QVBoxLayout()  # Сохраняем ссылку на layout

        # Информационное сообщение
        self.info_label = QtWidgets.QLabel("""Убедитесь в том, что в лекции нет слайдов с анимацией\n
Программа использует время клика, равное 0,3 секунды\n
Если в лекции присутствует анимация, замедляющая полный показ слайда, это может сказаться на конечном результате\n
Наличие анимации вы можете заподозрить по времени, указанном в правом нижнем углу презентации\n
Если презентации нет, то нажмите 'Подтвердить' в данном окне, программа продолжит свою работу\n
Если презентация есть поставьте галочку в вопросе 'Есть ли в презентации слайды для анимации?'\n
Быстро вручную пролистайте лекцию и определите номера слайдов, а также максимальное время действия анимации на одном слайде\n
В диалоговом окнах соответственно укажите номера слайдов через запятую (x1, x2,...)\n
и максимальное время анимации на одном слайде в секундах целым числом (огруглять в большую сторону)\n
После этого нажмите "Продолжить"                         
Введите пары: слайд и время""")
        self.layout.addWidget(self.info_label)

        # Чекбокс для анимации
        self.animation_checkbox = QtWidgets.QCheckBox("Есть ли анимация в лекции?")
        self.layout.addWidget(self.animation_checkbox)

        # Поле для ввода постоянного значения
        self.constant_value_input = QtWidgets.QLineEdit(self)
        self.constant_value_input.setPlaceholderText("Введите максимальное время анимации на 1 слайде (в секундах)")
        self.layout.addWidget(self.constant_value_input)

        # Поле для ввода нескольких значений
        self.values_input = QtWidgets.QLineEdit(self)
        self.values_input.setPlaceholderText("Введите номера слайдов с анимацией через запятую (x1, x2, ...)")
        self.layout.addWidget(self.values_input)

        # Кнопка подтверждения
        self.confirm_button = QtWidgets.QPushButton("Подтвердить")
        self.confirm_button.clicked.connect(self.on_confirm)
        self.layout.addWidget(self.confirm_button)

        self.setLayout(self.layout)

        # Скрываем поля для ввода, если анимации нет
        self.update_input_fields()

        # Подключаем сигнал чекбокса
        self.animation_checkbox.stateChanged.connect(self.update_input_fields)

    def update_input_fields(self):
        # Включаем или отключаем поля для ввода в зависимости от состояния чекбокса
        is_checked = self.animation_checkbox.isChecked()
        self.constant_value_input.setEnabled(is_checked)
        self.values_input.setEnabled(is_checked)

        if not is_checked:
            # Если анимации нет, очищаем поля ввода
            self.constant_value_input.clear()
            self.values_input.clear()


    def on_confirm(self):
        # Проверяем, есть ли анимация
        if not self.animation_checkbox.isChecked():
            self.accept()  # Если анимации нет, просто закрываем окно
            return

        # Получаем значение из первого поля ввода
        constant_value = self.constant_value_input.text().strip()
        
        # Проверяем, является ли постоянное значение целым числом
        if not constant_value.isdigit():
            QtWidgets.QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите корректное постоянное числовое значение.")
            return

        # Получаем значения из второго поля ввода
        values_input = self.values_input.text().strip()

        # Разделяем значения по запятой и проверяем
        values = values_input.split(',')
        filtered_values = [v.strip() for v in values if v.strip().isdigit()]

        if not filtered_values:
            QtWidgets.QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите хотя бы одно целое положительное число.")
            return

        # Формируем словарь
        data = {int(val): int(constant_value) for val in filtered_values}

        # Выводим результат в консоль или используйте его как вам нужно
        print(data)
        self.accept()  # Закрываем окно
        return(data)
        

    
    def closeEvent(self, event):
        """Обработка закрытия окна"""
        print("Программа закрыта пользователем.")  # Вывод сообщения в консоль (по желанию)
        QtWidgets.QApplication.quit()  # Закрываем всю программу


          

class InfoDialog(QtWidgets.QDialog):
    def __init__(self, width_ratio=0.25, height_ratio=0.2, x_offset_ratio=0.1, y_offset_ratio=0.5):
        super().__init__()
        self.setWindowTitle("Информация")

        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
        
        screen_geometry = QtWidgets.QDesktopWidget().screenGeometry()
        width = int(screen_geometry.width() * width_ratio)
        height = int(screen_geometry.height() * height_ratio)

        # Calculate position based on offset ratios
        x_pos = int(screen_geometry.width() * x_offset_ratio)
        y_pos = int(screen_geometry.height() * y_offset_ratio)

        self.setGeometry(x_pos, y_pos, width, height)

        layout = QtWidgets.QVBoxLayout()

        message = ("Данный шаг программы позволяет сделать скриншот области экрана \n"
                   "на основе двух кликов мыши.\n"
                   "После нажатия кнопки 'Понятно' программа ожидает 3 клика: \n"
                   "1. Верхняя левая часть слайда.\n"
                   "2. Нижняя правая часть слайда.\n"
                   "3. Кнопка перелистывания слайда слева в углу презентации\n"
                   "После этого программа продолжит выполнение.")
        
        label = QtWidgets.QLabel(message)
        layout.addWidget(label)

        button = QtWidgets.QPushButton("Понятно", self)
        button.clicked.connect(self.accept)
        layout.addWidget(button)
        
        self.setLayout(layout)
        self.exec_()

def get_slide_count():
    window = QtWidgets.QDialog()

    window.setWindowFlags(window.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
    
    window.setWindowTitle("Ввод количества слайдов")

    screen_geometry = QtWidgets.QDesktopWidget().screenGeometry()
    width_ratio = 0.25
    height_ratio = 0.2
    x_offset_ratio = 0.1
    y_offset_ratio = 0.5

    width = int(screen_geometry.width() * width_ratio)
    height = int(screen_geometry.height() * height_ratio)

    # Calculate position based on offset ratios
    x_pos = int(screen_geometry.width() * x_offset_ratio)
    y_pos = int(screen_geometry.height() * y_offset_ratio)


    window.setGeometry(x_pos, y_pos, width, height)
    layout = QtWidgets.QVBoxLayout()

    label = QtWidgets.QLabel('''Убедитесь, что лекция в самом начале\n
Программа требует на вход количество слайдов(целое положительное число)\n
После нажатия кнопки "Подтвердить" она продолжит свою работу\n
Введите количество слайдов:''')
    layout.addWidget(label)
    
    spin_box = QtWidgets.QSpinBox()
    spin_box.setMinimum(1)
    layout.addWidget(spin_box)
    
    submit_btn = QtWidgets.QPushButton("Подтвердить")
    layout.addWidget(submit_btn)
    
    window.setLayout(layout)

    def on_submit():
        window.close()

    submit_btn.clicked.connect(on_submit)

    window.exec_()
    
    return spin_box.value()

def save_document(doc):
    app = QtWidgets.QApplication(sys.argv)
    file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
        None,
        "Сохранить как",
        "",
        "Microsoft Word Documents (*.docx);;PDF Documents (*.pdf)"
    )

    print(f"Выбранный путь к файлу: {file_path}")

    if file_path:
        if file_path.lower().endswith('.pdf'):
            docx_temp_path = os.path.splitext(file_path)[0] + ".docx"
            doc.save(docx_temp_path)

            if not os.path.exists(docx_temp_path):
                print(f"Ошибка: временный файл не был создан: {docx_temp_path}")
                return

            print(f"Временный файл сохранен как: {docx_temp_path}")

            word = win32com.client.Dispatch("Word.Application")
            time.sleep(0.1)
            try:
                doc_word = word.Documents.Open(os.path.abspath(docx_temp_path))
                doc_word.SaveAs(os.path.abspath(file_path), FileFormat=17)  # FileFormat=17 означает PDF
                doc_word.Close()
                print(f"Документ сохранен как PDF: {file_path}")
            except Exception as e:
                print(f"Ошибка при сохранении в PDF: {e}")
            finally:
                word.Quit()

            if os.path.exists(docx_temp_path):
                os.remove(docx_temp_path)
                print('Временный файл удален')
        else:
            doc.save(file_path)
            print(f"Документ сохранен как DOCX: {file_path}")

        if os.path.exists(file_path):
            os.startfile(file_path)
        else:
            print(f"Файл не найден: {file_path}")
    else:
        print("Сохранение отменено")

# Main Execution
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    
    # Show the animation dialog first
    animation_dialog = AnimationDialog()
    if animation_dialog.exec_() == QtWidgets.QDialog.Accepted:
        slide_data = animation_dialog.on_confirm()

        if slide_data is not None:
            print("Полученные данные слайдов:", slide_data)
        else:
            print("Слайды для анимации не указаны.")  # Сообщение, если слайды не указаны

    # Show the info dialog
    InfoDialog()

    # Create instance for capturing clicks
    click_positions = ClickPositions()

    # Start the mouse listener
    with mouse.Listener(on_click=click_positions.on_click) as listener:
        listener.join()

    # Get positions after stopping the listener
    pos1, pos2, pos3 = click_positions.positions
    x1, y1 = pos1
    x2, y2 = pos2
    x3, y3 = pos3
    print(x1, y1, x2, y2, x3, y3)

    # Get slide count from user
    n = get_slide_count()
    print(f'got slide count - {n}')

    if n is None or n <= 0:
        print("Ошибка: введено некорректное число слайдов.")
    else:
        print(f'словарь - {type(slide_data)}')
        for i in range(1, n+1):
            screenshot_param = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            print(f'screenshot {i} parametred')
            screenshot_param.save(r'C:screenshot.png')
            print(f'screenshot {i} saved')
            doc.add_picture(r'C:screenshot.png', width=docx.shared.Cm(14.99))
            print(f'screenshot {i} added to doc')
            if slide_data == None:
                pyautogui.click(x3, y3, interval=0.3)
                print('clicked for 0.3 sec')
            else:
                if (i+1) in slide_data:
                    pyautogui.click(x3, y3, interval=int(slide_data[i+1]))
                    print(f'clicked for {slide_data[i+1]} sec')
                else:
                    pyautogui.click(x3, y3, interval=0.3)
                    print('clicked for 0.3 sec')
            os.remove(r'C:screenshot.png')
            print(f'screenshot {i} removed')

    save_document(doc)
