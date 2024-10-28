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
            if len(self.positions) == 2:
                return False

class InfoDialog(QtWidgets.QDialog):
    def __init__(self, width_ratio=0.25, height_ratio=0.2, x_offset_ratio=0.1, y_offset_ratio=0.5):
        super().__init__()
        self.setWindowTitle("Информация")
        
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
                   "После нажатия кнопки 'Понятно' программа ожидает 2 клика: \n"
                   "1. Верхняя левая часть слайда.\n"
                   "2. Нижняя правая часть слайда.\n"
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
Особенно проверьте первый слайд\n
Первый слайд должен перелистываться за 1 клик в центр\n
После нажатия кнопки "Подтвердить" программа продолжит свою работу\n
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
    
    # Show the info dialog
    InfoDialog()

    # Create instance for capturing clicks
    click_positions = ClickPositions()

    # Start the mouse listener
    with mouse.Listener(on_click=click_positions.on_click) as listener:
        listener.join()

    # Get positions after stopping the listener
    pos1, pos2 = click_positions.positions
    x1, y1 = pos1
    x2, y2 = pos2
    print(x1, y1, x2, y2)

    # Get slide count from user
    n = get_slide_count()
    print(f'got slide count - {n}')

    if n is None or n <= 0:
        print("Ошибка: введено некорректное число слайдов.")
    else:
        for i in range(n):
            screenshot_param = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            print('screenshot parametred')
            screenshot_param.save(r'C:screenshot.png')
            print('screenshot saved')
            doc.add_picture(r'C:screenshot.png', width=docx.shared.Cm(14.99))
            print('doc added')
            pyautogui.click(x=(x1+x2)/2, y=(y1+y2)/2, interval=0.3)
            print('clicked')
            os.remove(r'C:screenshot.png')
            print('screenshot removed')

    save_document(doc)
