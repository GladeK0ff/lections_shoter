import docx
# import pyscreenshot
from PIL import ImageGrab
import pyautogui
from pynput import mouse
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import time

doc = docx.Document()

class ClickPositions:
    def __init__(self):
        self.positions = []

    def on_click(self, x, y, _button, pressed):
        if pressed:
            # Сохраняем координаты
            self.positions.append((x, y))
            # Останавливаем слушатель после захвата двух позиций
            if len(self.positions) == 2:
                return False

def show_info_dialog():
    root = tk.Tk()
    root.title("Информация")

    # Получите размеры экрана
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # Настройка размеров и положения окна
    width = int(screen_width * 0.25)   # 25% от ширины экрана
    height = int(screen_height * 0.2)  # 20% от высоты экрана
    x_pos = int((screen_width - width) / 9)   # Центрирование по горизонтали
    y_pos = int((screen_height - height) / 2)  # Центрирование по вертикали
    
    root.geometry(f"{width}x{height}+{x_pos}+{y_pos}")

    message = ("Данный шаг программы позволяет сделать скриншот области экрана \n"
               "на основе двух кликов мыши.\n"
               "После нажатия кнопки 'Понятно' программа ожидает 2 клика: \n"
               "1. Верхняя левая часть слайда.\n"
               "2. Нижняя правая часть слайда.\n"
               "После этого программа продолжит выполнение.")
    
    label = tk.Label(root, text=message, padx=20, pady=20)
    label.pack()

    button = tk.Button(root, text="Понятно", command=root.destroy)
    button.pack(pady=10)

    root.mainloop()

# Показать окно с информацией пользователю
show_info_dialog()

# Создаем экземпляр класса для захвата кликов
click_positions = ClickPositions()

# Установка слушателя
with mouse.Listener(on_click=click_positions.on_click) as listener:
    listener.join()

# Получаем позиции после остановки слушателя
pos1, pos2 = click_positions.positions
x1, y1 = pos1
x2, y2 = pos2

print(x1, y1)
print(x2, y2)

def get_slide_count():
    root = tk.Tk()
    root.title("Ввод количества слайдов")
    
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    width = int(screen_width * 0.25)
    height = int(screen_height * 0.2)
    x_pos = int((screen_width - width) / 9)
    y_pos = int((screen_height - height) / 2)

    root.geometry(f"{width}x{height}+{x_pos}+{y_pos}")
    root.attributes("-topmost", True)

    slide_count = None

    label = tk.Label(root, text='''Убедитесь, что лекция в самом начале\nОсобенно проверьте первый слайд\nПервый слайд должен перелистываться за 1 клик в центр\nПосле нажатия кнопки "Подтвердить" программа продолжит свою работу\nВведите количество слайдов:''')
    label.pack(pady=10)
    print('window showed')
    entry = tk.Entry(root)
    entry.pack(pady=5)
    entry.focus_force()

    def on_submit(event=None):
        nonlocal slide_count
        try:
            slide_count = int(entry.get())
            print('got number')
            if slide_count <= 0:
                raise ValueError("Число должно быть положительным.")
            root.destroy()
        except ValueError:
            messagebox.showerror("Ошибка", "Пожалуйста, введите корректное число слайдов.")

    submit_btn = tk.Button(root, text="Подтвердить", command=on_submit)
    entry.bind("<Return>", on_submit)
    submit_btn.pack(pady=10)

    root.deiconify()
    root.mainloop()

    return slide_count if slide_count is not None and slide_count > 0 else None

n = get_slide_count()
print(f'got slide count - {n}')

if n is None or n <= 0:
    print("Ошибка: введено некорректное число слайдов.")
else:
    for i in range(n):
        screenshot_parametres = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        print('screenshot parametred')
        screenshot_parametres.save(r'C:screenshot.png')
        print('screenshot saved')
        doc.add_picture(r'C:screenshot.png', width=docx.shared.Cm(14.99))
        print('doc added')
        pyautogui.click(x=(x1+x2)/2, y=(y1+y2)/2, interval=0.15)
        print('clicked')
        os.remove(r'C:screenshot.png')
        print('scr removed')

def save_document(doc):
    root = tk.Tk()
    root.withdraw()  # Скрытие основного окна

    file_path = filedialog.asksaveasfilename(
        title="Сохранить как",
        defaultextension=".docx",
        filetypes=[("Microsoft Word Documents", "*.docx"),
                   ("PDF Documents", "*.pdf")]
    )

    print(f"Выбранный путь к файлу: {file_path}")

    if file_path:  # Проверяем, выбрали ли файл
        if file_path.lower().endswith('.pdf'):
            docx_temp_path = os.path.normpath(os.path.splitext(file_path)[0] + ".docx")
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

save_document(doc)
