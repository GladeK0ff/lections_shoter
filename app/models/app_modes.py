# app_modes.py
from PyQt5 import QtWidgets, QtCore
from handlers.dialog_handler import InfoDialog, SaveFileDialog, SlideCountDialog
from handlers.document_handler import DocumentHandler
from handlers.click_handler import ClickHandler
import keyboard
import threading
import time
import os
import sys
from PIL import ImageGrab
import pyautogui
import tempfile
from logs.logger import logger
from pynput import mouse
from docx.shared import Cm


class AnimationApp(QtCore.QObject):
    screenshot_signal = QtCore.pyqtSignal(str)
    save_signal = QtCore.pyqtSignal()

    def __init__(self):
        super().__init__()
        self.doc = DocumentHandler.create_document()
        self.stop_app = False
        self.info_dialog = None
        self.screenshot_count = 0
        self.screenshot_message = ""

        exit_thread = threading.Thread(target=self.check_exit_key, daemon=True)
        exit_thread.start()

        self.screenshot_signal.connect(self.add_screenshot_to_doc)
        self.save_signal.connect(self.save_and_exit)

    def check_exit_key(self):
        keyboard.wait('q')
        self.stop_app = True
        if self.info_dialog:
            QtWidgets.QApplication.postEvent(
                self.info_dialog,
                QtCore.QEvent(QtCore.QEvent.Close)
            )
        logger.info("Программа завершена по нажатию клавиши 'Q'.")

    def start_capture(self):
        info_message = (
            "Программа работает в режиме с анимацией.\n"
            "После нажатия 'Понятно', программа ожидает два клика:\n"
            "1. Верхняя левая точка слайда.\n"
            "2. Нижняя правая точка слайда."
        )
        InfoDialog(info_message).exec_()

        click_handler = ClickHandler(mode="animation")
        with mouse.Listener(on_click=click_handler.on_click) as listener:
            listener.join()

        pos1, pos2 = click_handler.get_positions()[:2]
        x1, y1 = pos1
        x2, y2 = pos2

        logger.info(f"Координаты скриншота - {pos1}, {pos2}")

        self.show_info_dialog()

        capture_thread = threading.Thread(target=self.capture_loop, args=(x1, y1, x2, y2), daemon=True)
        capture_thread.start()

        while not self.stop_app:
            QtWidgets.QApplication.processEvents()
            time.sleep(0.01)

    def show_info_dialog(self):
        if not self.info_dialog or not self.info_dialog.isVisible():
            info_message = (
                "Инструкции:\n"
                "- Нажмите 'E', чтобы сделать скриншот.\n"
                "- Нажмите 'S', чтобы сохранить файл.\n"
                "- Нажмите 'Q', чтобы выйти из программы."
            )
            self.info_dialog = InfoDialog(info_message + "\n\n" + self.screenshot_message, show_button=False)
            self.info_dialog.show()

    def capture_loop(self, x1, y1, x2, y2):
        while not self.stop_app:
            if keyboard.is_pressed('e'):
                screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                screenshot_path = os.path.join(tempfile.gettempdir(), "temp.png")
                screenshot.save(screenshot_path)

                self.screenshot_signal.emit(screenshot_path)

                self.screenshot_count += 1
                self.screenshot_message = f"Скриншот {self.screenshot_count} сделан"
                if self.info_dialog:
                    current_text = self.info_dialog.label.text().split("\n\n")[0]
                    updated_text = f"{current_text}\n\n{self.screenshot_message}"
                    self.info_dialog.label.setText(updated_text)

                logger.info(f"Сделан скриншот {self.screenshot_count}")

            if keyboard.is_pressed('s'):
                logger.info("Нажата клавиша 'S'. Начинаем сохранение файла.")
                self.save_signal.emit()
                break

            time.sleep(0.01)


    @QtCore.pyqtSlot(str)
    def add_screenshot_to_doc(self, screenshot_path):
        try:
            self.doc.add_picture(screenshot_path, width=Cm(21), height=Cm(13))
            logger.info(f"Скриншот успешно добавлен: {screenshot_path}")
        except Exception as e:
            logger.error(f"Ошибка при добавлении скриншота: {e}")
        finally:
            if os.path.exists(screenshot_path):
                os.remove(screenshot_path)

    @QtCore.pyqtSlot()
    def save_and_exit(self):
        """Сохраняет документ и завершает программу."""
        save_dialog = SaveFileDialog()
        result = save_dialog.exec_()  # Показываем диалоговое окно и проверяем результат
        if result:  # Если пользователь нажал "Сохранить"
            file_path = save_dialog.get_save_file_path()
            if file_path:
                DocumentHandler.save_document(self.doc, file_path)
                if os.path.exists(file_path):  # Проверяем, существует ли файл
                    os.startfile(file_path)  # Открываем файл
                self.stop_app = True  # Завершаем программу
                logger.info(f"Документ сохранен: {file_path}")

                if self.info_dialog:  # Закрываем информационное окно
                    QtWidgets.QApplication.postEvent(
                        self.info_dialog,
                        QtCore.QEvent(QtCore.QEvent.Close)
                    )
        else:  # Если пользователь нажал "Отмена"
            logger.info("Сохранение отменено пользователем")
            # Продолжаем работу программы


class NoAnimationApp:
    def __init__(self):
        self.click_positions = ClickHandler(mode="no_animation")
        self.stop_app = False

        exit_thread = threading.Thread(target=self.check_exit_key, daemon=True)
        exit_thread.start()

    def check_exit_key(self):
        keyboard.wait('q')
        self.stop_app = True
        logger.info("Программа завершена по нажатию клавиши 'Q'.")

    def start_capture(self):
    # Показываем диалоговое окно с инструкциями перед началом захвата
        info_message = (
            "Программа работает в режиме без анимации.\n"
            "После нажатия 'Понятно', программа ожидает три клика:\n"
            "1. Верхняя левая часть слайда.\n"
            "2. Нижняя правая часть слайда.\n"
            "3. Кнопка проигрывания слайда."
        )
        InfoDialog(info_message, show_button=True).exec_()

        with mouse.Listener(on_click=self.click_positions.on_click) as listener:
            listener.join()

        pos1, pos2, pos3 = self.click_positions.positions
        x1, y1 = pos1
        x2, y2 = pos2

        logger.info(f"Координаты скриншота - {pos1}, {pos2}. Координаты перемотки слайда - {pos3}")

        slide_count_dialog = SlideCountDialog()
        if slide_count_dialog.exec_():
            slide_count = slide_count_dialog.get_slide_count()
            logger.info(f"Количество слайдов: {slide_count}")

            doc = DocumentHandler.create_document()

            for i in range(slide_count):
                if self.stop_app:
                    break

                screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                screenshot_path = f"slide_{i + 1}.png"
                screenshot.save(screenshot_path)

                doc.add_picture(screenshot_path, width=Cm(21), height=Cm(13))
                logger.info(f"Слайд {i + 1} добавлен в файл")
                os.remove(screenshot_path)

                pyautogui.click(x=pos3[0], y=pos3[1])  # Клик для перехода к следующему слайду
                time.sleep(0.2)  # Пауза между скриншотами

            file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                None,
                "Сохранить как",
                "",
                "Microsoft Word Documents (*.docx);;PDF Documents (*.pdf)"
            )
            if file_path:  # Проверяем, был ли выбран путь к файлу
                DocumentHandler.save_document(doc, file_path)
                if os.path.exists(file_path):  # Проверяем, существует ли файл
                    os.startfile(file_path)  # Открываем файл
                logger.info(f"Документ сохранен: {file_path}")
                sys.exit()  # Закрываем программу
            else:
                logger.info("Сохранение отменено пользователем")
                # Продолжаем работу программы, если пользователь нажал "Отмена"