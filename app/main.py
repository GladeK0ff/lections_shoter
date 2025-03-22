from PyQt5 import QtWidgets
from models.app_modes import AnimationApp, NoAnimationApp
from handlers.dialog_handler import QuestionDialog
import sys
from logs.logger import logger

if __name__ == "__main__":
    logger.info("Программа начала свою работу")
    app = QtWidgets.QApplication(sys.argv)

    # Задаем вопрос пользователю: есть ли анимация?
    question_dialog = QuestionDialog()
    if question_dialog.exec_():
        if question_dialog.result == "with_animation":
            logger.info("Выбран вариант с анимацией")
            animation_app = AnimationApp()
            animation_app.start_capture()

        elif question_dialog.result == "without_animation":
            logger.info("Выбран вариант без анимации")
            no_animation_app = NoAnimationApp()
            no_animation_app.start_capture()

    logger.info("Программа закончила свою работу")
    sys.exit(app.exec_())
