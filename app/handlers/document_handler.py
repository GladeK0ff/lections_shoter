# document_handler.py
import os
import win32com.client
from docx import Document
from docx.shared import Cm
from logs.logger import logger


class DocumentHandler:
    @staticmethod
    def save_document(doc, file_path):
        if file_path.lower().endswith('.pdf'):
            temp_docx_path = os.path.splitext(file_path)[0] + ".docx"
            doc.save(temp_docx_path)

            word = win32com.client.Dispatch("Word.Application")
            try:
                doc_word = word.Documents.Open(os.path.abspath(temp_docx_path))
                doc_word.SaveAs(os.path.abspath(file_path), FileFormat=17)
                doc_word.Close()
                logger.info(f"Документ сохранен как PDF: {file_path}")
            except Exception as e:
                logger.error(f"Ошибка при сохранении в PDF: {e}")
            finally:
                word.Quit()

            if os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
        else:
            doc.save(file_path)
            logger.info(f"Документ сохранен как DOCX: {file_path}")

    @staticmethod
    def create_document():
        doc = Document()

        section = doc.sections[0]
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(0.5)
        section.right_margin = Cm(0.5)

        return doc
