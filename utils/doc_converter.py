import os
import win32com.client as win32
import logging

logger = logging.getLogger(__name__)

def convert_doc_to_docx(folder_path: str):
    if not os.path.isdir(folder_path):
        logger.error("Папка не существует.")
        return

    word = None
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
    except Exception as e:
        logger.error(f"Ошибка Word: {e}")
        return

    converted_count = 0
    try:
        for root, _, filenames in os.walk(folder_path):
            for filename in filenames:
                if filename.lower().endswith('.doc') and not filename.lower().endswith(
                        '.docx') and not filename.startswith('~$'):
                    doc_path = os.path.abspath(os.path.join(root, filename))
                    docx_path = doc_path + 'x'

                    if os.path.exists(docx_path):
                        continue

                    try:
                        doc = word.Documents.Open(doc_path)
                        doc.SaveAs2(docx_path, FileFormat=16)
                        doc.Close()
                        os.remove(doc_path)
                        converted_count += 1
                    except Exception as e:
                        logger.error(f"Ошибка {filename}: {e}")
    finally:
        if word:
            word.Quit()
        logger.info(f"Конвертировано файлов: {converted_count}")