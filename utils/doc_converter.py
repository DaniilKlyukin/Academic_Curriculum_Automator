import os
import win32com.client as win32
from tqdm import tqdm
import logging

logging.basicConfig(level=logging.ERROR)
logger = logging.getLogger(__name__)


def convert_doc_to_docx(folder_path: str):
    if not os.path.isdir(folder_path):
        print("Папка не существует.")
        return

    print("Сканирование папок...")
    files_to_convert = []

    def fast_scan(path):
        try:
            with os.scandir(path) as it:
                for entry in it:
                    if entry.is_dir():
                        fast_scan(entry.path)
                    elif entry.is_file():
                        name = entry.name.lower()
                        if name.endswith('.doc') and not name.endswith('.docx') and not name.startswith('~$'):
                            doc_path = os.path.abspath(entry.path)
                            docx_path = doc_path + 'x'
                            if not os.path.exists(docx_path):
                                files_to_convert.append((doc_path, docx_path, entry.name))
        except PermissionError:
            pass

    fast_scan(folder_path)

    if not files_to_convert:
        print("Файлов для конвертации не найдено.")
        return

    word = None
    converted_count = 0

    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        word.ScreenUpdating = False

        word.Options.CheckGrammarAsYouType = False
        word.Options.CheckSpellingAsYouType = False
        word.Options.BackgroundSave = False

        pbar = tqdm(files_to_convert, unit="file", dynamic_ncols=True)

        for doc_path, docx_path, filename in pbar:
            pbar.set_description(f" {filename[:30]}")

            try:
                doc = word.Documents.Open(doc_path, AddToRecentFiles=False, ReadOnly=True, Visible=False)
                doc.SaveAs2(docx_path, FileFormat=16)
                doc.Close(0)

                os.remove(doc_path)
                converted_count += 1
            except Exception as e:
                tqdm.write(f"Ошибка в {filename}: {e}")

    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")
    finally:
        if word:
            word.ScreenUpdating = True
            word.Quit()
        print(f"\nГотово! Обработано файлов: {converted_count}")