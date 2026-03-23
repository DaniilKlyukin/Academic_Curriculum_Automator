import os
import win32com.client as win32


def convert_doc_to_docx():
    # Запрашиваем путь до папки
    folder_path = input("Введите путь до папки с файлами .doc: ").strip()

    # Убираем кавычки, если путь был скопирован вместе с ними
    folder_path = folder_path.strip('\'"')

    # Проверяем, существует ли папка
    if not os.path.isdir(folder_path):
        print("Ошибка: Указанная папка не существует.")
        return

    print("\nЗапуск Microsoft Word в фоновом режиме...")
    try:
        # Открываем Word в фоновом режиме
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
    except Exception as e:
        print(f"Ошибка при запуске Word: {e}")
        print("Убедитесь, что Microsoft Word установлен на этом компьютере.")
        return

    converted_count = 0

    try:
        # Перебираем все файлы в папке
        for filename in os.listdir(folder_path):
            # Ищем файлы .doc (игнорируем .docx и временные файлы ~$...)
            if filename.lower().endswith('.doc') and not filename.lower().endswith('.docx') and not filename.startswith(
                    '~$'):

                # Формируем абсолютные пути
                doc_path = os.path.abspath(os.path.join(folder_path, filename))
                docx_path = doc_path + 'x'  # Превращаем .doc в .docx

                # Если файл .docx уже существует, пропускаем (чтобы не удалить оригинал без конвертации)
                if os.path.exists(docx_path):
                    print(f"Пропуск: {filename} (файл .docx уже существует)")
                    continue

                print(f"Конвертация: {filename} -> {filename + 'x'} ...", end=" ")

                try:
                    # Открываем документ
                    doc = word.Documents.Open(doc_path)
                    # Сохраняем как .docx (FileFormat=16)
                    doc.SaveAs2(docx_path, FileFormat=16)
                    doc.Close()

                    # --- НОВЫЙ БЛОК: УДАЛЕНИЕ ---
                    os.remove(doc_path)
                    # ----------------------------

                    print("Успешно! (исходник удален)")
                    converted_count += 1
                except Exception as e:
                    print(f"Ошибка при обработке файла! ({e})")

    finally:
        # Обязательно закрываем процесс Word
        word.Quit()
        print(f"\nГотово! Конвертировано и удалено файлов: {converted_count}")


if __name__ == "__main__":
    convert_doc_to_docx()