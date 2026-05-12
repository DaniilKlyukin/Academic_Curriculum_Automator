import logging
from services.doc_converter import convert_doc_to_docx

logging.basicConfig(
    filename='app_errors.log',
    filemode='w',
    level=logging.ERROR,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding='utf-8'
)

def main():
    folder_path = input("Путь к папке с .doc: ").strip().strip('"')
    if not folder_path:
        return
    convert_doc_to_docx(folder_path)

if __name__ == "__main__":
    main()