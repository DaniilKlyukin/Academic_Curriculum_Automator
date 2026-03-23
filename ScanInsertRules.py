from WordDocumentEditor import WordDocumentEditor

from typing import Optional


# Примеры правил вставки
def first_page_resolver(editor: WordDocumentEditor) -> int:
    return 1


def second_page_resolver(editor: WordDocumentEditor) -> int:
    return 2


def coordination_page_resolver(editor: WordDocumentEditor) -> Optional[int]:
    return editor.find_page_with_text("Лист согласования")
