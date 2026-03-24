import os
from services.signature_processor import process_docx_signatures


def main():
    print("=== Утилита замены ФИО в зонах подписей (РПД) — Рекурсивный поиск ===")

    dir_path = input("Введите путь к папке с файлами .docx: ").strip().strip('"')
    if not os.path.isdir(dir_path):
        print("Ошибка: Путь не найден.")
        return

    old_fio = input("Введите ФИО, которое нужно заменить: ").strip()
    new_fio = input("Введите новое ФИО: ").strip()

    if not old_fio or not new_fio:
        print("Ошибка: ФИО не могут быть пустыми.")
        return

    stats = {"ok": 0, "skip": 0, "err": 0}

    for root, dirs, files in os.walk(dir_path):
        for filename in files:
            if filename.lower().endswith('.docx') and not filename.startswith('~$'):
                full_path = os.path.join(root, filename)

                rel_path = os.path.relpath(full_path, dir_path)

                success, message = process_docx_signatures(full_path, old_fio, new_fio)

                if success:
                    print(f"[OK] {rel_path}")
                    stats["ok"] += 1
                else:
                    if "не найдено" in message:
                        print(f"[SKIP] {rel_path} (ФИО не обнаружено)")
                        stats["skip"] += 1
                    else:
                        print(f"[ERR] {rel_path}: {message}")
                        stats["err"] += 1

    print(f"\nЗавершено!")
    print(f"Обновлено: {stats['ok']}")
    print(f"Пропущено: {stats['skip']}")
    print(f"Ошибок: {stats['err']}")


if __name__ == "__main__":
    main()