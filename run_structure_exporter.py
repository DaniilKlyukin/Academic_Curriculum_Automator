import os
from utils.structure_exporter import generate_tree


def main():
    print("--- Генератор структуры папок для ИИ (эффективный режим) ---")

    # 1. Путь
    target_path = input("Введите путь к папке (Enter для текущей): ").strip()
    if not target_path:
        target_path = os.getcwd()

    if not os.path.isdir(target_path):
        print(f"Ошибка: Путь '{target_path}' не найден.")
        return

    # 2. Глубина
    try:
        depth_str = input("Глубина поиска (по умолчанию 2): ").strip()
        max_depth = int(depth_str) if depth_str else 2
    except ValueError:
        max_depth = 2

    # 3. Лимит файлов
    try:
        limit_str = input("Макс. файлов в одной папке (по умолчанию 10): ").strip()
        max_files = int(limit_str) if limit_str else 10
    except ValueError:
        max_files = 10

    print(f"\nАнализирую структуру: {target_path}...")

    # Генерируем дерево
    tree_text = generate_tree(target_path, max_depth, max_files_per_dir=max_files)

    output_filename = "project_structure_for_ai.txt"
    save_path = os.path.join(target_path, output_filename)

    try:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(tree_text)

        print("-" * 40)
        print(f"✅ Готово!")
        print(f"📁 Файл сохранен: {save_path}")
        print(f"📄 Строк в отчете: {len(tree_text.splitlines())}")
        print(f"💡 Если файлов было слишком много, они были сокращены до {max_files} на папку.")
        print("-" * 40)
    except Exception as e:
        print(f"Ошибка при записи файла: {e}")


if __name__ == "__main__":
    main()