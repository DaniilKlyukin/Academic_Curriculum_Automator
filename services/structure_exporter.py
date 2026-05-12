import os
import sys
import fnmatch
from typing import Set, List, Any


def generate_tree(
        start_dir: str,
        max_depth: int,
        max_files_per_dir: int = 10,
        indent_str: str = "  "
) -> str:
    """
    Рекурсивно строит текстовое представление дерева папок и файлов.
    Поддерживает фильтрацию системных директорий, скрытых файлов и ограничение глубины обхода.
    """

    exclude_dirs: Set[str] = {
        '$RECYCLE.BIN', 'System Volume Information', '.Spotlight-V100', '.Trashes',
        'RECYCLER', 'lost+found', '.git', '.svn', '.hg', 'node_modules',
        'bower_components', '__pycache__', '.venv', 'venv', 'env', '.env_old',
        'virtualenv', 'vendor', 'packages', '.nuget', '.gradle', '.maven',
        'target', 'out', 'build', 'dist', 'bin', 'obj', '.idea', '.vscode',
        '.vs', '.history', '.settings', '.ipynb_checkpoints', 'wandb', 'runs',
        '.matplotlib', '.local', 'CMakeFiles', 'cmake-build-debug', 'cmake-build-release',
        '.docker', '.vagrant', '.cache', '.pytest_cache', '.mypy_cache',
    }

    exclude_patterns: Set[str] = {
        '.ds_store', 'thumbs.db', 'desktop.ini', 'icon\r', 'package-lock.json',
        'yarn.lock', 'pnpm-lock.yaml', 'composer.lock', 'poetry.lock',
        'gemfile.lock', '.python-version', '.gitignore', '*.o', '*.obj',
        '*.so', '*.a', '*.lib', '*.dll', '*.exe', '*.out', '*.pyc', '*.pyo',
        '*.pyd', '*.class', '*.jar', '*.aux', '*.log', '*.toc', '*.fdb_latexmk',
        '*.fls', '*.synctex.gz', '*.blg', '*.bbl', '*.pdfsync', '*.nav',
        '*.snm', '*.out', '*.swp', '*.swo', '*.bak', '*.old', '*.tmp',
        'npm-debug.log*', 'yarn-debug.log*', 'yarn-error.log*', '.bash_history',
        '.zsh_history', '.python_history', '.ssh', '.Xauthority', '.wget-hsts', '.lesshst'
    }

    def is_hidden(entry: os.DirEntry) -> bool:
        """
        Проверяет, является ли файл или директория скрытой (поддерживает Linux и Windows).
        """
        if entry.name.startswith('.'):
            return entry.name != '.env'
        if sys.platform == 'win32':
            try:
                import ctypes
                attrs: int = ctypes.windll.kernel32.GetFileAttributesW(entry.path)
                return attrs != -1 and (bool(attrs & 2) or bool(attrs & 4))
            except Exception:
                return False
        return False

    def should_exclude(name: str) -> bool:
        """
        Проверяет, попадает ли имя файла/папки в список исключений или паттернов.
        """
        name_lower: str = name.lower()
        if name_lower in exclude_dirs:
            return True
        for pattern in exclude_patterns:
            if fnmatch.fnmatch(name_lower, pattern):
                return True
        return False

    output: List[str] = []
    abs_start_dir: str = os.path.abspath(start_dir)
    root_name: str = os.path.basename(abs_start_dir) or abs_start_dir
    output.append(f"{root_name}/")

    def walk(current_dir: str, current_depth: int) -> None:
        """
        Внутренняя рекурсивная функция для обхода дерева директорий.
        """
        if current_depth >= max_depth:
            return

        try:
            entries: List[os.DirEntry] = []
            with os.scandir(current_dir) as it:
                for entry in it:
                    if not entry.is_symlink() and not is_hidden(entry) and not should_exclude(entry.name):
                        entries.append(entry)

            entries.sort(key=lambda e: (not e.is_dir(), e.name.lower()))
        except PermissionError:
            output.append(f"{indent_str * (current_depth + 1)}[Ошибка: Доступ запрещен]")
            return
        except Exception as e:
            output.append(f"{indent_str * (current_depth + 1)}[Ошибка: {e}]")
            return

        dirs: List[os.DirEntry] = [e for e in entries if e.is_dir()]
        files: List[os.DirEntry] = [e for e in entries if not e.is_dir()]

        for d in dirs:
            prefix: str = indent_str * (current_depth + 1)
            output.append(f"{prefix}{d.name}/")
            walk(d.path, current_depth + 1)

        num_files: int = len(files)
        prefix_files: str = indent_str * (current_depth + 1)

        if num_files > max_files_per_dir:
            for f in files[:max_files_per_dir]:
                output.append(f"{prefix_files}{f.name}")
            remaining: int = num_files - max_files_per_dir
            output.append(f"{prefix_files}... [и еще {remaining} файл(ов)]")
        else:
            for f in files:
                output.append(f"{prefix_files}{f.name}")

    walk(abs_start_dir, 0)
    return "\n".join(output)