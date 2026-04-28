import os
import sys
import fnmatch


def generate_tree(start_dir, max_depth, max_files_per_dir=10, indent_str="  "):
    """
    Рекурсивно строит дерево папок и файлов с ограничением количества файлов в одной директории.
    """
    exclude_dirs = {
        '$RECYCLE.BIN', 'System Volume Information', '.Spotlight-V100', '.Trashes',
        'RECYCLER', 'lost+found', '.git', '.svn', '.hg', 'node_modules',
        'bower_components', '__pycache__', '.venv', 'venv', 'env', '.env_old',
        'virtualenv', 'vendor', 'packages', '.nuget', '.gradle', '.maven',
        'target', 'out', 'build', 'dist', 'bin', 'obj', '.idea', '.vscode',
        '.vs', '.history', '.settings', '.ipynb_checkpoints', 'wandb', 'runs',
        '.matplotlib', '.local', 'CMakeFiles', 'cmake-build-debug', 'cmake-build-release',
        '.docker', '.vagrant', '.cache', '.pytest_cache', '.mypy_cache',
    }

    exclude_patterns = {
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

    def is_hidden(entry):
        if entry.name.startswith('.'):
            return entry.name != '.env'
        if sys.platform == 'win32':
            try:
                import ctypes
                attrs = ctypes.windll.kernel32.GetFileAttributesW(entry.path)
                return attrs != -1 and (attrs & 2 or attrs & 4)
            except:
                return False
        return False

    def should_exclude(name):
        name_lower = name.lower()
        if name_lower in exclude_dirs:
            return True
        for pattern in exclude_patterns:
            if fnmatch.fnmatch(name_lower, pattern):
                return True
        return False

    output = []
    start_dir = os.path.abspath(start_dir)
    root_name = os.path.basename(start_dir) or start_dir
    output.append(f"{root_name}/")

    def walk(current_dir, current_depth):
        if current_depth >= max_depth:
            return

        try:
            # Получаем все элементы и фильтруем их
            entries = []
            for entry in os.scandir(current_dir):
                if not entry.is_symlink() and not is_hidden(entry) and not should_exclude(entry.name):
                    entries.append(entry)

            # Сортируем: сначала папки, потом файлы
            entries.sort(key=lambda e: (not e.is_dir(), e.name.lower()))
        except PermissionError:
            output.append(f"{indent_str * (current_depth + 1)}[Ошибка: Доступ запрещен]")
            return
        except Exception as e:
            output.append(f"{indent_str * (current_depth + 1)}[Ошибка: {e}]")
            return

        # Разделяем на папки и файлы для удобного лимитирования
        dirs = [e for e in entries if e.is_dir()]
        files = [e for e in entries if not e.is_dir()]

        # Сначала обрабатываем все папки (их обычно не так много, как файлов)
        for d in dirs:
            prefix = indent_str * (current_depth + 1)
            output.append(f"{prefix}{d.name}/")
            walk(d.path, current_depth + 1)

        # Затем обрабатываем файлы с учетом лимита
        num_files = len(files)
        prefix = indent_str * (current_depth + 1)

        if num_files > max_files_per_dir:
            # Показываем только первые N файлов
            for f in files[:max_files_per_dir]:
                output.append(f"{prefix}{f.name}")
            # Добавляем инфо-строку о пропущенных файлах
            remaining = num_files - max_files_per_dir
            output.append(f"{prefix}... [и еще {remaining} файл(ов)]")
        else:
            for f in files:
                output.append(f"{prefix}{f.name}")

    walk(start_dir, 0)
    return "\n".join(output)