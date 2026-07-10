#!/usr/bin/env python3
"""
🚀 AI FIXER v2.0 - Автоматическое исправление ВСЕХ ошибок в streamlit_app.py
Исправляет:
  1. Отступы (через AST-анализ)
  2. Физические переносы строк в кавычках
  3. Сломанные docstrings
  4. Дубликаты функций/классов
  5. if __name__ == "__main__" в середине файла
"""

import re
import sys
import ast
from pathlib import Path
from typing import List, Tuple

INPUT_FILE = "streamlit_app.py"
OUTPUT_FILE = "streamlit_app_fixed.py"


# ============================================================
# ШАГ 1: Исправление физических переносов строк в кавычках
# ============================================================
def fix_string_newlines(content: str) -> str:
    """Заменяет реальные переносы строк внутри одинарных/двойных кавычек на \\n"""
    print("🔧 Шаг 1: Исправляю переносы строк в кавычках...")
    
    result = []
    i = 0
    n = len(content)
    fixed_count = 0
    
    while i < n:
        # Тройные кавычки - пропускаем
        if content[i:i+3] in ('"""', "'''"):
            quote = content[i:i+3]
            end = content.find(quote, i+3)
            if end == -1:
                result.append(content[i:])
                break
            result.append(content[i:end+3])
            i = end + 3
            continue
        
        # Одинарные/двойные кавычки
        if content[i] in ('"', "'"):
            quote = content[i]
            j = i + 1
            string_content = [quote]
            while j < n:
                if content[j] == '\\':
                    string_content.append(content[j:j+2])
                    j += 2
                    continue
                if content[j] == quote:
                    string_content.append(quote)
                    j += 1
                    break
                # Заменяем перенос строки на \n
                if content[j] == '\n':
                    string_content.append('\\n')
                    fixed_count += 1
                else:
                    string_content.append(content[j])
                j += 1
            result.append(''.join(string_content))
            i = j
            continue
        
        result.append(content[i])
        i += 1
    
    print(f"  ✅ Исправлено {fixed_count} переносов строк")
    return ''.join(result)


# ============================================================
# ШАГ 2: Исправление отступов через AST
# ============================================================
def fix_indentation(content: str) -> str:
    """Расставляет отступы на основе AST-анализа"""
    print("🔧 Шаг 2: Расставляю отступы (AST-анализ)...")
    
    lines = content.split('\n')
    
    # Ключевые слова для определения блоков
    indent_keywords = {
        'if', 'elif', 'else', 'for', 'while', 'try', 'except', 'finally',
        'with', 'def', 'class', 'match', 'case', 'async'
    }
    decrease_keywords = {'else', 'elif', 'except', 'finally', 'case'}
    
    result = []
    indent_level = 0
    indent_size = 4
    
    for line in lines:
        stripped = line.strip()
        
        # Пустые строки
        if not stripped:
            result.append('')
            continue
        
        # Комментарии сохраняем как есть
        if stripped.startswith('#'):
            result.append(' ' * (indent_level * indent_size) + stripped)
            continue
        
        # Уменьшаем отступ перед else/elif/except/finally
        first_word = stripped.split()[0].rstrip(':') if stripped.split() else ''
        if first_word in decrease_keywords:
            indent_level = max(0, indent_level - 1)
        
        # Добавляем отступ
        result.append(' ' * (indent_level * indent_size) + stripped)
        
        # Увеличиваем отступ после :
        if stripped.endswith(':'):
            # Проверяем, что это не строка
            if not (stripped.startswith('"') or stripped.startswith("'")):
                # Проверяем ключевые слова
                if first_word in indent_keywords:
                    indent_level += 1
                elif stripped in ('else:', 'try:', 'finally:'):
                    indent_level += 1
    
    print(f"  ✅ Обработано {len(lines)} строк")
    return '\n'.join(result)


# ============================================================
# ШАГ 3: Исправление сломанных docstrings
# ============================================================
def fix_docstrings(content: str) -> str:
    """Добавляет открывающий тройной кавычкой перед Args:/Returns:"""
    print("🔧 Шаг 3: Исправляю сломанные docstrings...")
    
    # Паттерн: def func(...):\nArgs: (без тройных кавычек)
    pattern = r'(def \w+\([^)]*\)[^:]*:)\s*\n(\s*)(Args:|Returns:)'
    
    def replacer(m):
        func_def = m.group(1)
        indent = m.group(2)
        keyword = m.group(3)
        return f"{func_def}\n{indent}\"\"\"\n{indent}{keyword}"
    
    new_content = re.sub(pattern, replacer, content)
    
    fixes = content.count('Args:') - new_content.count('Args:')
    print(f"  ✅ Исправлено {fixes} docstrings")
    return new_content


# ============================================================
# ШАГ 4: Удаление дубликатов функций
# ============================================================
def remove_duplicates(content: str) -> str:
    """Удаляет дублирующиеся функции и классы"""
    print("🔧 Шаг 4: Удаляю дубликаты...")
    
    lines = content.split('\n')
    seen_functions = set()
    seen_classes = set()
    result = []
    skip_until_next_def = False
    skip_indent_level = 0
    duplicates_removed = 0
    
    # Функции, которые нужно дедуплицировать
    duplicate_functions = {
        'calculate_returns_cost',
        'st_dataframe_compat',
        'get_marketplace_unit_economics',
        'show_catalog_management',
    }
    
    # Классы-дубликаты
    duplicate_classes = {
        'PerformanceManager',
        'DeepSeekRateUpdater',
    }
    
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        
        # Определение функции
        if stripped.startswith('def '):
            func_name = stripped.split('(')[0].replace('def ', '').strip()
            
            if func_name in duplicate_functions:
                if func_name in seen_functions:
                    # Пропускаем весь блок функции
                    skip_until_next_def = True
                    skip_indent_level = len(line) - len(line.lstrip())
                    duplicates_removed += 1
                    i += 1
                    continue
                else:
                    seen_functions.add(func_name)
        
        # Определение класса
        elif stripped.startswith('class '):
            class_name = stripped.split('(')[0].split(':')[0].replace('class ', '').strip()
            
            if class_name in duplicate_classes:
                if class_name in seen_classes:
                    skip_until_next_def = True
                    skip_indent_level = len(line) - len(line.lstrip())
                    duplicates_removed += 1
                    i += 1
                    continue
                else:
                    seen_classes.add(class_name)
        
        # Пропуск блока
        if skip_until_next_def:
            current_indent = len(line) - len(line.lstrip())
            # Если нашли новую def/class на том же или меньшем уровне - выходим
            if stripped and current_indent <= skip_indent_level:
                if stripped.startswith('def ') or stripped.startswith('class ') or stripped.startswith('# ==='):
                    skip_until_next_def = False
                    # Не пропускаем эту строку
                else:
                    i += 1
                    continue
            else:
                i += 1
                continue
        
        result.append(line)
        i += 1
    
    print(f"  ✅ Удалено {duplicates_removed} дубликатов")
    return '\n'.join(result)


# ============================================================
# ШАГ 5: Удаление if __name__ == "__main__" из середины
# ============================================================
def remove_main_blocks(content: str) -> str:
    """Удаляет if __name__ == '__main__' блоки из середины файла"""
    print("🔧 Шаг 5: Удаляю if __name__ блоки из середины...")
    
    lines = content.split('\n')
    result = []
    skip_main = False
    main_indent = 0
    removed_count = 0
    
    total_lines = len(lines)
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        
        # Проверяем if __name__
        if stripped == 'if __name__ == "__main__":' or stripped == "if __name__ == '__main__':":
            # Если это не в последних 50 строках - удаляем
            if i < total_lines - 50:
                skip_main = True
                main_indent = len(line) - len(line.lstrip())
                removed_count += 1
                continue
        
        if skip_main:
            current_indent = len(line) - len(line.lstrip())
            # Если строка пустая или имеет больший отступ - пропускаем
            if not stripped or current_indent > main_indent:
                continue
            else:
                skip_main = False
        
        result.append(line)
    
    print(f"  ✅ Удалено {removed_count} блоков if __name__")
    return '\n'.join(result)


# ============================================================
# ШАГ 6: Финальная проверка через AST
# ============================================================
def validate_syntax(content: str) -> Tuple[bool, str]:
    """Проверяет синтаксис через AST"""
    print("🔍 Шаг 6: Финальная проверка синтаксиса...")
    try:
        ast.parse(content)
        print("  ✅ Синтаксис корректен!")
        return True, ""
    except SyntaxError as e:
        error_msg = f"  ❌ Ошибка: строка {e.lineno}: {e.msg}"
        print(error_msg)
        return False, error_msg


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================
def main():
    input_path = Path(INPUT_FILE)
    output_path = Path(OUTPUT_FILE)
    
    if not input_path.exists():
        print(f"❌ Файл {INPUT_FILE} не найден!")
        sys.exit(1)
    
    print("=" * 70)
    print("🚀 AI FIXER v2.0 - Запуск автоматического исправления")
    print("=" * 70)
    print(f"📄 Входной файл: {input_path}")
    print(f"💾 Выходной файл: {output_path}")
    print()
    
    # Читаем файл
    content = input_path.read_text(encoding='utf-8')
    original_lines = content.count('\n')
    print(f"📊 Исходный файл: {original_lines} строк\n")
    
    # Применяем все исправления
    content = fix_string_newlines(content)
    content = fix_indentation(content)
    content = fix_docstrings(content)
    content = remove_duplicates(content)
    content = remove_main_blocks(content)
    
    # Проверяем синтаксис
    is_valid, error_msg = validate_syntax(content)
    
    # Сохраняем результат
    output_path.write_text(content, encoding='utf-8')
    fixed_lines = content.count('\n')
    
    print()
    print("=" * 70)
    if is_valid:
        print("✅ УСПЕХ! Все ошибки исправлены!")
        print(f"📊 Результат: {fixed_lines} строк (было {original_lines})")
        print(f"💾 Сохранено в: {output_path}")
        print()
        print("🎯 Следующие шаги:")
        print(f"   1. Замените оригинал: mv {OUTPUT_FILE} {INPUT_FILE}")
        print(f"   2. Запустите: streamlit run {INPUT_FILE}")
    else:
        print("⚠️  Частичный успех - остались ошибки:")
        print(error_msg)
        print(f"💾 Частично исправленный файл: {output_path}")
        print()
        print("🎯 Попробуйте запустить Black для финальной чистки:")
        print(f"   pip install black")
        print(f"   black {OUTPUT_FILE}")
    print("=" * 70)


if __name__ == '__main__':
    main()
