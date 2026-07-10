#!/usr/bin/env python3
"""
Автоматическое исправление всех ошибок в streamlit_app.py
- Исправляет отступы
- Заменяет физические переносы строк в кавычках на \n
- Исправляет docstrings
- Удаляет дубликаты
"""

import re
from pathlib import Path

def fix_all_issues(input_file: str, output_file: str):
    """Главная функция исправления"""
    print(f"📖 Читаю файл: {input_file}")
    
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # ШАГ 1: Исправляем физические переносы строк в строковых литералах
    print("🔧 Шаг 1: Исправляю переносы строк в кавычках...")
    
    # Функция для замены переносов внутри строк
    def fix_string_newlines(match):
        full_match = match.group(0)
        # Если это тройные кавычки - не трогаем
        if full_match.startswith('"""') or full_match.startswith("'''"):
            return full_match
        # Заменяем реальные переносы на \n
        return full_match.replace('\n', '\\n')
    
    # Обрабатываем строки с одинарными и двойными кавычками
    content = re.sub(
        r'(?<!["\'])"(?:[^"\\]|\\.)*"',
        fix_string_newlines,
        content,
        flags=re.DOTALL
    )
    
    content = re.sub(
        r"(?<![\"'])'(?:[^'\\]|\\.)*'",
        fix_string_newlines,
        content,
        flags=re.DOTALL
    )
    
    # ШАГ 2: Исправляем отступы
    print("🔧 Шаг 2: Исправляю отступы...")
    
    lines = content.split('\n')
    result = []
    indent_level = 0
    indent_size = 4
    
    # Ключевые слова для управления отступами
    indent_increase = ['if ', 'elif ', 'else:', 'for ', 'while ', 'try:', 
                       'except:', 'finally:', 'with ', 'def ', 'class ', 
                       'match ', 'case ']
    indent_decrease = ['else:', 'elif ', 'except:', 'finally:', 'case ']
    
    for i, line in enumerate(lines):
        original_line = line
        stripped = line.strip()
        
        # Пустые строки
        if not stripped:
            result.append('')
            continue
        
        # Уменьшаем отступ перед else/except/finally
        if any(stripped.startswith(kw) for kw in indent_decrease):
            indent_level = max(0, indent_level - 1)
        
        # Добавляем отступ
        fixed_line = ' ' * (indent_level * indent_size) + stripped
        result.append(fixed_line)
        
        # Увеличиваем отступ после :
        if stripped.endswith(':') and not stripped.startswith('#'):
            if any(stripped.startswith(kw) for kw in indent_increase):
                indent_level += 1
    
    content = '\n'.join(result)
    
    # ШАГ 3: Исправляем сломанные docstrings
    print("🔧 Шаг 3: Исправляю docstrings...")
    
    # Ищем паттерн: def func(...):\nArgs: (без открывающего """)
    content = re.sub(
        r'(def \w+\([^)]*\)[^:]*:)\s*\n(\s*)(Args:)',
        r'\1\n\2"""\n\2\3',
        content
    )
    
    # ШАГ 4: Удаляем дубликаты функций
    print("🔧 Шаг 4: Удаляю дубликаты...")
    
    # Удаляем дубликат calculate_returns_cost (оставляем первый)
    lines = content.split('\n')
    seen_functions = set()
    result = []
    skip_until_next_def = False
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        
        # Проверяем определение функции
        if stripped.startswith('def '):
            func_name = stripped.split('(')[0].replace('def ', '').strip()
            
            # Специальные случаи: удаляем дубликаты
            if func_name in ['calculate_returns_cost', 'st_dataframe_compat', 
                            'get_marketplace_unit_economics']:
                if func_name in seen_functions:
                    skip_until_next_def = True
                    continue
                else:
                    seen_functions.add(func_name)
                    skip_until_next_def = False
            else:
                skip_until_next_def = False
        
        # Пропускаем строки до следующего def
        if skip_until_next_def:
            if stripped.startswith('def ') or stripped.startswith('class '):
                skip_until_next_def = False
            else:
                continue
        
        result.append(line)
    
    content = '\n'.join(result)
    
    # ШАГ 5: Удаляем if __name__ == "__main__" из середины файла
    print("🔧 Шаг 5: Удаляю if __name__ из середины файла...")
    
    lines = content.split('\n')
    result = []
    skip_main_block = False
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        
        # Проверяем if __name__ == "__main__"
        if stripped == 'if __name__ == "__main__":':
            # Проверяем, не в конце ли файла
            remaining_lines = [l for l in lines[i+1:] if l.strip()]
            if len(remaining_lines) > 50:  # Если после него много кода - удаляем
                skip_main_block = True
                continue
        
        # Восстанавливаем пропуск после следующего блока
        if skip_main_block and stripped and not stripped.startswith('#'):
            if not stripped.startswith(('print', 'loader', 'result', 'comparison', 
                                       'stats', 'for batch', 'prefetch')):
                skip_main_block = False
        
        if not skip_main_block:
            result.append(line)
    
    content = '\n'.join(result)
    
    # Сохраняем результат
    print(f"💾 Сохраняю результат: {output_file}")
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✅ Готово! Исправленный файл: {output_file}")
    print(f"📊 Обработано строк: {len(lines)}")
    
    return output_file

if __name__ == '__main__':
    # Используем файлы
    input_file = 'streamlit_app.py'
    output_file = 'streamlit_app_fixed.py'
    
    try:
        fix_all_issues(input_file, output_file)
        print("\n🎉 Теперь замените оригинальный файл на исправленный!")
    except FileNotFoundError:
        print(f"❌ Файл {input_file} не найден!")
        print("💡 Убедитесь, что файл streamlit_app.py находится в той же папке")
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
