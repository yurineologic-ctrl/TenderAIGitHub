#!/usr/bin/env python3
# Скрипт для удаления дублированного кода

with open('lenovo_parser.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Найти первый return r (конец новой функции)
first_return_idx = None
for i in range(len(lines)):
    if i > 100 and lines[i].strip() == 'return r':
        first_return_idx = i
        break

# Найти 'def fetch_detail_page_text'
fetch_func_idx = None
for i in range(len(lines)):
    if 'def fetch_detail_page_text' in lines[i]:
        fetch_func_idx = i
        break

print(f"First return r at line {first_return_idx + 1}")
print(f"fetch_detail_page_text at line {fetch_func_idx + 1}")

# Сохранить только нужные части
with open('lenovo_parser.py', 'w', encoding='utf-8') as f:
    # Пишем до и включая первый return r
    f.writelines(lines[:first_return_idx + 2])
    # Добавляем пустую строку
    f.write('\n')
    # Пишем с функции fetch_detail_page_text до конца
    f.writelines(lines[fetch_func_idx:])

print("Fixed!")
