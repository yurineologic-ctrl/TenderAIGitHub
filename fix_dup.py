#!/usr/bin/env python3
# Fix lenovo_parser.py - remove duplicate code

with open('lenovo_parser.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find first 'return r' (end of new parse_specs)
lines = content.split('\n')
first_return_idx = None
for i, line in enumerate(lines):
    if i > 300 and line.strip() == 'return r':
        first_return_idx = i
        break

# Find 'def fetch_detail_page_text'
fetch_idx = None
for i, line in enumerate(lines):
    if 'def fetch_detail_page_text' in line:
        fetch_idx = i
        break

print(f"First return at line {first_return_idx + 1}")
print(f"fetch_detail at line {fetch_idx + 1}")

# Reconstruct
new_lines = lines[:first_return_idx + 1] + ['', ''] + lines[fetch_idx:]

with open('lenovo_parser.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(new_lines))

print("Fixed!")
