import py_compile
import sys

try:
    py_compile.compile('lenovo_parser.py', doraise=True)
    print("✓ Syntax OK!")
    sys.exit(0)
except py_compile.PyCompileError as e:
    print("✗ Syntax Error:")
    print(e)
    sys.exit(1)
