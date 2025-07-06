import sys
import codecs

if len(sys.argv) != 3:
    print("Usage: python convert_encoding.py <file_path> <target_encoding>")
    sys.exit(1)

file_path = sys.argv[1]
target_encoding = sys.argv[2]

try:
    # Try reading with 'utf-8' first, then 'cp932' (Shift-JIS for Windows)
    try:
        with codecs.open(file_path, 'r', 'utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        with codecs.open(file_path, 'r', 'cp932') as f:
            content = f.read()

    with codecs.open(file_path, 'w', target_encoding) as f:
        f.write(content)
    print(f"Successfully converted {file_path} to {target_encoding}")
except Exception as e:
    print(f"Error converting {file_path}: {e}")
    sys.exit(1)