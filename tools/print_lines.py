import sys

def main():
    if len(sys.argv) < 2:
        print("usage: print_lines.py <file> [start] [end]")
        return
    path = sys.argv[1]
    start = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    end = int(sys.argv[3]) if len(sys.argv) > 3 else 200
    with open(path, 'rb') as f:
        data = f.read()
    try:
        s = data.decode('utf-8')
    except Exception:
        s = data.decode('latin1')
    lines = s.splitlines()
    for i in range(max(1,start), min(end, len(lines))+1):
        # Avoid console encoding issues by escaping non-ascii
        line = lines[i-1]
        try:
            print(f"{i:4}: {line}")
        except UnicodeEncodeError:
            safe = line.encode('unicode_escape').decode('ascii')
            print(f"{i:4}: {safe}")

if __name__ == '__main__':
    main()
