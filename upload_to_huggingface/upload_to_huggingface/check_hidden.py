with open(r'c:\Users\Hp\Desktop\ORXAN\anbar_botu\inventory_app.py', encoding='utf-8') as f:
    lines = f.readlines()

found = False
for idx in range(3400, len(lines)):
    line = lines[idx]
    non_ascii = [c for c in line if ord(c) > 127]
    if non_ascii:
        safe_line = line.strip().encode('ascii', 'backslashreplace').decode('ascii')
        print(f"Line {idx+1}: {safe_line}")
        found = True

if not found:
    print("Verification success: No non-ASCII characters found in JS block of MOBILE_HTML_CONTENT!")
