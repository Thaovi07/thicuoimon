with open('ref_app.mjs', encoding='utf-8') as f:
    code = f.read()

# Let's inspect the main navigation items and functions
import re
print("Functions in ref_app.mjs:")
funcs = re.findall(r'function\s+(\w+)\s*\(', code)
print(funcs[:30])

# Look for export logic
exports = re.findall(r'export|download|pdf|docx|html|csv', code, re.IGNORECASE)
print("Export related terms count:", len(exports))

# Find lines with download or export
for line in code.split('\n'):
    if any(k in line.lower() for k in ['export', 'download', 'pdf', 'docx', 'html', 'print']):
        if len(line.strip()) < 120:
            print("Export line:", line.strip())
