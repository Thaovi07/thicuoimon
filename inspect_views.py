with open('ref_app.mjs', encoding='utf-8') as f:
    code = f.read()

import re

for view in ['contextForm', 'sourceView', 'cleanView', 'analysisView', 'exportView', 'discoveryPanel']:
    m = re.search(r'function\s+' + view + r'\s*\([^)]*\)\s*\{', code)
    if m:
        start = m.start()
        print(f"=== VIEW: {view} ===")
        print(code[start:start+1800])
        print("...\n")
