with open('ref_styles.css', encoding='utf-8') as f:
    css = f.read()

import re
root_match = re.search(r':root\s*\{([^}]+)\}', css)
if root_match:
    print("CSS Variables:")
    print(root_match.group(1)[:1200])
