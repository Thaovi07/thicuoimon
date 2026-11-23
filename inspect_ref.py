import re

with open('ref_app.mjs', encoding='utf-8') as f:
    code = f.read()

print("--- Searching for stages / views / sections ---")
matches = re.findall(r'const\s+(\w+View|STEPS|STAGES|VIEWS|NAV_ITEMS)\s*=', code)
print("Const matches:", matches)

# Find steps array
for line in code.split('\n'):
    if 'label:' in line or 'title:' in line or 'step:' in line or 'stage' in line.lower():
        if len(line.strip()) < 100:
            print(line.strip())
