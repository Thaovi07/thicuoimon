import urllib.request
import json

base_url = 'http://localhost:8000'

def post(endpoint, data={}):
    req = urllib.request.Request(
        f'{base_url}{endpoint}',
        data=json.dumps(data).encode('utf-8'),
        headers={'Content-Type': 'application/json'}
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode('utf-8'))

def get(endpoint):
    req = urllib.request.Request(f'{base_url}{endpoint}')
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode('utf-8'))

print("1. Status:")
print(get('/api/status'))

print("\n2. Load sample:")
s = post('/api/load-sample', {'filename': 'du_lieu_kinh_doanh_mau.csv'})
print("Sample loaded:", s.get('success'), s.get('filename'), s.get('total_rows'))

print("\n3. Clean detect:")
c = post('/api/clean-detect')
print("Clean detect:", c.get('success'), "proposals count:", len(c.get('proposals', [])))

print("\n4. Apply clean:")
ac = post('/api/clean-apply', {
    "strip_space": True,
    "drop_empty": True,
    "drop_duplicate": True,
    "fill_strategy": "mean",
    "selected_ops": ["strip_space", "drop_empty", "drop_duplicate", "fill_missing"]
})
print("Clean apply:", ac.get('success'), "clean rows:", ac.get('clean_rows'))

print("\n5. Audit diff:")
ad = post('/api/audit-diff', {})
print("Audit diff:", ad.get('success'), "diff rows count:", len(ad.get('diff_rows', [])))

print("\n6. Step 4 Analyze:")
an = post('/api/analyze', {})
print("Analyze:", an.get('success'), "business:", an.get('business', {}).get('industry'))

print("\n7. Step 5 Dashboard:")
db = post('/api/filter-dashboard', {})
print("Dashboard filter:", db.get('success'), "kpis count:", len(db.get('kpis', [])))

print("\n8. Step 6 Executive report:")
rp = get('/api/executive-report')
print("Executive report:", rp.get('success'), "charts:", list(rp.get('charts', {}).keys()))
