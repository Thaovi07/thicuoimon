import urllib.request
import json

csv_content = """Mã đơn,Ngày,Kênh,Khu vực,Doanh thu,Chi phí
ORD01,2026-01-01,Website,Miền Bắc,10000000,6000000
ORD02,2026-01-02,Shopee,Miền Nam,15000000,9000000
ORD03,2026-01-03,Cửa hàng,Miền Trung,8000000,5000000
ORD04,2026-01-04,Tiki,Miền Nam,12000000,7000000
"""

payload = {
    "filename": "test_sales.csv",
    "content": csv_content,
    "type": "csv",
    "business": {}
}

data = json.dumps(payload).encode('utf-8')
req = urllib.request.Request('http://localhost:8000/api/upload-raw', data=data, headers={'Content-Type': 'application/json'})

try:
    with urllib.request.urlopen(req) as resp:
        print("Upload-raw code:", resp.getcode())
        content = resp.read().decode('utf-8')
        print("Upload-raw response:", content[:400])
except Exception as e:
    print("Upload-raw error:", e)
    if hasattr(e, 'read'):
        print("Error content:", e.read().decode('utf-8'))
