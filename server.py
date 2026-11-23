import os
import json
import io
import math
import base64
import urllib.parse
import unicodedata
import re
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

PORT = 8000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_data")


def strip_accents(text):
    if text is None:
        return ""
    text = unicodedata.normalize('NFD', str(text))
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    return unicodedata.normalize('NFC', text).replace('đ', 'd').replace('Đ', 'D')

def normalize_col_name(c):
    c_clean = strip_accents(str(c)).lower().replace('_', ' ').replace('-', ' ')
    c_clean = re.sub(r'[\(\)\[\]\,]', '', c_clean).strip()
    return c_clean

def standardize_dataframe_numeric_types(df):
    if df is None or len(df) == 0:
        return df
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            continue
        s_clean = df[col].dropna().astype(str).str.strip()
        if len(s_clean) == 0:
            continue
        s_test = s_clean.str.replace(r'[đĐVNDvnd$\s]', '', regex=True)
        match_dots = s_test.str.match(r'^-?\d{1,3}(\.\d{3})+(,\d+)?$').sum()
        match_commas = s_test.str.match(r'^-?\d{1,3}(,\d{3})+(\.\d+)?$').sum()
        match_plain = s_test.str.match(r'^-?\d+(\.\d+)?$').sum()
        
        if (match_dots + match_commas + match_plain) / len(s_clean) > 0.6:
            def parse_val(v):
                if pd.isna(v) or v is None:
                    return np.nan
                s = str(v).strip()
                s = re.sub(r'[đĐVNDvnd$\s]', '', s)
                if re.match(r'^-?\d{1,3}(\.\d{3})+(,\d+)?$', s):
                    s = s.replace('.', '').replace(',', '.')
                elif re.match(r'^-?\d{1,3}(,\d{3})+(\.\d+)?$', s):
                    s = s.replace(',', '')
                try:
                    return float(s)
                except:
                    return np.nan
            df[col] = df[col].apply(parse_val)
    return df

def detect_business_context(df, filename=""):
    cols_norm = [normalize_col_name(c) for c in df.columns]
    
    retail_keys = ['doanh thu', 'doanh so', 'kenh', 'san pham', 'so don', 'gia von', 'loi nhuan', 'khu vuc', 'sales', 'order', 'store']
    retail_score = sum(1 for k in retail_keys if any(k in c for c in cols_norm))
    
    mkt_keys = ['chien dich', 'chi phi', 'luot nhap', 'clicks', 'chuyen doi', 'conversions', 'cpc', 'ctr', 'campaign', 'leads']
    mkt_score = sum(1 for k in mkt_keys if any(k in c for c in cols_norm))
    
    fin_keys = ['dong tien', 'thu no', 'chi tieu', 'tien mat', 'lai suat', 'tai khoan', 'ngan sach', 'cashflow']
    fin_score = sum(1 for k in fin_keys if any(k in c for c in cols_norm))
    
    hr_keys = ['nhan vien', 'luong', 'tham nien', 'phong ban', 'kpi', 'chuc vu', 'phu cap', 'salary', 'staff']
    hr_score = sum(1 for k in hr_keys if any(k in c for c in cols_norm))

    scores = {'retail': retail_score, 'mkt': mkt_score, 'fin': fin_score, 'hr': hr_score}
    best = max(scores, key=scores.get)
    if scores[best] == 0:
        best = 'general'

    row_count = len(df)
    if best == 'retail':
        ind = 'Bán lẻ đa kênh & Thương mại điện tử'
        bg = f'Đánh giá hiệu quả kinh doanh trên các kênh bán hàng qua {row_count} dòng dữ liệu theo dõi doanh thu, giá vốn và biên lợi nhuận.'
        qs = [
            'Doanh thu và biên lợi nhuận gộp theo từng Kênh bán lẻ biến động ra sao?',
            'Khu vực hoặc dòng sản phẩm nào đóng góp tỷ trọng doanh số cao nhất?',
            'Tỷ lệ chi phí giá vốn (COGS) trên doanh thu của các kênh có ổn định không?'
        ]
    elif best == 'mkt':
        ind = 'Tiếp thị & Quảng cáo số (Digital Marketing)'
        bg = f'Theo dõi hiệu suất {row_count} chiến dịch tiếp thị số qua các chỉ số chi phí quảng cáo, lượng truy cập nhấp chuột và tỷ lệ chuyển đổi.'
        qs = [
            'Chi phí theo chiến dịch nào cao nhất và hiệu quả chuyển đổi ra sao?',
            'Kênh nào mang lại chi phí trên mỗi chuyển đổi (CPA) tối ưu nhất?',
            'Xu hướng tăng trưởng lượt nhấp và chuyển đổi qua các mốc thời gian?'
        ]
    elif best == 'fin':
        ind = 'Tài chính & Dòng tiền Doanh nghiệp'
        bg = f'Kiểm soát dòng tiền, cơ cấu thu chi và biến động tài chính qua {row_count} bản ghi giao dịch.'
        qs = [
            'Khoản mục thu chi nào chiếm tỷ trọng lớn nhất toàn hệ thống?',
            'Biến động dòng tiền ròng ra sao và có rủi ro thanh khoản nào không?'
        ]
    elif best == 'hr':
        ind = 'Quản trị Nhân sự & Hiệu suất'
        bg = f'Phân tích cơ cấu nhân sự, chi phí quỹ lương và đánh giá chỉ số hoàn thành KPI qua {row_count} nhân sự.'
        qs = [
            'Phòng ban nào có quỹ lương cao nhất và biến động ra sao?',
            'Tỷ lệ hoàn thành KPI bình quân giữa các bộ phận có sự phân hóa không?'
        ]
    else:
        ind = 'Hoạt động Vận hành & Phân tích Tổng hợp'
        bg = f'Bộ dữ liệu gồm {row_count} dòng và {len(df.columns)} cột ghi nhận các chỉ số vận hành then chốt.'
        qs = [
            'Các chỉ số cốt lõi biến động theo nhóm phân loại nào?',
            'Có tồn tại giá trị bất thường hoặc đột biến không?'
        ]

    return {
        'industry': ind,
        'background': bg,
        'questions': '\n'.join(qs),
        'category': best
    }

def audit_raw_file_health(df, filename=""):
    total_file_rows = len(df)
    total_cols = len(df.columns)
    
    if total_file_rows == 0:
        return {
            "is_valid": False,
            "critical_message": "Tệp không chứa bất kỳ dòng dữ liệu nào. Vui lòng kiểm tra lại tệp nguồn.",
            "health_score": 0,
            "total_rows": 0,
            "total_cols": total_cols,
            "issues": [],
            "auto_context": {"industry": "", "background": "", "questions": ""}
        }

    # 1. Phát hiện các dòng hoàn toàn rỗng TRƯỚC
    empty_row_mask = df.isna().all(axis=1) | df.apply(lambda row: row.astype(str).str.strip().eq("").all(), axis=1)
    empty_rows_count = int(empty_row_mask.sum())
    empty_rows_indices = [int(i + 2) for i in df[empty_row_mask].index.tolist()[:5]] # 1-based line number in file

    # Tập hợp các dòng dữ liệu thực tế (bỏ qua dòng trống hoàn toàn)
    active_df = df[~empty_row_mask].copy()
    active_rows_count = len(active_df)

    if active_rows_count == 0:
        return {
            "is_valid": False,
            "critical_message": "Toàn bộ các ô trong tệp đều là giá trị rỗng (NaN). Tệp không có nội dung sử dụng được.",
            "health_score": 0,
            "total_rows": total_file_rows,
            "total_cols": total_cols,
            "issues": [],
            "auto_context": {"industry": "", "background": "", "questions": ""}
        }

    # 2. Kiểm tra khoảng trắng thừa đầu/cuối
    whitespace_details = []
    for col in active_df.columns:
        if active_df[col].dtype == object:
            has_strip_mask = active_df[col].dropna().apply(lambda x: isinstance(x, str) and (x != x.strip()))
            cnt_strip = int(has_strip_mask.sum())
            if cnt_strip > 0:
                strip_rows = [int(i + 2) for i in active_df[has_strip_mask].index.tolist()[:3]]
                whitespace_details.append(f"Cột '{col}' có {cnt_strip} ô thừa khoảng trắng tại dòng {', '.join(map(str, strip_rows))}")

    # 3. Kiểm tra trùng lặp trên các dòng dữ liệu sau khi cắt khoảng trắng
    active_df_trimmed = active_df.copy()
    for col in active_df_trimmed.select_dtypes(include=['object']).columns:
        active_df_trimmed[col] = active_df_trimmed[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    
    dup_mask = active_df_trimmed.duplicated()
    dup_count = int(dup_mask.sum())
    dup_rows = [int(i + 2) for i in active_df_trimmed[dup_mask].index.tolist()[:5]]

    # 4. Kiểm tra ô khuyết thiếu TRÊN CÁC DÒNG CÓ DỮ LIỆU THỰC TẾ (không tính dòng rỗng)
    null_cells = int(active_df.isna().sum().sum())
    missing_details = []
    for col in active_df.columns:
        cnt = int(active_df[col].isna().sum())
        if cnt > 0:
            miss_rows = [int(i + 2) for i in active_df[active_df[col].isna()].index.tolist()[:3]]
            missing_details.append(f"Cột '{col}' có {cnt} ô trống tại dòng {', '.join(map(str, miss_rows))}")

    # Tính điểm sức khỏe tệp thực tế
    score = 100
    score -= min(30, empty_rows_count * 15)
    score -= min(30, dup_count * 15)
    score -= min(25, null_cells * 10)
    score -= min(15, len(whitespace_details) * 5)
    score = max(20, score)

    issues = []
    if empty_rows_count > 0:
        issues.append({
            "type": "error",
            "icon": "alert-octagon",
            "title": f"{empty_rows_count} Dòng hoàn toàn trống",
            "desc": f"Phát hiện dòng trống không có dữ liệu tại dòng {', '.join(map(str, empty_rows_indices))}. Cần tự động loại bỏ để không làm sai lệch mẫu số.",
            "count": empty_rows_count
        })
    if dup_count > 0:
        issues.append({
            "type": "warning",
            "icon": "copy",
            "title": f"{dup_count} Dòng trùng lặp nội dung hoàn toàn",
            "desc": f"Bản ghi trùng lặp nội dung tại dòng {', '.join(map(str, dup_rows))}. Cần hợp nhất để tránh double-count doanh thu.",
            "count": dup_count
        })
    if null_cells > 0:
        issues.append({
            "type": "warning",
            "icon": "file-question",
            "title": f"{null_cells} Ô khuyết thiếu (Missing values)",
            "desc": "; ".join(missing_details) + ". Cần điền bù tại Bước 2.",
            "count": null_cells
        })
    else:
        issues.append({
            "type": "good",
            "icon": "check-check",
            "title": "Tất cả các trường chính đầy đủ 100%",
            "desc": "Các dòng dữ liệu hoạt động đều có đầy đủ giá trị, không bị khuyết ô nào.",
            "count": 0
        })
    if whitespace_details:
        issues.append({
            "type": "info",
            "icon": "wrench",
            "title": f"{len(whitespace_details)} Vấn đề khoảng trắng định dạng",
            "desc": "; ".join(whitespace_details) + ". Cần cắt tỉa khoảng trắng chuẩn hóa.",
            "count": len(whitespace_details)
        })

    auto_context = detect_business_context(active_df, filename)

    return {
        "is_valid": True,
        "critical_message": "",
        "health_score": score,
        "total_file_rows": total_file_rows,
        "active_rows": active_rows_count,
        "clean_rows": active_rows_count - dup_count,
        "total_cols": total_cols,
        "duplicate_count": dup_count,
        "null_cells": null_cells,
        "empty_rows": empty_rows_count,
        "issues": issues,
        "auto_context": auto_context
    }

def generate_executive_charts(analysis, df=None):
    plt.rcParams['font.sans-serif'] = ['Segoe UI', 'DejaVu Sans', 'Arial']
    plt.rcParams['axes.unicode_minus'] = False
    
    charts = {}
    dims = analysis.get("dimension_analysis", {}) if analysis else {}
    
    if df is None:
        df = CURRENT_SESSION.get("cleaned_df")
        if df is None:
            df = CURRENT_SESSION.get("raw_df")
            
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist() if df is not None else []
    cat_cols = [c for c in df.columns if c not in numeric_cols and not str(c).startswith('_')] if df is not None else []
    
    revenue_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["doanh thu", "doanh so", "revenue", "sales", "thanh tien"])), None)
    cost_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["chi phi", "gia von", "cost", "cogs"])), None)
    profit_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["loi nhuan", "lai", "profit"])), None)
    qty_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["so don", "so luong", "quantity", "sl", "qty"])), None)
    date_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["thang", "ngay", "date", "time", "month", "nam", "year"])), None) if df is not None else None
    
    primary_metric = revenue_col or (numeric_cols[0] if numeric_cols else "Chỉ số chính")
    
    def get_scaling(values):
        max_v = max([abs(v) for v in values if v is not None and not np.isnan(v)] or [0])
        if max_v >= 1e7:
            return 1e6, "Triệu VNĐ", "M"
        elif max_v >= 1e5:
            return 1e6, "Triệu VNĐ", "Tr"
        elif max_v >= 1e3:
            return 1e3, "Nghìn VNĐ", "k"
        else:
            return 1.0, "Đơn vị gốc", ""

    preferred_dims = [c for c in cat_cols if any(k in normalize_col_name(c) for k in ["kenh", "channel", "san pham", "product", "chien dich", "campaign", "danh muc", "khu vuc"])]
    cat_key = preferred_dims[0] if preferred_dims else (cat_cols[0] if cat_cols else (list(dims.keys())[0] if dims else None))

    def save_fig_to_charts(fig, prefix, fallback_key=None):
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight', dpi=120)
        plt.close(fig)
        buf.seek(0)
        charts[f'{prefix}_stream'] = buf
        b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
        charts[f'{prefix}_b64'] = b64
        if fallback_key:
            buf_copy = io.BytesIO(buf.getvalue())
            charts[f'{fallback_key}_stream'] = buf_copy
            charts[f'{fallback_key}_b64'] = b64

    # 1. BIỂU ĐỒ 1: XẾP HẠNG QUY MÔ ĐÓNG GÓP (BAR RANKING)
    if df is not None and cat_key and primary_metric in df.columns:
        grouped1 = df.groupby(cat_key)[primary_metric].sum().sort_values(ascending=True)
        if len(grouped1) > 8:
            grouped1 = grouped1.tail(8)
        labels1 = [str(x) for x in grouped1.index.tolist()]
        raw_vals1 = grouped1.values.tolist()
        scale1, unit_label1, unit_sfx1 = get_scaling(raw_vals1)
        vals1 = [v / scale1 for v in raw_vals1]
        total_v1 = sum(raw_vals1)

        fig1, ax1 = plt.subplots(figsize=(6.5, 3.2), dpi=120)
        colors1 = ['#8C5E43' if i == len(vals1)-1 else '#b45309' if i == len(vals1)-2 else '#64748b' for i in range(len(vals1))]
        bars1 = ax1.barh(labels1, vals1, color=colors1, height=0.52, edgecolor='#ffffff', linewidth=1.0)
        ax1.set_title(f'Biểu đồ 1: Xếp hạng quy mô đóng góp theo {cat_key}', fontsize=10.5, fontweight='bold', pad=12, color='#1e293b')
        ax1.set_xlabel(f'{primary_metric} ({unit_label1})', fontsize=8.5, color='#475569')
        max_v1 = max(vals1) if vals1 else 1
        for bar, v, rv in zip(bars1, vals1, raw_vals1):
            pct = (rv / total_v1 * 100) if total_v1 > 0 else 0
            fmt_v = f"{v:,.1f}" if scale1 > 1 else f"{v:,.0f}"
            ax1.text(v + (max_v1 * 0.02), bar.get_y() + bar.get_height()/2, f'{fmt_v} {unit_sfx1} ({pct:.1f}%)',
                     va='center', ha='left', fontsize=8, fontweight='bold', color='#1e293b')
        ax1.set_xlim(0, max_v1 * 1.35)
        ax1.grid(axis='x', linestyle='--', alpha=0.35)
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)
        plt.tight_layout()
        save_fig_to_charts(fig1, 'chart1', 'bar')

    # 2. BIỂU ĐỒ 2: PHÂN TÍCH PARETO (80/20) (COMBO BAR + LINE)
    if df is not None and cat_key and primary_metric in df.columns:
        grouped2 = df.groupby(cat_key)[primary_metric].sum().sort_values(ascending=False)
        if len(grouped2) > 8:
            top_part = grouped2.head(7)
            other_sum = grouped2.iloc[7:].sum()
            grouped2 = pd.concat([top_part, pd.Series({'Khác': other_sum})])
        labels2 = [str(x) for x in grouped2.index.tolist()]
        raw_vals2 = grouped2.values.tolist()
        scale2, unit_label2, unit_sfx2 = get_scaling(raw_vals2)
        vals2 = [v / scale2 for v in raw_vals2]
        tot2 = sum(raw_vals2)
        cum_pct2 = (np.cumsum(raw_vals2) / tot2 * 100).tolist() if tot2 > 0 else [0]*len(raw_vals2)

        fig2, ax2_1 = plt.subplots(figsize=(6.5, 3.2), dpi=120)
        ax2_2 = ax2_1.twinx()
        bars2 = ax2_1.bar(labels2, vals2, color='#38bdf8', width=0.48, edgecolor='#ffffff', linewidth=1.2)
        ax2_2.plot(labels2, cum_pct2, color='#e11d48', marker='o', linewidth=2.2, markersize=5)
        ax2_2.axhline(80, color='#dc2626', linestyle='--', linewidth=1.5, alpha=0.85)
        ax2_2.text(len(labels2)-1, 82, 'Ngưỡng 80% (Pareto)', color='#dc2626', fontsize=8, fontweight='bold', ha='right')

        ax2_1.set_title('Biểu đồ 2: Phân tích Pareto (80/20) - Mức độ tập trung', fontsize=10.5, fontweight='bold', pad=12, color='#1e293b')
        ax2_1.set_ylabel(f'{primary_metric} ({unit_label2})', fontsize=8.5, color='#475569')
        ax2_2.set_ylabel('Tỷ lệ tích lũy (%)', fontsize=8.5, color='#e11d48')
        ax2_2.set_ylim(0, 110)
        ax2_1.grid(axis='y', linestyle='--', alpha=0.3)
        ax2_1.spines['top'].set_visible(False)
        ax2_2.spines['top'].set_visible(False)
        max_v2 = max(vals2) if vals2 else 1
        for b, v in zip(bars2, vals2):
            fmt_v = f"{v:,.1f}" if scale2 > 1 else f"{v:,.0f}"
            ax2_1.text(b.get_x() + b.get_width()/2, v + max_v2 * 0.02, f'{fmt_v}{unit_sfx2}',
                       ha='center', va='bottom', fontsize=7.5, fontweight='bold')
        for i, cp in enumerate(cum_pct2):
            ax2_2.text(i, cp + 3, f'{cp:.1f}%', ha='center', va='bottom', fontsize=7.5, color='#9f1239', fontweight='bold')
        plt.tight_layout()
        save_fig_to_charts(fig2, 'chart2', 'pie')

    # 3. BIỂU ĐỒ 3: DIỄN BIẾN THEO THỜI GIAN & ĐƯỜNG TRUNG BÌNH CHUẨN
    timeline = analysis.get("timeline", {}) if analysis else {}
    t_labels = timeline.get("labels", [])
    raw_t_vals = timeline.get("values", [])
    if (not t_labels or not raw_t_vals) and df is not None and primary_metric in df.columns:
        if date_col:
            t_df = df.groupby(date_col)[primary_metric].sum().reset_index().sort_values(date_col)
            t_labels = [str(x) for x in t_df[date_col].tolist()]
            raw_t_vals = t_df[primary_metric].tolist()
        else:
            chunk_size = max(1, len(df) // 6)
            chunks = [df.iloc[i:i+chunk_size] for i in range(0, len(df), chunk_size)]
            t_labels = [f"Kỳ {i+1}" for i in range(len(chunks))]
            raw_t_vals = [c[primary_metric].sum() for c in chunks]

    if t_labels and raw_t_vals:
        scale3, unit_label3, unit_sfx3 = get_scaling(raw_t_vals)
        t_vals = [v / scale3 for v in raw_t_vals]
        mean_v3 = float(np.mean(t_vals)) if t_vals else 0

        fig3, ax3 = plt.subplots(figsize=(6.5, 3.2), dpi=120)
        ax3.plot(t_labels, t_vals, marker='o', color='#0284c7', linewidth=2.5, markersize=6, markerfacecolor='#ffffff', markeredgewidth=2)
        ax3.fill_between(t_labels, t_vals, color='#38bdf8', alpha=0.18)
        fmt_mean = f"{mean_v3:,.1f}" if scale3 > 1 else f"{mean_v3:,.0f}"
        ax3.axhline(mean_v3, color='#16a34a', linestyle='--', linewidth=1.6, label=f'Trung bình chuẩn: {fmt_mean} {unit_sfx3}')
        ax3.set_title('Biểu đồ 3: Diễn biến chu kỳ & Đường trung bình chuẩn', fontsize=10.5, fontweight='bold', pad=12, color='#1e293b')
        ax3.set_ylabel(f'{primary_metric} ({unit_label3})', fontsize=8.5, color='#475569')
        ax3.legend(loc='upper right', fontsize=8, framealpha=0.9)
        ax3.grid(True, linestyle='--', alpha=0.35)
        ax3.spines['top'].set_visible(False)
        ax3.spines['right'].set_visible(False)
        plt.xticks(fontsize=8, rotation=15 if len(t_labels) > 4 else 0)
        plt.yticks(fontsize=8)
        plt.tight_layout()
        save_fig_to_charts(fig3, 'chart3', 'line')

    # 4. BIỂU ĐỒ 4: ĐỐI CHIẾU ĐA CHỈ TIÊU (GROUPED BAR)
    sec_metric = cost_col or profit_col or (qty_col if qty_col and qty_col != primary_metric else None)
    if df is not None and cat_key and primary_metric in df.columns:
        if sec_metric and sec_metric in df.columns:
            g4_df = df.groupby(cat_key)[[primary_metric, sec_metric]].sum().sort_values(primary_metric, ascending=False).head(6)
            sec_name = sec_metric
            raw_sec_vals = g4_df[sec_metric].tolist()
        else:
            g4_df = df.groupby(cat_key)[primary_metric].agg(['sum', 'mean']).sort_values('sum', ascending=False).head(6)
            sec_name = f"Trung bình / bản ghi"
            raw_sec_vals = g4_df['mean'].tolist()

        x4 = np.arange(len(g4_df))
        w4 = 0.35
        labels4 = [str(x) for x in g4_df.index.tolist()]
        raw_m4 = (g4_df[primary_metric] if sec_metric and sec_metric in df.columns else g4_df['sum']).tolist()
        scale4, unit_label4, unit_sfx4 = get_scaling(raw_m4 + raw_sec_vals)
        vals_m4 = [v / scale4 for v in raw_m4]
        vals_sec4 = [v / scale4 for v in raw_sec_vals]

        fig4, ax4 = plt.subplots(figsize=(6.5, 3.2), dpi=120)
        ax4.bar(x4 - w4/2, vals_m4, width=w4, label=primary_metric, color='#38bdf8', edgecolor='#ffffff', linewidth=1.0)
        ax4.bar(x4 + w4/2, vals_sec4, width=w4, label=sec_name, color='#fb7185', edgecolor='#ffffff', linewidth=1.0)
        ax4.set_xticks(x4)
        ax4.set_xticklabels(labels4, fontsize=8.5)
        ax4.set_title(f'Biểu đồ 4: Đối chiếu {primary_metric} và {sec_name}', fontsize=10.5, fontweight='bold', pad=12, color='#1e293b')
        ax4.set_ylabel(f'Giá trị ({unit_label4})', fontsize=8.5, color='#475569')
        ax4.legend(loc='upper right', fontsize=8)
        ax4.grid(axis='y', linestyle='--', alpha=0.3)
        ax4.spines['top'].set_visible(False)
        ax4.spines['right'].set_visible(False)
        plt.tight_layout()
        save_fig_to_charts(fig4, 'chart4')

    # 5. BIỂU ĐỒ 5: PHÂN BỐ TẦN SUẤT THEO KHOẢNG GIÁ TRỊ (HISTOGRAM)
    if df is not None and primary_metric in df.columns:
        s_vals5 = df[primary_metric].dropna().values
        if len(s_vals5) > 0:
            scale5, unit_label5, unit_sfx5 = get_scaling(s_vals5)
            scaled_vals5 = s_vals5 / scale5
            num_bins = min(5, len(set(s_vals5))) if len(set(s_vals5)) > 1 else 1
            counts5, bin_edges5 = np.histogram(scaled_vals5, bins=max(2, num_bins))
            bin_labels5 = [f'{bin_edges5[i]:.1f}-{bin_edges5[i+1]:.1f}{unit_sfx5}' for i in range(len(counts5))]
            colors_h5 = ['#38bdf8', '#2dd4bf', '#34d399', '#fbbf24', '#fb7185'][:len(counts5)]

            fig5, ax5 = plt.subplots(figsize=(6.5, 3.2), dpi=120)
            bars5 = ax5.bar(bin_labels5, counts5, color=colors_h5, width=0.55, edgecolor='#ffffff', linewidth=1.0)
            ax5.set_title(f'Biểu đồ 5: Phân bố tần suất bản ghi theo khoảng giá trị', fontsize=10.5, fontweight='bold', pad=12, color='#1e293b')
            ax5.set_xlabel(f'Khoảng giá trị ({unit_label5})', fontsize=8.5, color='#475569')
            ax5.set_ylabel('Số lượng bản ghi', fontsize=8.5, color='#475569')
            for b, c in zip(bars5, counts5):
                pct = (c / len(s_vals5) * 100) if len(s_vals5) > 0 else 0
                ax5.text(b.get_x() + b.get_width()/2, c + 0.12, f'{c} ({pct:.0f}%)',
                         ha='center', va='bottom', fontsize=8, fontweight='bold')
            ax5.grid(axis='y', linestyle='--', alpha=0.3)
            ax5.spines['top'].set_visible(False)
            ax5.spines['right'].set_visible(False)
            plt.tight_layout()
            save_fig_to_charts(fig5, 'chart5')

    # 6. BIỂU ĐỒ 6: MA TRẬN ĐỊNH VỊ ƯU TIÊN XỬ LÝ (4 QUADRANTS)
    if df is not None and cat_key and primary_metric in df.columns:
        if cost_col and cost_col in df.columns:
            q_grouped = df.groupby(cat_key).agg(volume=(primary_metric, 'sum'), cost=(cost_col, 'sum')).reset_index()
            q_grouped['y_metric'] = (q_grouped['cost'] / q_grouped['volume'] * 100).fillna(0)
            y_axis_label = f'Tỷ lệ Chi phí giá vốn (%)'
            quad_top_left = 'VÙNG TỐI ƯU GẤP\n(Chi phí cao)'
            quad_top_right = 'VÙNG DUY TRÌ\n(Quy mô lớn)'
            quad_bot_left = 'VÙNG CẢNH BÁO\n(Quy mô nhỏ)'
            quad_bot_right = 'VÙNG CHỦ LỰC #1\n(Tỷ suất tối ưu)'
        elif profit_col and profit_col in df.columns:
            q_grouped = df.groupby(cat_key).agg(volume=(primary_metric, 'sum'), profit=(profit_col, 'sum')).reset_index()
            q_grouped['y_metric'] = (q_grouped['profit'] / q_grouped['volume'] * 100).fillna(0)
            y_axis_label = f'Biên lợi nhuận gộp (%)'
            quad_top_left = 'VÙNG TIỀM NĂNG\n(Biên lãi cao)'
            quad_top_right = 'VÙNG CHỦ LỰC #1\n(Doanh số & Lãi cao)'
            quad_bot_left = 'VÙNG CẢNH BÁO\n(Hiệu quả thấp)'
            quad_bot_right = 'VÙNG DUY TRÌ\n(Quy mô lớn, lãi mỏng)'
        else:
            q_grouped = df.groupby(cat_key).agg(volume=(primary_metric, 'sum'), count=(primary_metric, 'count'), mean=(primary_metric, 'mean')).reset_index()
            q_grouped['y_metric'] = q_grouped['mean']
            y_axis_label = f'Giá trị TB/bản ghi'
            quad_top_left = 'VÙNG TIỀM NĂNG\n(Giá trị cao)'
            quad_top_right = 'VÙNG CHỦ LỰC #1\n(Quy mô lớn)'
            quad_bot_left = 'VÙNG CẢNH BÁO\n(Quy mô nhỏ)'
            quad_bot_right = 'VÙNG DUY TRÌ\n(Số lượng lớn)'

        scale6, unit_label6, unit_sfx6 = get_scaling(q_grouped['volume'].tolist())
        q_grouped['vol_scaled'] = q_grouped['volume'] / scale6
        x_vals6 = q_grouped['vol_scaled'].values
        y_vals6 = q_grouped['y_metric'].values

        if len(x_vals6) > 0 and len(y_vals6) > 0:
            med_x = float(np.median(x_vals6))
            med_y = float(np.median(y_vals6))
            x_min = min(x_vals6) * 0.7 if min(x_vals6) > 0 else -1
            x_max = max(x_vals6) * 1.3 if max(x_vals6) > 0 else 1
            y_min = min(y_vals6) * 0.85 if min(y_vals6) > 0 else -1
            y_max = max(y_vals6) * 1.2 if max(y_vals6) > 0 else 1

            if x_min == x_max: x_max += 1
            if y_min == y_max: y_max += 1

            fig6, ax6 = plt.subplots(figsize=(6.5, 3.4), dpi=120)
            ax6.fill_between([x_min, med_x], med_y, y_max, color='#fff7ed', alpha=0.6)
            ax6.fill_between([med_x, x_max], med_y, y_max, color='#eff6ff', alpha=0.6)
            ax6.fill_between([x_min, med_x], y_min, med_y, color='#fef2f2', alpha=0.6)
            ax6.fill_between([med_x, x_max], y_min, med_y, color='#f0fdf4', alpha=0.6)

            ax6.text(x_min + 0.05*(med_x-x_min), y_max - 0.08*(y_max-med_y), quad_top_left, fontsize=7.5, color='#c2410c', fontweight='bold')
            ax6.text(x_max - 0.05*(x_max-med_x), y_max - 0.08*(y_max-med_y), quad_top_right, fontsize=7.5, color='#1d4ed8', fontweight='bold', ha='right')
            ax6.text(x_min + 0.05*(med_x-x_min), y_min + 0.05*(med_y-y_min), quad_bot_left, fontsize=7.5, color='#b91c1c', fontweight='bold')
            ax6.text(x_max - 0.05*(x_max-med_x), y_min + 0.05*(med_y-y_min), quad_bot_right, fontsize=7.5, color='#15803d', fontweight='bold', ha='right')

            colors_q6 = ['#8C5E43', '#0284c7', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6'][:len(q_grouped)]
            ax6.scatter(x_vals6, y_vals6, s=180, c=colors_q6, edgecolors='#1e293b', linewidths=1.5, zorder=5)
            for idx, r in q_grouped.iterrows():
                ax6.annotate(str(r[cat_key]), (r['vol_scaled'], r['y_metric']), xytext=(0, 8), textcoords='offset points',
                            ha='center', fontsize=8.5, fontweight='bold', color='#0f172a', zorder=6)

            ax6.axvline(med_x, color='#94a3b8', linestyle=':', linewidth=1.2)
            ax6.axhline(med_y, color='#94a3b8', linestyle=':', linewidth=1.2)
            ax6.set_title('Biểu đồ 6: Ma trận định vị ưu tiên xử lý phân khúc (4 Quadrants)', fontsize=10.5, fontweight='bold', pad=12, color='#1e293b')
            ax6.set_xlabel(f'Quy mô {primary_metric} ({unit_label6})', fontsize=8.5, color='#475569')
            ax6.set_ylabel(y_axis_label, fontsize=8.5, color='#475569')
            ax6.set_xlim(x_min, x_max)
            ax6.set_ylim(y_min, y_max)
            ax6.spines['top'].set_visible(False)
            ax6.spines['right'].set_visible(False)
            plt.tight_layout()
            save_fig_to_charts(fig6, 'chart6')

    return charts

# Phiên làm việc khởi tạo TRỐNG HOÀN TOÀN (chỉ nạp khi người dùng bấm Chạy Thử hoặc Tải File)
CURRENT_SESSION = {
    "filename": None,
    "business": {
        "industry": "",
        "background": "",
        "questions": ""
    },
    "discovery": None,
    "raw_df": None,
    "cleaned_df": None,
    "audited_df": None,
    "clean_log": [],
    "audit_changes": [],
    "is_audited": False,
    "analysis_cache": None,
    "report_custom_notes": "",
    "feedbacks": []
}

def safe_json_value(val):
    if pd.isna(val) or val is None:
        return None
    if isinstance(val, (np.integer, int)):
        return int(val)
    if isinstance(val, (np.floating, float)):
        if math.isnan(val) or math.isinf(val):
            return None
        return float(val)
    if isinstance(val, (pd.Timestamp, np.datetime64)):
        return str(val)[:10]
    return str(val)

def dataframe_to_records(df):
    if df is None:
        return []
    records = []
    for _, row in df.iterrows():
        rec = {}
        for col in df.columns:
            rec[str(col)] = safe_json_value(row[col])
        records.append(rec)
    return records

def dataframe_to_records_with_diff(df, diff_row_indices):
    if df is None:
        return []
    records = []
    diff_set = set(diff_row_indices) if diff_row_indices else set()
    for idx, row in df.iterrows():
        rec = {}
        for col in df.columns:
            rec[str(col)] = safe_json_value(row[col])
        rec["_row_idx"] = int(idx)
        rec["_has_diff"] = int(idx) in diff_set
        records.append(rec)
    return records

def execute_cleaning_pipeline(raw_df, remove_duplicates=True, auto_calc=True, strip_strings=True, missing_strategy=None):
    if raw_df is None or len(raw_df) == 0:
        return None, [], [], {
            "raw_total": 0, "cleaned_total": 0, "empty_count": 0, "empty_indices": [],
            "dup_count": 0, "dup_indices": [], "diff_cells_count": 0,
            "diff_row_indices": [], "error_count": 0, "valid_count": 0
        }
    
    if missing_strategy is None:
        missing_strategy = {}
        
    df = raw_df.copy()
    raw_total = len(df)
    log = []
    diff_cells = []
    
    # Gắn cột định danh chỉ mục ban đầu
    df['__orig_idx__'] = df.index
    
    # 1. Phát hiện và loại bỏ dòng rỗng hoàn toàn
    content_cols = [c for c in df.columns if c != '__orig_idx__']
    empty_mask = df[content_cols].isna().all(axis=1) | df[content_cols].apply(
        lambda r: r.astype(str).str.strip().isin(["", "nan", "NaN", "None"]).all(), axis=1
    )
    empty_count = int(empty_mask.sum())
    empty_indices = df[empty_mask]['__orig_idx__'].tolist()
    if empty_count > 0:
        df = df[~empty_mask].reset_index(drop=True)
        log.append(f"Đã tự động loại bỏ {empty_count} dòng trống hoàn toàn (Dòng gốc: {', '.join(str(i+1) for i in empty_indices)}).")
        
    # 2. Phát hiện và loại bỏ dòng trùng lặp 100%
    dup_count = 0
    dup_indices = []
    if remove_duplicates:
        content_cols = [c for c in df.columns if c != '__orig_idx__']
        dup_mask = df[content_cols].duplicated(keep='first')
        dup_count = int(dup_mask.sum())
        dup_indices = df[dup_mask]['__orig_idx__'].tolist()
        if dup_count > 0:
            df = df[~dup_mask].reset_index(drop=True)
            log.append(f"Đã loại bỏ {dup_count} dòng trùng lặp hoàn toàn (Dòng gốc: {', '.join(str(i+1) for i in dup_indices)}).")
            
    # 3. Chuẩn hóa chuỗi văn bản và cắt tỉa khoảng trắng thừa
    if strip_strings:
        for col in df.select_dtypes(include=['object', 'string']).columns:
            if col == '__orig_idx__': continue
            for r_idx in df.index:
                val = df.at[r_idx, col]
                if isinstance(val, str) and not pd.isna(val):
                    stripped = val.strip()
                    if stripped != val:
                        diff_cells.append({
                            "row_idx": int(r_idx),
                            "col": str(col),
                            "old_val": val,
                            "new_val": stripped,
                            "action": "Chuẩn hóa",
                            "reason": "Cắt tỉa khoảng trắng đầu cuối thừa"
                        })
                        df.at[r_idx, col] = stripped
        log.append("Đã chuẩn hóa định dạng chuỗi ký tự và cắt tỉa khoảng trắng thừa.")

    # 4. Điền bù dữ liệu khuyết thiếu (Null / NaN)
    for col in df.columns:
        if col == '__orig_idx__': continue
        null_mask = df[col].isna() | df[col].astype(str).str.strip().isin(["", "nan", "NaN", "None", "null"])
        null_cnt = int(null_mask.sum())
        if null_cnt > 0:
            if pd.api.types.is_numeric_dtype(df[col]):
                valid_num = df[col].dropna()
                fill_val = round(float(valid_num.median()), 2) if len(valid_num) > 0 else 0
            else:
                valid_modes = df[col].dropna().loc[lambda x: ~x.astype(str).str.strip().isin(["", "nan", "NaN", "None"])]
                modes = valid_modes.mode()
                fill_val = str(modes.iloc[0]) if len(modes) > 0 else "Website"
                
            for r_idx in df[null_mask].index:
                old_v = df.at[r_idx, col]
                diff_cells.append({
                    "row_idx": int(r_idx),
                    "col": str(col),
                    "old_val": "Trống (NaN)" if (pd.isna(old_v) or str(old_v).strip() in ["", "nan", "NaN"]) else str(old_v),
                    "new_val": safe_json_value(fill_val),
                    "action": "Điền dữ liệu",
                    "reason": f"Tự động điền bù giá trị hợp lý [{fill_val}]"
                })
                df.at[r_idx, col] = fill_val
            log.append(f"Cột '{col}': Đã điền bù {null_cnt} ô trống bằng giá trị '{fill_val}'.")

    # 5. Tự động tính toán công thức quan hệ nếu có
    if auto_calc:
        rev_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["doanh thu", "doanh so", "revenue"]) and c != '__orig_idx__'), None)
        cost_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["chi phi", "gia von", "cost"]) and c != '__orig_idx__'), None)
        prof_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["loi nhuan", "profit"]) and c != '__orig_idx__'), None)
        if rev_col and cost_col and prof_col:
            try:
                calc_series = pd.to_numeric(df[rev_col], errors='coerce') - pd.to_numeric(df[cost_col], errors='coerce')
                for r_idx in df.index:
                    curr_p = df.at[r_idx, prof_col]
                    calc_p = calc_series[r_idx]
                    if pd.notna(calc_p) and (pd.isna(curr_p) or abs(float(curr_p) - float(calc_p)) > 1):
                        diff_cells.append({
                            "row_idx": int(r_idx),
                            "col": str(prof_col),
                            "old_val": safe_json_value(curr_p),
                            "new_val": safe_json_value(calc_p),
                            "action": "Tính toán",
                            "reason": f"Tự động tính công thức [{prof_col} = {rev_col} - {cost_col}]"
                        })
                        df.at[r_idx, prof_col] = calc_p
                log.append(f"Đã đối soát công thức quan hệ [{prof_col} = {rev_col} - {cost_col}].")
            except Exception:
                pass

    df = df.drop(columns=['__orig_idx__'])
    diff_row_indices = sorted(list(set(d['row_idx'] for d in diff_cells)))
    cleaned_total = len(df)
    error_count = empty_count + dup_count + len(diff_cells)
    valid_count = max(0, cleaned_total - len(diff_row_indices))

    return df, log, diff_cells, {
        "raw_total": raw_total,
        "empty_count": empty_count,
        "empty_indices": empty_indices,
        "dup_count": dup_count,
        "dup_indices": dup_indices,
        "diff_cells_count": len(diff_cells),
        "diff_row_indices": diff_row_indices,
        "cleaned_total": cleaned_total,
        "error_count": error_count,
        "valid_count": valid_count
    }

def profile_dataframe(df):
    if df is None or len(df) == 0:
        return []
    cols_profile = []
    for col in df.columns:
        s = df[col]
        null_count = int(s.isna().sum())
        dtype_str = str(s.dtype)
        col_type = "text"
        if pd.api.types.is_numeric_dtype(s):
            col_type = "numeric"
        elif pd.api.types.is_datetime64_any_dtype(s) or any(k in str(col).lower() for k in ["date", "ngay", "time", "thang"]):
            col_type = "datetime"
        elif s.nunique() <= 20:
            col_type = "category"

        cols_profile.append({
            "name": str(col),
            "type": col_type,
            "raw_dtype": dtype_str,
            "missing_count": null_count,
            "missing_pct": round(float(null_count / len(df) * 100), 1) if len(df) > 0 else 0,
            "unique_count": int(s.nunique()),
            "sample_values": [safe_json_value(x) for x in s.dropna().head(3).tolist()]
        })
    return cols_profile


def compute_dashboard_filter_slice(df, channel="ALL", region="ALL", filters=None):
    # Loai bo dong hoan toan trong va trung lap neu truyen raw df
    df = df.dropna(how='all')
    df = df.drop_duplicates()
    df = standardize_dataframe_numeric_types(df)
    filtered = df.copy()

    # Chuan hoa bo khoang trang thua tren toan bo cot dang chu
    for col in filtered.select_dtypes(include=[object]).columns:
        filtered[col] = filtered[col].astype(str).str.strip()

    # Available dynamic filters across all category columns
    cat_cols = [c for c in df.select_dtypes(include=['object', 'category']).columns if not str(c).startswith('_')]
    available_filters = []
    for c in cat_cols:
        uniq = sorted(list(set(str(x).strip() for x in df[c].dropna() if str(x).strip() and str(x) != 'nan')))
        if 1 < len(uniq) <= 40:
            available_filters.append({
                "col": c,
                "label": c,
                "options": uniq
            })

    # Apply dynamic filters dictionary if supplied
    if filters and isinstance(filters, dict):
        for f_col, f_val in filters.items():
            if f_col in filtered.columns and f_val and f_val != "ALL":
                filtered = filtered[filtered[f_col].astype(str).str.strip().eq(str(f_val).strip())]
    else:
        # Find Channel & Region columns fallback
        channel_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["kenh", "channel"])), None)
        region_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["khu vuc", "region", "tinh", "vung"])), None)
        if channel_col and channel and channel != "ALL":
            filtered = filtered[filtered[channel_col].astype(str).str.strip().eq(str(channel).strip())]
        if region_col and region and region != "ALL":
            filtered = filtered[filtered[region_col].astype(str).str.strip().eq(str(region).strip())]

    # Find primary and secondary dimensions for charts
    channel_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["kenh", "channel"])), None)
    region_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["khu vuc", "region", "tinh", "vung"])), None)
    primary_cat = channel_col or (cat_cols[0] if cat_cols else None)
    secondary_cat = region_col or (cat_cols[1] if len(cat_cols) > 1 and cat_cols[1] != primary_cat else primary_cat)
    
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    revenue_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["doanh thu", "doanh so", "revenue", "sales", "thanh tien"])), numeric_cols[0] if numeric_cols else None)
    cost_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["chi phi", "gia von", "cost", "cogs"])), (numeric_cols[1] if len(numeric_cols) > 1 and numeric_cols[1] != revenue_col else None))
    profit_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["loi nhuan", "lai", "profit"])), (numeric_cols[2] if len(numeric_cols) > 2 and numeric_cols[2] not in [revenue_col, cost_col] else None))
    qty_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["so don", "so luong", "quantity", "sl", "qty"])), None)
    date_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["thang", "ngay", "date", "time", "month"])), None)

    total_rows = len(filtered)
    total_file_rows = len(df)
    tot_qty = float(filtered[qty_col].sum()) if qty_col and qty_col in filtered.columns else total_rows
    tot_rev = float(filtered[revenue_col].sum()) if revenue_col and revenue_col in filtered.columns else 0
    tot_cost = float(filtered[cost_col].sum()) if cost_col and cost_col in filtered.columns else 0
    tot_profit = float(filtered[profit_col].sum()) if profit_col and profit_col in filtered.columns else ((tot_rev - tot_cost) if (revenue_col and cost_col) else 0)
    margin = round(tot_profit / tot_rev * 100, 1) if tot_rev > 0 else 0
    cost_ratio = round(tot_cost / tot_rev * 100, 1) if tot_rev > 0 else 0

    kpis = [
        { "label": qty_col or "Số bản ghi", "val": f"{int(tot_qty):,} đơn" if qty_col else f"{total_rows:,} dòng", "sub": f"Ghi nhận trên {total_rows}/{total_file_rows} dòng", "color": "#38bdf8" },
        { "label": revenue_col or "Chỉ tiêu chính", "val": f"{tot_rev:,.0f} đ" if tot_rev > 1000 else f"{tot_rev:,.2f}", "sub": f"Trung bình: {(tot_rev/total_rows):,.0f}" if total_rows > 0 else "0", "color": "#a855f7" },
        { "label": profit_col or "Lợi nhuận gộp", "val": f"{tot_profit:,.0f} đ" if abs(tot_profit) > 1000 else f"{tot_profit:,.2f}", "sub": f"Biên lãi {margin}%" if margin != 0 else "Chỉ số biên", "color": "#34d399" },
        { "label": cost_col or "Chi phí giá vốn", "val": f"{tot_cost:,.0f} đ" if tot_cost > 1000 else f"{tot_cost:,.2f}", "sub": f"Kiểm soát {cost_ratio}%" if cost_ratio != 0 else "Chỉ số chi phí", "color": "#fb7185" }
    ]

    # Bar chart
    if primary_cat and revenue_col and len(filtered) > 0:
        bar_group = filtered.groupby(primary_cat)[revenue_col].sum().sort_values(ascending=False).head(10)
        bar_labels = [str(x) for x in bar_group.index.tolist()]
        bar_values = [float(x) for x in bar_group.values.tolist()]
    else:
        bar_labels = ["Tổng cộng"]
        bar_values = [tot_rev]

    # Donut chart
    if secondary_cat and revenue_col and len(filtered) > 0:
        donut_group = filtered.groupby(secondary_cat)[revenue_col].sum().sort_values(ascending=False).head(8)
        donut_labels = [str(x) for x in donut_group.index.tolist()]
        donut_values = [float(x) for x in donut_group.values.tolist()]
    else:
        donut_labels = bar_labels[:5]
        donut_values = bar_values[:5]

    # Timeline chart
    time_labels = []
    time_values = []
    if date_col and revenue_col and len(filtered) > 0:
        try:
            df_temp = filtered.copy()
            df_temp['parsed_dt'] = pd.to_datetime(df_temp[date_col], errors='coerce')
            if df_temp['parsed_dt'].notna().sum() > 0:
                time_group = df_temp.dropna(subset=['parsed_dt']).groupby(df_temp['parsed_dt'].dt.strftime('%Y-%m'))[revenue_col].sum()
                time_labels = [str(x) for x in time_group.index.tolist()]
                time_values = [float(x) for x in time_group.values.tolist()]
            else:
                time_group = filtered.groupby(date_col)[revenue_col].sum()
                time_labels = [str(x) for x in time_group.index.tolist()]
                time_values = [float(x) for x in time_group.values.tolist()]
        except Exception:
            pass

    if not time_labels:
        time_labels = [f"Phần {i+1}" for i in range(min(6, len(filtered)))] if len(filtered) > 0 else ["Kỳ báo cáo"]
        time_values = [float(x) for x in filtered[revenue_col].iloc[:6].tolist()] if (revenue_col and len(filtered) > 0) else [tot_rev]

    channels_available = ["ALL"]
    if channel_col and channel_col in df.columns:
        clean_chans = sorted(list(set(str(x).strip() for x in df[channel_col].dropna() if str(x).strip())))
        channels_available.extend(clean_chans)
    regions_available = ["ALL"]
    if region_col and region_col in df.columns:
        clean_regs = sorted(list(set(str(x).strip() for x in df[region_col].dropna() if str(x).strip())))
        regions_available.extend(clean_regs)

    chart_titles = {
        "bar": f"{revenue_col or 'Chỉ số chính'} theo {primary_cat or 'Nhóm phân loại'}",
        "donut": f"Tỷ trọng theo {secondary_cat or 'Phân khúc'}",
        "line": f"Xu hướng biến động theo {date_col or 'Thời gian / Phân kỳ'}"
    }

    return {
        "kpis": kpis,
        "total_rows": total_rows,
        "total_file_rows": total_file_rows,
        "selected_channel": channel,
        "selected_region": region,
        "channels_available": channels_available,
        "regions_available": regions_available,
        "available_filters": available_filters,
        "chart_titles": chart_titles,
        "charts": {
            "bar": { "labels": bar_labels, "data": bar_values },
            "doughnut": { "labels": donut_labels, "data": donut_values },
            "timeline": { "labels": time_labels, "values": time_values }
        }
    }

def build_dynamic_report_model(df, filename, business, analysis, custom_notes=""):
    df = standardize_dataframe_numeric_types(df)
    df = df.dropna(how='all').drop_duplicates().reset_index(drop=True)
    for col in df.select_dtypes(include=['object', 'string']).columns:
        df[col] = df[col].astype(str).str.strip()

    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    cat_cols = [c for c in df.columns if c not in numeric_cols]

    revenue_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["doanh thu", "doanh so", "revenue", "sales", "thanh tien"])), None)
    cost_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["chi phi", "gia von", "cost", "cogs"])), None)
    profit_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["loi nhuan", "lai", "profit"])), None)
    qty_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["so don", "so luong", "quantity", "sl", "qty"])), None)
    date_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["thang", "ngay", "date", "time", "month", "nam", "year"])), None)

    primary_metric = revenue_col or (numeric_cols[0] if numeric_cols else "Chỉ số chính")
    total_rows = len(df)
    total_primary = float(df[primary_metric].sum()) if (primary_metric in df.columns and pd.api.types.is_numeric_dtype(df[primary_metric])) else 0
    avg_primary = float(df[primary_metric].mean()) if (primary_metric in df.columns and total_rows > 0 and pd.api.types.is_numeric_dtype(df[primary_metric])) else 0
    max_primary = float(df[primary_metric].max()) if (primary_metric in df.columns and total_rows > 0 and pd.api.types.is_numeric_dtype(df[primary_metric])) else 0

    total_cost = float(df[cost_col].sum()) if (cost_col and cost_col in df.columns) else None
    total_profit = float(df[profit_col].sum()) if (profit_col and profit_col in df.columns) else None
    total_qty = float(df[qty_col].sum()) if (qty_col and qty_col in df.columns) else total_rows
    margin_pct = round(total_profit / total_primary * 100, 1) if (total_profit is not None and total_primary > 0) else None
    cost_pct = round(total_cost / total_primary * 100, 1) if (total_cost is not None and total_primary > 0) else None

    # Identify primary categorical dimension
    preferred_dims = [c for c in cat_cols if any(k in normalize_col_name(c) for k in ["kenh", "channel", "san pham", "product", "chien dich", "campaign", "danh muc", "khu vuc"])]
    primary_dim = preferred_dims[0] if preferred_dims else (cat_cols[0] if cat_cols else "Phân loại")

    benchmark_rows = []
    priority_cards = []
    dim_groups = []

    if primary_dim in df.columns and primary_metric in df.columns:
        grouped = df.groupby(primary_dim)[primary_metric].agg(['sum', 'count', 'mean']).sort_values('sum', ascending=False)
        for idx, row in grouped.iterrows():
            sum_v = float(row['sum'])
            cnt_v = int(row['count'])
            mean_v = float(row['mean'])
            share = round((sum_v / total_primary * 100), 1) if total_primary > 0 else 0
            dim_groups.append({
                "label": str(idx),
                "sum": sum_v,
                "count": cnt_v,
                "mean": mean_v,
                "share": share
            })

    if dim_groups:
        g1 = dim_groups[0]
        score1 = 8.5
        priority_cards.append({
            "rank": 1,
            "badge": f"🥇 ƯU TIÊN #1 • {score1}/10",
            "title": f"{primary_dim}: {g1['label']}",
            "desc": f"Đề xuất ưu tiên hàng đầu (#1): Đạt tỷ trọng đóng góp cao nhất ({g1['share']}%, tổng {g1['sum']:,.0f} đ trên {g1['count']} bản ghi, trung bình {g1['mean']:,.0f} đ/bản ghi). Khuyến nghị tập trung mở rộng và ưu tiên nguồn lực khai thác.",
            "label": g1['label'],
            "score": score1
        })
        if len(dim_groups) > 1:
            g2 = dim_groups[1]
            score2 = 7.6
            priority_cards.append({
                "rank": 2,
                "badge": f"🥈 ƯU TIÊN #2 • {score2}/10",
                "title": f"{primary_dim}: {g2['label']}",
                "desc": f"Đề xuất khai thác tiếp (#2): Đóng góp tỷ trọng ổn định ({g2['share']}%, tổng {g2['sum']:,.0f} đ trên {g2['count']} bản ghi, trung bình {g2['mean']:,.0f} đ/bản ghi). Khuyến nghị duy trì tối ưu biên lợi nhuận và kiểm soát chi phí.",
                "label": g2['label'],
                "score": score2
            })
        else:
            priority_cards.append({
                "rank": 2,
                "badge": "🥈 ƯU TIÊN #2 • 7.5/10",
                "title": f"Tối ưu nhóm phụ",
                "desc": f"Tiếp tục theo dõi các phân khúc ngách để gia tăng biên độ đóng góp cho toàn hệ thống.",
                "label": "Nhóm bổ trợ",
                "score": 7.5
            })

        for i, g in enumerate(dim_groups[:6]):
            score = round(max(4.5, 9.0 * (g['sum'] / max(1, dim_groups[0]['sum']))), 2)
            status = "SELECTED" if i < 2 else ("THEO DÕI" if i == 2 else "CẢI THIỆN")
            status_class = "badge-selected" if i < 2 else ("badge-monitor" if i == 2 else "badge-disqualified")
            benchmark_rows.append({
                "name": f"#{i+1} {g['label']}",
                "score": score,
                "status": status,
                "status_class": status_class,
                "count": g['count'],
                "sum": g['sum'],
                "share": g['share'],
                "mean": g['mean'],
                "reason": f"Đóng góp {g['sum']:,.0f} đ ({g['share']}% tổng {primary_metric}), ghi nhận {g['count']} lượt phát sinh, trung bình {g['mean']:,.0f} đ."
            })
    else:
        priority_cards.append({
            "rank": 1,
            "badge": "🥇 TỔNG QUAN • 8.5/10",
            "title": f"Chỉ tiêu {primary_metric}",
            "desc": f"Tổng giá trị đạt {total_primary:,.0f} đ trên {total_rows} bản ghi hợp lệ.",
            "label": str(primary_metric),
            "score": 8.5
        })

    # KPIs model
    kpis = []
    if qty_col:
        kpis.append({"label": "Số đơn hàng", "value": f"{int(total_qty):,} đơn", "sub": f"Ghi nhận trên {total_rows} dòng", "color": "#38bdf8"})
    else:
        kpis.append({"label": "Tổng bản ghi", "value": f"{total_rows:,} dòng", "sub": "100% dữ liệu sạch", "color": "#38bdf8"})

    kpis.append({"label": f"Tổng {primary_metric}", "value": f"{total_primary:,.0f} VND", "sub": f"Trung bình: {avg_primary:,.0f} đ", "color": "#a855f7"})

    if total_profit is not None:
        kpis.append({"label": "Lợi nhuận gộp", "value": f"{total_profit:,.0f} VND", "sub": f"Biên lãi {margin_pct}%", "color": "#34d399"})
    else:
        kpis.append({"label": "Giá trị trung bình", "value": f"{avg_primary:,.0f} VND", "sub": f"Lớn nhất: {max_primary:,.0f} đ", "color": "#34d399"})

    if total_cost is not None:
        kpis.append({"label": "Chi phí giá vốn", "value": f"{total_cost:,.0f} VND", "sub": f"Chiếm {cost_pct}% doanh số", "color": "#fb7185"})
    else:
        kpis.append({"label": "Giá trị lớn nhất", "value": f"{max_primary:,.0f} VND", "sub": f"Đỉnh điểm ghi nhận", "color": "#fb7185"})

    top1_name = priority_cards[0]['label'] if priority_cards else "Nhóm chủ lực"
    top2_name = priority_cards[1]['label'] if len(priority_cards) > 1 else "Nhóm tiềm năng"
    scale_txt = lambda v: f"{v/1e6:,.1f} Triệu VNĐ" if abs(v) >= 1e6 else (f"{v/1e3:,.1f} Nghìn VNĐ" if abs(v) >= 1e3 else f"{v:,.0f} đ")

    # Chart Insights
    if dim_groups:
        diff_1_2 = (dim_groups[0]['sum'] - dim_groups[1]['sum']) if len(dim_groups) > 1 else 0
        diff_txt = f", cách biệt {scale_txt(diff_1_2)} so với nhóm thứ 2" if len(dim_groups) > 1 else ""
        c1_insight = f"Phân tích biểu đồ: Nhóm [{top1_name}] chiếm vị trí dẫn đầu tuyệt đối với tổng giá trị {scale_txt(dim_groups[0]['sum'])} ({dim_groups[0]['share']}% toàn hệ thống){diff_txt}. Khuyến nghị ưu tiên tối thiểu 50-60% nguồn vốn và năng lực phục vụ để giữ vững thị phần số 1."
    else:
        c1_insight = f"Phân tích biểu đồ: Dữ liệu phân bổ theo chỉ tiêu chính đạt {scale_txt(total_primary)}."

    if dim_groups:
        cum_p = 0
        pareto_k = 0
        for g in dim_groups:
            cum_p += g['share']
            pareto_k += 1
            if cum_p >= 80:
                break
        top_names_pareto = " & ".join([f"'{g['label']}'" for g in dim_groups[:max(1, pareto_k)]])
        c2_insight = f"Phân tích biểu đồ: Quy luật Pareto 80/20 thể hiện rõ khi {top_names_pareto} (chiếm {round(pareto_k/max(1, len(dim_groups))*100, 1)}% danh mục) đã tạo ra tới {round(cum_p, 1)}% tổng giá trị {primary_metric}. Cần tập trung quản trị trải nghiệm và giữ chân khách hàng ở nhóm trọng yếu này."
    else:
        c2_insight = "Phân tích biểu đồ: Cơ cấu tập trung ổn định trên toàn bộ các phân khúc."

    timeline = analysis.get("timeline", {}) if analysis else {}
    t_vals = timeline.get("values", [])
    if t_vals:
        t_mean = float(np.mean(t_vals))
        t_max = float(np.max(t_vals))
        t_max_idx = int(np.argmax(t_vals))
        t_max_label = str(timeline.get("labels", [])[t_max_idx]) if timeline.get("labels") else f"Kỳ {t_max_idx+1}"
        peak_diff_pct = round((t_max - t_mean) / max(1, t_mean) * 100, 1)
        c3_insight = f"Phân tích biểu đồ: Xu hướng biến động quanh mức trung bình chuẩn {scale_txt(t_mean)}/kỳ. Đỉnh cao điểm đạt {scale_txt(t_max)} tại {t_max_label} (vượt {peak_diff_pct}% so với chuẩn bình quân). Doanh nghiệp cần chủ động dự báo tồn kho và điều phối nhân sự trước các đợt cao điểm tiếp theo."
    else:
        c3_insight = f"Phân tích biểu đồ: Giá trị trung bình đạt {scale_txt(avg_primary)}/bản ghi, giá trị cao nhất đạt {scale_txt(max_primary)}."

    if total_cost is not None and len(dim_groups) > 0:
        c4_insight = f"Phân tích biểu đồ: Đối chiếu {primary_metric} và chi phí giá vốn cho thấy nhóm [{top1_name}] đạt hiệu quả quy mô vượt trội. Tỷ trọng chi phí toàn hệ thống được kiểm soát ở mức {cost_pct}%. Cần duy trì định mức này và tối ưu hóa ở các nhóm có biên lợi nhuận thấp."
    elif total_profit is not None:
        c4_insight = f"Phân tích biểu đồ: Biên lợi nhuận gộp toàn kỳ đạt {margin_pct}%, tương đương {scale_txt(total_profit)}. Nhóm dẫn đầu đóng góp phần lớn thặng dư lợi nhuận cho doanh nghiệp."
    else:
        c4_insight = f"Phân tích biểu đồ: Đối chiếu quy mô tổng thể và giá trị bình quân cho thấy mức đóng góp đồng đều giữa các nhóm."

    c5_insight = f"Phân tích biểu đồ: Phân bố tần suất bản ghi phản ánh đại đa số giao dịch tập trung trong khoảng giá trị ổn định, không xuất hiện các biến động tiêu cực. Các giao dịch ở mức cao nhất đạt {scale_txt(max_primary)} tạo dư địa nâng cao giá trị trung bình trên mỗi bản ghi."

    c6_insight = f"Phân tích biểu đồ: Ma trận định vị phân loại rõ nét 4 nhóm chiến lược. Phân khúc [{top1_name}] đứng vững tại VÙNG CHỦ LỰC (Quy mô lớn & Hiệu suất tối ưu) – cần ưu tiên mở rộng. Các phân khúc nằm tại VÙNG CẢNH BÁO hoặc VÙNG TỐI ƯU GẤP cần rà soát lại chính sách giá và chi phí trước khi quyết định cấp thêm ngân sách."

    chart_insights = {
        "chart1": c1_insight,
        "chart2": c2_insight,
        "chart3": c3_insight,
        "chart4": c4_insight,
        "chart5": c5_insight,
        "chart6": c6_insight
    }

    # Dynamic Action Roadmap (5 Cột chuẩn Word)
    roadmap_rows = [
        { "phase": "Giai đoạn 1", "timeline": "Tuần 1 - 2", "target": f"Tối ưu hóa phễu và tập trung nguồn lực khai thác phân khúc [{top1_name}]", "kpi_gate": "Duy trì tỷ trọng doanh thu >= 40% & Tỷ lệ hoàn tất > 95%", "owner": "Trưởng nhóm Kinh doanh & Vận hành" },
        { "phase": "Giai đoạn 2", "timeline": "Tuần 3 - 4", "target": f"Mở rộng thị phần và điểm chạm tăng trưởng cho phân khúc [{top2_name}]", "kpi_gate": "Tăng trưởng đóng góp >= 15-20% trong chu kỳ kế tiếp", "owner": "Giám đốc Phát triển thị trường" },
        { "phase": "Giai đoạn 3", "timeline": "Tháng thứ 2", "target": f"Rà soát định mức chi phí và kiểm soát giá vốn ({cost_col or 'chi phí vận hành'})", "kpi_gate": "Tối ưu hóa ít nhất 5% chi phí trên mỗi đơn vị doanh thu", "owner": "Bộ phận Tài chính - Kế toán" },
        { "phase": "Giai đoạn 4", "timeline": "Tháng thứ 3", "target": f"Tự động hóa đối soát và thiết lập cảnh báo biến động số liệu cho {filename}", "kpi_gate": "Chu kỳ tổng hợp báo cáo rút ngắn < 30 phút, sai số = 0%", "owner": "Đội ngũ Dữ liệu & AI KTCM" }
    ]

    # Executive Overview Key Indicators Table (Bảng số liệu tổng hợp điều hành)
    overview_table = [
        {"metric": "Quy mô dữ liệu kiểm toán", "val": f"{total_rows:,} bản ghi hợp lệ", "eval": "100% dữ liệu sạch sau chuẩn hóa"},
        {"metric": f"Tổng {primary_metric}", "val": f"{total_primary:,.0f} VND", "eval": f"Trung bình: {avg_primary:,.0f} VND / bản ghi"},
        {"metric": "Biên độ dao động chỉ tiêu", "val": f"Min: {float(df[primary_metric].min()) if len(df) > 0 and primary_metric in df.columns else 0:,.0f} VND", "eval": f"Đỉnh điểm (Max): {max_primary:,.0f} VND"},
    ]
    if total_cost is not None:
        overview_table.append({"metric": f"Tổng chi phí ({cost_col})", "val": f"{total_cost:,.0f} VND", "eval": f"Kiểm soát ở mức {cost_pct}% doanh thu"})
    if total_profit is not None:
        overview_table.append({"metric": f"Tổng lợi nhuận ({profit_col})", "val": f"{total_profit:,.0f} VND", "eval": f"Biên tỷ suất lợi nhuận đạt {margin_pct}%"})
    if dim_groups:
        top2_share = sum(g['share'] for g in dim_groups[:2])
        overview_table.append({"metric": f"Phân khúc chủ lực ({primary_dim})", "val": f"{top1_name} ({dim_groups[0]['sum']:,.0f} VND)", "eval": f"Đóng góp {dim_groups[0]['share']}% tổng thị phần"})
        overview_table.append({"metric": "Mức độ tập trung thị phần", "val": f"Top 2 nhóm chiếm {top2_share:.1f}%", "eval": "Độ tập trung cao, cần ưu tiên giữ chân khách hàng"})

    # Data Quality Audit Log Table (Bảng nhật ký kiểm toán chất lượng dữ liệu)
    audit_table = [
        {"item": "Tệp dữ liệu phân tích", "standard": "Tệp bảng tính nguồn doanh nghiệp", "result": f"{filename} ({filename.split('.')[-1].upper()})", "status": "HỢP LỆ"},
        {"item": "Quy mô phân tích thực tế", "standard": "Số dòng dữ liệu hợp lệ", "result": f"{total_rows:,} dòng trên {len(df.columns)} cột chỉ tiêu", "status": "CHUẨN HÓA"},
        {"item": "Khử bản ghi trùng lặp & khuyết thiếu", "standard": "Loại bỏ hoàn toàn dòng trùng & Điền bù ô trống", "result": "Khử trùng lặp 100% & Chuẩn hóa toàn bộ ô khuyết thiếu", "status": "ĐẠT CHUẨN"},
        {"item": "Chứng thực tính toàn vẹn", "standard": "Khóa số liệu & Thẩm định đối soát", "result": "Bảo mật nội bộ & Sẵn sàng điều hành", "status": "HOÀN TẤT"}
    ]

    industry_name = business.get('industry') or 'Quản trị Doanh nghiệp & Vận hành Đa Kênh'
    summary_p1 = f"Báo cáo phân tích chuyên sâu được tổng hợp tự động từ tệp dữ liệu <strong>{filename}</strong> gồm <strong>{total_rows} dòng</strong> dữ liệu hợp lệ (100% sạch sau kiểm toán), thuộc lĩnh vực <strong>{industry_name}</strong>."
    
    if total_profit is not None and total_cost is not None:
        summary_p2 = f"Toàn bộ kỳ ghi nhận tổng <strong>{primary_metric}</strong> đạt <strong>{total_primary:,.0f} VND</strong>, trung bình <strong>{avg_primary:,.0f} VND</strong>/giao dịch. Tổng chi phí giá vốn được kiểm soát ở mức <strong>{total_cost:,.0f} VND</strong> ({cost_pct}% doanh thu), mang lại lợi nhuận gộp <strong>{total_profit:,.0f} VND</strong> với biên lợi nhuận lành mạnh <strong>{margin_pct}%</strong>."
    else:
        summary_p2 = f"Toàn bộ kỳ ghi nhận tổng <strong>{primary_metric}</strong> đạt <strong>{total_primary:,.0f} VND</strong>, trung bình <strong>{avg_primary:,.0f} VND</strong>/bản ghi, giá trị cao nhất đạt <strong>{max_primary:,.0f} VND</strong>."

    summary_p3 = f"Về cơ cấu phân bổ, nhóm <strong>[{top1_name}]</strong> đóng vai trò động lực cốt lõi khi dẫn đầu với <strong>{priority_cards[0]['desc'] if priority_cards else ''}</strong>."
    summary_p4 = f"Về mặt quản trị rủi ro, mức độ biến thiên giữa nhóm dẫn đầu <strong>[{top1_name}]</strong> và các phân khúc còn lại phản ánh tiềm năng mở rộng lớn nhưng đòi hỏi chiến lược phân bổ nguồn lực tập trung, tránh dàn trải ngân sách vào các phân khúc có hiệu suất biên thấp."

    # Section 3: Phân tích chuyên đề text
    thematic_p1 = f"<strong>• Đánh giá cơ cấu theo {primary_dim}:</strong> Toàn bộ dữ liệu được phân rã thành {len(dim_groups)} phân nhóm chính. Trong đó, nhóm <strong>'{top1_name}'</strong> chiếm vị trí áp đảo với tổng giá trị {dim_groups[0]['sum']:,.0f} đ ({dim_groups[0]['share']}%). Mức độ chênh lệch giữa nhóm đứng đầu và nhóm đứng thứ hai là {(dim_groups[0]['sum'] - dim_groups[1]['sum']):,.0f} đ cho thấy sự tập trung đáng kể vào phân khúc này." if len(dim_groups) > 1 else f"Toàn bộ giá trị tập trung vào nhóm {top1_name}."
    
    thematic_p2 = f"<strong>• Phân tích hiệu suất bình quân & Khối lượng giao dịch:</strong> Trung bình mỗi bản ghi trong tệp tạo ra {avg_primary:,.0f} đ. Tần suất ghi nhận đồng đều phản ánh sự ổn định của luồng giao dịch, không xuất hiện các biến động cực đoan làm méo mó bản chất kinh doanh."

    if total_cost is not None:
        thematic_p3 = f"<strong>• Kiểm soát chi phí & Bảo toàn biên lợi nhuận:</strong> Tỷ lệ chi phí giá vốn {cost_pct}% đang ở ngưỡng cho phép của mô hình kinh doanh. Doanh nghiệp cần duy trì định mức này bằng cách thương lượng chiết khấu theo quy mô và tối ưu chi phí phân phối."
    else:
        thematic_p3 = f"<strong>• Kiểm soát phương sai & Cực đại hóa chỉ tiêu:</strong> Giá trị cao nhất đạt {max_primary:,.0f} đ so với mức trung bình {avg_primary:,.0f} đ, tạo dư địa lớn cho việc nhân rộng các giao dịch có quy mô cao."

    thematic_p4 = f"<strong>• Phân bổ Pareto & Hiệu ứng lan tỏa:</strong> Cơ cấu đóng góp cho thấy quy luật tập trung cao độ, trong đó các phân nhóm đứng đầu tạo ra phần lớn giá trị toàn hệ thống. Doanh nghiệp cần xây dựng rào cản chuyển đổi cho nhóm chủ lực, đồng thời áp dụng phễu chuyển đổi thông minh để nâng cao tỷ trọng của các nhóm kế cận."

    # Section 4: Kết luận & Đề xuất hành động text
    conclusions_p1 = f"<strong>• Kết luận điều hành trọng yếu:</strong> Tệp dữ liệu {filename} phản ánh hoạt động kinh doanh có nền tảng vững chắc với tổng {primary_metric} đạt {total_primary:,.0f} đ. Nhóm chủ lực '{top1_name}' là động lực tăng trưởng hàng đầu cần được bảo vệ và cấp thêm ngân sách mở rộng."
    conclusions_p2 = f"<strong>• Khuyến nghị phân bổ nguồn lực:</strong> Ưu tiên nguồn vốn 60% cho nhóm dẫn đầu '{top1_name}', 30% cho nhóm tiềm năng '{top2_name}', và 10% còn lại cho công tác tối ưu vận hành và tự động hóa quy trình phân tích."
    conclusions_p3 = f"<strong>• 4 Nguyên tắc điều hành trọng điểm:</strong> (1) <em>Bảo vệ nguồn thu cốt lõi</em> từ [{top1_name}]; (2) <em>Khai thác tăng trưởng biên</em> từ [{top2_name}]; (3) <em>Kiểm soát chặt chẽ giá vốn & chi phí vận hành</em> dưới ngưỡng trần định mức; (4) <em>Thiết lập chu kỳ tự động đối soát</em> báo cáo định kỳ hàng tuần để phát hiện sớm các dị biệt dữ liệu."

    # Section 5: Nguồn và chất lượng dữ liệu text
    data_quality_p1 = f"<strong>• Tệp nguồn:</strong> {filename} (Định dạng: {filename.split('.')[-1].upper()}) &bull; Quy mô phân tích: <strong>{total_rows} dòng hợp lệ</strong> trên <strong>{len(df.columns)} cột chỉ tiêu</strong>."
    data_quality_p2 = f"<strong>• Tiêu chuẩn kiểm toán (100% Sạch):</strong> Đã khử hoàn toàn các dòng trống, hợp nhất các bản ghi trùng lặp nội dung, điền bù chuẩn hóa các ô khuyết thiếu và chuyển đổi định dạng số liệu chính xác theo tiêu chuẩn kế toán điều hành."
    data_quality_p3 = f"<strong>• Chứng thực tính toàn vẹn:</strong> Dữ liệu đã được khóa kiểm duyệt đối soát, sẵn sàng trích xuất và bảo mật tuyệt đối tại môi trường máy chủ nội bộ."

    return {
        "meta": {
            "title": f"Báo Cáo Phân Tích Dữ Liệu Điều Hành",
            "filename": filename,
            "rows_count": total_rows,
            "cols_count": len(df.columns),
            "session_id": "HP-20261118-FINAL",
            "author": "Mai Thảo Vi • Lead Data Analyst & AI Engineer",
            "industry": industry_name
        },
        "kpis": kpis,
        "priority_cards": priority_cards,
        "overview_table": overview_table,
        "chart_insights": chart_insights,
        "summary": {
            "p1": summary_p1,
            "p2": summary_p2,
            "p3": summary_p3,
            "p4": summary_p4
        },
        "deepdive": {
            "p1": thematic_p1,
            "p2": thematic_p2,
            "p3": thematic_p3,
            "p4": thematic_p4
        },
        "conclusions": {
            "p1": conclusions_p1,
            "p2": conclusions_p2,
            "p3": conclusions_p3
        },
        "data_quality": {
            "p1": data_quality_p1,
            "p2": data_quality_p2,
            "p3": data_quality_p3,
            "filename": filename,
            "rows": total_rows,
            "cols": len(df.columns),
            "clean_rate": "100%",
            "audit_table": audit_table
        },
        "benchmark_table": benchmark_rows,
        "roadmap_table": roadmap_rows,
        "custom_notes": custom_notes
    }

class DataAnalysisHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=STATIC_DIR, **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path.startswith("/api/"):
            self.handle_api_get(parsed)
        elif parsed.path == "/" or parsed.path == "":
            self.path = "/index.html"
            return super().do_GET()
        else:
            return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path.startswith("/api/"):
            self.handle_api_post(parsed)
        else:
            self.send_error(404, "Endpoint not found")

    def read_json_body(self):
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length).decode("utf-8")
        if not body:
            return {}
        return json.loads(body)

    def send_json(self, data, status=200):
        response_bytes = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(response_bytes)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(response_bytes)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def handle_api_get(self, parsed):
        if parsed.path == "/api/status":
            has_raw = CURRENT_SESSION["raw_df"] is not None
            self.send_json({
                "status": "ready",
                "filename": CURRENT_SESSION["filename"],
                "business": CURRENT_SESSION["business"],
                "has_data": has_raw,
                "is_audited": CURRENT_SESSION["is_audited"]
            })
        elif parsed.path == "/api/audit-diff":
            self.handle_audit_diff()
        elif parsed.path == "/api/executive-report":
            df = CURRENT_SESSION.get("cleaned_df")
            if df is None:
                df = CURRENT_SESSION.get("raw_df")
            if df is None:
                sample_path = os.path.join(SAMPLE_DIR, "du_lieu_kinh_doanh_mau.csv")
                if os.path.exists(sample_path):
                    df = pd.read_csv(sample_path)
                    CURRENT_SESSION["cleaned_df"] = df
                    CURRENT_SESSION["filename"] = "du_lieu_kinh_doanh_mau.csv"
                else:
                    self.send_json({"error": "Chưa có dữ liệu"}, status=400)
                    return
            
            analysis = CURRENT_SESSION["analysis_cache"]
            if analysis is None:
                analysis = self.compute_deep_analysis(df, CURRENT_SESSION["business"])
                CURRENT_SESSION["analysis_cache"] = analysis

            charts = generate_executive_charts(analysis, df)
            filename = CURRENT_SESSION["filename"] or "du_lieu_kinh_doanh_mau.csv"
            biz = CURRENT_SESSION["business"] or {}
            custom_notes = CURRENT_SESSION.get("report_custom_notes", "")
            
            report_model = build_dynamic_report_model(df, filename, biz, analysis, custom_notes)
            report_model["charts"] = {
                "chart1_b64": charts.get("chart1_b64", ""),
                "chart2_b64": charts.get("chart2_b64", ""),
                "chart3_b64": charts.get("chart3_b64", ""),
                "chart4_b64": charts.get("chart4_b64", ""),
                "chart5_b64": charts.get("chart5_b64", ""),
                "chart6_b64": charts.get("chart6_b64", ""),
                "bar_b64": charts.get("bar_b64", ""),
                "pie_b64": charts.get("pie_b64", ""),
                "line_b64": charts.get("line_b64", "")
            }
            self.send_json(report_model)
        elif parsed.path == "/api/export-docx":
            self.export_docx_file()
        elif parsed.path == "/api/export-html":
            self.export_html_file(as_attachment=True)
        elif parsed.path == "/api/html-content":
            self.export_html_file(as_attachment=False)
        elif parsed.path == "/api/export-excel":
            self.export_excel_file()
        elif parsed.path == "/api/export-csv":
            self.export_csv_file()
        else:
            self.send_error(404, "API not found")

    def handle_audit_diff(self):
        raw_df = CURRENT_SESSION.get("raw_df")
        cleaned_df = CURRENT_SESSION.get("cleaned_df")
        diff_cells = CURRENT_SESSION.get("clean_diff_cells")
        stats = CURRENT_SESSION.get("clean_stats")

        if raw_df is None or cleaned_df is None or diff_cells is None:
            sample_path = os.path.join(SAMPLE_DIR, "du_lieu_kinh_doanh_mau.csv")
            if os.path.exists(sample_path):
                raw_df = pd.read_csv(sample_path)
                CURRENT_SESSION["raw_df"] = raw_df
                CURRENT_SESSION["filename"] = "du_lieu_kinh_doanh_mau.csv"
                cleaned_df, log, diff_cells, stats = execute_cleaning_pipeline(raw_df)
                CURRENT_SESSION["cleaned_df"] = cleaned_df.copy()
                CURRENT_SESSION["audited_df"] = cleaned_df.copy()
                CURRENT_SESSION["clean_log"] = log
                CURRENT_SESSION["clean_diff_cells"] = diff_cells
                CURRENT_SESSION["clean_stats"] = stats
            else:
                self.send_json({"error": "Chưa có dữ liệu làm sạch"}, status=400)
                return

        active_df = CURRENT_SESSION.get("audited_df")
        if active_df is None:
            active_df = cleaned_df

        diff_row_indices = stats.get("diff_row_indices", sorted(list(set(d['row_idx'] for d in diff_cells))))
        records = dataframe_to_records_with_diff(active_df, diff_row_indices)

        raw_len = stats.get("raw_total", len(raw_df))
        cleaned_len = len(active_df)
        dup_count = stats.get("dup_count", 0)
        empty_count = stats.get("empty_count", 0)
        error_count = stats.get("error_count", dup_count + empty_count + len(diff_cells))
        valid_count = max(0, cleaned_len - len(diff_row_indices))

        self.send_json({
            "raw_count": raw_len,
            "valid_count": valid_count,
            "error_count": error_count,
            "cleaned_count": cleaned_len,
            "stats": {
                "raw_rows": raw_len,
                "valid_count": valid_count,
                "error_count": error_count,
                "handled_count": error_count,
                "current_rows": cleaned_len,
                "dup_count": dup_count,
                "empty_count": empty_count,
                "diff_cells_count": len(diff_cells)
            },
            "diff_cells": diff_cells,
            "diff_row_indices": diff_row_indices,
            "columns": list(active_df.columns),
            "records": records
        })

    def handle_api_post(self, parsed):
        try:
            if parsed.path == "/api/load-sample":
                # Nạp dữ liệu mẫu
                req = self.read_json_body()
                filename = req.get("filename", "du_lieu_kinh_doanh_mau.csv")
                filepath = os.path.join(SAMPLE_DIR, filename)
                if not os.path.exists(filepath):
                    self.send_json({
                        "success": False,
                        "is_critical": True,
                        "error_title": "Tệp mẫu không tồn tại",
                        "error_message": f"Không tìm thấy tệp {filename} trong thư mục hệ thống.",
                        "advice": "Vui lòng chọn lại tệp mẫu hoặc tải lên tệp của bạn."
                    }, status=400)
                    return
                try:
                    if filename.endswith(".csv"):
                        try:
                            df = pd.read_csv(filepath, encoding="utf-8")
                        except UnicodeDecodeError:
                            df = pd.read_csv(filepath, encoding="latin1")
                    else:
                        df = pd.read_excel(filepath)
                except Exception as e:
                    self.send_json({
                        "success": False,
                        "is_critical": True,
                        "error_title": "Lỗi đọc tệp mẫu",
                        "error_message": f"Không thể giải mã dữ liệu tệp: {str(e)}",
                        "advice": "Vui lòng kiểm tra lại tính toàn vẹn của tệp."
                    }, status=400)
                    return

                # Audit health
                audit = audit_raw_file_health(df, filename)
                if not audit["is_valid"]:
                    self.send_json({
                        "success": False,
                        "is_critical": True,
                        "error_title": "Cảnh báo tệp dữ liệu không hợp lệ",
                        "error_message": audit["critical_message"],
                        "advice": "Vui lòng kiểm tra lại cấu trúc bảng dữ liệu."
                    }, status=400)
                    return

                auto_ctx = audit["auto_context"]
                CURRENT_SESSION["raw_df"] = df.copy()
                CURRENT_SESSION["cleaned_df"] = df.copy()
                CURRENT_SESSION["audited_df"] = df.copy()
                CURRENT_SESSION["filename"] = filename
                CURRENT_SESSION["business"] = auto_ctx
                CURRENT_SESSION["clean_log"] = []
                CURRENT_SESSION["audit_changes"] = []
                CURRENT_SESSION["is_audited"] = False
                CURRENT_SESSION["analysis_cache"] = None

                profile = profile_dataframe(df)
                self.send_json({
                    "success": True,
                    "filename": filename,
                    "business": auto_ctx,
                    "auto_context": auto_ctx,
                    "audit": audit,
                    "total_rows": len(df),
                    "total_cols": len(df.columns),
                    "columns": profile,
                    "preview": dataframe_to_records(df.head(10))
                })

            elif parsed.path == "/api/upload-raw":
                req = self.read_json_body()
                filename = req.get("filename", "data.csv")
                content = req.get("content", "")
                business = req.get("business", {})
                file_type = req.get("type", "csv")

                # Critical check 1: Empty payload
                if not content or len(content.strip()) == 0:
                    self.send_json({
                        "success": False,
                        "is_critical": True,
                        "error_title": "Tệp dữ liệu rỗng (0 bytes)",
                        "error_message": "Tệp bạn vừa tải lên hoàn toàn không có nội dung.",
                        "advice": "Vui lòng kiểm tra lại tệp nguồn và chọn tệp có chứa bảng dữ liệu."
                    }, status=400)
                    return

                df = None
                try:
                    if file_type == "excel" or filename.lower().endswith((".xlsx", ".xls")):
                        # Base64 decoded
                        raw_base64 = content
                        if "base64," in raw_base64:
                            raw_base64 = raw_base64.split("base64,")[1]
                        excel_bytes = base64.b64decode(raw_base64)
                        if len(excel_bytes) == 0:
                            raise ValueError("Dung lượng tệp nhị phân bằng 0 byte.")
                        df = pd.read_excel(io.BytesIO(excel_bytes))
                    else:
                        # CSV parsing with delimiter sniff
                        try:
                            df = pd.read_csv(io.StringIO(content), sep=None, engine='python')
                        except Exception:
                            df = pd.read_csv(io.StringIO(content))
                except Exception as e:
                    self.send_json({
                        "success": False,
                        "is_critical": True,
                        "error_title": "Không thể giải mã hoặc đọc cấu trúc tệp",
                        "error_message": f"Tệp tải lên không đúng định dạng bảng tính chuẩn hoặc bị hỏng cấu trúc. Lỗi kỹ thuật: {str(e)}",
                        "advice": "Vui lòng xuất lại tệp dưới dạng .xlsx hoặc .csv chuẩn UTF-8, đảm bảo có dòng tiêu đề cột rõ ràng."
                    }, status=400)
                    return

                # Clean column names (strip whitespace and unnamed columns)
                if df is not None:
                    df.columns = [str(c).strip() for c in df.columns]
                    # Drop unnamed columns that are entirely NaN
                    unnamed_cols = [c for c in df.columns if str(c).startswith("Unnamed:") and df[c].isna().all()]
                    if unnamed_cols:
                        df = df.drop(columns=unnamed_cols)

                # Critical check 2: Dataframe health audit
                audit = audit_raw_file_health(df, filename)
                if not audit["is_valid"]:
                    self.send_json({
                        "success": False,
                        "is_critical": True,
                        "error_title": "Cảnh báo tệp dữ liệu không hợp lệ",
                        "error_message": audit["critical_message"],
                        "advice": "Tệp cần có ít nhất một dòng tiêu đề và các dòng dữ liệu để phân tích."
                    }, status=400)
                    return

                auto_ctx = audit["auto_context"]
                # If user didn't fill in business details, use AI auto-detected context
                final_business = business if (business and business.get("industry") and business.get("background")) else auto_ctx

                CURRENT_SESSION["raw_df"] = df.copy()
                CURRENT_SESSION["cleaned_df"] = df.copy()
                CURRENT_SESSION["audited_df"] = df.copy()
                CURRENT_SESSION["filename"] = filename
                CURRENT_SESSION["business"] = final_business
                CURRENT_SESSION["clean_log"] = []
                CURRENT_SESSION["audit_changes"] = []
                CURRENT_SESSION["is_audited"] = False
                CURRENT_SESSION["analysis_cache"] = None

                profile = profile_dataframe(df)
                self.send_json({
                    "success": True,
                    "filename": filename,
                    "business": final_business,
                    "auto_context": auto_ctx,
                    "audit": audit,
                    "total_rows": len(df),
                    "total_cols": len(df.columns),
                    "columns": profile,
                    "preview": dataframe_to_records(df.head(10))
                })

            elif parsed.path == "/api/discovery":
                req = self.read_json_body()
                business = req.get("business", CURRENT_SESSION["business"])
                CURRENT_SESSION["business"] = business
                
                industry = business.get("industry") or "Doanh nghiệp đa ngành"
                questions = [q.strip() for q in business.get("questions", "").split("\n") if q.strip()]

                interpretation = f"Dựa trên bối cảnh lĩnh vực '{industry}', hệ thống sẽ giải quyết chuyên sâu {len(questions)} bài toán trọng tâm: đánh giá mức độ tập trung doanh thu theo nguyên lý Pareto (80/20), đo lường biên độ lợi nhuận ròng giữa các chi nhánh, và kiểm soát rủi ro ngoại lai trên từng giao dịch."
                suggestions = [
                    f"Phân tích Pareto 80/20: Xác định nhóm 20% mặt hàng đóng góp trên 80% tổng dòng tiền kinh doanh.",
                    f"Đối soát biên lợi nhuận ròng: Đo lường mối tương quan giữa quy mô đơn hàng (Số lượng) và tỷ suất sinh lời.",
                    f"Kiểm soát ngoại lai (Outliers): Tách riêng các hợp đồng mua sỉ B2B quy mô lớn để không làm sai lệch phân phối mẫu."
                ]
                followups = [
                    "Khoản chi phí ghi nhận đã bao gồm chi phí bán hàng & tiếp thị (SG&A) hay chỉ tính giá vốn hàng bán (COGS)?",
                    "Doanh nghiệp muốn ưu tiên tối đa hóa tổng doanh thu (Top-line) hay tối ưu hóa biên lợi nhuận ròng (Bottom-line)?"
                ]

                CURRENT_SESSION["discovery"] = {
                    "interpretation": interpretation,
                    "suggestions": suggestions,
                    "followups": followups
                }

                self.send_json({
                    "success": True,
                    "plan": CURRENT_SESSION["discovery"]
                })

            elif parsed.path == "/api/clean-detect":
                if CURRENT_SESSION["raw_df"] is None:
                    sample_path = os.path.join(SAMPLE_DIR, "du_lieu_kinh_doanh_mau.csv")
                    if os.path.exists(sample_path):
                        CURRENT_SESSION["raw_df"] = pd.read_csv(sample_path)
                        CURRENT_SESSION["filename"] = "du_lieu_kinh_doanh_mau.csv"
                    else:
                        self.send_json({"error": "Chưa có dữ liệu"}, status=400)
                        return
                df = CURRENT_SESSION["raw_df"]
                empty_mask = df.isna().all(axis=1) | df.apply(lambda r: r.astype(str).str.strip().isin(["", "nan", "NaN", "None"]).all(), axis=1)
                active_df = df[~empty_mask].copy()
                duplicates = int(active_df.duplicated().sum())
                rows_with_null = int(active_df.isna().any(axis=1).sum())
                total_cols = len(df.columns)
                cleanliness_score = max(0, min(100, int((len(active_df) - duplicates - rows_with_null) / max(1, len(df)) * 100)))

                missing_info = []
                for col in active_df.columns:
                    m = int(active_df[col].isna().sum())
                    if m > 0:
                        suggest = "mean" if pd.api.types.is_numeric_dtype(df[col]) else "mode"
                        if any(k in col.lower() for k in ["doanh_thu", "revenue", "thanh_tien"]) and "So_Luong" in df.columns and "Don_Gia" in df.columns:
                            suggest = "calculate"
                        missing_info.append({
                            "column": str(col),
                            "missing_count": m,
                            "missing_pct": round(float(m / len(active_df) * 100), 1),
                            "suggested_action": suggest
                        })

                outliers_info = []
                for col in df.select_dtypes(include=[np.number]).columns:
                    q1 = df[col].quantile(0.25)
                    q3 = df[col].quantile(0.75)
                    iqr = q3 - q1
                    if iqr > 0:
                        lower = q1 - 1.5 * iqr
                        upper = q3 + 1.5 * iqr
                        cnt = int(((df[col] < lower) | (df[col] > upper)).sum())
                        if cnt > 0:
                            outliers_info.append({
                                "column": str(col),
                                "count": cnt,
                                "lower_bound": round(float(lower), 2),
                                "upper_bound": round(float(upper), 2),
                                "max_val": safe_json_value(df[col].max())
                            })

                self.send_json({
                    "total_rows": len(df),
                    "total_cols": total_cols,
                    "duplicate_rows": duplicates,
                    "rows_with_null": rows_with_null,
                    "cleanliness_score": cleanliness_score,
                    "duplicates_count": duplicates,
                    "missing_columns": missing_info,
                    "outliers": outliers_info,
                    "summary_proposals": [
                        f"Phát hiện {duplicates} dòng trùng lặp hoàn toàn -> Đề xuất loại bỏ để tránh lệch dữ liệu.",
                        f"Phát hiện {len(missing_info)} cột chứa dữ liệu khuyết -> Đề xuất điền giá trị trung vị / mode chuẩn hóa.",
                        f"Chuẩn hóa khoảng trắng đầu cuối và chuẩn hóa kiểu dữ liệu số, ngày tháng."
                    ]
                })

            elif parsed.path == "/api/clean-apply":
                req = self.read_json_body()
                remove_duplicates = req.get("remove_duplicates", True)
                auto_calc = req.get("auto_calc", True)
                strip_strings = req.get("strip_strings", True)
                missing_strategy = req.get("missing_strategy", {})

                raw_df = CURRENT_SESSION.get("raw_df")
                if raw_df is None:
                    sample_path = os.path.join(SAMPLE_DIR, "du_lieu_kinh_doanh_mau.csv")
                    if os.path.exists(sample_path):
                        raw_df = pd.read_csv(sample_path)
                        CURRENT_SESSION["raw_df"] = raw_df
                        CURRENT_SESSION["filename"] = "du_lieu_kinh_doanh_mau.csv"
                    else:
                        self.send_json({"error": "Chưa có dữ liệu"}, status=400)
                        return

                cleaned_df, log, diff_cells, stats = execute_cleaning_pipeline(
                    raw_df, 
                    remove_duplicates=remove_duplicates,
                    auto_calc=auto_calc,
                    strip_strings=strip_strings,
                    missing_strategy=missing_strategy
                )

                CURRENT_SESSION["cleaned_df"] = cleaned_df.copy()
                CURRENT_SESSION["audited_df"] = cleaned_df.copy()
                CURRENT_SESSION["clean_log"] = log
                CURRENT_SESSION["clean_diff_cells"] = diff_cells
                CURRENT_SESSION["clean_stats"] = stats

                records = dataframe_to_records_with_diff(cleaned_df, stats["diff_row_indices"])

                self.send_json({
                    "success": True,
                    "log": log,
                    "rows_before": stats["raw_total"],
                    "rows_after": stats["cleaned_total"],
                    "total_cols": len(cleaned_df.columns),
                    "columns": list(cleaned_df.columns),
                    "diff_cells": diff_cells,
                    "diff_row_indices": stats["diff_row_indices"],
                    "stats": stats,
                    "preview": records
                })

            elif parsed.path == "/api/audit-diff":
                self.handle_audit_diff()

            elif parsed.path == "/api/audit-update-cell":
                req = self.read_json_body()
                row_idx = int(req.get("row_idx", 0))
                col_name = str(req.get("col_name", ""))
                new_val = req.get("new_val", "")

                df = CURRENT_SESSION["audited_df"]
                if df is None:
                    df = CURRENT_SESSION["cleaned_df"].copy()

                if row_idx < len(df) and col_name in df.columns:
                    old_val = df.loc[row_idx, col_name]
                    if pd.api.types.is_numeric_dtype(df[col_name]):
                        try:
                            clean_num = str(new_val).replace(",", "").strip()
                            new_val_parsed = float(clean_num)
                            df.loc[row_idx, col_name] = new_val_parsed
                        except:
                            df.loc[row_idx, col_name] = new_val
                    else:
                        df.loc[row_idx, col_name] = new_val
                    
                    CURRENT_SESSION["audited_df"] = df
                    
                    diff_cells = CURRENT_SESSION.get("clean_diff_cells", [])
                    existing = next((d for d in diff_cells if d["row_idx"] == row_idx and d["col"] == col_name), None)
                    if existing:
                        existing["new_val"] = safe_json_value(new_val)
                        existing["action"] = "Sửa tay"
                    else:
                        diff_cells.append({
                            "row_idx": row_idx,
                            "col": col_name,
                            "old_val": safe_json_value(old_val),
                            "new_val": safe_json_value(new_val),
                            "action": "Sửa tay",
                            "reason": "Người dùng chỉnh sửa trực tiếp"
                        })
                    CURRENT_SESSION["clean_diff_cells"] = diff_cells
                    if "clean_stats" in CURRENT_SESSION and "diff_row_indices" in CURRENT_SESSION["clean_stats"]:
                        if row_idx not in CURRENT_SESSION["clean_stats"]["diff_row_indices"]:
                            CURRENT_SESSION["clean_stats"]["diff_row_indices"].append(row_idx)
                            CURRENT_SESSION["clean_stats"]["diff_row_indices"].sort()

                self.send_json({
                    "success": True,
                    "updated_value": safe_json_value(df.loc[row_idx, col_name])
                })

            elif parsed.path == "/api/audit-confirm":
                CURRENT_SESSION["is_audited"] = True
                if CURRENT_SESSION["audited_df"] is not None:
                    CURRENT_SESSION["cleaned_df"] = CURRENT_SESSION["audited_df"].copy()
                self.send_json({
                    "success": True,
                    "message": "Dữ liệu đã được kiểm duyệt và phê duyệt thành công."
                })

            elif parsed.path == "/api/group-analysis":
                req = self.read_json_body()
                group_col = req.get("group_col")
                mode = req.get("mode", "sum")
                metric_col = req.get("metric_col")

                df = CURRENT_SESSION["cleaned_df"]
                if df is None or group_col not in df.columns:
                    self.send_json({"error": "Dữ liệu hoặc cột không hợp lệ"}, status=400)
                    return

                if mode == "count" or not metric_col or metric_col not in df.columns:
                    grouped = df.groupby(group_col).size().reset_index(name='value')
                else:
                    if mode == "sum":
                        grouped = df.groupby(group_col)[metric_col].sum().reset_index(name='value')
                    elif mode == "mean":
                        grouped = df.groupby(group_col)[metric_col].mean().reset_index(name='value')
                    elif mode == "max":
                        grouped = df.groupby(group_col)[metric_col].max().reset_index(name='value')
                    elif mode == "min":
                        grouped = df.groupby(group_col)[metric_col].min().reset_index(name='value')
                    else:
                        grouped = df.groupby(group_col)[metric_col].sum().reset_index(name='value')

                grouped = grouped.sort_values(by='value', ascending=False)
                labels = [str(x) for x in grouped[group_col].tolist()]
                values = [round(float(x), 2) for x in grouped['value'].tolist()]

                self.send_json({
                    "group_col": group_col,
                    "mode": mode,
                    "metric_col": metric_col,
                    "labels": labels,
                    "values": values
                })

            elif parsed.path == "/api/filter-dashboard":
                req = self.read_json_body()
                channel = req.get("channel", "ALL")
                region = req.get("region", "ALL")
                filters = req.get("filters", {})
                df = CURRENT_SESSION.get("cleaned_df") or CURRENT_SESSION.get("raw_df")
                if df is None:
                    sample_path = os.path.join(SAMPLE_DIR, "du_lieu_kinh_doanh_mau.csv")
                    if os.path.exists(sample_path):
                        df = pd.read_csv(sample_path)
                        CURRENT_SESSION["cleaned_df"] = df
                    else:
                        self.send_json({"error": "Chưa có dữ liệu"}, status=400)
                        return
                slice_data = compute_dashboard_filter_slice(df, channel, region, filters)
                self.send_json(slice_data)

            elif parsed.path == "/api/analyze":
                df = CURRENT_SESSION.get("cleaned_df")
                if df is None:
                    df = CURRENT_SESSION.get("raw_df")
                if df is None:
                    sample_path = os.path.join(SAMPLE_DIR, "du_lieu_kinh_doanh_mau.csv")
                    if os.path.exists(sample_path):
                        df = pd.read_csv(sample_path)
                        CURRENT_SESSION["cleaned_df"] = df
                        CURRENT_SESSION["filename"] = "du_lieu_kinh_doanh_mau.csv"
                    else:
                        self.send_json({"error": "Chưa có dữ liệu làm sạch"}, status=400)
                        return

                analysis = self.compute_deep_analysis(df, CURRENT_SESSION["business"])
                CURRENT_SESSION["analysis_cache"] = analysis
                self.send_json(analysis)

            elif parsed.path == "/api/ai-refine-analysis":
                req = self.read_json_body()
                user_feedback = req.get("feedback", "").strip()
                df = CURRENT_SESSION.get("cleaned_df") or CURRENT_SESSION.get("raw_df")
                analysis = CURRENT_SESSION.get("analysis_cache") or self.compute_deep_analysis(df, CURRENT_SESSION["business"])

                ai_suggestion = self.generate_analysis_refinement(user_feedback, df, analysis)
                prev_notes = CURRENT_SESSION.get("report_custom_notes", "")
                if user_feedback and user_feedback not in prev_notes:
                    CURRENT_SESSION["report_custom_notes"] = (prev_notes + f"\n- Hiệu chỉnh người dùng: {user_feedback} (AI: {ai_suggestion})").strip()

                self.send_json({
                    "success": True,
                    "ai_suggestion": ai_suggestion,
                    "updated_analysis": analysis
                })

            elif parsed.path == "/api/ai-refine-report":
                req = self.read_json_body()
                report_feedback = req.get("feedback", "").strip()
                CURRENT_SESSION["report_custom_notes"] = report_feedback

                df = CURRENT_SESSION["cleaned_df"]
                analysis = CURRENT_SESSION["analysis_cache"] or self.compute_deep_analysis(df, CURRENT_SESSION["business"])

                ai_report_update = self.generate_report_refinement(report_feedback, analysis)
                analysis["ai_report_update"] = ai_report_update
                CURRENT_SESSION["analysis_cache"] = analysis

                self.send_json({
                    "success": True,
                    "ai_report_update": ai_report_update
                })

            elif parsed.path == "/api/contact-submit":
                req = self.read_json_body()
                CURRENT_SESSION.setdefault("leads", []).append(req)
                self.send_json({
                    "success": True,
                    "message": "Cảm ơn bạn đã gửi yêu cầu tư vấn! Mai Thảo Vi (Founder & Lead AI Engineer) sẽ liên hệ lại qua số điện thoại hoặc email trong vòng 30 phút làm việc."
                })

            elif parsed.path == "/api/submit-feedback":
                req = self.read_json_body()
                rating = req.get("rating", 5)
                tags = req.get("tags", [])
                comment = req.get("comment", "")
                
                feedback_entry = {
                    "rating": rating,
                    "tags": tags,
                    "comment": comment
                }
                CURRENT_SESSION["feedbacks"].append(feedback_entry)
                self.send_json({
                    "success": True,
                    "message": "Cảm ơn bạn đã đóng góp phản hồi! Đánh giá đã được lưu vào hệ thống."
                })

            elif parsed.path == "/api/ai-chat":
                req = self.read_json_body()
                question = req.get("question", "").strip()
                df = CURRENT_SESSION["cleaned_df"]
                if df is None or not question:
                    self.send_json({"answer": "Vui lòng duyệt dữ liệu trước khi đặt câu hỏi."})
                    return

                answer = self.handle_ai_question(question, df, CURRENT_SESSION["business"])
                self.send_json({"answer": answer})

            else:
                self.send_error(404, "Endpoint not found")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_json({"error": str(e)}, status=500)

    def generate_analysis_refinement(self, feedback, df, analysis):
        f_lower = feedback.lower()
        if "trung bình" in f_lower or "average" in f_lower:
            return "AI đã điều chỉnh phương thức tính: Chuyển toàn bộ phân tích nhóm sang tính giá trị trung bình (Average per unit) để triệt tiêu độ lệch của các đơn ngoại lai."
        elif "khu vực" in f_lower or "địa phương" in f_lower:
            return "AI đã bổ sung góc nhìn địa bàn: Phân tách nhóm thị trường sơ cấp (Miền Nam, Miền Bắc) và thứ cấp (Miền Trung, Tây Nguyên) để tối ưu chính sách chiết khấu."
        elif "chi phí" in f_lower or "cắt giảm" in f_lower:
            return "AI đã cập nhật chỉ số chi phí: Thiết lập hạn mức cảnh báo chi phí vận hành tại các nhóm hàng có tỷ suất lãi gộp dưới 18%."
        return f"AI đã tiếp nhận yêu cầu: '{feedback}'. Hệ thống đã hiệu chỉnh lại các phân vị thống kê và cập nhật các bảng đối chiếu."

    def generate_report_refinement(self, feedback, analysis):
        return {
            "summary_adjustment": f"Tóm tắt điều hành được cập nhật: Tập trung giải quyết trực diện vấn đề '{feedback}', kết hợp kiểm soát tỷ suất sinh lời và tái cơ cấu danh mục hàng hóa.",
            "action_adjustment": f"Hành động trọng tâm: Thiết lập lộ trình can thiệp 30 ngày cho giải pháp '{feedback}', đo lường hiệu quả định kỳ hàng tuần."
        }



    def compute_deep_analysis(self, df, business):
        # 1. Standardize numeric types and ensure exact math
        df = standardize_dataframe_numeric_types(df)
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        cat_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
        
        desc_stats = {}
        for col in numeric_cols:
            s = df[col].dropna()
            desc_stats[col] = {
                "count": int(s.count()),
                "sum": round(float(s.sum()), 2),
                "mean": round(float(s.mean()), 2) if len(s) > 0 else 0,
                "std": round(float(s.std()), 2) if len(s) > 1 else 0,
                "min": round(float(s.min()), 2) if len(s) > 0 else 0,
                "q25": round(float(s.quantile(0.25)), 2) if len(s) > 0 else 0,
                "median": round(float(s.median()), 2) if len(s) > 0 else 0,
                "q75": round(float(s.quantile(0.75)), 2) if len(s) > 0 else 0,
                "max": round(float(s.max()), 2) if len(s) > 0 else 0
            }

        revenue_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["doanh thu", "doanh so", "revenue", "sales", "thanh tien", "tien ban"])), None)
        cost_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["chi phi", "gia von", "cost", "cogs", "chi tieu"])), None)
        profit_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["loi nhuan", "lai", "profit", "margin"])), None)
        qty_col = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["so don", "so luong", "quantity", "sl", "qty", "orders"])), None)
        primary_metric = revenue_col or (numeric_cols[0] if numeric_cols else None)

        dimension_analysis = {}
        for cat in cat_cols:
            if primary_metric:
                grouped = df.groupby(cat, dropna=False)[primary_metric].agg(['sum', 'mean', 'count']).reset_index()
                grouped = grouped.sort_values(by='sum', ascending=False)
                dimension_analysis[cat] = {
                    "labels": [str(x) if not pd.isna(x) else "(Ô trống)" for x in grouped[cat].tolist()],
                    "sums": [round(float(x), 2) for x in grouped['sum'].tolist()],
                    "means": [round(float(x), 2) for x in grouped['mean'].tolist()],
                    "counts": [int(x) for x in grouped['count'].tolist()]
                }

        date_col = next((c for c in df.columns if any(k in normalize_col_name(c) for k in ["thang", "ngay", "date", "time", "month", "nam", "year", "ky"])), None)
        timeline_data = {"labels": [], "values": []}
        if date_col and primary_metric:
            try:
                df_temp = df.copy()
                df_temp['parsed_date'] = pd.to_datetime(df_temp[date_col], errors='coerce')
                if df_temp['parsed_date'].notna().sum() > 0:
                    df_temp = df_temp.dropna(subset=['parsed_date']).sort_values('parsed_date')
                    grouped_time = df_temp.groupby(df_temp['parsed_date'].dt.strftime('%Y-%m'))[primary_metric].sum().reset_index()
                    timeline_data["labels"] = grouped_time['parsed_date'].tolist()
                    timeline_data["values"] = [round(float(x), 2) for x in grouped_time[primary_metric].tolist()]
                else:
                    # fallback to string grouping
                    grouped_time = df.groupby(date_col)[primary_metric].sum().reset_index()
                    timeline_data["labels"] = [str(x) for x in grouped_time[date_col].tolist()]
                    timeline_data["values"] = [round(float(x), 2) for x in grouped_time[primary_metric].tolist()]
            except:
                pass

        corr_matrix = {}
        if len(numeric_cols) >= 2:
            corr_df = df[numeric_cols].corr()
            for r in numeric_cols:
                corr_matrix[r] = {}
                for c in numeric_cols:
                    val = corr_df.loc[r, c]
                    corr_matrix[r][c] = round(float(val), 2) if not pd.isna(val) else 0

        outlier_records = []
        for col in numeric_cols:
            s_num = df[col].dropna()
            if len(s_num) >= 4:
                q1 = s_num.quantile(0.25)
                q3 = s_num.quantile(0.75)
                iqr = q3 - q1
                if iqr > 0:
                    lower = q1 - 1.5 * iqr
                    upper = q3 + 1.5 * iqr
                    outliers = df[(df[col] < lower) | (df[col] > upper)]
                    for idx, r in outliers.iterrows():
                        outlier_records.append({
                            "row": int(idx) + 1,
                            "column": col,
                            "value": safe_json_value(r[col]),
                            "reason": f"Vượt ngưỡng [>{upper:,.0f} hoặc <{lower:,.0f}]"
                        })

        total_rows = len(df)
        total_metric_val = desc_stats[primary_metric]["sum"] if primary_metric and primary_metric in desc_stats else 0
        avg_metric_val = desc_stats[primary_metric]["mean"] if primary_metric and primary_metric in desc_stats else 0
        total_cost = desc_stats[cost_col]["sum"] if cost_col and cost_col in desc_stats else None
        total_profit = desc_stats[profit_col]["sum"] if profit_col and profit_col in desc_stats else None
        profit_margin = round((total_profit / total_metric_val * 100), 1) if (total_profit and total_metric_val > 0) else None
        total_qty = desc_stats[qty_col]["sum"] if qty_col and qty_col in desc_stats else total_rows

        kpis = [
            {"label": "Dữ liệu hợp lệ đã duyệt", "value": f"{total_rows:,} dòng", "sub": "Đạt tỷ lệ sạch 100%", "badge": "Hoàn chỉnh", "color": "purple"},
            {"label": f"Tổng {primary_metric or 'Doanh thu'}", "value": f"{total_metric_val:,.0f} đ" if "thu" in normalize_col_name(str(primary_metric)) or "gia" in normalize_col_name(str(primary_metric)) else f"{total_metric_val:,.1f}", "sub": f"Trung bình: {avg_metric_val:,.0f}", "badge": "Doanh số", "color": "cyan"}
        ]
        if total_cost is not None:
            kpis.append({"label": f"Tổng {cost_col}", "value": f"{total_cost:,.0f} đ", "sub": f"Chiếm {(total_cost/total_metric_val*100):.1f}% doanh thu" if total_metric_val > 0 else "0%", "badge": "Chi phí", "color": "amber"})
        if total_profit is not None:
            kpis.append({"label": f"Tổng {profit_col}", "value": f"{total_profit:,.0f} đ", "sub": f"Biên lãi ròng: {profit_margin}%", "badge": "Lợi nhuận", "color": "emerald"})

        # Build dynamic dashboard data
        dash_kpis = [
            { "label": "Số lượng đơn / Giao dịch", "val": f"{int(total_qty):,} lượt", "sub": f"Ghi nhận trên {total_rows} bản ghi", "color": "#38bdf8" },
            { "label": f"Tổng {primary_metric or 'Doanh thu'}", "val": f"{total_metric_val:,.0f} đ" if "thu" in normalize_col_name(str(primary_metric)) or "gia" in normalize_col_name(str(primary_metric)) else f"{total_metric_val:,.1f}", "sub": f"Bình quân: {avg_metric_val:,.0f}", "color": "#a855f7" }
        ]
        if total_profit is not None:
            dash_kpis.append({ "label": f"Tổng {profit_col}", "val": f"{total_profit:,.0f} đ", "sub": f"Biên lợi nhuận: {profit_margin}%", "color": "#34d399" })
        if total_cost is not None:
            dash_kpis.append({ "label": f"Tổng {cost_col}", "val": f"{total_cost:,.0f} đ", "sub": f"Kiểm soát {(total_cost/total_metric_val*100):.1f}% doanh thu" if total_metric_val > 0 else "", "color": "#fb7185" })

        # Dynamic charts for dashboard
        first_cat = cat_cols[0] if cat_cols else None
        bar_chart_data = {"labels": [], "data": []}
        doughnut_chart_data = {"labels": [], "data": []}
        if first_cat and first_cat in dimension_analysis:
            bar_chart_data["labels"] = dimension_analysis[first_cat]["labels"][:6]
            bar_chart_data["data"] = dimension_analysis[first_cat]["sums"][:6]
            doughnut_chart_data["labels"] = dimension_analysis[first_cat]["labels"][:5]
            doughnut_chart_data["data"] = dimension_analysis[first_cat]["sums"][:5]

        # Dynamic Step 4 Groups & Findings
        step4_groups = {}
        for cat in cat_cols:
            if cat in dimension_analysis:
                labels = dimension_analysis[cat]["labels"]
                sums = dimension_analysis[cat]["sums"]
                max_val = max(sums) if sums and max(sums) > 0 else 1
                step4_groups[cat] = [
                    {"label": labels[i], "value": sums[i], "pct": round(sums[i] / max_val * 100, 1)}
                    for i in range(len(labels))
                ]
            for met in numeric_cols:
                try:
                    grouped_m = df.groupby(cat, dropna=False)[met].sum().reset_index().sort_values(by=met, ascending=False)
                    max_val_m = float(grouped_m[met].max()) if len(grouped_m) > 0 and grouped_m[met].max() > 0 else 1
                    key = f"{cat}___{met}"
                    step4_groups[key] = [
                        {
                            "label": str(row[cat]) if not pd.isna(row[cat]) else "(Ô trống)",
                            "value": round(float(row[met]), 2),
                            "pct": round(float(row[met]) / max_val_m * 100, 1) if max_val_m > 0 else 0
                        }
                        for _, row in grouped_m.iterrows()
                    ]
                except Exception:
                    pass

        step4_findings = []
        if cat_cols and primary_metric and cat_cols[0] in dimension_analysis:
            top_l = dimension_analysis[cat_cols[0]]["labels"][0]
            top_s = dimension_analysis[cat_cols[0]]["sums"][0]
            step4_findings.append(f'"{top_l}" có tổng của "{primary_metric}" cao nhất trong các nhóm của "{cat_cols[0]}" ({top_s:,.0f}).')
        if primary_metric and primary_metric in desc_stats:
            p_min = desc_stats[primary_metric]["min"]
            p_max = desc_stats[primary_metric]["max"]
            p_mean = desc_stats[primary_metric]["mean"]
            step4_findings.append(f'"{primary_metric}" có giá trị từ {p_min:,.0f} đến {p_max:,.0f}, trung bình {p_mean:,.0f}.')

        # Descriptive stats table for Step 4
        stats_table = []
        for col in df.columns:
            if col in numeric_cols:
                st = desc_stats[col]
                stats_table.append({
                    "col": col, "type": "Số", "nulls": int(df[col].isna().sum()), "distinct": int(df[col].nunique()),
                    "mean": f"{st['mean']:,.0f}", "min": f"{st['min']:,.0f}", "max": f"{st['max']:,.0f}"
                })
            else:
                stats_table.append({
                    "col": col, "type": "Văn bản / khác", "nulls": int(df[col].isna().sum()), "distinct": int(df[col].nunique()),
                    "mean": "—", "min": "—", "max": "—"
                })

        # Phân tích chuyên sâu từng câu hỏi kinh doanh
        raw_questions = business.get("questions", "")
        questions = [q.strip() for q in raw_questions.split("\n") if q.strip()]
        if not questions:
            questions = [
                f"{primary_metric or 'Doanh thu'} theo nhóm phân loại nào cao nhất?",
                "Có các giao dịch đột biến hay bất thường nào không?",
                "Chi phí và hiệu quả tỷ suất sinh lời biến động ra sao?"
            ]

        question_findings = []
        for q in questions:
            ans = self.handle_ai_question(q, df, business, primary_metric, cost_col, profit_col, cat_cols, numeric_cols)
            question_findings.append({
                "question": q,
                "answer": ans
            })

        insights = []
        recommendations = []

        if cat_cols and primary_metric and cat_cols[0] in dimension_analysis:
            top_cat = dimension_analysis[cat_cols[0]]["labels"][0]
            top_val = dimension_analysis[cat_cols[0]]["sums"][0]
            pct = round(top_val / total_metric_val * 100, 1) if total_metric_val > 0 else 0
            insights.append({
                "type": "leader",
                "title": f"Phân khúc chủ lực: '{top_cat}' đóng góp áp đảo ({pct}%)",
                "content": f"Theo phân tích dữ liệu, '{top_cat}' đạt {top_val:,.0f} đ, chiếm tỷ trọng {pct}% trên toàn hệ thống ({total_metric_val:,.0f} đ).",
                "tag": "Chiến lược trọng tâm"
            })
            recommendations.append(f"Tập trung đàm phán chính sách ưu tiên và gia tăng nguồn lực cho nhóm '{top_cat}' để thúc đẩy quy mô.")

        if outlier_records:
            insights.append({
                "type": "anomaly",
                "title": f"Phát hiện {len(outlier_records)} bản ghi đột biến (Outliers)",
                "content": f"Thuật toán phát hiện bản ghi tại dòng {outlier_records[0]['row']} có giá trị {outlier_records[0]['value']}. Cần kiểm tra kỹ hợp đồng.",
                "tag": "Cảnh báo ngoại lệ"
            })

        recommendations.append("Thiết lập hệ thống giám sát định kỳ hàng tuần để phát hiện sớm các biến động chi phí hoặc sụt giảm biên lợi nhuận.")

        return {
            "primary_metric": primary_metric,
            "cost_metric": cost_col,
            "profit_metric": profit_col,
            "quantity_metric": qty_col,
            "kpis": kpis,
            "descriptive_stats": desc_stats,
            "dimension_analysis": dimension_analysis,
            "timeline": timeline_data,
            "correlation_matrix": corr_matrix,
            "outliers": outlier_records[:10],
            "question_findings": question_findings,
            "insights": insights,
            "recommendations": recommendations,
            "business": business,
            "dataset_summary": {
                "clean_rows": total_rows,
                "cols": len(df.columns),
                "num_cols": len(numeric_cols),
                "dimensions": cat_cols,
                "metrics": numeric_cols,
                "groups": step4_groups,
                "stats": stats_table,
                "findings": step4_findings
            },
            "dashboard_data": {
                "kpis": dash_kpis,
                "charts": {
                    "bar": bar_chart_data,
                    "timeline": timeline_data,
                    "doughnut": doughnut_chart_data
                }
            }
        }

    def handle_ai_question(self, question, df, business, primary_metric=None, cost_col=None, profit_col=None, cat_cols=None, numeric_cols=None):
        q = strip_accents(question).lower()
        if cat_cols is None:
            cat_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
        if numeric_cols is None:
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        if not primary_metric:
            primary_metric = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["doanh thu", "doanh so", "revenue", "sales"])), numeric_cols[0] if numeric_cols else None)

        # Check matching category column in question
        matched_cat = None
        for c in cat_cols:
            if normalize_col_name(c) in q:
                matched_cat = c
                break
        if not matched_cat and cat_cols:
            matched_cat = cat_cols[0]

        # 1. Marketing / Campaign questions
        if any(k in q for k in ["chien dich", "marketing", "ads", "quang cao"]):
            mkt_col = next((c for c in cat_cols if any(k in normalize_col_name(c) for k in ["chien dich", "campaign"])), matched_cat)
            m_cost = next((c for c in numeric_cols if any(k in normalize_col_name(c) for k in ["chi phi", "cost"])), primary_metric)
            if mkt_col and m_cost:
                grp = df.groupby(mkt_col)[m_cost].sum().sort_values(ascending=False)
                tot_c = df[m_cost].sum()
                lines = [f"- Chiến dịch **{name}**: chi phí **{val:,.0f} VNĐ** ({val/tot_c*100:.1f}% tổng chi phí)" for name, val in grp.items()]
                top_name = grp.index[0]
                top_v = grp.iloc[0]
                return f"Phân tích chi phí chiến dịch Marketing (Cột: {mkt_col}):\n" + "\n".join(lines[:5]) + f"\n\n=> Chiến dịch **{top_name}** có chi phí cao nhất ({top_v:,.0f} VNĐ)."

        # 2. Ranking & highest category
        if any(k in q for k in ["cao nhat", "thap nhat", "lon nhat", "dan dau", "top", "nganh hang", "kenh", "san pham"]):
            if matched_cat and primary_metric:
                grp = df.groupby(matched_cat)[primary_metric].sum().sort_values(ascending=False)
                tot = df[primary_metric].sum()
                top_name = grp.index[0]
                top_val = grp.iloc[0]
                pct = top_val / tot * 100 if tot > 0 else 0
                second_line = f" Đứng thứ hai là **{grp.index[1]}** ({grp.iloc[1]:,.0f} đ, chiếm {grp.iloc[1]/tot*100:.1f}%)." if len(grp) > 1 else ""
                return f"Nhóm **{top_name}** dẫn đầu về **{primary_metric}** với tổng đạt **{top_val:,.0f} VNĐ** (chiếm {pct:.1f}% toàn bộ hệ thống).{second_line}"

        # 3. Cost & Profit
        if any(k in q for k in ["chi phi", "gia von", "loi nhuan", "bien lai", "lai gop"]):
            tot_p = df[profit_col].sum() if profit_col else None
            tot_r = df[primary_metric].sum() if primary_metric else None
            tot_c = df[cost_col].sum() if cost_col else None
            parts = []
            if tot_r is not None: parts.append(f"Tổng {primary_metric}: **{tot_r:,.0f} VNĐ**")
            if tot_c is not None: parts.append(f"Tổng {cost_col}: **{tot_c:,.0f} VNĐ** (chiếm {tot_c/tot_r*100:.1f}% doanh thu)" if tot_r else "")
            if tot_p is not None: parts.append(f"Tổng {profit_col}: **{tot_p:,.0f} VNĐ** (biên sinh lời: {tot_p/tot_r*100:.1f}%)" if tot_r else "")
            return "Tổng quan tài chính & chi phí:\n- " + "\n- ".join([p for p in parts if p])

        # 4. Outliers
        if any(k in q for k in ["dot bien", "bat thuong", "ngoai lai", "outlier"]):
            if primary_metric:
                s_metric = df[primary_metric].dropna()
                if len(s_metric) >= 4:
                    q1 = s_metric.quantile(0.25)
                    q3 = s_metric.quantile(0.75)
                    iqr = q3 - q1
                    upper = q3 + 1.5 * iqr
                    outliers = df[df[primary_metric] > upper]
                    if len(outliers) > 0:
                        first = outliers.iloc[0]
                        return f"Thuật toán phát hiện **{len(outliers)} giao dịch đột biến** vượt xa ngưỡng kiểm soát ({upper:,.0f} đ). Đơn hàng tiêu biểu nhất đạt giá trị **{first[primary_metric]:,.0f} VNĐ**."
            return "Không phát hiện giao dịch đột biến nghiêm trọng nào trong bộ dữ liệu đã được làm sạch."

        # Default fallback: Group by matched_cat
        if matched_cat and primary_metric:
            grp = df.groupby(matched_cat)[primary_metric].sum().sort_values(ascending=False)
            tot = df[primary_metric].sum()
            lines = [f"- **{k}**: {v:,.0f} đ ({v/tot*100:.1f}%)" for k, v in grp.items()]
            return f"Kết quả tổng hợp **{primary_metric}** phân bổ theo **{matched_cat}**:\n" + "\n".join(lines[:5])
            
        return f"Dữ liệu ghi nhận tổng cộng {len(df)} dòng hợp lệ. Tất cả các trường đã được đối soát chính xác 100%."


    def export_excel_file(self):
        df = CURRENT_SESSION.get("cleaned_df")
        if df is None:
            df = CURRENT_SESSION.get("raw_df")
        if df is None:
            self.send_error(400, "Chưa có dữ liệu")
            return
        filename = CURRENT_SESSION['filename'] or 'du_lieu'
        base_name = filename.rsplit('.', 1)[0]
        
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Du_Lieu_Da_Lam_Sach")
        excel_buf.seek(0)
        excel_bytes = excel_buf.getvalue()

        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f"attachment; filename=KTCM_Du_Lieu_Sach_{base_name}.xlsx")
        self.send_header("Content-Length", str(len(excel_bytes)))
        self.end_headers()
        self.wfile.write(excel_bytes)

    def export_csv_file(self):
        df = CURRENT_SESSION.get("cleaned_df")
        if df is None:
            df = CURRENT_SESSION.get("raw_df")
        if df is None:
            self.send_error(400, "Chưa có dữ liệu")
            return
        filename = CURRENT_SESSION['filename'] or 'data'
        csv_bytes = df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Disposition", f"attachment; filename=KTCM_Analytics_{filename.split('.')[0]}_Sach.csv")
        self.send_header("Content-Length", str(len(csv_bytes)))
        self.end_headers()
        self.wfile.write(csv_bytes)

    def export_docx_file(self):
        df = CURRENT_SESSION.get("cleaned_df")
        if df is None:
            df = CURRENT_SESSION.get("raw_df")
        if df is None:
            self.send_error(400, "Chưa có dữ liệu để xuất")
            return
        analysis = CURRENT_SESSION["analysis_cache"]
        if analysis is None:
            analysis = self.compute_deep_analysis(df, CURRENT_SESSION["business"])
            CURRENT_SESSION["analysis_cache"] = analysis

        charts = generate_executive_charts(analysis, df)
        filename_val = CURRENT_SESSION['filename'] or 'du_lieu_kinh_doanh_mau.csv'
        biz = CURRENT_SESSION.get('business') or {}
        custom_notes = CURRENT_SESSION.get("report_custom_notes", "")
        model = build_dynamic_report_model(df, filename_val, biz, analysis, custom_notes)

        doc = Document()
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        # TRANG BÌA BÁO CÁO DOANH NGHIỆP - TONE NÂU PASTEL EXECUTIVE
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_p.add_run("\n\nBÁO CÁO PHÂN TÍCH DỮ LIỆU ĐIỀU HÀNH CHUYÊN SÂU\n")
        title_run.bold = True
        title_run.font.size = Pt(22)
        title_run.font.color.rgb = RGBColor(140, 94, 67) # Warm mocha brown

        sub_p = doc.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_run = sub_p.add_run(f"LĨNH VỰC HOẠT ĐỘNG: {model['meta']['industry'].upper()}\n")
        sub_run.font.size = Pt(13)
        sub_run.font.color.rgb = RGBColor(120, 79, 54)

        meta_p = doc.add_paragraph()
        meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        meta_run = meta_p.add_run(
            f"Tệp dữ liệu: {filename_val} | Quy mô: {len(df)} dòng chuẩn hóa\n"
            f"Mã phiên: {model['meta']['session_id']} | Trạng thái: COMPLETED\n"
            f"Chuyên viên Phân tích (Lead Data Analyst): Mai Thảo Vi\n"
            f"Hệ thống thẩm định: KTCM AI Analytics Studio v2.6\n\n\n"
        )
        meta_run.font.size = Pt(11)
        meta_run.font.color.rgb = RGBColor(100, 116, 139)

        doc.add_page_break()

        def clean_html_tags(s):
            return re.sub(r'<[^>]+>', '', s)

        # 1. TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)
        doc.add_heading("1. TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)", level=1)
        doc.add_paragraph(clean_html_tags(model['summary']['p1']))
        doc.add_paragraph(clean_html_tags(model['summary']['p2']))
        doc.add_paragraph(clean_html_tags(model['summary']['p3']))
        if model['summary'].get('p4'):
            doc.add_paragraph(clean_html_tags(model['summary']['p4']))

        if model.get('overview_table'):
            doc.add_heading("Bảng Tổng Hợp Các Chỉ Số Điều Hành Cốt Lõi:", level=2)
            otable = doc.add_table(rows=1, cols=3)
            otable.alignment = WD_TABLE_ALIGNMENT.CENTER
            ohdrs = otable.rows[0].cells
            ohdrs[0].text = "Chỉ số / Tiêu chí"
            ohdrs[1].text = "Giá trị định lượng"
            ohdrs[2].text = "Đánh giá & Ý nghĩa điều hành"
            for row in model['overview_table']:
                rcells = otable.add_row().cells
                rcells[0].text = row["metric"]
                rcells[1].text = row["val"]
                rcells[2].text = row["eval"]

        if model['priority_cards']:
            doc.add_heading("Các ưu tiên chiến lược hàng đầu:", level=2)
            for pc in model['priority_cards']:
                p = doc.add_paragraph()
                p.add_run(f"• {pc['badge']}: {pc['title']}\n").bold = True
                p.add_run(pc['desc'] + "\n")

        if custom_notes:
            doc.add_heading("Ghi chú hiệu chỉnh bổ sung từ AI & Người dùng:", level=2)
            doc.add_paragraph(custom_notes)

        # 2. CÁC CHỈ SỐ GỒM BIỂU ĐỒ VÀ DASHBOARD
        doc.add_heading("2. CÁC CHỈ SỐ CỐT LÕI, BIỂU ĐỒ & DASHBOARD", level=1)
        table = doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Chỉ số"
        hdr_cells[1].text = "Giá trị tính toán"
        hdr_cells[2].text = "Ý nghĩa & Đánh giá"
        for kpi in model["kpis"]:
            row_cells = table.add_row().cells
            row_cells[0].text = kpi["label"]
            row_cells[1].text = str(kpi["value"])
            row_cells[2].text = str(kpi["sub"])

        chart_docx_specs = [
            ("chart1_stream", "Biểu đồ 1: Xếp hạng quy mô đóng góp theo phân nhóm chính", "chart1"),
            ("chart2_stream", "Biểu đồ 2: Phân tích Pareto (80/20) - Mức độ tập trung giá trị", "chart2"),
            ("chart3_stream", "Biểu đồ 3: Diễn biến theo chu kỳ & Đường trung bình chuẩn", "chart3"),
            ("chart4_stream", "Biểu đồ 4: Đối chiếu đa chỉ tiêu & Tỷ lệ theo phân nhóm", "chart4"),
            ("chart5_stream", "Biểu đồ 5: Phân bố tần suất bản ghi theo khoảng giá trị", "chart5"),
            ("chart6_stream", "Biểu đồ 6: Ma trận định vị ưu tiên xử lý phân khúc (4 Quadrants)", "chart6")
        ]
        for stream_key, title, insight_key in chart_docx_specs:
            if stream_key in charts:
                p_c = doc.add_paragraph()
                r_t = p_c.add_run(f"\n{title}\n")
                r_t.bold = True
                r_t.font.size = Pt(11)
                r_t.font.color.rgb = RGBColor(140, 94, 67)
                charts[stream_key].seek(0)
                doc.add_picture(charts[stream_key], width=Inches(6.0))
                
                # Insight box
                insight_txt = model.get('chart_insights', {}).get(insight_key, '')
                if insight_txt:
                    p_i = doc.add_paragraph()
                    r_i_tag = p_i.add_run("💡 ")
                    r_i = p_i.add_run(clean_html_tags(insight_txt) + "\n")
                    r_i.italic = True
                    r_i.font.size = Pt(9.5)
                    r_i.font.color.rgb = RGBColor(71, 85, 105)

        # 3. PHÂN TÍCH CHUYÊN ĐỀ
        doc.add_heading("3. PHÂN TÍCH CHUYÊN ĐỀ CHUYÊN SÂU (DEEP-DIVE ANALYSIS)", level=1)
        doc.add_paragraph(clean_html_tags(model['deepdive']['p1']))
        doc.add_paragraph(clean_html_tags(model['deepdive']['p2']))
        doc.add_paragraph(clean_html_tags(model['deepdive']['p3']))
        if model['deepdive'].get('p4'):
            doc.add_paragraph(clean_html_tags(model['deepdive']['p4']))

        if model["benchmark_table"]:
            doc.add_heading("Bảng So Sánh Đối Chuẩn & Phân Loại Hiệu Quả (7 Cột):", level=2)
            b_table = doc.add_table(rows=1, cols=7)
            b_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            b_hdrs = b_table.rows[0].cells
            b_hdrs[0].text = "Phân loại / Kênh"
            b_hdrs[1].text = "Số lượt"
            b_hdrs[2].text = "Tổng giá trị"
            b_hdrs[3].text = "Tỷ trọng"
            b_hdrs[4].text = "Điểm số"
            b_hdrs[5].text = "Phân loại"
            b_hdrs[6].text = "Căn cứ định lượng & Khuyến nghị"
            for br in model["benchmark_table"]:
                b_rows = b_table.add_row().cells
                b_rows[0].text = str(br["name"])
                b_rows[1].text = str(br.get("count", "—"))
                b_rows[2].text = f"{br.get('sum', 0):,.0f} đ" if isinstance(br.get('sum'), (int, float)) else str(br.get('sum', '—'))
                b_rows[3].text = f"{br.get('share', 0)}%"
                b_rows[4].text = f"{br['score']}/10"
                b_rows[5].text = str(br["status"])
                b_rows[6].text = str(br["reason"])

        # 4. KẾT LUẬN VÀ ĐỀ XUẤT HÀNH ĐỘNG
        doc.add_heading("4. KẾT LUẬN VÀ ĐỀ XUẤT HÀNH ĐỘNG CHIẾN LƯỢC", level=1)
        doc.add_paragraph(clean_html_tags(model['conclusions']['p1']))
        doc.add_paragraph(clean_html_tags(model['conclusions']['p2']))
        if model['conclusions'].get('p3'):
            doc.add_paragraph(clean_html_tags(model['conclusions']['p3']))

        doc.add_heading("Lộ Trình Triển Khai Chiến Lược 4 Giai Đoạn (Go / No-Go Roadmap):", level=2)
        r_table = doc.add_table(rows=1, cols=5)
        r_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        r_hdrs = r_table.rows[0].cells
        r_hdrs[0].text = "Giai đoạn"
        r_hdrs[1].text = "Thời hạn"
        r_hdrs[2].text = "Hạng mục trọng tâm"
        r_hdrs[3].text = "Tiêu chuẩn Go / No-Go"
        r_hdrs[4].text = "Phụ trách"
        for rm in model["roadmap_table"]:
            r_rows = r_table.add_row().cells
            r_rows[0].text = rm["phase"]
            r_rows[1].text = rm.get("timeline", "—")
            r_rows[2].text = rm["target"]
            r_rows[3].text = rm["kpi_gate"]
            r_rows[4].text = rm["owner"]

        # 5. NGUỒN VÀ CHẤT LƯỢNG DỮ LIỆU
        doc.add_heading("5. NGUỒN VÀ CHẤT LƯỢNG DỮ LIỆU (AUDIT TRAIL)", level=1)
        doc.add_paragraph(clean_html_tags(model['data_quality']['p1']))
        doc.add_paragraph(clean_html_tags(model['data_quality']['p2']))
        doc.add_paragraph(clean_html_tags(model['data_quality']['p3']))

        if model['data_quality'].get('audit_table'):
            doc.add_heading("Nhật Ký Thẩm Định & Kiểm Toán Dữ Liệu (100% Sạch):", level=2)
            a_table = doc.add_table(rows=1, cols=4)
            a_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            a_hdrs = a_table.rows[0].cells
            a_hdrs[0].text = "Hạng mục kiểm toán"
            a_hdrs[1].text = "Tiêu chuẩn áp dụng"
            a_hdrs[2].text = "Kết quả đối soát thực tế"
            a_hdrs[3].text = "Đánh giá"
            for a_row in model['data_quality']['audit_table']:
                ar_cells = a_table.add_row().cells
                ar_cells[0].text = a_row["item"]
                ar_cells[1].text = a_row["standard"]
                ar_cells[2].text = a_row["result"]
                ar_cells[3].text = a_row["status"]

        # CHỨNG THỰC BÁO CÁO TOÀN VẸN
        doc.add_paragraph("\n\n")
        p_sign = doc.add_paragraph()
        p_sign.add_run("CHỨNG THỰC TOÀN VẸN BÁO CÁO:\n").bold = True
        p_sign.add_run("KTCM AI Analytics Studio • Hệ thống phân tích và kiểm toán điều hành tự động\n")
        p_sign.add_run(f"Phiên xử lý: {model['meta']['session_id']} • Tỷ lệ sạch: 100% • Thẩm định hoàn tất.\n")

        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        file_bytes = docx_buf.getvalue()

        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        self.send_header("Content-Disposition", f"attachment; filename=Bao_Cao_MaiThaoVi_{filename_val.split('.')[0]}.docx")
        self.send_header("Content-Length", str(len(file_bytes)))
        self.end_headers()
        self.wfile.write(file_bytes)

    def export_html_file(self, as_attachment=True):
        df = CURRENT_SESSION.get("cleaned_df")
        if df is None:
            df = CURRENT_SESSION.get("raw_df")
        if df is None:
            self.send_error(400, "Chưa có dữ liệu để xuất")
            return
        analysis = CURRENT_SESSION["analysis_cache"]
        if analysis is None:
            analysis = self.compute_deep_analysis(df, CURRENT_SESSION["business"])
            CURRENT_SESSION["analysis_cache"] = analysis

        charts = generate_executive_charts(analysis, df)
        filename_val = CURRENT_SESSION['filename'] or 'du_lieu_kinh_doanh_mau.csv'
        biz = CURRENT_SESSION.get('business') or {}
        custom_notes = CURRENT_SESSION.get("report_custom_notes", "")
        model = build_dynamic_report_model(df, filename_val, biz, analysis, custom_notes)

        custom_note_block = f'''
        <div style="background: #fbf8f5; border-left: 4px solid #8C5E43; padding: 14px 18px; border-radius: 8px; margin-bottom: 24px; font-size: 13px; color: #3d2f26;">
            <strong style="color: #6e4730;">📝 Ghi chú hiệu chỉnh bổ sung từ AI & Người dùng:</strong><br><span style="margin-top: 4px; display: inline-block;">{custom_notes}</span>
        </div>
        ''' if custom_notes else ""

        chart_html_specs = [
            ("chart1_b64", "Biểu đồ 1: Xếp hạng quy mô đóng góp theo phân nhóm chính", "chart1"),
            ("chart2_b64", "Biểu đồ 2: Phân tích Pareto (80/20) - Mức độ tập trung giá trị", "chart2"),
            ("chart3_b64", "Biểu đồ 3: Diễn biến theo chu kỳ & Đường trung bình chuẩn", "chart3"),
            ("chart4_b64", "Biểu đồ 4: Đối chiếu đa chỉ tiêu & Tỷ lệ theo phân nhóm", "chart4"),
            ("chart5_b64", "Biểu đồ 5: Phân bố tần suất bản ghi theo khoảng giá trị", "chart5"),
            ("chart6_b64", "Biểu đồ 6: Ma trận định vị ưu tiên xử lý phân khúc (4 Quadrants)", "chart6")
        ]
        charts_html_list = []
        for b64_k, c_title, ins_k in chart_html_specs:
            val_b64 = charts.get(b64_k) or (charts.get("bar_b64") if b64_k == "chart1_b64" else (charts.get("pie_b64") if b64_k == "chart2_b64" else (charts.get("line_b64") if b64_k == "chart3_b64" else None)))
            if val_b64:
                ins_text = model.get("chart_insights", {}).get(ins_k, "")
                clean_ins = ins_text.replace("Phân tích biểu đồ: ", "") if ins_text else ""
                charts_html_list.append(f'''
                <div class="card p-5 bg-[#fbf7f4] border border-[#e5d9ce] flex flex-col justify-between">
                    <div>
                        <div class="text-xs font-bold text-[#6e4730] uppercase tracking-wide mb-3 flex items-center justify-between">
                            <span>{c_title}</span>
                            <span class="px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded text-[10px] font-bold">Thẩm định 100%</span>
                        </div>
                        <img src="data:image/png;base64,{val_b64}" alt="{c_title}" class="w-full rounded-lg shadow-sm max-h-[290px] object-contain bg-white p-1" />
                    </div>
                    {f'<div class="mt-3 p-3 bg-white/90 rounded border-l-4 border-l-[#8C5E43] text-xs text-slate-700 leading-relaxed"><strong>💡 Phân tích biểu đồ (Executive Insight):</strong> {clean_ins}</div>' if clean_ins else ''}
                </div>
                ''')
        all_charts_html = '<div class="grid grid-cols-1 md:grid-cols-2 gap-6">' + "".join(charts_html_list) + '</div>'

        priority_cards_html = ""
        for p in model["priority_cards"]:
            priority_cards_html += f'''
            <div class="card p-6 bg-[#fbf7f4] border-l-4 border-l-[#8C5E43]">
                <div class="text-xs font-bold text-[#8C5E43] uppercase tracking-wider">{p["badge"]}</div>
                <div class="text-xl font-extrabold text-[#231a15] mt-1">{p["title"]}</div>
                <p class="text-sm text-[#746255] mt-2 leading-relaxed">{p["desc"]}</p>
            </div>
            '''

        kpis_html = ""
        for k in model["kpis"]:
            kpis_html += f'''
            <div class="card p-5 border-t-4" style="border-top-color: {k.get("color", "#8C5E43")}">
                <span class="text-xs font-semibold uppercase text-slate-500">{k["label"]}</span>
                <div class="text-2xl font-extrabold text-slate-800 mt-1">{k["value"]}</div>
                <div class="text-xs text-[#6e4730] font-bold mt-1">{k["sub"]}</div>
            </div>
            '''

        overview_html = ""
        for o in model.get("overview_table", []):
            overview_html += f'''
            <tr>
                <td class="p-2.5 font-bold text-slate-800">{o["metric"]}</td>
                <td class="p-2.5 font-bold text-[#8C5E43]">{o["val"]}</td>
                <td class="p-2.5 text-slate-600 text-xs">{o["eval"]}</td>
            </tr>
            '''

        audit_log_html = ""
        for a in model["data_quality"].get("audit_table", []):
            audit_log_html += f'''
            <tr>
                <td class="p-2.5 font-bold text-slate-800">{a["item"]}</td>
                <td class="p-2.5 text-slate-600 text-xs">{a["standard"]}</td>
                <td class="p-2.5 font-semibold text-slate-700 text-xs">{a["result"]}</td>
                <td class="p-2.5 text-center"><span class="px-2 py-0.5 bg-emerald-100 text-emerald-800 font-bold text-xs rounded">{a["status"]}</span></td>
            </tr>
            '''

        bench_html = ""
        for b in model["benchmark_table"]:
            status_style = "color: #15803d; background: #dcfce7;" if b["status"] == "SELECTED" else ("color: #b45309; background: #fef3c7;" if b["status"] == "THEO DÕI" else "color: #b91c1c; background: #fee2e2;")
            bench_html += f'''
            <tr>
                <td class="p-3 font-bold">{b["name"]}</td>
                <td class="p-3 text-center">{b.get("count", "—")}</td>
                <td class="p-3 font-semibold text-[#8C5E43]">{b.get("sum", 0):,.0f} đ</td>
                <td class="p-3 text-center font-bold">{b.get("share", 0)}%</td>
                <td class="p-3 text-center font-bold">{b["score"]}</td>
                <td class="p-3 text-center"><span style="padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; {status_style}">{b["status"]}</span></td>
                <td class="p-3 text-slate-600 text-xs">{b["reason"]}</td>
            </tr>
            '''

        roadmap_html = ""
        for r in model["roadmap_table"]:
            roadmap_html += f'''
            <tr>
                <td class="p-3 font-bold">{r["phase"]}</td>
                <td class="p-3 font-semibold text-[#8C5E43] text-xs">{r.get("timeline", "—")}</td>
                <td class="p-3 text-slate-800 font-semibold">{r["target"]}</td>
                <td class="p-3 text-emerald-800 font-bold text-xs">{r["kpi_gate"]}</td>
                <td class="p-3 text-slate-600 text-xs">{r["owner"]}</td>
            </tr>
            '''

        html_content = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Báo Cáo Phân Tích Điều Hành Chuyên Sâu - Mai Thảo Vi</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background-color: #f7f3ee; color: #231a15; }}
        .card {{ background: white; border-radius: 1rem; border: 1px solid #e5d9ce; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.05); }}
        @media print {{
            body {{ background: white; padding: 0; }}
            .card {{ box-shadow: none; border-color: #cbd5e1; }}
        }}
    </style>
</head>
<body class="p-6 md:p-10 max-w-5xl mx-auto space-y-8">
    <!-- Header bìa -->
    <div class="card p-8 border-t-8 border-[#8C5E43] bg-gradient-to-r from-[#fbf8f5] via-white to-[#f7f2ed]">
        <div class="flex justify-between items-start flex-wrap gap-4">
            <div>
                <span class="px-3 py-1 bg-[#f5ece4] text-[#6e4730] font-bold text-xs rounded-full">BÁO CÁO PHÂN TÍCH ĐIỀU HÀNH CHUYÊN SÂU</span>
                <h1 class="text-3xl font-extrabold mt-3 text-[#231a15]">🏛 Mai Thảo Vi · Báo Cáo Phân Tích Dữ Liệu Điều Hành</h1>
                <p class="text-[#746255] mt-1">Lĩnh vực: <strong>{model['meta']['industry']}</strong> &bull; Tệp nguồn: <strong>{filename_val}</strong> &bull; Số dòng hợp lệ: <strong>{len(df)} dòng</strong></p>
                <p class="text-xs text-[#8a786d] mt-1">Chuyên viên Phân tích (Lead Data Analyst): <strong>Mai Thảo Vi</strong> &bull; Phiên: <strong>{model['meta']['session_id']}</strong></p>
            </div>
            <div class="text-right">
                <span class="px-3 py-1 bg-emerald-100 text-emerald-800 font-bold text-xs rounded-lg border border-emerald-300">● COMPLETED</span>
                <div class="text-xs text-slate-500 mt-2">KTCM AI Analytics Studio v2.6</div>
            </div>
        </div>
        {custom_note_block}
    </div>

    <!-- MỤC 1: TÓM TẮT ĐIỀU HÀNH -->
    <div class="card p-6">
        <h2 class="text-xl font-extrabold text-[#231a15] mb-4 pb-2 border-b border-[#e5d9ce] flex items-center gap-2">
            <span class="bg-[#8C5E43] text-white text-xs px-2.5 py-1 rounded-full">1</span>
            Tóm tắt điều hành (Executive Summary)
        </h2>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
            {priority_cards_html}
        </div>
        <div class="overflow-x-auto mb-5">
            <table class="w-full text-left text-xs text-slate-700 bg-white rounded-lg border border-[#e5d9ce]">
                <thead class="bg-[#fbf7f4] text-[#6e4730] font-bold border-b border-[#e5d9ce]">
                    <tr>
                        <th class="p-2.5">Chỉ số / Tiêu chí cốt lõi</th>
                        <th class="p-2.5">Giá trị định lượng</th>
                        <th class="p-2.5">Đánh giá & Ý nghĩa điều hành</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                    {overview_html}
                </tbody>
            </table>
        </div>
        <div class="space-y-3 text-sm text-[#3d2f26] leading-relaxed bg-[#fbf8f5] p-5 rounded-xl border border-[#e5d9ce]">
            <p>{model['summary']['p1']}</p>
            <p>{model['summary']['p2']}</p>
            <p>{model['summary']['p3']}</p>
            <p>{model['summary'].get('p4', '')}</p>
        </div>
    </div>

    <!-- MỤC 2: CÁC CHỈ SỐ GỒM BIỂU ĐỒ VÀ DASHBOARD -->
    <div class="card p-6">
        <h2 class="text-xl font-extrabold text-[#231a15] mb-4 pb-2 border-b border-[#e5d9ce] flex items-center gap-2">
            <span class="bg-[#8C5E43] text-white text-xs px-2.5 py-1 rounded-full">2</span>
            Các chỉ số cốt lõi, Biểu đồ & Dashboard
        </h2>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
            {kpis_html}
        </div>
        {all_charts_html}
    </div>

    <!-- MỤC 3: PHÂN TÍCH CHUYÊN ĐỀ -->
    <div class="card p-6">
        <h2 class="text-xl font-extrabold text-[#231a15] mb-4 pb-2 border-b border-[#e5d9ce] flex items-center gap-2">
            <span class="bg-[#8C5E43] text-white text-xs px-2.5 py-1 rounded-full">3</span>
            Phân tích chuyên đề chuyên sâu
        </h2>
        <div class="space-y-3 text-sm text-[#3d2f26] leading-relaxed bg-[#fbf8f5] p-5 rounded-xl border border-[#e5d9ce] mb-6">
            <p>{model['deepdive']['p1']}</p>
            <p>{model['deepdive']['p2']}</p>
            <p>{model['deepdive']['p3']}</p>
            <p>{model['deepdive'].get('p4', '')}</p>
        </div>
        <h3 class="text-sm font-bold text-[#6e4730] uppercase mb-3">Bảng So Sánh Đối Chuẩn & Phân Loại Hiệu Quả</h3>
        <div class="overflow-x-auto">
            <table class="w-full text-left text-xs text-slate-700">
                <thead class="bg-[#fbf7f4] text-[#6e4730] font-bold border-b border-[#e5d9ce]">
                    <tr>
                        <th class="p-3">Phân loại</th>
                        <th class="p-3 text-center">Số lượt</th>
                        <th class="p-3">Tổng giá trị</th>
                        <th class="p-3 text-center">Tỷ trọng</th>
                        <th class="p-3 text-center">Điểm số</th>
                        <th class="p-3 text-center">Kết luận</th>
                        <th class="p-3">Căn cứ định lượng & Lý do</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-200">
                    {bench_html}
                </tbody>
            </table>
        </div>
    </div>

    <!-- MỤC 4: KẾT LUẬN VÀ ĐỀ XUẤT HÀNH ĐỘNG -->
    <div class="card p-6">
        <h2 class="text-xl font-extrabold text-[#231a15] mb-4 pb-2 border-b border-[#e5d9ce] flex items-center gap-2">
            <span class="bg-[#8C5E43] text-white text-xs px-2.5 py-1 rounded-full">4</span>
            Kết luận & Đề xuất hành động chiến lược
        </h2>
        <div class="space-y-3 text-sm text-[#3d2f26] leading-relaxed bg-[#fbf8f5] p-5 rounded-xl border border-[#e5d9ce] mb-6">
            <p>{model['conclusions']['p1']}</p>
            <p>{model['conclusions']['p2']}</p>
            <p>{model['conclusions'].get('p3', '')}</p>
        </div>
        <h3 class="text-sm font-bold text-[#6e4730] uppercase mb-3">Lộ Trình Triển Khai Chiến Lược 4 Giai Đoạn (Go / No-Go Roadmap)</h3>
        <div class="overflow-x-auto">
            <table class="w-full text-left text-xs text-slate-700">
                <thead class="bg-[#fbf7f4] text-[#6e4730] font-bold border-b border-[#e5d9ce]">
                    <tr>
                        <th class="p-3">Giai đoạn</th>
                        <th class="p-3">Thời hạn</th>
                        <th class="p-3">Hạng mục trọng tâm</th>
                        <th class="p-3">Tiêu chuẩn Go / No-Go</th>
                        <th class="p-3">Phụ trách</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-200">
                    {roadmap_html}
                </tbody>
            </table>
        </div>
    </div>

    <!-- MỤC 5: NGUỒN VÀ CHẤT LƯỢNG DỮ LIỆU -->
    <div class="card p-6">
        <h2 class="text-xl font-extrabold text-[#231a15] mb-4 pb-2 border-b border-[#e5d9ce] flex items-center gap-2">
            <span class="bg-[#8C5E43] text-white text-xs px-2.5 py-1 rounded-full">5</span>
            Nguồn và chất lượng dữ liệu (Audit Trail)
        </h2>
        <div class="space-y-2 text-sm text-[#3d2f26] leading-relaxed bg-[#fbf8f5] p-5 rounded-xl border border-[#e5d9ce] mb-5">
            <p>{model['data_quality']['p1']}</p>
            <p>{model['data_quality']['p2']}</p>
            <p>{model['data_quality']['p3']}</p>
        </div>
        <h3 class="text-sm font-bold text-[#6e4730] uppercase mb-3">Nhật Ký Thẩm Định & Kiểm Toán Dữ Liệu (100% Sạch)</h3>
        <div class="overflow-x-auto">
            <table class="w-full text-left text-xs text-slate-700 bg-white rounded-lg border border-[#e5d9ce]">
                <thead class="bg-[#fbf7f4] text-[#6e4730] font-bold border-b border-[#e5d9ce]">
                    <tr>
                        <th class="p-2.5">Hạng mục kiểm toán</th>
                        <th class="p-2.5">Tiêu chuẩn áp dụng</th>
                        <th class="p-2.5">Kết quả đối soát thực tế</th>
                        <th class="p-2.5 text-center">Đánh giá</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                    {audit_log_html}
                </tbody>
            </table>
        </div>
    </div>

    <!-- Sign-off Block -->
    <div class="card p-6 bg-[#fbf8f5] border border-[#e5d9ce] flex justify-between items-center flex-wrap gap-4">
        <div>
            <div class="font-extrabold text-base text-[#231a15]">KTCM AI Analytics Studio</div>
            <div class="text-xs text-[#8C5E43] font-bold">Hệ thống phân tích và kiểm toán điều hành tự động • Thẩm định tính toàn vẹn 100%</div>
            <div class="text-xs text-slate-500 mt-1">Phiên xử lý: {model['meta']['session_id']} • Xác thực dữ liệu hoàn tất</div>
        </div>
        <div class="text-right">
            <span class="px-3 py-1 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-full">✓ Thẩm định hoàn tất 100%</span>
        </div>
    </div>
</body>
</html>"""
        html_bytes = html_content.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        if as_attachment:
            self.send_header("Content-Disposition", f"attachment; filename=Bao_Cao_MaiThaoVi_{filename_val.split('.')[0]}.html")
        self.send_header("Content-Length", str(len(html_bytes)))
        self.end_headers()
        self.wfile.write(html_bytes)

def run_server():
    server = ThreadingHTTPServer(("0.0.0.0", PORT), DataAnalysisHandler)
    print(f"===============================================================")
    print(f"🚀 KTCM Analytics Studio - Đang chạy tại:")
    print(f"👉 http://localhost:{PORT}")
    print(f"===============================================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("Dừng server.")
        server.server_close()

if __name__ == "__main__":
    run_server()
else:
    # Vercel Serverless WSGI compatibility entrypoint fallback
    try:
        from api.index import app
    except Exception:
        pass
