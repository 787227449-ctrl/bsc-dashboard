#!/usr/bin/env python3
"""
用法: python3 update_data.py <excel_path> [html_path]
默认 html_path = ./june_exam.html

从 Excel 数据更新 Dashboard HTML 中的 JS 变量:
  - const D={...}
  - const TREND = {...}
  - const HISTORY = [...]
  - const TOP10_DATA = {...}
  - BSC_DATA 保持不变
"""

import sys
import os
import re
import json
import math
import pandas as pd
import numpy as np

# ============================================================
# 0. 参数
# ============================================================
if len(sys.argv) < 2:
    print("用法: python3 update_data.py <excel_path> [html_path]")
    sys.exit(1)

EXCEL_PATH = sys.argv[1]
HTML_PATH = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(__file__) or '.', 'june_exam.html')

META = {
    "date": "2026-06-17",
    "exam_days": 9,
    "exam_total_days": 24
}

# ============================================================
# Helpers
# ============================================================
def safe_float(v, default=0.0):
    """Safely convert to float, handling #REF!, #N/A, None, strings."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return default
    if isinstance(v, str):
        v = v.strip()
        if v in ('', '-', '#REF!', '#N/A', '#VALUE!', '#DIV/0!', '#NAME?'):
            return default
        try:
            return float(v)
        except ValueError:
            return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default

def safe_int(v, default=0):
    return int(safe_float(v, default))

def safe_bool(v, default=0):
    """Convert 0/1/True/False/string to 0 or 1."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return default
    if isinstance(v, bool):
        return 1 if v else 0
    if isinstance(v, str):
        v = v.strip()
        if v in ('', '-', '#REF!', '#N/A'):
            return default
        try:
            return 1 if float(v) >= 1 else 0
        except ValueError:
            return default
    try:
        return 1 if float(v) >= 1 else 0
    except (ValueError, TypeError):
        return default

def safe_str(v, default=''):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return default
    return str(v).strip()

def r4(v):
    """Round to 4 decimals."""
    return round(v, 4)

def r2(v):
    return round(v, 2)

# ============================================================
# 1. Read Excel
# ============================================================
print("Reading Excel...")
xls = pd.ExcelFile(EXCEL_PATH)

df_roster = pd.read_excel(xls, '花名册')
df_cat = pd.read_excel(xls, '品类诊断明细')
df_goods = pd.read_excel(xls, '商品诊断明细')
df_yesterday_goods = pd.read_excel(xls, '昨日商品诊断')
df_score_yesterday = pd.read_excel(xls, '昨日得分', header=None)
df_score_lastweek = pd.read_excel(xls, '上周得分', header=None)

# ============================================================
# 2. Build roster mapping (考核==1)
# ============================================================
print("Building roster...")
roster = {}
for _, row in df_roster.iterrows():
    if safe_int(row.get('考核')) != 1:
        continue
    hid = safe_str(row.get('蜂窝id'))
    if not hid:
        continue
    roster[hid] = {
        'name': safe_str(row.get('蜂窝名称')),
        'bd': safe_str(row.get('蜂窝挂靠人')),
        'group': safe_str(row.get('联络点')),
    }
print(f"  考核蜂窝数: {len(roster)}")

# ============================================================
# 3. Build category data from 品类诊断明细
# ============================================================
print("Building category data...")
cat_data = {}  # hive_id -> {cat_name -> {...}}
for _, row in df_cat.iterrows():
    hid = safe_str(row.get('蜂窝ID'))
    if hid not in roster:
        continue
    cat_name = safe_str(row.get('宽前端三级品类'))
    if not cat_name:
        continue

    # Determine if 必上菜
    is_must = False
    if '是否必上菜' in df_cat.columns:
        val = safe_str(row.get('是否必上菜'))
        if val and val not in ('0', 'nan', ''):
            is_must = True
    if not is_must and '是否核心品类' in df_cat.columns:
        if safe_int(row.get('是否核心品类')) == 1:
            is_must = True
    if not is_must and '做工类型' in df_cat.columns:
        val = safe_str(row.get('做工类型'))
        if '必上菜' in val:
            is_must = True
    if not is_must and '品类识别' in df_cat.columns:
        # check if it's in the expected category
        val = safe_str(row.get('品类识别'))
        if val and val not in ('0', 'nan', ''):
            is_must = True

    if not is_must:
        continue

    if hid not in cat_data:
        cat_data[hid] = {}

    cat_data[hid][cat_name] = {
        'threshold': safe_float(row.get('标杆阈值')),
        'is_quality': safe_bool(row.get('质量是否达标')),
        'has_bm': safe_bool(row.get('是否有标杆')),
        'base_has_bm': safe_bool(row.get('是否有标杆-基期')),
        'has_houdu': safe_bool(row.get('厚度是否达标')),
        'real_bm': safe_bool(row.get('实时标杆')) if '实时标杆' in df_cat.columns else 0,
    }

# ============================================================
# 4. Build goods detail from 商品诊断明细
# ============================================================
print("Building goods detail...")

def build_goods_from_df(df, roster_hids=None):
    """Parse goods dataframe and return {hive_id -> {cat -> [items]}}."""
    goods = {}
    for _, row in df.iterrows():
        hid = safe_str(row.get('蜂窝ID'))
        if roster_hids and hid not in roster_hids:
            continue
        cat = safe_str(row.get('宽前端三级品类'))
        if not cat:
            continue

        daily_nat = safe_float(row.get('日均单产（自然日）'))
        daily_plain = safe_float(row.get('日均单产'))
        daily = daily_nat if daily_nat > 0 else daily_plain
        threshold = safe_float(row.get('标杆单产阈值'))
        diff = safe_float(row.get('距离标杆阈值差距'))
        total_orders = safe_float(row.get('拼好饭订单'))
        is_bm = safe_bool(row.get('[标杆]达标商品'))
        is_quality = safe_bool(row.get('[质量]达标商品'))

        # Status
        if is_bm:
            status = "bm"
        elif diff >= -2 and diff < 0:
            status = "near"
        elif total_orders == 0 or daily == 0:
            status = "zero"
        else:
            status = "active"

        # need_daily
        if diff < 0:
            need_daily = r2(daily + abs(diff) * 1.05)
        else:
            need_daily = r2(daily)

        # pinxiao
        pinxiao = r2(daily / threshold) if threshold > 0 else 0.0

        score = safe_float(row.get('商品分'))
        settle_price = safe_float(row.get('查询截止日SPU最低sku录入结算价'))
        yesterday_orders = safe_int(row.get('（辅助列）昨日订单'))

        item = {
            'merchant': safe_str(row.get('商家名称')),
            'merchant_id': safe_str(row.get('商家ID')),
            'spu': safe_str(row.get('SPU名称', '-')),
            'spu_id': safe_str(row.get('SPU-ID')),
            'daily': r2(daily),
            'threshold': r2(threshold),
            'diff': r2(diff),
            'total_orders': r2(total_orders),
            'settle_price': r2(settle_price),
            'need_daily': need_daily,
            'status': status,
            'is_bm': is_bm,
            'is_quality': is_quality,
            'has_ad': safe_bool(row.get('是否买广告')),
            'has_pindan': safe_bool(row.get('是否开通拼单宝')),
            'score': r2(score),
            'liangchu': safe_bool(row.get('商家亮厨是否达标')),
            'is_baoping': safe_bool(row.get('周期内是否爆品')),
            'has_houdu': safe_bool(row.get('厚度是否达标')) if '厚度是否达标' in df.columns else 0,
            'yesterday_orders': yesterday_orders,
            'pinxiao': pinxiao,
            'is_new': safe_bool(row.get('是否周期内新增')),
        }

        # Change tracking vs yesterday
        if '昨日是否标杆' in df.columns:
            was_bm = safe_bool(row.get('昨日是否标杆'))
            item['bp_chg'] = is_bm - was_bm
        else:
            item['bp_chg'] = 0

        # Price change detection
        item['price_chg'] = 0
        item['ad_chg'] = 0

        # Check for pickup
        if '是否开通到店自取' in df.columns:
            item['has_pickup'] = safe_bool(row.get('是否开通到店自取'))

        if hid not in goods:
            goods[hid] = {}
        if cat not in goods[hid]:
            goods[hid][cat] = []
        goods[hid][cat].append(item)

    # Sort each category by daily desc
    for hid in goods:
        for cat in goods[hid]:
            goods[hid][cat].sort(key=lambda x: -x['daily'])

    return goods

goods_today = build_goods_from_df(df_goods, set(roster.keys()))
goods_yesterday = build_goods_from_df(df_yesterday_goods, set(roster.keys()))

# ============================================================
# 5. Build D object
# ============================================================
print("Building D object...")

def build_hive_data(hive_id, roster_entry, cat_info, goods_map):
    """Build the data structure for one hive."""
    cats_available = cat_info.get(hive_id, {})
    goods_for_hive = goods_map.get(hive_id, {})

    # Get the list of must-have categories
    cat_names = sorted(cats_available.keys())
    cat_status = {}

    for cat_name in cat_names:
        cat_meta = cats_available[cat_name]
        items = goods_for_hive.get(cat_name, [])

        # Count BM SPUs
        current_bm_count = sum(1 for it in items if it['is_bm'])

        # Determine is_dabiao
        base_has_bm = cat_meta.get('base_has_bm', 0)
        if base_has_bm:
            is_dabiao = current_bm_count >= 2
        else:
            is_dabiao = current_bm_count >= 1

        # is_quality from category level
        is_quality = bool(cat_meta.get('is_quality', 0))

        # dabiao_detail text
        if not items:
            dabiao_detail = "无数据"
        elif base_has_bm:
            dabiao_detail = f"基期有标杆,达标SPU数={current_bm_count}"
        else:
            if current_bm_count >= 1:
                dabiao_detail = "基期无标杆,当前有标杆"
            else:
                dabiao_detail = "基期无标杆,当前无标杆"

        # Head SPU calculations: use BM SPUs if any, else top 1
        bm_items = [it for it in items if it['is_bm']]
        if bm_items:
            head_items = bm_items
        elif items:
            head_items = [items[0]]  # top by daily
        else:
            head_items = []

        head_daily_sum = r2(sum(it['daily'] for it in head_items))
        head_thresh_sum = r2(sum(it['threshold'] for it in head_items))
        pinxiao_multiple = r4(head_daily_sum / head_thresh_sum) if head_thresh_sum > 0 else 0

        # Build goods_detail
        goods_detail = []
        for it in items:
            goods_detail.append(it)

        cat_status[cat_name] = {
            'is_dabiao': is_dabiao,
            'is_quality': is_quality,
            'base_has_bm': base_has_bm,
            'current_bm_count': current_bm_count,
            'dabiao_detail': dabiao_detail,
            'head_daily_sum': head_daily_sum,
            'head_thresh_sum': head_thresh_sum,
            'pinxiao_multiple': pinxiao_multiple,
            'goods_detail': goods_detail,
        }

    return {
        'id': hive_id,
        'cats': cat_names,
        'cat_count': len(cat_names),
        'bd': roster_entry['bd'],
        'group': roster_entry['group'],
        'cat_status': cat_status,
    }


D = {}
for hive_id, info in roster.items():
    hive_name = info['name']
    hive_data = build_hive_data(hive_id, info, cat_data, goods_today)
    D[hive_name] = hive_data

# Add meta - D must have structure {"hives": {...}, "meta": {...}}
D_with_meta = {"hives": D, "meta": META}

print(f"  D蜂窝数: {len(D)}")

# ============================================================
# 6. Calculate scores
# ============================================================
print("Calculating scores...")

def calc_score(hive_data):
    """Calculate finalScore for a hive."""
    cats = hive_data.get('cat_status', {})
    cat_count = len(cats)
    if cat_count == 0:
        return 0.0, 0.0, 1.0, 0

    dabiao_count = sum(1 for c in cats.values() if c.get('is_dabiao'))
    quality_count = sum(1 for c in cats.values() if c.get('is_quality'))

    if dabiao_count == 0:
        score = r4(quality_count / max(cat_count, 1) * 0.2)
        return score, 0.0, 1.0, dabiao_count

    dabiao_rate = min(dabiao_count / 3, 4/3)

    dabiao_cats = [c for c in cats.values() if c.get('is_dabiao')]
    pinxiao_list = [(c['head_daily_sum'], c['head_thresh_sum']) for c in dabiao_cats]

    if dabiao_count > 3:
        pinxiao_list.sort(key=lambda x: x[0] / max(x[1], 0.001), reverse=True)
        pinxiao_list = pinxiao_list[:4]

    total_daily = sum(p[0] for p in pinxiao_list)
    total_thresh = sum(p[1] for p in pinxiao_list)

    if total_thresh > 0:
        pinxiao_multiple = total_daily / total_thresh
    else:
        pinxiao_multiple = 1.0

    if pinxiao_multiple >= 2.0:
        coeff = 1.8
    elif pinxiao_multiple >= 1.5:
        coeff = 1.5
    elif pinxiao_multiple >= 1.2:
        coeff = 1.3
    elif pinxiao_multiple >= 1.0:
        coeff = 1.0
    else:
        coeff = 1.0

    final_score = r4(dabiao_rate * coeff)
    return final_score, r4(dabiao_rate), coeff, dabiao_count


# Today's scores
today_scores = {}
for hive_name, hive_data in D.items():
    fs, dr, coeff, dab = calc_score(hive_data)
    today_scores[hive_data['id']] = {
        'name': hive_name,
        'bd': hive_data['bd'],
        'group': hive_data['group'],
        'finalScore': fs,
        'dabiaoRate': dr,
        'pinxiaoCoeff': coeff,
        'dabiao': dab,
        'catCount': hive_data['cat_count'],
    }

# Yesterday's scores (build from yesterday goods)
yesterday_scores = {}
for hive_id, info in roster.items():
    hive_data_y = build_hive_data(hive_id, info, cat_data, goods_yesterday)
    fs, dr, coeff, dab = calc_score(hive_data_y)
    yesterday_scores[hive_id] = {
        'name': info['name'],
        'bd': info['bd'],
        'group': info['group'],
        'finalScore': fs,
        'dabiaoRate': dr,
        'pinxiaoCoeff': coeff,
        'dabiao': dab,
        'catCount': hive_data_y['cat_count'],
    }

# ============================================================
# 7. Calculate TREND (DoD and MoM)
# ============================================================
print("Calculating TREND...")

def calc_dod(today_val, yesterday_val):
    if yesterday_val == 0 and today_val == 0:
        return 0.0
    if yesterday_val == 0:
        return r2(today_val * 100) if today_val > 0 else 0.0
    return r2((today_val - yesterday_val) / abs(yesterday_val) * 100)

# Parse 昨日得分 for MoM reference (标杆达标率)
# Row 39+ has hive-level data: BD, 联络点, 蜂窝, ..., col9=标杆达标率
# We'll also get group-level from rows 5-8 and overall from row 4
def parse_score_sheet(df):
    """Parse 昨日得分/上周得分 sheet for 标杆达标率 at hive, group, overall levels."""
    result = {'overall': 0.0, 'groups': {}, 'hives': {}}

    # Overall: row 4, col 9 (标杆达标率)
    if len(df) > 4:
        result['overall'] = safe_float(df.iloc[4, 9])

    # Groups: rows 5-8
    for i in range(5, min(9, len(df))):
        group_name = safe_str(df.iloc[i, 0])
        if group_name:
            result['groups'][group_name] = safe_float(df.iloc[i, 9])

    # Hives: rows 41+
    for i in range(41, len(df)):
        hive_name = safe_str(df.iloc[i, 2])
        if hive_name:
            result['hives'][hive_name] = safe_float(df.iloc[i, 9])

    return result

yesterday_sheet = parse_score_sheet(df_score_yesterday)
lastweek_sheet = parse_score_sheet(df_score_lastweek)

# Build TREND
trend_hives = {}
for hid, ts in today_scores.items():
    ys = yesterday_scores.get(hid, {})
    dod = calc_dod(ts['finalScore'], ys.get('finalScore', 0))

    # MoM: compare today's dabiao rate with last week's 标杆达标率
    hive_name = ts['name']
    lw_val = lastweek_sheet['hives'].get(hive_name, 0)
    today_dabiao_rate = ts['dabiaoRate']
    mom = calc_dod(today_dabiao_rate, lw_val) if lw_val else (r2(today_dabiao_rate * 100) if today_dabiao_rate > 0 else 0.0)

    trend_hives[hid] = {'dod': dod, 'mom': mom}

# Group-level TREND
groups = set(ts['group'] for ts in today_scores.values())
trend_groups = {}
for g in groups:
    g_hids = [h for h, ts in today_scores.items() if ts['group'] == g]
    if not g_hids:
        continue
    avg_today = sum(today_scores[h]['finalScore'] for h in g_hids) / len(g_hids)
    avg_yesterday = sum(yesterday_scores.get(h, {}).get('finalScore', 0) for h in g_hids) / len(g_hids)
    dod = calc_dod(avg_today, avg_yesterday)

    # MoM from group scores
    lw_val = lastweek_sheet['groups'].get(g, 0)
    avg_dabiao = sum(today_scores[h]['dabiaoRate'] for h in g_hids) / len(g_hids)
    mom = calc_dod(avg_dabiao, lw_val) if lw_val else 0.0

    trend_groups[g] = {'dod': dod, 'mom': mom}

# Overall TREND
all_hids = list(today_scores.keys())
overall_today = sum(today_scores[h]['finalScore'] for h in all_hids) / max(len(all_hids), 1)
overall_yesterday = sum(yesterday_scores.get(h, {}).get('finalScore', 0) for h in all_hids) / max(len(all_hids), 1)
overall_dod = calc_dod(overall_today, overall_yesterday)
overall_dabiao = sum(today_scores[h]['dabiaoRate'] for h in all_hids) / max(len(all_hids), 1)
lw_overall = lastweek_sheet.get('overall', 0)
overall_mom = calc_dod(overall_dabiao, lw_overall) if lw_overall else 0.0

TREND = {
    'overall': {'dod': overall_dod, 'mom': overall_mom},
    'groups': trend_groups,
    'hives': trend_hives,
}

# ============================================================
# 8. Build TOP10_DATA
# ============================================================
print("Building TOP10_DATA...")

# Collect all goods across all hives, grouped by 宽前端三级品类
all_cat_goods = {}
for _, row in df_goods.iterrows():
    cat = safe_str(row.get('宽前端三级品类'))
    if not cat:
        continue

    daily_nat = safe_float(row.get('日均单产（自然日）'))
    daily_plain = safe_float(row.get('日均单产'))
    daily = daily_nat if daily_nat > 0 else daily_plain
    if daily <= 0:
        continue

    hid = safe_str(row.get('蜂窝ID'))
    hive_name = safe_str(row.get('蜂窝名称'))
    bd = safe_str(row.get('BD姓名'))
    group = safe_str(row.get('联络点'))
    std = safe_str(row.get('标准菜名称', ''))

    item = {
        'std': std,
        'shop': safe_str(row.get('商家名称')),
        'spu': safe_str(row.get('SPU名称')),
        'bd': bd,
        'hive': hive_name,
        'group': group,
        'output': r2(daily),
        'price': r2(safe_float(row.get('查询截止日SPU最低sku录入结算价'))),
        'ad': safe_bool(row.get('是否买广告')),
        'pdb': safe_bool(row.get('是否开通拼单宝')),
        'pickup': safe_bool(row.get('是否开通到店自取')),
        'hot': safe_bool(row.get('周期内是否爆品')),
        'score': r2(safe_float(row.get('商品分'))),
        'benchmark': safe_bool(row.get('[标杆]达标商品')),
    }

    if cat not in all_cat_goods:
        all_cat_goods[cat] = []
    all_cat_goods[cat].append(item)

TOP10_DATA = {}
for cat, items in all_cat_goods.items():
    items.sort(key=lambda x: -x['output'])
    top10 = items[:10]
    std_names = list(dict.fromkeys(it['std'] for it in items if it['std']))  # dedup preserving order
    TOP10_DATA[cat] = {
        'std_names': std_names,
        'items': top10,
    }

print(f"  TOP10品类数: {len(TOP10_DATA)}")

# ============================================================
# 9. Build HISTORY entry
# ============================================================
print("Building HISTORY entry...")

history_entry = {
    'date': META['date'],
    'exam_days': META['exam_days'],
    'hives': today_scores,
    'overall': {'dod': overall_dod, 'mom': overall_mom},
}

# ============================================================
# 10. Replace in HTML
# ============================================================
print("Replacing in HTML...")

with open(HTML_PATH, 'r', encoding='utf-8') as f:
    html = f.read()

def js_dumps(obj):
    """Convert Python object to JS-compatible string (using json.dumps)."""
    return json.dumps(obj, ensure_ascii=False, separators=(',', ':'))

# 10.1 Replace TREND
# Pattern: const TREND = { ... };
trend_js = f"const TREND = {json.dumps(TREND, ensure_ascii=False, indent=2)};"
# Use regex to find and replace
pattern = re.compile(r'const TREND\s*=\s*\{.*?\};', re.DOTALL)
if pattern.search(html):
    html = pattern.sub(trend_js, html, count=1)
    print("  ✓ TREND replaced")
else:
    print("  ✗ TREND pattern not found!")

# 10.2 Append to HISTORY
# Find const HISTORY = [...];
pattern_hist = re.compile(r'(const HISTORY\s*=\s*\[)(.*?)(\];)', re.DOTALL)
match_hist = pattern_hist.search(html)
if match_hist:
    existing = match_hist.group(2).rstrip().rstrip(',')
    new_entry = json.dumps(history_entry, ensure_ascii=False)
    new_history = f"{match_hist.group(1)}{existing},{new_entry}];"
    html = html[:match_hist.start()] + new_history + html[match_hist.end():]
    print("  ✓ HISTORY appended")
else:
    print("  ✗ HISTORY pattern not found!")

# 10.3 Replace TOP10_DATA
pattern_top10 = re.compile(r'const TOP10_DATA\s*=\s*\{.*?\};\s*(?=const |function |//|<)', re.DOTALL)
top10_js = f"const TOP10_DATA = {json.dumps(TOP10_DATA, ensure_ascii=False)};\n"
if pattern_top10.search(html):
    html = pattern_top10.sub(top10_js, html, count=1)
    print("  ✓ TOP10_DATA replaced")
else:
    # Try alternative: find from 'const TOP10_DATA' to next 'const '
    print("  Trying alternative TOP10_DATA replacement...")
    idx_start = html.find('const TOP10_DATA')
    if idx_start >= 0:
        # Find the end: look for next 'const ' after this
        rest = html[idx_start + 16:]  # skip 'const TOP10_DATA'
        # Find the semicolon that ends this statement
        # The data is a big object literal, we need to find the matching end
        idx_next_const = rest.find('\nconst ')
        if idx_next_const < 0:
            idx_next_const = rest.find('\n\nconst ')
        if idx_next_const >= 0:
            html = html[:idx_start] + top10_js + html[idx_start + 16 + idx_next_const:]
            print("  ✓ TOP10_DATA replaced (alt method)")
        else:
            print("  ✗ Could not find TOP10_DATA end!")
    else:
        print("  ✗ TOP10_DATA not found!")

# 10.4 Replace D
# Use brace-counting to find exact end of D object
d_js = f"const D={json.dumps(D_with_meta, ensure_ascii=False)};\n"
idx_d_start = html.find('const D={')
if idx_d_start < 0:
    idx_d_start = html.find('const D =')
if idx_d_start >= 0:
    brace_start = html.find('{', idx_d_start)
    depth = 0
    in_string = False
    escape_next = False
    d_end = -1
    for i in range(brace_start, len(html)):
        ch = html[i]
        if escape_next:
            escape_next = False
            continue
        if ch == '\\':
            if in_string:
                escape_next = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                d_end = i + 1  # include the '}'
                break
    if d_end > 0:
        # Include the ';' after '}'
        if html[d_end] == ';':
            d_end += 1
        html = html[:idx_d_start] + d_js + html[d_end:]
        print("  \u2713 D replaced")
    else:
        print("  \u2717 Could not find D closing brace!")
else:
    print("  \u2717 D variable not found!")

# 11. Write output
with open(HTML_PATH, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"\n✅ Done! Updated {HTML_PATH}")
print(f"   File size: {len(html.encode('utf-8')):,} bytes")
print(f"   蜂窝数: {len(D_with_meta.get('hives', D_with_meta))}")
print(f"   TOP10品类: {len(TOP10_DATA)}")
print(f"   日期: {META['date']}, 考核天数: {META['exam_days']}")
