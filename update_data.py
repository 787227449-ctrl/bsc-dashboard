#!/usr/bin/env python3
"""
用法: python3 update_data.py <当前excel> <基期excel> [html_path]
默认 html_path = ./june_exam.html

从 Excel 数据更新 Dashboard HTML 中的 JS 变量:
  - const D={...}   — 保留原始 HTML 中的蜂窝列表、品类列表、SPU ID 列表，只更新数值
  - const TREND = {...}
  - const HISTORY = [...]
  - const TOP10_DATA = {...}
  - BSC_DATA 保持不变

日均单产计算方式：(当前订单 - 基期订单) / 考核天数
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
if len(sys.argv) < 3:
    print("用法: python3 update_data.py <当前excel> <基期excel> [html_path]")
    sys.exit(1)

EXCEL_PATH = sys.argv[1]       # 当前数据 Excel（如6月17日）
BASE_EXCEL_PATH = sys.argv[2]  # 基期数据 Excel（如6月8日）
HTML_PATH = sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.path.dirname(__file__) or '.', 'june_exam.html')

META = {
    "date": "2026-06-17",
    "exam_days": 8,
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
# 0.5 Parse original D from HTML (preserve structure)
# ============================================================
print("Parsing original D from HTML...")

def extract_d_from_html(html_path):
    """Extract the D object from the HTML file using brace counting."""
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    idx = html.find('const D={')
    if idx < 0:
        idx = html.find('const D =')
    if idx < 0:
        raise ValueError("Cannot find 'const D=' in HTML file")

    brace_start = html.find('{', idx)
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
                d_end = i + 1
                break

    if d_end < 0:
        raise ValueError("Cannot find closing brace for D object")

    d_str = html[brace_start:d_end]
    return json.loads(d_str)

orig_D = extract_d_from_html(HTML_PATH)
print(f"  原始蜂窝数: {len(orig_D.get('hives', {}))}")
for hname, hdata in orig_D.get('hives', {}).items():
    print(f"    {hname}: {hdata['cat_count']} cats")

# Build lookup: hive_id -> hive_name (from original D)
orig_hive_id_to_name = {}
for hname, hdata in orig_D['hives'].items():
    orig_hive_id_to_name[str(hdata['id'])] = hname

# ============================================================
# 1. Read Excel (当前 + 基期)
# ============================================================
print("\nReading current Excel...")
xls = pd.ExcelFile(EXCEL_PATH)

df_roster = pd.read_excel(xls, '花名册')
df_cat = pd.read_excel(xls, '品类诊断明细')
df_goods = pd.read_excel(xls, '商品诊断明细')
df_yesterday_goods = pd.read_excel(xls, '昨日商品诊断')
df_score_yesterday = pd.read_excel(xls, '昨日得分', header=None)
df_score_lastweek = pd.read_excel(xls, '上周得分', header=None)

print("Reading base period Excel...")
xls_base = pd.ExcelFile(BASE_EXCEL_PATH)
df_goods_base = pd.read_excel(xls_base, '商品诊断明细')

# Build base period SPU-ID -> 拼好饭订单 mapping
base_orders_map = {}  # {spu_id_str: orders}
for _, row in df_goods_base.iterrows():
    spu_id = safe_str(row.get('SPU-ID'))
    if spu_id:
        orders = safe_float(row.get('拼好饭订单'))
        base_orders_map[spu_id] = orders
print(f"  基期 SPU 数: {len(base_orders_map)}")

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
# 3. Build category-level data from 品类诊断明细 (indexed by hive_id + cat)
# ============================================================
print("Building category data from Excel...")
cat_data_excel = {}  # hive_id -> {cat_name -> {...}}
for _, row in df_cat.iterrows():
    hid = safe_str(row.get('蜂窝ID'))
    if not hid:
        continue
    cat_name = safe_str(row.get('宽前端三级品类'))
    if not cat_name:
        continue

    if hid not in cat_data_excel:
        cat_data_excel[hid] = {}

    cat_data_excel[hid][cat_name] = {
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

def build_goods_from_df(df, roster_hids=None, use_base_deduction=False):
    """Parse goods dataframe and return {hive_id -> {cat -> [items]}}.
    
    If use_base_deduction=True, daily is calculated as:
      (current_orders - base_orders) / META['exam_days']
    instead of using the Excel '日均单产（自然日）' field.
    """
    goods = {}
    for _, row in df.iterrows():
        hid = safe_str(row.get('蜂窝ID'))
        if roster_hids and hid not in roster_hids:
            continue
        cat = safe_str(row.get('宽前端三级品类'))
        if not cat:
            continue

        total_orders = safe_float(row.get('拼好饭订单'))
        threshold = safe_float(row.get('标杆单产阈值'))

        if use_base_deduction:
            # 考核期日均单产 = (当前订单 - 基期订单) / 考核天数
            spu_id_str = safe_str(row.get('SPU-ID'))
            base_orders = base_orders_map.get(spu_id_str, 0.0)
            exam_orders = total_orders - base_orders
            daily = exam_orders / META['exam_days'] if META['exam_days'] > 0 else 0.0
            daily = max(daily, 0.0)  # 防止负数
        else:
            daily_nat = safe_float(row.get('日均单产（自然日）'))
            daily_plain = safe_float(row.get('日均单产'))
            daily = daily_nat if daily_nat > 0 else daily_plain

        # Recalculate diff and is_bm based on (possibly new) daily
        diff = r2(daily - threshold) if threshold > 0 else 0.0
        is_bm = 1 if (daily >= threshold and threshold > 0) else 0
        is_quality = safe_bool(row.get('[质量]达标商品'))

        # Status
        if is_bm:
            status = "bm"
        elif diff >= -2 and diff < 0:
            status = "near"
        elif daily == 0:
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

# Collect all hive IDs that appear in original D
orig_hive_ids = set(str(h['id']) for h in orig_D['hives'].values())
goods_today = build_goods_from_df(df_goods, orig_hive_ids, use_base_deduction=True)
goods_yesterday = build_goods_from_df(df_yesterday_goods, orig_hive_ids, use_base_deduction=False)

# ============================================================
# 5. Build D object — preserving original structure
# ============================================================
print("Building D object (preserving original cats/spu_ids)...")

def update_hive_data(orig_hive, cat_excel_for_hive, goods_for_hive):
    """
    Update a hive's cat_status data using Excel data,
    while preserving the original cats list, cat_count, and SPU IDs.
    """
    hive_id = str(orig_hive['id'])
    orig_cats = orig_hive['cats']
    orig_cat_count = orig_hive['cat_count']
    orig_cat_status = orig_hive.get('cat_status', {})

    new_cat_status = {}

    for cat_name in orig_cats:
        orig_cat_data = orig_cat_status.get(cat_name, {})
        excel_cat = cat_excel_for_hive.get(cat_name, {})
        excel_goods = goods_for_hive.get(cat_name, [])

        orig_goods = orig_cat_data.get('goods_detail', [])
        orig_spu_ids = [g['spu_id'] for g in orig_goods]
        orig_goods_by_spu = {g['spu_id']: g for g in orig_goods}
        excel_goods_by_spu = {g['spu_id']: g for g in excel_goods}

        merged_goods = []
        for spu_id in orig_spu_ids:
            if spu_id in excel_goods_by_spu:
                merged_goods.append(excel_goods_by_spu[spu_id])
            else:
                merged_goods.append(orig_goods_by_spu[spu_id])

        for spu_id, item in excel_goods_by_spu.items():
            if spu_id not in orig_goods_by_spu:
                merged_goods.append(item)

        merged_goods.sort(key=lambda x: -x.get('daily', 0))

        items = merged_goods
        current_bm_count = sum(1 for it in items if it.get('is_bm'))

        base_has_bm = excel_cat.get('base_has_bm', orig_cat_data.get('base_has_bm', 0))
        is_quality = bool(excel_cat.get('is_quality', orig_cat_data.get('is_quality', 0)))

        if base_has_bm:
            is_dabiao = current_bm_count >= 2
        else:
            is_dabiao = current_bm_count >= 1

        if not items:
            dabiao_detail = "无数据"
        elif base_has_bm:
            dabiao_detail = f"基期有标杆,达标SPU数={current_bm_count}"
        else:
            if current_bm_count >= 1:
                dabiao_detail = "基期无标杆,当前有标杆"
            else:
                dabiao_detail = "基期无标杆,当前无标杆"

        bm_items = [it for it in items if it.get('is_bm')]
        if bm_items:
            head_items = bm_items
        elif items:
            head_items = [items[0]]
        else:
            head_items = []

        head_daily_sum = r2(sum(it.get('daily', 0) for it in head_items))
        head_thresh_sum = r2(sum(it.get('threshold', 0) for it in head_items))
        pinxiao_multiple = r4(head_daily_sum / head_thresh_sum) if head_thresh_sum > 0 else 0

        new_cat_status[cat_name] = {
            'is_dabiao': is_dabiao,
            'is_quality': is_quality,
            'base_has_bm': base_has_bm,
            'current_bm_count': current_bm_count,
            'dabiao_detail': dabiao_detail,
            'head_daily_sum': head_daily_sum,
            'head_thresh_sum': head_thresh_sum,
            'pinxiao_multiple': pinxiao_multiple,
            'goods_detail': merged_goods,
        }

    return {
        'id': orig_hive['id'],
        'cats': orig_cats,
        'cat_count': orig_cat_count,
        'bd': orig_hive.get('bd', ''),
        'group': orig_hive.get('group', ''),
        'cat_status': new_cat_status,
    }


D = {}
for hive_name, orig_hive in orig_D['hives'].items():
    hive_id = str(orig_hive['id'])
    cat_excel = cat_data_excel.get(hive_id, {})
    goods_hive = goods_today.get(hive_id, {})
    D[hive_name] = update_hive_data(orig_hive, cat_excel, goods_hive)

D_with_meta = dict(orig_D)
D_with_meta['hives'] = D
D_with_meta['meta'] = META

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

yesterday_scores = {}
for hive_name, orig_hive in orig_D['hives'].items():
    hive_id = str(orig_hive['id'])
    cat_excel = cat_data_excel.get(hive_id, {})
    goods_hive_y = goods_yesterday.get(hive_id, {})
    hive_data_y = update_hive_data(orig_hive, cat_excel, goods_hive_y)
    fs, dr, coeff, dab = calc_score(hive_data_y)
    yesterday_scores[hive_id] = {
        'name': hive_name,
        'bd': orig_hive.get('bd', ''),
        'group': orig_hive.get('group', ''),
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

def parse_score_sheet(df):
    """Parse 昨日得分/上周得分 sheet for 标杆达标率."""
    result = {'overall': 0.0, 'groups': {}, 'hives': {}}
    if len(df) > 4:
        result['overall'] = safe_float(df.iloc[4, 9])
    for i in range(5, min(9, len(df))):
        group_name = safe_str(df.iloc[i, 0])
        if group_name:
            result['groups'][group_name] = safe_float(df.iloc[i, 9])
    for i in range(41, len(df)):
        hive_name = safe_str(df.iloc[i, 2])
        if hive_name:
            result['hives'][hive_name] = safe_float(df.iloc[i, 9])
    return result

yesterday_sheet = parse_score_sheet(df_score_yesterday)
lastweek_sheet = parse_score_sheet(df_score_lastweek)

trend_hives = {}
for hid, ts in today_scores.items():
    ys = yesterday_scores.get(str(hid), {})
    dod = calc_dod(ts['finalScore'], ys.get('finalScore', 0))
    hive_name = ts['name']
    lw_val = lastweek_sheet['hives'].get(hive_name, 0)
    today_dabiao_rate = ts['dabiaoRate']
    mom = calc_dod(today_dabiao_rate, lw_val) if lw_val else (r2(today_dabiao_rate * 100) if today_dabiao_rate > 0 else 0.0)
    trend_hives[str(hid)] = {'dod': dod, 'mom': mom}

groups = set(ts['group'] for ts in today_scores.values())
trend_groups = {}
for g in groups:
    g_hids = [h for h, ts in today_scores.items() if ts['group'] == g]
    if not g_hids:
        continue
    avg_today = sum(today_scores[h]['finalScore'] for h in g_hids) / len(g_hids)
    avg_yesterday = sum(yesterday_scores.get(str(h), {}).get('finalScore', 0) for h in g_hids) / len(g_hids)
    dod = calc_dod(avg_today, avg_yesterday)
    lw_val = lastweek_sheet['groups'].get(g, 0)
    avg_dabiao = sum(today_scores[h]['dabiaoRate'] for h in g_hids) / len(g_hids)
    mom = calc_dod(avg_dabiao, lw_val) if lw_val else 0.0
    trend_groups[g] = {'dod': dod, 'mom': mom}

all_hids = list(today_scores.keys())
overall_today = sum(today_scores[h]['finalScore'] for h in all_hids) / max(len(all_hids), 1)
overall_yesterday = sum(yesterday_scores.get(str(h), {}).get('finalScore', 0) for h in all_hids) / max(len(all_hids), 1)
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

# Use base deduction for daily calculation (same as D object)
all_cat_goods = {}
for _, row in df_goods.iterrows():
    cat = safe_str(row.get('宽前端三级品类'))
    if not cat:
        continue

    # 考核期日均单产 = (当前订单 - 基期订单) / 考核天数
    spu_id_str = safe_str(row.get('SPU-ID'))
    current_orders = safe_float(row.get('拼好饭订单'))
    base_orders = base_orders_map.get(spu_id_str, 0.0)
    exam_orders = current_orders - base_orders
    daily = max(exam_orders / META['exam_days'], 0.0) if META['exam_days'] > 0 else 0.0
    if daily <= 0:
        continue

    threshold = safe_float(row.get('标杆单产阈值'))
    is_bm = 1 if (daily >= threshold and threshold > 0) else 0

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
        'benchmark': is_bm,
    }

    if cat not in all_cat_goods:
        all_cat_goods[cat] = []
    all_cat_goods[cat].append(item)

TOP10_DATA = {}
for cat, items in all_cat_goods.items():
    items.sort(key=lambda x: -x['output'])
    top10 = items[:10]
    std_names = list(dict.fromkeys(it['std'] for it in items if it['std']))
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

# Helper: replace between placeholder comments
def replace_placeholder(html, var_decl, ph_start, ph_end, new_value):
    """Replace content between /*PH_START*/ and /*PH_END*/ comments."""
    ph_pattern = re.compile(
        re.escape(ph_start) + r'.*?' + re.escape(ph_end),
        re.DOTALL
    )
    replacement = ph_start + new_value + ph_end
    new_html, count = ph_pattern.subn(replacement, html, count=1)
    return new_html, count > 0

# 10.1 Replace TREND
trend_js_val = json.dumps(TREND, ensure_ascii=False, indent=2)
trend_js = f"const TREND = /*TREND_PLACEHOLDER*/{trend_js_val}/*END_TREND*/;"
# Try placeholder approach first
html, ok = replace_placeholder(html, 'const TREND = ', '/*TREND_PLACEHOLDER*/', '/*END_TREND*/', trend_js_val)
if ok:
    print("  ✓ TREND replaced (placeholder)")
else:
    # Fallback: old regex
    pattern = re.compile(r'const TREND\s*=\s*\{.*?\};', re.DOTALL)
    if pattern.search(html):
        html = pattern.sub(f"const TREND = /*TREND_PLACEHOLDER*/{trend_js_val}/*END_TREND*/;", html, count=1)
        print("  ✓ TREND replaced (regex fallback)")
    else:
        print("  ✗ TREND pattern not found!")

# 10.2 Append to HISTORY
# Try placeholder approach: replace the content between placeholders
hist_ph_start = '/*HISTORY_PLACEHOLDER*/'
hist_ph_end = '/*END_HISTORY*/'
ph_pattern_hist = re.compile(re.escape(hist_ph_start) + r'(.*?)' + re.escape(hist_ph_end), re.DOTALL)
match_hist_ph = ph_pattern_hist.search(html)
new_entry = json.dumps(history_entry, ensure_ascii=False)
if match_hist_ph:
    existing_raw = match_hist_ph.group(1).strip()
    # existing_raw is a JSON array content (without [ and ])
    if existing_raw == '[]' or existing_raw == '':
        existing_raw = ''
    else:
        # strip outer brackets if present
        if existing_raw.startswith('[') and existing_raw.endswith(']'):
            existing_raw = existing_raw[1:-1]
        existing_raw = existing_raw.rstrip().rstrip(',')
    new_array = '[' + (existing_raw + ',' if existing_raw else '') + new_entry + ']'
    html = html[:match_hist_ph.start()] + hist_ph_start + new_array + hist_ph_end + html[match_hist_ph.end():]
    print("  ✓ HISTORY appended (placeholder)")
else:
    # Fallback: old regex
    pattern_hist = re.compile(r'(const HISTORY\s*=\s*\[)(.*?)(\];)', re.DOTALL)
    match_hist = pattern_hist.search(html)
    if match_hist:
        existing = match_hist.group(2).rstrip().rstrip(',')
        new_history = f"{match_hist.group(1)}{existing},{new_entry}];"
        html = html[:match_hist.start()] + new_history + html[match_hist.end():]
        print("  ✓ HISTORY appended (regex fallback)")
    else:
        print("  ✗ HISTORY pattern not found!")

# 10.3 Replace TOP10_DATA
top10_val = json.dumps(TOP10_DATA, ensure_ascii=False)
top10_js = f"const TOP10_DATA = /*TOP10_PLACEHOLDER*/{top10_val}/*END_TOP10*/;\n"
html, ok = replace_placeholder(html, 'const TOP10_DATA = ', '/*TOP10_PLACEHOLDER*/', '/*END_TOP10*/', top10_val)
if ok:
    print("  ✓ TOP10_DATA replaced (placeholder)")
else:
    # Fallback: old approaches
    pattern_top10 = re.compile(r'const TOP10_DATA\s*=\s*\{.*?\};\s*(?=const |function |//|<)', re.DOTALL)
    if pattern_top10.search(html):
        html = pattern_top10.sub(f"const TOP10_DATA = /*TOP10_PLACEHOLDER*/{top10_val}/*END_TOP10*/;\n", html, count=1)
        print("  ✓ TOP10_DATA replaced (regex fallback)")
    else:
        idx_start = html.find('const TOP10_DATA')
        if idx_start >= 0:
            rest = html[idx_start + 16:]
            idx_next_const = rest.find('\nconst ')
            if idx_next_const < 0:
                idx_next_const = rest.find('\n\nconst ')
            if idx_next_const >= 0:
                html = html[:idx_start] + f"const TOP10_DATA = /*TOP10_PLACEHOLDER*/{top10_val}/*END_TOP10*/;\n" + html[idx_start + 16 + idx_next_const:]
                print("  ✓ TOP10_DATA replaced (alt method)")
            else:
                print("  ✗ Could not find TOP10_DATA end!")
        else:
            print("  ✗ TOP10_DATA not found!")

# 10.4 Replace D
d_val_js = json.dumps(D_with_meta, ensure_ascii=False)
d_js = f"const D = /*D_PLACEHOLDER*/{d_val_js}/*END_D*/;\n"
# Try placeholder approach first
html, ok = replace_placeholder(html, 'const D = ', '/*D_PLACEHOLDER*/', '/*END_D*/', d_val_js)
if ok:
    print("  ✓ D replaced (placeholder)")
else:
    # Fallback: brace counting
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
                    d_end = i + 1
                    break
        if d_end > 0:
            if d_end < len(html) and html[d_end] == ';':
                d_end += 1
            html = html[:idx_d_start] + d_js + html[d_end:]
            print("  ✓ D replaced (brace fallback)")
        else:
            print("  ✗ Could not find D closing brace!")
    else:
        print("  ✗ D variable not found!")

# ============================================================
# 10.5 Build and replace SEARCH_DATA
# ============================================================
print("Building SEARCH_DATA...")

# Build lookup: spu_id -> extra fields from current Excel (商品诊断明细)
extra_fields_map = {}  # spu_id_str -> dict
for _, row in df_goods.iterrows():
    spu_id = safe_str(row.get('SPU-ID'))
    if not spu_id:
        continue
    extra_fields_map[spu_id] = {
        'liang_chu': safe_bool(row.get('商家亮厨是否达标')),
        'ad': safe_bool(row.get('是否买广告')),
        'pin_dan_bao': safe_bool(row.get('是否开通拼单宝')),
        'zi_qu': safe_bool(row.get('是否开通到店自取')),
        'score': r2(safe_float(row.get('商品分'))),
        'bd_name': safe_str(row.get('BD姓名')),
        'merchant_name': safe_str(row.get('商家名称')),
        'hive_name': safe_str(row.get('蜂窝名称')),
        'group': safe_str(row.get('联络点')),
        'cat': safe_str(row.get('宽前端三级品类')),
    }

# Build SEARCH_DATA: all SPU rows from 商品诊断明细
SEARCH_DATA = []
for _, row in df_goods.iterrows():
    spu_id = safe_str(row.get('SPU-ID'))
    if not spu_id:
        continue

    hive_name = safe_str(row.get('蜂窝名称'))
    group = safe_str(row.get('联络点'))
    cat = safe_str(row.get('宽前端三级品类'))

    current_orders = safe_float(row.get('拼好饭订单'))
    base_orders = base_orders_map.get(spu_id, 0.0)
    exam_orders = current_orders - base_orders
    daily = r2(max(exam_orders / META['exam_days'], 0.0) if META['exam_days'] > 0 else 0.0)

    threshold = r2(safe_float(row.get('标杆单产阈值')))
    is_bm = 1 if (daily >= threshold and threshold > 0) else 0
    diff = r2(daily - threshold) if threshold > 0 else 0.0

    entry = {
        'spu_id': spu_id,
        'spu_name': safe_str(row.get('SPU名称', '')),
        'hive_name': hive_name,
        'cat': cat,
        'group': group,
        'bd_name': safe_str(row.get('BD姓名')),
        'merchant_name': safe_str(row.get('商家名称')),
        'daily': daily,
        'threshold': threshold,
        'is_bm': is_bm,
        'diff': diff,
        'liang_chu': safe_bool(row.get('商家亮厨是否达标')),
        'ad': safe_bool(row.get('是否买广告')),
        'pin_dan_bao': safe_bool(row.get('是否开通拼单宝')),
        'zi_qu': safe_bool(row.get('是否开通到店自取')),
        'score': r2(safe_float(row.get('商品分'))),
    }
    SEARCH_DATA.append(entry)

print(f"  SEARCH_DATA 条数: {len(SEARCH_DATA)}")

search_val_js = json.dumps(SEARCH_DATA, ensure_ascii=False)
search_js = f"const SEARCH_DATA = /*SEARCH_PLACEHOLDER*/{search_val_js}/*END_SEARCH*/;"

# Replace SEARCH_DATA - try placeholder first
html, ok = replace_placeholder(html, 'const SEARCH_DATA = ', '/*SEARCH_PLACEHOLDER*/', '/*END_SEARCH*/', search_val_js)
if ok:
    print("  ✓ SEARCH_DATA replaced (placeholder)")
else:
    # Fallback: regex
    pattern_search = re.compile(r'const SEARCH_DATA\s*=\s*\[.*?\];', re.DOTALL)
    if pattern_search.search(html):
        html = pattern_search.sub(search_js, html, count=1)
        print("  ✓ SEARCH_DATA replaced (regex fallback)")
    else:
        insert_marker = '// END_DATA_BLOCK'
        if insert_marker in html:
            html = html.replace(insert_marker, search_js + '\n' + insert_marker)
            print("  ✓ SEARCH_DATA inserted at END_DATA_BLOCK")
        else:
            print("  ✗ SEARCH_DATA placeholder not found, skipping")

# 11. Write output
with open(HTML_PATH, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"\n✅ Done! Updated {HTML_PATH}")
print(f"   File size: {len(html.encode('utf-8')):,} bytes")
print(f"   蜂窝数: {len(D_with_meta.get('hives', {}))}")
print(f"   TOP10品类: {len(TOP10_DATA)}")
print(f"   SEARCH_DATA 条数: {len(SEARCH_DATA)}")
print(f"   日期: {META['date']}, 考核天数: {META['exam_days']}")
