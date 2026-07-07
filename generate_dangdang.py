"""
当当网渠道周报数据生成脚本
用法：python generate_dangdang.py [Excel路径]
默认：D:/Desktop/0706.xlsx
输出：dangdang_data.js
"""
import pandas as pd
import json
import sys
from datetime import datetime, timedelta

# ============================================================
# 0. 参数
# ============================================================
EXCEL_PATH = sys.argv[1] if len(sys.argv) > 1 else 'D:/Desktop/0706.xlsx'

# ============================================================
# 1. 读取数据
# ============================================================
df_ld = pd.read_excel(EXCEL_PATH, sheet_name='leads明细')
df_hm = pd.read_excel(EXCEL_PATH, sheet_name='活码加微明细')

# 日期列标准化
df_ld['日期'] = pd.to_datetime(df_ld['leads购课日期'])
df_hm['日期'] = pd.to_datetime(df_hm['加微日期'])

# ============================================================
# 2. 筛选当当数据
# ============================================================
hot_mask = df_ld['leads计划名称'].isin(['当当网', '当当双科阅写体验课'])
sf_mask = df_ld['leads图书SPU名称'].astype(str) == '当当课包'
dd_mask = hot_mask | sf_mask

dd = df_ld[dd_mask].copy()
dd_hm = df_hm[df_hm['图书项目组'] == '当当网'].copy()

# 字段反转
dd['结课'] = dd['leads转正人次'].fillna(0).astype(int)
dd['转正'] = dd['leads结课人次'].fillna(0).astype(int)
dd['leads'] = dd['leads人次'].fillna(0).astype(int)

def classify_course(t):
    t = str(t)
    if '0元' in t: return '0元'
    if '低价' in t or '1元' in t: return '低价'
    return '其他'

dd['价位'] = dd['leads类型'].apply(classify_course)

# 来源标记（用loc避免warning）
dd.loc[hot_mask[dd_mask], '来源'] = '热点'
dd.loc[sf_mask[dd_mask], '来源'] = '索粉'

# ============================================================
# 3. 日期范围
# ============================================================
max_date = max(dd['日期'].max(), dd_hm['日期'].max())

def fmt_date(d):
    return f"{d.month}/{d.day}"

def fmt_ymd(d):
    return f"{d.year}/{d.month}/{d.day}"

data_cutoff = fmt_ymd(max_date)

this_week_end = max_date
this_week_start = this_week_end - timedelta(days=6)

last_week_end = this_week_start - timedelta(days=1)
last_week_start = last_week_end - timedelta(days=6)

tw_label = f"{fmt_date(this_week_start)}-{fmt_date(this_week_end)}"
lw_label = f"{fmt_date(last_week_start)}-{fmt_date(last_week_end)}"

print(f"数据截止: {data_cutoff}")
print(f"本周: {tw_label}  |  上周: {lw_label}")

# ============================================================
# 4. 辅助函数
# ============================================================
def rate(a, b):
    if b == 0: return '--'
    return f"{a/b*100:.1f}%"

def rate_val(a, b):
    if b == 0: return 0
    return round(a/b*100, 1)

def change(this_v, last_v):
    if last_v == 0:
        return ('+Inf', 'up') if this_v > 0 else ('--', 'flat')
    pct = (this_v - last_v) / last_v * 100
    if abs(pct) < 0.05: return '--', 'flat'
    return (f"+{pct:.1f}%", 'up') if pct > 0 else (f"{pct:.1f}%", 'down')

def get_week_range(d):
    return d - timedelta(days=d.weekday())

# ============================================================
# 5. 整体汇总（不分品，当当只有一本书）
# ============================================================
def week_stats(dd, dd_hm, ws, we):
    w_dd = dd[(dd['日期'] >= ws) & (dd['日期'] <= we)]
    w_hm = dd_hm[(dd_hm['日期'] >= ws) & (dd_hm['日期'] <= we)]
    jw = int(w_hm['当日加微人数(包含删除好友)'].sum())
    leads = int(w_dd['leads'].sum())
    jk = int(w_dd['结课'].sum())
    zz = int(w_dd['转正'].sum())
    lk = rate(w_dd[w_dd['价位'].isin(['0元','低价'])]['leads'].sum(), jw)
    jkzz = rate(zz, jk)
    return {'jw': jw, 'leads': leads, 'jk': jk, 'zz': zz, 'lkRate': lk, 'jkzzRate': jkzz}

tw = week_stats(dd, dd_hm, this_week_start, this_week_end)
lw = week_stats(dd, dd_hm, last_week_start, last_week_end)

# 累计
cum_jw = int(dd_hm['当日加微人数(包含删除好友)'].sum())
cum_leads = int(dd['leads'].sum())
cum_jk = int(dd['结课'].sum())
cum_zz = int(dd['转正'].sum())

ly_total = int(dd[dd['价位'] == '0元']['leads'].sum())
ly_zz_total = int(dd[dd['价位'] == '0元']['转正'].sum())
dj_total = int(dd[dd['价位'] == '低价']['leads'].sum())
dj_zz_total = int(dd[dd['价位'] == '低价']['转正'].sum())
other_total = int(dd[dd['价位'] == '其他']['leads'].sum())
other_zz_total = int(dd[dd['价位'] == '其他']['转正'].sum())

cum_lk_rate = rate(ly_total + dj_total, cum_jw)
cum_jkzz_rate = rate(cum_zz, cum_jk)

# 环比
jw_chg, jw_dir = change(tw['jw'], lw['jw'])
leads_chg, leads_dir = change(tw['leads'], lw['leads'])
jk_chg, jk_dir = change(tw['jk'], lw['jk'])
zz_chg, zz_dir = change(tw['zz'], lw['zz'])

tw_lk_val = rate_val(tw['leads'], tw['jw'])
lw_lk_val = rate_val(lw['leads'], lw['jw'])
lk_chg, lk_dir = change(tw_lk_val, lw_lk_val)

tw_jkzz_val = rate_val(tw['zz'], tw['jk']) if tw['jk'] > 0 else 0
lw_jkzz_val = rate_val(lw['zz'], lw['jk']) if lw['jk'] > 0 else 0
jkzz_chg, jkzz_dir = change(tw_jkzz_val, lw_jkzz_val)

# ============================================================
# 6. 周明细（近4周）
# ============================================================
weeks_data = []
current_week_start = get_week_range(this_week_start)
for i in range(3, -1, -1):
    ws = current_week_start - timedelta(weeks=i)
    we = ws + timedelta(days=6)
    w_label = f"{fmt_date(ws)}-{fmt_date(we)}"

    w_dd = dd[(dd['日期'] >= ws) & (dd['日期'] <= we)]
    w_hm = dd_hm[(dd_hm['日期'] >= ws) & (dd_hm['日期'] <= we)]

    w_jw = int(w_hm['当日加微人数(包含删除好友)'].sum())
    w_ly = int(w_dd[w_dd['价位'] == '0元']['leads'].sum())
    w_dj = int(w_dd[w_dd['价位'] == '低价']['leads'].sum())
    w_sk = int(w_dd[w_dd['leads类型'].astype(str).str.contains('双科', na=False)]['leads'].sum())
    w_hot = int(w_dd[w_dd['来源'] == '热点']['leads'].sum())
    w_sf = int(w_dd[w_dd['来源'] == '索粉']['leads'].sum())
    w_zz = int(w_dd['转正'].sum())

    ly_rate_val_w = rate_val(w_ly, w_jw)
    ly_outlier = ly_rate_val_w > 100

    weeks_data.append({
        'label': w_label,
        'jw': w_jw,
        'ly_leads': w_ly, 'ly_rate': rate(w_ly, w_jw), 'ly_outlier': ly_outlier,
        'dj_leads': w_dj, 'dj_rate': rate(w_dj, w_jw),
        'sk_leads': w_sk,
        'hot': w_hot, 'sufen': w_sf, 'zz': w_zz,
    })

w4_jw = sum(w['jw'] for w in weeks_data)
w4_ly = sum(w['ly_leads'] for w in weeks_data)
w4_dj = sum(w['dj_leads'] for w in weeks_data)
w4_hot = sum(w['hot'] for w in weeks_data)
w4_sf = sum(w['sufen'] for w in weeks_data)
w4_zz = sum(w['zz'] for w in weeks_data)

# ============================================================
# 7. 月明细
# ============================================================
dd['月份'] = dd['日期'].dt.to_period('M')
monthly_data = []
for m in sorted(dd['月份'].dropna().unique()):
    m_str = str(m)
    m_dd = dd[dd['月份'] == m]
    m_hm = dd_hm[dd_hm['日期'].dt.to_period('M') == m]

    m_jw = int(m_hm['当日加微人数(包含删除好友)'].sum())
    m_ly = int(m_dd[m_dd['价位'] == '0元']['leads'].sum())
    m_dj = int(m_dd[m_dd['价位'] == '低价']['leads'].sum())
    m_sk = int(m_dd[m_dd['leads类型'].astype(str).str.contains('双科', na=False)]['leads'].sum())
    m_ly_zz = int(m_dd[m_dd['价位'] == '0元']['转正'].sum())
    m_dj_zz = int(m_dd[m_dd['价位'] == '低价']['转正'].sum())
    m_sk_zz = int(m_dd[m_dd['leads类型'].astype(str).str.contains('双科', na=False)]['转正'].sum())

    monthly_data.append({
        'month': m_str, 'jw': m_jw,
        'ly_leads': m_ly, 'ly_rate': rate(m_ly, m_jw), 'ly_zz': m_ly_zz,
        'dj_leads': m_dj, 'dj_rate': rate(m_dj, m_jw), 'dj_zz': m_dj_zz,
        'sk_leads': m_sk, 'sk_zz': m_sk_zz,
    })

# ============================================================
# 8. 累计结构（拆三张卡）
# ============================================================

# --- 卡1：加微 → 课型漏斗 ---
sk_total = int(dd[dd['leads类型'].astype(str).str.contains('双科', na=False)]['leads'].sum())
sk_zz_total = int(dd[dd['leads类型'].astype(str).str.contains('双科', na=False)]['转正'].sum())

# --- 卡2：热点详细 ---
hot_dd = dd[dd['来源'] == '热点']
hot_ly = int(hot_dd[hot_dd['价位'] == '0元']['leads'].sum())
hot_dj = int(hot_dd[hot_dd['价位'] == '低价']['leads'].sum())
hot_sk = int(hot_dd[hot_dd['leads类型'].astype(str).str.contains('双科', na=False)]['leads'].sum())
hot_other = int(hot_dd['leads'].sum()) - hot_ly - hot_dj - hot_sk
hot_ly_zz = int(hot_dd[hot_dd['价位'] == '0元']['转正'].sum())
hot_sk_zz = int(hot_dd[hot_dd['leads类型'].astype(str).str.contains('双科', na=False)]['转正'].sum())
hot_other_zz = int(hot_dd['转正'].sum()) - hot_ly_zz - hot_sk_zz

# --- 卡3：索粉详细 ---
sf_dd = dd[dd['来源'] == '索粉']
sf_total = int(sf_dd['leads'].sum())
sf_zz_total = int(sf_dd['转正'].sum())
sf_ly = int(sf_dd[sf_dd['价位'] == '0元']['leads'].sum())

hot_total = int(dd[dd['来源'] == '热点']['leads'].sum())
hot_zz_all = int(dd[dd['来源'] == '热点']['转正'].sum())

# ============================================================
# 9. 课型拆分
# ============================================================
course_types = []
for ct in dd['leads类型'].dropna().unique():
    ct_dd = dd[dd['leads类型'] == ct]
    course_types.append({
        'name': ct,
        'count': int(ct_dd['leads'].sum()),
        'pct': rate(ct_dd['leads'].sum(), cum_leads),
        'zz': int(ct_dd['转正'].sum()),
    })
course_types.sort(key=lambda x: -x['count'])

# ============================================================
# 10. 学科分布
# ============================================================
subjects = []
for subj in dd['leads学科'].dropna().unique():
    subj_dd = dd[dd['leads学科'] == subj]
    subjects.append({
        'name': subj,
        'count': int(subj_dd['leads'].sum()),
        'pct': rate(subj_dd['leads'].sum(), cum_leads),
        'zz': int(subj_dd['转正'].sum()),
    })
subjects.sort(key=lambda x: -x['count'])

# ============================================================
# 11. TOP热点来源
# ============================================================
top_sources = []
for unit in hot_dd['leads单元名称'].dropna().unique():
    unit_dd = hot_dd[hot_dd['leads单元名称'] == unit]
    top_sources.append({
        'name': unit,
        'count': int(unit_dd['leads'].sum()),
        'zz': int(unit_dd['转正'].sum()),
    })
top_sources.sort(key=lambda x: -x['count'])
top_sources = top_sources[:10]

# ============================================================
# 12. 索粉TOP来源
# ============================================================
sufen_sources = []
for plan in sf_dd['leads计划名称'].dropna().unique():
    plan_dd = sf_dd[sf_dd['leads计划名称'] == plan]
    sufen_sources.append({
        'name': plan,
        'count': int(plan_dd['leads'].sum()),
        'zz': int(plan_dd['转正'].sum()),
    })
sufen_sources.sort(key=lambda x: -x['count'])

# ============================================================
# 13. 活码资源位（新版活码加微表无资源位列，跳过）
# ============================================================
resources = []  # 新版数据无资源位列，留空

# ============================================================
# 14. 活码名称
# ============================================================
qr_names = []
for qr in dd_hm['活码名称'].dropna().unique():
    qr_hm = dd_hm[dd_hm['活码名称'] == qr]
    qr_names.append({
        'name': qr,
        'count': int(qr_hm['当日加微人数(包含删除好友)'].sum()),
        'pct': rate(qr_hm['当日加微人数(包含删除好友)'].sum(), cum_jw),
    })
qr_names.sort(key=lambda x: -x['count'])

# ============================================================
# 15. 图表数据
# ============================================================
chart_labels = [w['label'] for w in weeks_data]
chart_jw = [w['jw'] for w in weeks_data]
chart_ly_rate = [rate_val(w['ly_leads'], w['jw']) if w['jw'] > 0 else 0 for w in weeks_data]
chart_dj_rate = [rate_val(w['dj_leads'], w['jw']) if w['jw'] > 0 else 0 for w in weeks_data]

# ============================================================
# 16. 组装 JS 数据
# ============================================================
data = {
    'meta': {
        'title': '当当网渠道数据周报',
        'thisWeekLabel': tw_label,
        'lastWeekLabel': lw_label,
        'dataDate': data_cutoff,
        'productName': '阅读写作双提升课',
        'singleProduct': True,
    },
    'summary': {
        'thisWeek': tw,
        'lastWeek': lw,
        'cumulative': {
            'jw': cum_jw, 'leads': cum_leads, 'jk': cum_jk, 'zz': cum_zz,
            'lkRate': cum_lk_rate, 'jkzzRate': cum_jkzz_rate,
        },
        'changes': {
            'jw': jw_chg, 'jwDir': jw_dir,
            'leads': leads_chg, 'leadsDir': leads_dir,
            'jk': jk_chg, 'jkDir': jk_dir,
            'zz': zz_chg, 'zzDir': zz_dir,
            'lkRate': lk_chg, 'lkRateDir': lk_dir,
            'jkzzRate': jkzz_chg, 'jkzzRateDir': jkzz_dir,
        },
        # 累计漏斗用的0元+低价合计
        'cumLkBase': ly_total + dj_total,
    },
    'weeks': weeks_data,
    'weeksTotal': {
        'jw': w4_jw, 'ly_leads': w4_ly, 'ly_rate': rate(w4_ly, w4_jw),
        'dj_leads': w4_dj, 'dj_rate': rate(w4_dj, w4_jw),
        'hot': w4_hot, 'sufen': w4_sf, 'zz': w4_zz,
    },
    'monthly': monthly_data,
    # 三张累计卡
    'cumulativeCards': {
        'funnel': {
            'jw': cum_jw,
            'items': [
                {'label': '0元课', 'leads': ly_total, 'rate': rate(ly_total, cum_jw), 'zz': ly_zz_total, 'zzRate': rate(ly_zz_total, ly_total)},
                {'label': '低价课', 'leads': dj_total, 'rate': rate(dj_total, cum_jw), 'zz': dj_zz_total, 'zzRate': rate(dj_zz_total, dj_total) if dj_total > 0 else '--'},
                {'label': '双科爆品课', 'leads': sk_total, 'rate': rate(sk_total, cum_jw), 'zz': sk_zz_total, 'zzRate': rate(sk_zz_total, sk_total) if sk_total > 0 else '--'},
                {'label': '其他课型', 'leads': other_total - sk_total + sk_total,
                 'rate': rate(other_total, cum_jw), 'zz': other_zz_total},
            ],
            'note': '前几个月加微少但有领课——因早期走的不是加微链路（索粉/历史队列析出），近期才切换到加微链路',
        },
        'hot': {
            'total': hot_total, 'totalPct': rate(hot_total, cum_leads),
            'items': [
                {'label': '0元爆品课', 'leads': hot_ly, 'zz': hot_ly_zz},
                {'label': '双科爆品课', 'leads': hot_sk, 'zz': hot_sk_zz},
                {'label': '低价爆品课', 'leads': hot_dj, 'zz': 0},
                {'label': '其他课型', 'leads': hot_other, 'zz': hot_other_zz},
            ],
            'totalZz': hot_zz_all,
        },
        'sufen': {
            'total': sf_total, 'totalPct': rate(sf_total, cum_leads),
            'totalZz': sf_zz_total,
            'note': '索粉暂无转正，均为0元课 leads',
        },
    },
    'courseTypes': course_types,
    'subjects': subjects,
    'topSources': top_sources,
    'sufenSources': sufen_sources,
    'resources': resources,
    'qrNames': qr_names,
    'footnotes': {
        'weekly': '前几周加微基数小（1-5人），率超高不代表正常水平；6月起加微放量后率趋于稳定',
        'monthly': '前几个月加微少但有领课——因早期走索粉/历史队列析出链路，非加微链路；近期切换为加微链路。异常率值保留供参考',
        'summary': f'累计领课率以 0元＋低价课口径（{ly_total}+{dj_total}）/{cum_jw}={cum_lk_rate}；双科爆品课等非0元/低价 leads 单独呈现',
    },
    'chart': {
        'labels': chart_labels,
        'jw': chart_jw,
        'lyRate': chart_ly_rate,
        'djRate': chart_dj_rate,
    },
}

# ============================================================
# 17. 写出 dangdang_data.js
# ============================================================
js_content = f"// 当当网渠道周报数据 | 自动生成于 {datetime.now().strftime('%Y-%m-%d %H:%M')}\n// 数据源: {EXCEL_PATH}\nconst DD_DATA = {json.dumps(data, ensure_ascii=False, indent=2)};\n"

with open('dangdang_data.js', 'w', encoding='utf-8') as f:
    f.write(js_content)

print(f"\n[OK] dangdang_data.js 已生成")
print(f"   累计加微: {cum_jw} | leads: {cum_leads} | 转正: {cum_zz}")
print(f"   近4周: {w4_jw}加微 / {w4_ly + w4_dj}leads(0元+低价)")
