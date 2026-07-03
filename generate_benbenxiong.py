"""
生成笨笨熊下册周报数据 JSON 和 HTML
用法: python generate_benbenxiong.py [Excel路径]
默认: D:/Desktop/0630.xlsx
输出: benbenxiong_data.json, 笨笨熊周报.html
"""
import pandas as pd
import json
import sys
import re
from pathlib import Path
from datetime import datetime, timedelta

EXCEL = sys.argv[1] if len(sys.argv) > 1 else "D:/Desktop/0630.xlsx"
OUTPUT_JSON = "benbenxiong_data.json"
OUTPUT_HTML = "笨笨熊周报.html"
TEMPLATE = "笨笨熊周报_template.html"

# ===================== 工具函数 =====================

def safe_int(v, default=0):
    try:
        return int(float(v))
    except:
        return default

def pct(a, b):
    if b == 0:
        return 0
    return round(a / b * 100, 2)

def fmt_pct(v):
    return f"{v:.2f}%"

def fmt_pct1(v):
    return f"{v:.1f}%"

def num(v):
    return f"{v:,}"

def huanbi(curr, prev):
    if prev == 0:
        return None if curr == 0 else float('inf')
    return round((curr - prev) / prev * 100, 1)

def hb_str(v):
    if v is None or v == float('inf'):
        return '--'
    if v == -100:
        return '&darr; 100%'
    if v > 0:
        return f'&uarr; {v}%'
    if v < 0:
        return f'&darr; {abs(v)}%'
    return '--'

def hb_class(v):
    if v is None or v == float('inf'):
        return 'flat'
    if v > 0:
        return 'up'
    if v < 0:
        return 'down'
    return 'flat'

# ===================== 读取数据 =====================
df_week = pd.read_excel(EXCEL, sheet_name='分渠道周数据')
df_month = pd.read_excel(EXCEL, sheet_name='分渠道月数据')
df_total = pd.read_excel(EXCEL, sheet_name='分渠道总数据')
df_leads = pd.read_excel(EXCEL, sheet_name='leads明细')
df_hm = pd.read_excel(EXCEL, sheet_name='活码加微明细1')

# ===================== 整体数据 =====================
total_row = df_total[df_total['渠道'] == '笨笨熊'].iloc[0]
week_data = df_week[df_week['渠道'] == '笨笨熊'].sort_values('周开始时间', ascending=False)
month_data = df_month[df_month['渠道'] == '笨笨熊'].sort_values('月份', ascending=False)

overall = {
    '累计加微': safe_int(total_row['加微信号人数']),
    '累计leads': safe_int(total_row['leads人次']),
    '累计结课': safe_int(total_row['结课leads人次']),
    '累计转正': safe_int(total_row['结课转正人次']),
    '领课率': pct(safe_int(total_row['leads人次']), safe_int(total_row['加微信号人数'])),
    '结课转正率': pct(safe_int(total_row['结课转正人次']), safe_int(total_row['结课leads人次'])),
}

# ===================== 版本拆分（热点：先筛选计划名称==笨笨熊，再按单元名分类） =====================
def hot_version(unit_name):
    u = str(unit_name)
    if '三味' in u: return '三味版'
    if '公版' in u: return '公版薄版'
    if '蓝鲸' in u: return '蓝鲸版'
    if '快乐读书吧' in u: return '蓝鲸版'
    # 带年级数字（二年级/三年级等）的单元，属于蓝鲸版
    for g in ['一', '二', '三', '四', '五', '六']:
        if f'{g}年级' in u:
            return '蓝鲸版'
    return '通用/其他'

def sf_version(spu_name):
    s = str(spu_name)
    if '三味' in s: return '三味版'
    if '公版' in s: return '公版薄版'
    if '蓝鲸' in s: return '蓝鲸版'
    return '其他'

# 热点：计划名称==笨笨熊，再根据单元名称划分版本
hot = df_leads[df_leads['leads计划名称'] == '笨笨熊'].copy()
hot['版本'] = hot['leads单元名称'].apply(hot_version)

# 索粉：SPU含笨笨熊且计划名称不是笨笨熊
sf = df_leads[(df_leads['leads图书SPU名称'].str.contains('笨笨熊', na=False)) &
              (df_leads['leads计划名称'] != '笨笨熊')].copy()
sf['版本'] = sf['leads图书SPU名称'].apply(sf_version)

versions = {}
for ver in ['三味版', '公版薄版', '蓝鲸版']:
    hot_leads = safe_int(hot[hot['版本'] == ver]['leads人次'].sum())
    sf_leads = safe_int(sf[sf['版本'] == ver]['leads人次'].sum())
    hot_jieke = safe_int(hot[hot['版本'] == ver]['leads转正人次'].sum())
    hot_zz = safe_int(hot[hot['版本'] == ver]['leads结课人次'].sum())
    sf_jieke = safe_int(sf[sf['版本'] == ver]['leads转正人次'].sum())
    sf_zz = safe_int(sf[sf['版本'] == ver]['leads结课人次'].sum())
    versions[ver] = {
        '热点leads': hot_leads, '索粉leads': sf_leads,
        'leads合计': hot_leads + sf_leads,
        '结课': hot_jieke + sf_jieke,
        '转正': hot_zz + sf_zz,
        '领课率': pct(hot_leads + sf_leads, overall['累计加微']),
        '转正率': pct(hot_zz + sf_zz, hot_jieke + sf_jieke),
    }

other_hot = safe_int(hot[hot['版本'] == '通用/其他']['leads人次'].sum())
other_jk = safe_int(hot[hot['版本'] == '通用/其他']['leads转正人次'].sum())
other_zz = safe_int(hot[hot['版本'] == '通用/其他']['leads结课人次'].sum())
versions['未分版热点'] = {
    '热点leads': other_hot, '索粉leads': 0, 'leads合计': other_hot,
    '结课': other_jk, '转正': other_zz, '转正率': pct(other_zz, other_jk),
}

hot_total = safe_int(hot['leads人次'].sum())
hot_jk = safe_int(hot['leads转正人次'].sum())
hot_zz = safe_int(hot['leads结课人次'].sum())
sf_total = safe_int(sf['leads人次'].sum())
sf_jk = safe_int(sf['leads转正人次'].sum())
sf_zz = safe_int(sf['leads结课人次'].sum())

structure = {
    '热点': {'leads': hot_total, '结课': hot_jk, '转正': hot_zz, '转正率': pct(hot_zz, hot_jk)},
    '索粉': {'leads': sf_total, '结课': sf_jk, '转正': sf_zz, '转正率': pct(sf_zz, sf_jk)},
}

# ===================== 近5周数据（明细表+图表用） =====================
WEEK_N = 5
recent_weeks = week_data.head(WEEK_N).iloc[::-1]  # 按时间升序
weeks_detail = []
for _, row in recent_weeks.iterrows():
    jw = safe_int(row['加微信号人数'])
    leads = safe_int(row['leads人次'])
    jk = safe_int(row['结课leads人次'])
    zz = safe_int(row['结课转正人次'])
    leads0 = safe_int(row['0元_leads人次'])
    jk0 = safe_int(row['0元_结课leads人次'])
    zz0 = safe_int(row['0元_结课转正人次'])
    leads1 = safe_int(row['低价_leads人次'])
    jk1 = safe_int(row['低价_结课leads人次'])
    zz1 = safe_int(row['低价_结课转正人次'])
    start = pd.Timestamp(row['周开始时间'])
    end = start + timedelta(days=6)
    weeks_detail.append({
        '周开始': start.strftime('%m/%d'),
        '周范围': f"{start.strftime('%m/%d')}-{end.strftime('%m/%d')}",
        '加微': jw, 'leads': leads, '结课': jk, '转正': zz,
        '领课率': pct(leads, jw), '转正率': pct(zz, jk),
        '0元leads': leads0, '0元结课': jk0, '0元转正': zz0, '0元率': pct(leads0, jw),
        '低价leads': leads1, '低价结课': jk1, '低价转正': zz1, '低价率': pct(leads1, jw),
    })

# 本周/上周：用最后两个完整周
this_week = weeks_detail[-2] if len(weeks_detail) >= 2 else weeks_detail[-1] if weeks_detail else {}
last_week_orig = weeks_detail[-3] if len(weeks_detail) >= 3 else {}
lw_jw = last_week_orig.get('加微', 0)
lw_leads = last_week_orig.get('leads', 0)
lw_jk = last_week_orig.get('结课', 0)
lw_zz = last_week_orig.get('转正', 0)
lw_lk = pct(lw_leads, lw_jw)
lw_jkzz = pct(lw_zz, lw_jk)

recent_overall = {
    '近一周加微': this_week.get('加微', 0),
    '近一周leads': this_week.get('leads', 0),
    '近一周结课': this_week.get('结课', 0),
    '近一周转正': this_week.get('转正', 0),
    '近一周领课率': pct(this_week.get('leads', 0), this_week.get('加微', 0)),
    '近一周转正率': pct(this_week.get('转正', 0), this_week.get('结课', 0)),
    '加微环比': huanbi(this_week.get('加微', 0), lw_jw),
    'leads环比': huanbi(this_week.get('leads', 0), lw_leads),
    '转正环比': huanbi(this_week.get('转正', 0), lw_zz),
    '领课率环比': huanbi(pct(this_week.get('leads', 0), this_week.get('加微', 0)), lw_lk),
    '转正率环比': huanbi(pct(this_week.get('转正', 0), this_week.get('结课', 0)), lw_jkzz),
}

# ===================== 近5周（带0元/低价转正明细，替代月度） =====================
recent_5weeks = week_data.head(WEEK_N).iloc[::-1]
weeks5_detail = []
for _, row in recent_5weeks.iterrows():
    jw = safe_int(row['加微信号人数'])
    leads0 = safe_int(row['0元_leads人次'])
    zz0 = safe_int(row['0元_结课转正人次'])
    leads1 = safe_int(row['低价_leads人次'])
    zz1 = safe_int(row['低价_结课转正人次'])
    start = pd.Timestamp(row['周开始时间'])
    end = start + timedelta(days=6)
    weeks5_detail.append({
        '周范围': f"{start.strftime('%m/%d')}-{end.strftime('%m/%d')}",
        '加微': jw,
        '0元leads': leads0, '0元转正': zz0, '0元率': pct(leads0, jw),
        '低价leads': leads1, '低价转正': zz1, '低价率': pct(leads1, jw),
    })

# 课型拆分用全时段累计数据（分渠道总数据），避免与合计行全时段口径不一致
cumul_course = {
    '0元leads': safe_int(total_row['0元_leads人次']),
    '0元转正': safe_int(total_row['0元_结课转正人次']),
    '0元率': pct(safe_int(total_row['0元_leads人次']), overall['累计加微']),
    '0元转正率': pct(safe_int(total_row['0元_结课转正人次']), safe_int(total_row['0元_leads人次'])),
    '低价leads': safe_int(total_row['低价_leads人次']),
    '低价转正': safe_int(total_row['低价_结课转正人次']),
    '低价率': pct(safe_int(total_row['低价_leads人次']), overall['累计加微']),
    '低价转正率': pct(safe_int(total_row['低价_结课转正人次']), safe_int(total_row['低价_leads人次'])),
}

# ===================== 月度数据 =====================
recent_months = month_data.head(7).iloc[::-1]
months_detail = []
for _, row in recent_months.iterrows():
    jw = safe_int(row['加微信号人数'])
    leads0 = safe_int(row['0元_leads人次'])
    zz0 = safe_int(row['0元_结课转正人次'])
    leads1 = safe_int(row['低价_leads人次'])
    zz1 = safe_int(row['低价_结课转正人次'])
    months_detail.append({
        '月份': str(pd.Timestamp(row['月份']).strftime('%Y-%m')),
        '加微': jw,
        '0元leads': leads0, '0元转正': zz0, '0元率': pct(leads0, jw),
        '低价leads': leads1, '低价转正': zz1, '低价率': pct(leads1, jw),
    })

# ===================== TOP来源 =====================
# 热点：计划名称都是笨笨熊，展示具体热点名称
hot_grouped = hot.groupby('leads热点名称').agg(
    leads=('leads人次', 'sum'), jk=('leads转正人次', 'sum'), zz=('leads结课人次', 'sum')
).sort_values('leads', ascending=False).head(10)
top_hot = [{'name': str(n), 'leads': safe_int(r['leads']), '转正': safe_int(r['zz'])} for n, r in hot_grouped.iterrows()]

# 索粉：展示计划名称（不同来源的计划）
sf_grouped = sf.groupby('leads计划名称').agg(
    leads=('leads人次', 'sum'), jk=('leads转正人次', 'sum'), zz=('leads结课人次', 'sum')
).sort_values('leads', ascending=False).head(10)
top_sf = [{'name': str(n) if pd.notna(n) else '未知', 'leads': safe_int(r['leads']), '转正': safe_int(r['zz'])} for n, r in sf_grouped.iterrows()]

# ===================== 活码资源位拆分 =====================
# 过滤掉上册的笨笨熊（另开笨笨熊上册报告），并排除资源位=0/NaN
bm_hm_all = df_hm[df_hm['渠道'] == '笨笨熊']
bm_hm = bm_hm_all[~bm_hm_all['图书'].str.contains('上册', na=False)]
bm_hm_res = bm_hm[bm_hm['资源位'].notna() & (bm_hm['资源位'] != 0) & (bm_hm['资源位'] != '0')]
res_grouped = bm_hm_res.groupby('资源位')['当日加微人数(包含删除好友)'].sum().sort_values(ascending=False)
top_res = [{'name': str(r), '加微': safe_int(g)} for r, g in res_grouped.items()]

# ===================== 年级拆分（从活码名称提取） =====================
# 活码名称格式: 笨笨熊《西游记》语文五年级-内页音频
def extract_grade(hm_name):
    s = str(hm_name)
    for n in ['一', '二', '三', '四', '五', '六']:
        if f'{n}年级' in s:
            return f'{n}年级'
        if f'语文{n}' in s and '年级' not in s:
            return f'{n}年级'
    # fallback: 试试年级列
    return '未知'

bm_hm_grade = bm_hm.copy()
bm_hm_grade['年级分类'] = bm_hm_grade['活码名称'].apply(extract_grade)
gd_grouped = bm_hm_grade.groupby('年级分类')['当日加微人数(包含删除好友)'].sum().sort_values(ascending=False)
top_grade = [{'name': str(g), '加微': safe_int(v)} for g, v in gd_grouped.items()]

# ===================== 生成 JSON =====================
data = {
    'meta': {
        '渠道': '笨笨熊（下册）',
        '数据截止': '2026-06-30',
        '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M'),
        '说明': '包含三味版/公版薄版/蓝鲸版三个子版本，按一个品呈现',
    },
    'overall': overall,
    'recent_overall': recent_overall,
    'weeks': weeks_detail,
    'versions': versions,
    'structure': structure,
    'weeks5': weeks5_detail,
    'cumul_course': cumul_course,
    'months': months_detail,
    'top_hot': top_hot,
    'top_sf': top_sf,
    'top_res': top_res,
    'top_grade': top_grade,
}

with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# ===================== 生成 HTML =====================
html_template = Path(TEMPLATE).read_text(encoding='utf-8')
html_template = re.sub(r'\$([a-zA-Z_][a-zA-Z0-9_]*)', r'{{\1}}', html_template)

def version_table(ver, d, color, icon):
    return f"""<div>
    <div class="ext-col-title {color}">{icon} {ver}</div>
    <table class="ext-table">
      <thead><tr><th>指标</th><th>累计</th></tr></thead>
      <tbody>
        <tr><td>leads 领课人次</td><td class="num">{num(d['leads合计'])}</td></tr>
        <tr><td>结课人次</td><td class="num">{num(d['结课'])}</td></tr>
        <tr><td>转正人次</td><td class="num" style="color:#7c3aed;font-weight:700;">{d['转正']}</td></tr>
        <tr class="rate-row"><td>领课率</td><td class="cumulative">{fmt_pct(d['领课率'])}</td></tr>
        <tr class="rate-row"><td>结课转正率</td><td class="cumulative">{fmt_pct(d['转正率'])}</td></tr>
        <tr><td style="color:#64748b;font-size:11px;">热点 / 索粉</td><td class="num" style="font-size:12px;">{num(d['热点leads'])} / {num(d['索粉leads'])}</td></tr>
      </tbody>
    </table>
  </div>"""

def week_row(w):
    return f"<tr><td>{w['周范围']}</td><td class=\"num\">{num(w['加微'])}</td><td class=\"num\">{w['0元leads']}</td><td class=\"num wd-rate group-right\">{fmt_pct1(w['0元率'])}</td><td class=\"num group-left\">{w['低价leads']}</td><td class=\"num wd-rate-dim\">{fmt_pct1(w['低价率'])}</td><td class=\"num wd-zz\">{w['转正']}</td></tr>"

def week_summary(wks):
    sjw = sum(w['加微'] for w in wks)
    s0 = sum(w['0元leads'] for w in wks)
    s1 = sum(w['低价leads'] for w in wks)
    szz = sum(w['转正'] for w in wks)
    return f"<tr class=\"summary-row\"><td><strong>{len(wks)}周合计</strong></td><td class=\"num\"><strong>{num(sjw)}</strong></td><td class=\"num\"><strong>{num(s0)}</strong></td><td class=\"num wd-rate group-right\"><strong>{fmt_pct1(s0/sjw*100)}</strong></td><td class=\"num group-left\"><strong>{num(s1)}</strong></td><td class=\"num wd-rate-dim\"><strong>{fmt_pct1(s1/sjw*100)}</strong></td><td class=\"num wd-zz\"><strong>{szz}</strong></td></tr>"

def week5_row(w):
    return f"<tr><td>{w['周范围']}</td><td class=\"num\">{num(w['加微'])}</td><td class=\"num\">{num(w['0元leads'])}</td><td class=\"num wd-rate\">{fmt_pct1(w['0元率'])}</td><td class=\"num wd-zz group-right\">{w['0元转正']}</td><td class=\"num group-left\">{num(w['低价leads'])}</td><td class=\"num wd-rate-dim\">{fmt_pct1(w['低价率'])}</td><td class=\"num wd-zz\">{w['低价转正']}</td></tr>"

def month_row(m):
    return f"<tr><td>{m['月份']}</td><td class=\"num\">{num(m['加微'])}</td><td class=\"num\">{num(m['0元leads'])}</td><td class=\"num wd-rate\">{fmt_pct1(m['0元率'])}</td><td class=\"num wd-zz group-right\">{m['0元转正']}</td><td class=\"num group-left\">{num(m['低价leads'])}</td><td class=\"num wd-rate-dim\">{fmt_pct1(m['低价率'])}</td><td class=\"num wd-zz\">{m['低价转正']}</td></tr>"

def split_rows(rows):
    return '\n'.join(rows)

version_tables = '\n'.join([
    version_table('三味版', versions['三味版'], 'sw', '📖'),
    version_table('公版薄版', versions['公版薄版'], 'gb', '📘'),
    version_table('蓝鲸版', versions['蓝鲸版'], 'lj', '🐋'),
])

version_rows = [
    f"<tr><td><span class=\"badge badge-blue\">三味版</span></td><td class=\"num\">{num(versions['三味版']['leads合计'])}</td><td class=\"num\">{fmt_pct1(versions['三味版']['leads合计']/overall['累计leads']*100)}</td><td class=\"num wd-zz\">{versions['三味版']['转正']}</td></tr>",
    f"<tr><td><span class=\"badge badge-amber\">公版薄版</span></td><td class=\"num\">{num(versions['公版薄版']['leads合计'])}</td><td class=\"num\">{fmt_pct1(versions['公版薄版']['leads合计']/overall['累计leads']*100)}</td><td class=\"num wd-zz\">{versions['公版薄版']['转正']}</td></tr>",
    f"<tr><td><span class=\"badge badge-purple\">蓝鲸版</span></td><td class=\"num\">{num(versions['蓝鲸版']['leads合计'])}</td><td class=\"num\">{fmt_pct1(versions['蓝鲸版']['leads合计']/overall['累计leads']*100)}</td><td class=\"num wd-zz\">{versions['蓝鲸版']['转正']}</td></tr>",
    f"<tr><td>未分版热点</td><td class=\"num\">{num(versions['未分版热点']['leads合计'])}</td><td class=\"num\">{fmt_pct1(versions['未分版热点']['leads合计']/overall['累计leads']*100)}</td><td class=\"num wd-zz\">{versions['未分版热点']['转正']}</td></tr>",
    f"<tr style=\"background:#f8fafc;\"><td><strong>合计</strong></td><td class=\"num\"><strong>{num(overall['累计leads'])}</strong></td><td class=\"num\"><strong>100%</strong></td><td class=\"num wd-zz\"><strong>{overall['累计转正']}</strong></td></tr>",
]

course_rows = [
    f"<tr><td><span class=\"badge badge-blue\">0元爆品课</span></td><td class=\"num\">{num(cumul_course['0元leads'])}</td><td class=\"num\">{fmt_pct1(cumul_course['0元leads']/overall['累计leads']*100)}</td><td class=\"num wd-zz\">{cumul_course['0元转正']}</td></tr>",
    f"<tr><td><span class=\"badge badge-amber\">低价爆品课</span></td><td class=\"num\">{num(cumul_course['低价leads'])}</td><td class=\"num\">{fmt_pct1(cumul_course['低价leads']/overall['累计leads']*100)}</td><td class=\"num wd-zz\">{cumul_course['低价转正']}</td></tr>",
    f"<tr style=\"background:#f8fafc;\"><td><strong>合计</strong></td><td class=\"num\"><strong>{num(overall['累计leads'])}</strong></td><td class=\"num\"><strong>100%</strong></td><td class=\"num wd-zz\"><strong>{overall['累计转正']}</strong></td></tr>",
]

source_rows = [
    f"<tr><td><span class=\"badge badge-blue\">热点析出</span></td><td class=\"num\">{num(structure['热点']['leads'])}</td><td class=\"num\">{fmt_pct1(structure['热点']['leads']/overall['累计leads']*100)}</td><td class=\"num wd-zz\">{structure['热点']['转正']}</td></tr>",
    f"<tr><td><span class=\"badge badge-amber\">索粉析出</span></td><td class=\"num\">{num(structure['索粉']['leads'])}</td><td class=\"num\">{fmt_pct1(structure['索粉']['leads']/overall['累计leads']*100)}</td><td class=\"num wd-zz\">{structure['索粉']['转正']}</td></tr>",
    f"<tr style=\"background:#f8fafc;\"><td><strong>合计</strong></td><td class=\"num\"><strong>{num(overall['累计leads'])}</strong></td><td class=\"num\"><strong>100%</strong></td><td class=\"num wd-zz\"><strong>{overall['累计转正']}</strong></td></tr>",
]

# 热点TOP：热点名称（计划都是笨笨熊）
hotspot_rows = [f"<tr><td>{i+1}</td><td>{h['name']}</td><td class=\"num\">{num(h['leads'])}</td><td class=\"num wd-zz\">{h['转正']}</td></tr>" for i, h in enumerate(top_hot)]

# 索粉TOP：计划名称
sfsource_rows = [f"<tr><td>{i+1}</td><td>{s['name']}</td><td class=\"num\">{num(s['leads'])}</td><td class=\"num wd-zz\">{s['转正']}</td></tr>" for i, s in enumerate(top_sf)]

# 资源位拆分行
res_rows = [f"<tr><td>{r['name']}</td><td class=\"num\">{num(r['加微'])}</td><td class=\"num\">{fmt_pct1(r['加微']/overall['累计加微']*100)}</td></tr>" for r in top_res]

# 年级拆分行
gd_rows = [f"<tr><td>{g['name']}</td><td class=\"num\">{num(g['加微'])}</td><td class=\"num\">{fmt_pct1(g['加微']/overall['累计加微']*100)}</td></tr>" for g in top_grade]

# 关键发现
total_leads_all = overall['累计leads']
total_zz = overall['累计转正']
version_zz = {k: v['转正'] for k, v in versions.items()}
key_findings = (
    f"• 累计{num(total_leads_all)}条leads，转正{total_zz}人<br>"
    f"• 三味版转正{version_zz.get('三味版',0)}人、公版薄版{version_zz.get('公版薄版',0)}人、蓝鲸版{version_zz.get('蓝鲸版',0)}人<br>"
    f"• 0元课转正率（{fmt_pct(cumul_course['0元转正率'])}）vs 低价课转正率（{fmt_pct(cumul_course['低价转正率'])}）"
)

this_label = '06/24-06/30'
last_label = '06/17-06/23'

subst = {
    'this_week_label': this_label,
    'last_week_label': last_label,
    'this_5w_label': f"{weeks5_detail[0]['周范围']} ~ {weeks5_detail[-1]['周范围']}" if weeks5_detail else '',
    'cutoff': data['meta']['数据截止'],
    'this_jw': num(recent_overall['近一周加微']),
    'last_jw': num(lw_jw),
    'jw_hb': hb_str(recent_overall['加微环比']),
    'jw_hb_class': hb_class(recent_overall['加微环比']),
    'this_leads': num(recent_overall['近一周leads']),
    'last_leads': num(lw_leads),
    'leads_hb': hb_str(recent_overall['leads环比']),
    'leads_hb_class': hb_class(recent_overall['leads环比']),
    'this_jk': num(recent_overall['近一周结课']),
    'last_jk': num(lw_jk),
    'jk_hb': '--',
    'jk_hb_class': 'flat',
    'this_zz': num(recent_overall['近一周转正']),
    'last_zz': num(lw_zz),
    'zz_hb': hb_str(recent_overall['转正环比']),
    'zz_hb_class': hb_class(recent_overall['转正环比']),
    'this_lk': fmt_pct1(recent_overall['近一周领课率']),
    'last_lk': fmt_pct1(lw_lk),
    'lk_hb': hb_str(recent_overall['领课率环比']),
    'lk_hb_class': hb_class(recent_overall['领课率环比']),
    'cum_lk': fmt_pct1(overall['领课率']),
    'this_jkzz': fmt_pct1(recent_overall['近一周转正率']),
    'last_jkzz': fmt_pct1(lw_jkzz),
    'jkzz_hb': hb_str(recent_overall['转正率环比']),
    'jkzz_hb_class': hb_class(recent_overall['转正率环比']),
    'cum_jkzz': fmt_pct1(overall['结课转正率']),
    'cum_jw': num(overall['累计加微']),
    'cum_leads': num(overall['累计leads']),
    'cum_jk': num(overall['累计结课']),
    'cum_zz': overall['累计转正'],
    'version_tables': version_tables,
    'unversioned_leads': num(versions['未分版热点']['leads合计']),
    'unversioned_zz': versions['未分版热点']['转正'],
    'week_count': len(weeks_detail),
    'week_rows': '\n'.join([week_row(w) for w in weeks_detail] + [week_summary(weeks_detail)]),
    'week5_count': len(weeks5_detail),
    'week5_rows': '\n'.join([week5_row(w) for w in weeks5_detail]),
    'month_rows': '\n'.join([month_row(m) for m in months_detail]),
    'cum_0_leads': num(cumul_course['0元leads']),
    'cum_0_rate': fmt_pct(cumul_course['0元率']),
    'cum_0_zz': cumul_course['0元转正'],
    'cum_0_zz_rate': fmt_pct(cumul_course['0元转正率']),
    'cum_1_leads': num(cumul_course['低价leads']),
    'cum_1_rate': fmt_pct(cumul_course['低价率']),
    'cum_1_zz': cumul_course['低价转正'],
    'cum_1_zz_rate': fmt_pct(cumul_course['低价转正率']),
    'hot_leads': num(structure['热点']['leads']),
    'hot_pct': fmt_pct1(structure['热点']['leads']/overall['累计leads']*100),
    'hot_zz': structure['热点']['转正'],
    'sf_leads': num(structure['索粉']['leads']),
    'sf_pct': fmt_pct1(structure['索粉']['leads']/overall['累计leads']*100),
    'sf_zz': structure['索粉']['转正'],
    'key_findings': key_findings,
    'version_rows': split_rows(version_rows),
    'course_rows': split_rows(course_rows),
    'source_rows': split_rows(source_rows),
    'res_rows': split_rows(res_rows),
    'gd_rows': split_rows(gd_rows),
    'top_hot_rows': split_rows(hotspot_rows),
    'top_sf_rows': split_rows(sfsource_rows),
    'chart_weeks': ', '.join(["'" + w['周开始'] + "'" for w in weeks_detail]),
    'chart_jw': ', '.join([str(w['加微']) for w in weeks_detail]),
    'chart_rate0': ', '.join([str(w['0元率']) for w in weeks_detail]),
    'chart_rate1': ', '.join([str(w['低价率']) for w in weeks_detail]),
}

html_content = html_template
for k, v in subst.items():
    html_content = html_content.replace(f'{{{{{k}}}}}', str(v))

with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
    f.write(html_content)

# ===================== 打印摘要 =====================
sys.stdout.reconfigure(encoding='utf-8')

print("=" * 50)
print(f"[OK] 笨笨熊下册周报已生成")
print(f"   JSON文件: {OUTPUT_JSON}")
print(f"   HTML文件: {OUTPUT_HTML}")
print(f"   数据截止: 2026-06-30")
print(f"   累计加微: {overall['累计加微']:,}")
print(f"   累计leads: {overall['累计leads']:,}")
print(f"   累计转正: {overall['累计转正']}")
print(f"   领课率: {fmt_pct(overall['领课率'])}")
print(f"   结课转正率: {fmt_pct(overall['结课转正率'])}")
print(f"   周明细: {len(weeks_detail)}周 | 本周({this_label}): 加微={this_week['加微']} leads={this_week['leads']}")
print("--- 版本拆分 ---")
for ver in ['三味版', '公版薄版', '蓝鲸版']:
    v = versions[ver]
    print(f"   {ver}: 热点={v['热点leads']} 索粉={v['索粉leads']} leads={v['leads合计']} 转正={v['转正']}")
print(f"   未分版热点: leads={other_hot} 转正={other_zz}")
print("--- 资源位 ---")
for r in top_res:
    print(f"   {r['name']}: 加微={r['加微']:,}")
print("--- 年级 ---")
for g in top_grade:
    print(f"   {g['name']}: 加微={g['加微']:,}")
print("=" * 50)
