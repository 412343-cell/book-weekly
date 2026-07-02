"""
统一渠道周报生成器 v3 — 修复分隔线+月度拆分+热点名称
用法: python generate_channels.py [xiaochen|bbx_up|yangguang|hanzhijian|all]
"""
import pandas as pd
import json
import sys
import re
from pathlib import Path
from datetime import datetime, timedelta

EXCEL = "D:/Desktop/0630.xlsx"
WEEK_N = 5
TEMPLATE = "channel_template.html"

def safe_int(v, default=0):
    try: return int(float(v)) if pd.notna(v) else default
    except: return default

def safe_float(v, default=0.0):
    try: return float(v) if pd.notna(v) else default
    except: return default

def pct_str(a, b):
    if not b or b <= 0: return '--'
    return f"{a/b*100:.2f}%"

def pct_f1(a, b):
    if not b or b <= 0: return '--'
    return f"{a/b*100:.1f}%"

def pct_val(a, b):
    if not b or b <= 0: return 0.0
    return round(a/b*100, 2)

def num(v): return f"{v:,}"

def hb_calc(curr, prev):
    if prev is None or prev == 0:
        return 0, 'flat', '--'
    chg = round((curr - prev) / prev * 100, 1)
    if chg == 0: return chg, 'flat', '--'
    arrow = '↑' if chg > 0 else '↓'
    return chg, ('up' if chg > 0 else 'down'), f"{arrow}{abs(chg)}%"

def extract_grade(name_str):
    s = str(name_str)
    for n in ['一', '二', '三', '四', '五', '六']:
        if f'{n}年级' in s or f'语文{n}' in s or f'数学{n}' in s or f'英语{n}' in s:
            return f'{n}年级'
    return '未知'

def extract_subject(name_str):
    s = str(name_str)
    for subj in ['语文', '数学', '英语']:
        if subj in s: return subj
    return '未知'

# ===================== 渠道配置 =====================
CHANNELS = {
    'xiaochen': {
        'sheet_key': '小晨同学',
        'title': '小晨同学',
        'subtitle': '40篇童话故事速记1000个单词',
        'output_json': 'xiaochen_data.json',
        'output_html': '小晨同学周报.html',
        'hot_plan': '小晨同学',
        'hot_units': ['40篇童话故事速记1000个单词'],
        'sf_spus_all': ['小晨同学童话故事40篇'],
        'hm_filter': lambda df: df[df['图书'] == '40篇童话故事速记1000个单词'],
        'need_grade': True,
        'need_subject': False,
        'hot_top_label': '热点名称',
        'pins': [{
            'name': '40篇童话故事',
            'hot_units': ['40篇童话故事速记1000个单词'],
            'sf_spus': ['小晨同学童话故事40篇'],
            'hm_filter': lambda df: df[df['图书'] == '40篇童话故事速记1000个单词'],
        }],
    },
    'bbx_up': {
        'sheet_key': '蓝鲸',
        'title': '笨笨熊上册（蓝鲸）',
        'subtitle': '快乐读书吧·上册',
        'output_json': 'bbx_up_data.json',
        'output_html': '笨笨熊上册周报.html',
        'hot_plan': '蓝鲸',
        'hot_units': ['快乐读书吧-上册'],
        'sf_spus_all': ['蓝鲸快乐读书吧上册'],
        'hm_filter': lambda df: df[(df['图书'].str.contains('上册', na=False)) & 
                                    (df['活码名称'].str.contains('笨笨熊', na=False))],
        'need_grade': True,
        'need_subject': False,
        'hot_top_label': '热点名称',
        'pins': [{
            'name': '快乐读书吧-上册',
            'hot_units': ['快乐读书吧-上册'],
            'sf_spus': ['蓝鲸快乐读书吧上册'],
            'hm_filter': lambda df: df[(df['图书'].str.contains('上册', na=False)) & 
                                        (df['活码名称'].str.contains('笨笨熊', na=False))],
        }],
    },
    'yangguang': {
        'sheet_key': '阳光同学',
        'title': '阳光同学',
        'subtitle': '计算小达人 + 提优训练',
        'output_json': 'yangguang_data.json',
        'output_html': '阳光同学周报.html',
        'hot_plan': '阳光同学',
        'hot_units': ['阳光同学计算小达人', '阳光同学提优训练'],
        'sf_spus_all': ['阳光同学计算小达人上册', '阳光同学语文提优训练', '阳光同学数学提优训练', '阳光同学英语提优训练'],
        'hm_filter': lambda df: df[(df['图书'].str.contains('计算小达人|提优训练', na=False))],
        'need_grade': True,
        'need_subject': True,
        'pins': [
            {
                'name': '计算小达人',
                'hot_units': ['阳光同学计算小达人'],
                'sf_spus': ['阳光同学计算小达人上册'],
                'hm_filter': lambda df: df[df['图书'].str.contains('计算小达人', na=False)],
            },
            {
                'name': '提优训练',
                'hot_units': ['阳光同学提优训练'],
                'sf_spus': ['阳光同学语文提优训练', '阳光同学数学提优训练', '阳光同学英语提优训练'],
                'hm_filter': lambda df: df[df['图书'].str.contains('提优训练', na=False)],
            },
        ],
    },
    'hanzhijian': {
        'sheet_key': '汉知简',
        'title': '汉知简',
        'subtitle': '学霸笔记 + 同步作文',
        'output_json': 'hanzhijian_data.json',
        'output_html': '汉知简周报.html',
        'hot_plan': '汉知简',
        'hot_units_include': ['学霸笔记下册', '同步作文'],
        'hot_units_exclude': ['上册', '私域'],
        'sf_spus_all': ['汉知简学霸笔记语文', '汉知简学霸笔记数学', '汉知简学霸笔记英语'],
        'hm_filter': lambda df: df[(df['图书'].str.contains('学霸笔记|同步作文', na=False))],
        'need_grade': True,
        'need_subject': True,
        'pins': [
            {
                'name': '学霸笔记',
                'hot_unit_include': ['学霸笔记下册'],
                'sf_spus': ['汉知简学霸笔记语文', '汉知简学霸笔记数学', '汉知简学霸笔记英语'],
                'hm_filter': lambda df: df[df['图书'].str.contains('学霸笔记', na=False)],
            },
            {
                'name': '同步作文',
                'hot_unit_include': ['同步作文'],
                'sf_spus': [],
                'hm_filter': lambda df: df[df['图书'].str.contains('同步作文', na=False)],
            },
        ],
    },
}


# ===================== 辅助：leads按周分桶 =====================
def week_bucket(dt, week_starts):
    """将日期分配到最近的周开始时间桶"""
    if pd.isna(dt): return None
    for ws in week_starts:
        wend = ws + timedelta(days=6)
        if ws <= dt <= wend:
            return ws
    return None

def get_week_boundaries(week_data, n_weeks):
    """从分渠道周数据获取前n周的周开始时间列表（升序）"""
    wks = week_data.sort_values('周开始时间', ascending=False).head(n_weeks)
    starts = sorted([pd.Timestamp(w) for w in wks['周开始时间']])
    return starts


# ===================== 主生成函数 =====================
def generate_channel(ch_key):
    cfg = CHANNELS[ch_key]
    print(f"\n{'='*60}")
    print(f"生成: {cfg['title']}")

    # ========== 读取数据 ==========
    df_total = pd.read_excel(EXCEL, sheet_name='分渠道总数据')
    df_week = pd.read_excel(EXCEL, sheet_name='分渠道周数据')
    df_month = pd.read_excel(EXCEL, sheet_name='分渠道月数据')
    df_leads = pd.read_excel(EXCEL, sheet_name='leads明细')
    df_hm = pd.read_excel(EXCEL, sheet_name='活码加微明细1')

    # 字段反转
    df_leads['结课人次'] = df_leads['leads转正人次']
    df_leads['转正人次'] = df_leads['leads结课人次']

    # ========== 整体累计数据 ==========
    total_row = df_total[df_total['渠道'] == cfg['sheet_key']].iloc[0]
    c_jw = safe_int(total_row['加微信号人数'])
    c_ld = safe_int(total_row['leads人次'])
    c_jk = safe_int(total_row['结课leads人次'])
    c_zz = safe_int(total_row['结课转正人次'])

    # ========== 整体周数据（本周/上周） ==========
    week_full = df_week[df_week['渠道'] == cfg['sheet_key']].sort_values('周开始时间', ascending=False)
    recent_5 = week_full.head(WEEK_N).iloc[::-1]  # 升序，最近5周

    weeks_for_table = []
    for _, row in recent_5.iterrows():
        s = pd.Timestamp(row['周开始时间'])
        jw = safe_int(row['加微信号人数'])
        ld = safe_int(row['leads人次'])
        jk = safe_int(row['结课leads人次'])
        zz = safe_int(row['结课转正人次'])
        weeks_for_table.append({
            '范围': f"{s.strftime('%m/%d')}-{(s+timedelta(days=6)).strftime('%m/%d')}",
            '加微': jw, 'leads': ld, '结课': jk, '转正': zz,
        })

    this_idx = len(weeks_for_table) - 2  # 倒数第2个=本周
    last_idx = len(weeks_for_table) - 3  # 倒数第3个=上周
    if this_idx < 0: this_idx = 0
    if last_idx < 0: last_idx = 0

    tw = weeks_for_table[this_idx]
    lw = weeks_for_table[last_idx] if last_idx != this_idx else {'加微': 0, 'leads': 0, '结课': 0, '转正': 0}

    this_label = tw['范围']
    last_label = lw['范围']

    def overall_row(curr, prev):
        _, c, s = hb_calc(curr, prev)
        return {'curr': num(curr), 'prev': num(prev), 'hb_c': c, 'hb': s}

    # 构建整体替换变量
    repl = {}
    repl['title'] = cfg['title']
    repl['this_label'] = this_label
    repl['last_label'] = last_label
    repl['cutoff'] = '2026-06-30'

    for k, (c, p) in [('jw', (tw['加微'], lw['加微'])),
                        ('ld', (tw['leads'], lw['leads'])),
                        ('jk', (tw['结课'], lw['结课'])),
                        ('zz', (tw['转正'], lw['转正']))]:
        r = overall_row(c, p)
        repl[f't_{k}'] = r['curr']
        repl[f't_{k}_l'] = r['prev']
        repl[f't_{k}_hb_c'] = r['hb_c']
        repl[f't_{k}_hb'] = r['hb']

    # 比率
    for k, cfn in [('lk', lambda w: pct_val(w['leads'], w['加微'])),
                     ('jz', lambda w: pct_val(w['转正'], w['结课']))]:
        cv, pv = cfn(tw), cfn(lw)
        chg, hc, hs = hb_calc(cv, pv)
        repl[f't_{k}'] = f"{cv:.1f}%"
        repl[f't_{k}_l'] = f"{pv:.1f}%"
        repl[f't_{k}_hb_c'] = hc
        repl[f't_{k}_hb'] = hs

    # 累计
    repl['c_jw'] = num(c_jw)
    repl['c_ld'] = num(c_ld)
    repl['c_jk'] = num(c_jk)
    repl['c_zz'] = num(c_zz)
    repl['c_lk'] = pct_f1(c_ld, c_jw)
    repl['c_jz'] = pct_f1(c_zz, c_jk)

    # ========== 按品数据 ==========
    # 获取周边界
    week_starts = get_week_boundaries(week_full, WEEK_N)

    # 给leads加周标记
    df_leads['周'] = df_leads['leads购课日期'].apply(lambda d: week_bucket(pd.Timestamp(d) if pd.notna(d) else None, week_starts))

    pins_data = []
    chart_scripts = []

    for pi, pin_cfg in enumerate(cfg['pins']):
        pname = pin_cfg['name']
        print(f"  品: {pname}")

        # ---- 热点筛选 ----
        if 'hot_units' in pin_cfg:
            hot = df_leads[(df_leads['leads计划名称'] == cfg['hot_plan']) &
                           (df_leads['leads单元名称'].isin(pin_cfg['hot_units']))]
        elif 'hot_unit_include' in pin_cfg:
            m = df_leads['leads计划名称'] == cfg['hot_plan']
            for inc in pin_cfg['hot_unit_include']:
                m = m & df_leads['leads单元名称'].str.contains(inc, na=False)
            if 'hot_units_exclude' in cfg:
                for exc in cfg['hot_units_exclude']:
                    m = m & ~df_leads['leads单元名称'].str.contains(exc, na=False)
            hot = df_leads[m]
        else:
            hot = df_leads.iloc[:0]

        # ---- 索粉筛选 ----
        sf_spus = pin_cfg.get('sf_spus', [])
        if sf_spus:
            sf = df_leads[df_leads['leads图书SPU名称'].isin(sf_spus)]
        else:
            sf = df_leads.iloc[:0]

        # ---- 品级累计 ----
        hot_ld = safe_int(hot['leads人次'].sum())
        hot_jk = safe_int(hot['结课人次'].sum())
        hot_zz = safe_int(hot['转正人次'].sum())
        sf_ld = safe_int(sf['leads人次'].sum())
        sf_jk = safe_int(sf['结课人次'].sum())
        sf_zz = safe_int(sf['转正人次'].sum())
        pin_ld = hot_ld + sf_ld
        pin_jk = hot_jk + sf_jk
        pin_zz = hot_zz + sf_zz

        # 活码加微
        hm_pin = pin_cfg.get('hm_filter', lambda d: d.iloc[:0])(df_hm)
        pin_jw = safe_int(hm_pin['当日加微人数(包含删除好友)'].sum())

        # ---- 按周聚合（leads层级，非预聚合表） ----
        pin_weeks = []
        for ws in week_starts:
            wend = ws + timedelta(days=6)
            hot_w = hot[(hot['周'] == ws)]
            sf_w = sf[(sf['周'] == ws)]
            # 加微用活码日期聚合
            hm_w = hm_pin[(pd.to_datetime(hm_pin['加微日期']) >= ws) & (pd.to_datetime(hm_pin['加微日期']) <= wend)]
            jw_w = safe_int(hm_w['当日加微人数(包含删除好友)'].sum())

            hot_ld_w = safe_int(hot_w['leads人次'].sum())
            hot_zz_w = safe_int(hot_w['转正人次'].sum())
            sf_ld_w = safe_int(sf_w['leads人次'].sum())
            sf_zz_w = safe_int(sf_w['转正人次'].sum())
            ld_w = hot_ld_w + sf_ld_w
            zz_w = hot_zz_w + sf_zz_w

            pin_weeks.append({
                '周范围': f"{ws.strftime('%m/%d')}-{wend.strftime('%m/%d')}",
                '加微': jw_w,
                '0元leads': hot_ld_w,
                '0元率': pct_val(hot_ld_w, jw_w),
                '0元转正': hot_zz_w,
                '低价leads': sf_ld_w,
                '低价率': pct_val(sf_ld_w, jw_w),
                '低价转正': sf_zz_w,
                '热点': hot_ld_w,
                '索粉': sf_ld_w,
                '转正': zz_w,
            })

        # 本周/上周（品级）
        pw_this = pin_weeks[this_idx] if this_idx < len(pin_weeks) else {}
        pw_last = pin_weeks[last_idx] if last_idx < len(pin_weeks) else {}

        def pw_row(field, fmt='num'):
            c = pw_this.get(field, 0)
            p = pw_last.get(field, 0)
            if fmt == 'pct':
                r = overall_row(c, p)
                return {'curr': f"{c:.1f}%", 'prev': f"{p:.1f}%", 'hb_c': r['hb_c'], 'hb': r['hb']}
            else:
                r = overall_row(c, p)
                return {'curr': num(c), 'prev': num(p), 'hb_c': r['hb_c'], 'hb': r['hb']}

        pd_row = {}
        for fld in ['加微', 'leads', '结课', '转正']:
            r = pw_row(fld)
            pd_row[fld] = r
        for fld in ['领课率']:
            r = pw_row(fld, 'pct')
            pd_row[fld] = r
        for fld in ['转正率']:
            r = pw_row(fld, 'pct')
            pd_row[fld] = r

        # ---- 按周明细表（8列：周/加微/0元leads/0元率/0元转正/低价leads/低价率/低价转正）----
        # ⚠️ 分隔线规则：group-right 在 0元转正（0元组最右列），group-left 在 低价leads（低价组最左列）
        week_rows = []
        for w in pin_weeks:
            week_rows.append(
                f"<tr><td>{w['周范围']}</td>"
                f"<td class=\"num\">{num(w['加微'])}</td>"
                f"<td class=\"num\">{num(w['0元leads'])}</td>"
                f"<td class=\"num wd-rate\">{w['0元率']:.1f}%</td>"
                f"<td class=\"num wd-zz group-right\">{w['0元转正']}</td>"
                f"<td class=\"num group-left\">{num(w['低价leads'])}</td>"
                f"<td class=\"num wd-rate-dim\">{w['低价率']:.1f}%</td>"
                f"<td class=\"num wd-zz\">{w['低价转正']}</td></tr>"
            )
        # 合计行
        s0ld = sum(w['0元leads'] for w in pin_weeks)
        s0zz = sum(w['0元转正'] for w in pin_weeks)
        s1ld = sum(w['低价leads'] for w in pin_weeks)
        s1zz = sum(w['低价转正'] for w in pin_weeks)
        s_jw = sum(w['加微'] for w in pin_weeks)
        week_rows.append(
            f"<tr class=\"summary-row\"><td><strong>{len(pin_weeks)}周合计</strong></td>"
            f"<td class=\"num\"><strong>{num(s_jw)}</strong></td>"
            f"<td class=\"num\"><strong>{num(s0ld)}</strong></td>"
            f"<td class=\"num wd-rate\"><strong>{pct_val(s0ld,s_jw):.1f}%</strong></td>"
            f"<td class=\"num wd-zz group-right\"><strong>{s0zz}</strong></td>"
            f"<td class=\"num group-left\"><strong>{num(s1ld)}</strong></td>"
            f"<td class=\"num wd-rate-dim\"><strong>{pct_val(s1ld,s_jw):.1f}%</strong></td>"
            f"<td class=\"num wd-zz\"><strong>{s1zz}</strong></td></tr>"
        )

        # ---- Chart data ----
        chart_labels = ', '.join(["'" + w['周范围'][:5] + "'" for w in pin_weeks])
        chart_jw = ', '.join([str(w['加微']) for w in pin_weeks])
        chart_r0 = ', '.join([str(w['0元率']) for w in pin_weeks])
        chart_r1 = ', '.join([str(w['低价率']) for w in pin_weeks])

        bar_color = 'rgba(59,130,246,0.7)' if pi == 0 else 'rgba(217,119,6,0.7)'
        chart_id = f'pin{pi}Chart'
        chart_scripts.append(f"""
// {pname}
(function() {{
  var ctx = document.getElementById('{chart_id}');
  if (!ctx) return;
  new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: [{chart_labels}],
      datasets: [
        {{
          label: '加微数',
          type: 'bar',
          data: [{chart_jw}],
          backgroundColor: '{bar_color}',
          borderRadius: 4,
          yAxisID: 'y',
          order: 2
        }},
        {{
          label: '0元领课率',
          type: 'line',
          data: [{chart_r0}],
          borderColor: '#2563eb',
          fill: false,
          tension: 0.35,
          pointRadius: 4,
          pointBackgroundColor: '#2563eb',
          pointBorderColor: '#fff',
          pointBorderWidth: 2,
          yAxisID: 'y1',
          order: 1
        }},
        {{
          label: '低价领课率',
          type: 'line',
          data: [{chart_r1}],
          borderColor: '#d97706',
          fill: false,
          tension: 0.35,
          pointRadius: 4,
          pointBackgroundColor: '#d97706',
          pointBorderColor: '#fff',
          pointBorderWidth: 2,
          borderDash: [5,3],
          yAxisID: 'y1',
          order: 1
        }}
      ]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      interaction: {{ mode: 'index', intersect: false }},
      plugins: {{
        legend: {{ display: true, position: 'top', labels: {{ usePointStyle: true, padding: 20, font: {{ size: 11 }} }} }},
        tooltip: {{ callbacks: {{ label: ctx => ctx.dataset.label === '加微数' ? '加微: ' + ctx.raw.toLocaleString() : ctx.dataset.label + ': ' + ctx.raw.toFixed(1) + '%' }} }}
      }},
      scales: {{
        x: {{ grid: {{ display: false }} }},
        y: {{ type: 'linear', position: 'left', title: {{ display: true, text: '加微数', font: {{ size: 11 }}, color: '#64748b' }}, beginAtZero: true, grid: {{ color: '#f1f5f9' }} }},
        y1: {{ type: 'linear', position: 'right', title: {{ display: true, text: '领课率', font: {{ size: 11 }}, color: '#64748b' }}, beginAtZero: true, grid: {{ drawOnChartArea: false }}, ticks: {{ callback: v => v + '%' }} }}
      }}
    }}
  }});
}})();
""")

        # ---- TOP热点/索粉 ----
        hot_top = hot.groupby('leads单元名称').agg(
            leads=('leads人次', 'sum'), zz=('转正人次', 'sum')
        ).sort_values('leads', ascending=False).head(10)
        top_hot_rows = []
        for i, (n, r) in enumerate(hot_top.iterrows()):
            top_hot_rows.append(
                f"<tr><td>{i+1}</td><td>{n}</td>"
                f"<td class=\"num\">{num(safe_int(r['leads']))}</td>"
                f"<td class=\"num wd-zz\">{safe_int(r['zz'])}</td></tr>"
            )
        if not top_hot_rows:
            top_hot_rows = ['<tr><td colspan="4">暂无数据</td></tr>']

        sf_top = sf.groupby('leads计划名称').agg(
            leads=('leads人次', 'sum'), zz=('转正人次', 'sum')
        ).sort_values('leads', ascending=False).head(10)
        top_sf_rows = []
        for i, (n, r) in enumerate(sf_top.iterrows()):
            top_sf_rows.append(
                f"<tr><td>{i+1}</td><td>{n}</td>"
                f"<td class=\"num\">{num(safe_int(r['leads']))}</td>"
                f"<td class=\"num wd-zz\">{safe_int(r['zz'])}</td></tr>"
            )
        if not top_sf_rows:
            top_sf_rows = ['<tr><td colspan="4">暂无数据</td></tr>']

        # ---- 累计结构卡片 ----
        lk_rate = pct_val(pin_ld, pin_jw)
        zz_rate = pct_val(pin_zz, pin_jk)
        hot_zz_rate = pct_val(hot_zz, hot_jk)
        sf_zz_rate = pct_val(sf_zz, sf_jk)
        hot_pct = pct_val(hot_ld, pin_ld) if pin_ld else 0
        sf_pct = pct_val(sf_ld, pin_ld) if pin_ld else 0

        lk_line = f"领课率 {lk_rate:.2f}%，结课转正率 {zz_rate:.2f}%"
        if pi == 0:
            findings = f"• 累计{num(pin_ld)}条leads，转正{pin_zz}人<br>• 热点析出{num(hot_ld)}条({hot_pct:.1f}%)，索粉析出{num(sf_ld)}条({sf_pct:.1f}%)<br>• 0元转正率{pct_f1(hot_zz, hot_jk)} vs 低价转正率{pct_f1(sf_zz, sf_jk)}"
        else:
            findings = f"• 累计{num(pin_ld)}条leads，转正{pin_zz}人"

        pins_data.append({
            'name': pname,
            'bar_class': 'kantu' if pi == 0 else 'quwei',
            'cum_ld': num(pin_ld),
            'cum_zz': num(pin_zz),
            'week_count': len(pin_weeks),
            'week_rows': '\n'.join(week_rows),
            'chart_id': chart_id,
            'hot_top_label': cfg.get('hot_top_label', '热点名称'),
            'cumul': {
                'jw': num(pin_jw),
                'ld': num(pin_ld),
                'hot_ld': num(hot_ld), 'hot_zz': hot_zz, 'hot_zzl': pct_f1(hot_zz, hot_jk),
                'sf_ld': num(sf_ld), 'sf_zz': sf_zz, 'sf_zzl': pct_f1(sf_zz, sf_jk),
                'hot_pct': f"{hot_pct:.1f}%", 'sf_pct': f"{sf_pct:.1f}%",
                'lk': f"{lk_rate:.2f}%", 'zz_rate': f"{zz_rate:.2f}%",
                'findings': findings,
            },
            'top_hot_rows': '\n'.join(top_hot_rows),
            'top_sf_rows': '\n'.join(top_sf_rows),
            # 课型拆分
            'course_rows': [
                f"<tr><td><span class=\"badge badge-blue\">0元爆品课（热点）</span></td><td class=\"num\">{num(hot_ld)}</td><td class=\"num\">{pct_f1(hot_ld, pin_ld)}</td><td class=\"num wd-zz\">{hot_zz}</td></tr>",
                f"<tr><td><span class=\"badge badge-amber\">低价爆品课（索粉）</span></td><td class=\"num\">{num(sf_ld)}</td><td class=\"num\">{pct_f1(sf_ld, pin_ld)}</td><td class=\"num wd-zz\">{sf_zz}</td></tr>",
                f"<tr style=\"background:#f8fafc;\"><td><strong>合计</strong></td><td class=\"num\"><strong>{num(pin_ld)}</strong></td><td class=\"num\"><strong>100%</strong></td><td class=\"num wd-zz\"><strong>{pin_zz}</strong></td></tr>",
            ],
            'source_rows': [
                f"<tr><td><span class=\"badge badge-blue\">热点析出</span></td><td class=\"num\">{num(hot_ld)}</td><td class=\"num\">{pct_f1(hot_ld, pin_ld)}</td><td class=\"num wd-zz\">{hot_zz}</td></tr>",
                f"<tr><td><span class=\"badge badge-amber\">索粉析出</span></td><td class=\"num\">{num(sf_ld)}</td><td class=\"num\">{pct_f1(sf_ld, pin_ld)}</td><td class=\"num wd-zz\">{sf_zz}</td></tr>",
                f"<tr style=\"background:#f8fafc;\"><td><strong>合计</strong></td><td class=\"num\"><strong>{num(pin_ld)}</strong></td><td class=\"num\"><strong>100%</strong></td><td class=\"num wd-zz\"><strong>{pin_zz}</strong></td></tr>",
            ],
            # 对外并排表
            'outer_rows': [
                f"<tr><td>加微数</td><td class=\"num\">{pd_row['加微']['curr']}</td><td class=\"num\">{pd_row['加微']['prev']}</td><td class=\"{pd_row['加微']['hb_c']}\">{pd_row['加微']['hb']}</td><td class=\"num cumulative\">{num(pin_jw)}</td></tr>",
                f"<tr><td>leads 领课人次</td><td class=\"num\">{num(pin_weeks[this_idx].get('0元leads',0) + pin_weeks[this_idx].get('低价leads',0))}</td><td class=\"num\">{num(pin_weeks[last_idx].get('0元leads',0) + pin_weeks[last_idx].get('低价leads',0))}</td><td class=\"{pd_row['leads']['hb_c']}\">{pd_row['leads']['hb']}</td><td class=\"num cumulative\">{num(pin_ld)}</td></tr>",
                f"<tr><td>转正人次</td><td class=\"num\">{pin_weeks[this_idx].get('转正',0)}</td><td class=\"num\">{pin_weeks[last_idx].get('转正',0)}</td><td class=\"{pd_row['转正']['hb_c']}\">{pd_row['转正']['hb']}</td><td class=\"num cumulative\">{pin_zz}</td></tr>",
                f"<tr class=\"rate-row\"><td>领课率</td><td class=\"num\">{pct_val(pin_weeks[this_idx].get('0元leads',0) + pin_weeks[this_idx].get('低价leads',0), pin_weeks[this_idx].get('加微',1)):.1f}%</td><td class=\"num\">{pct_val(pin_weeks[last_idx].get('0元leads',0) + pin_weeks[last_idx].get('低价leads',0), max(pin_weeks[last_idx].get('加微',1),1)):.1f}%</td><td class=\"{pd_row['领课率']['hb_c']}\">{pd_row['领课率']['hb']}</td><td class=\"cumulative\">{lk_rate:.2f}%</td></tr>",
            ],
            'col_title_class': 'kt' if pi == 0 else 'qw',
            'icon': '📖' if pi == 0 else '📘',
            # 资源位
            'res_rows': '',
            # 年级
            'gd_rows': '',
            # 学科
            'sj_rows': '',
        })

    # ========== 活码资源位/年级/学科（按品拆分） ==========
    for pi, pin_cfg in enumerate(cfg['pins']):
        d_pin = pins_data[pi]
        hm_pin = pin_cfg.get('hm_filter', lambda df: df.iloc[:0])(df_hm)
        hm_pin_clean = hm_pin[hm_pin['资源位'].notna() & (hm_pin['资源位'] != 0) & (hm_pin['资源位'] != '0')]
        pin_jw = safe_int(hm_pin_clean['当日加微人数(包含删除好友)'].sum())

        # 资源位
        res_g = hm_pin_clean.groupby('资源位')['当日加微人数(包含删除好友)'].sum().sort_values(ascending=False)
        r_rows = []
        for n, v in res_g.items():
            r_rows.append(f"<tr><td>{n}</td><td class=\"num\">{num(safe_int(v))}</td><td class=\"num\">{pct_f1(safe_int(v), pin_jw)}</td></tr>")
        r_rows.append(f"<tr style=\"background:#f8fafc;\"><td><strong>合计</strong></td><td class=\"num\"><strong>{num(pin_jw)}</strong></td><td class=\"num\"><strong>100%</strong></td></tr>")
        d_pin['res_rows'] = '\n'.join(r_rows)

        # 年级
        if cfg['need_grade']:
            hm_gd = hm_pin.copy()
            hm_gd['年级分类'] = hm_gd['活码名称'].apply(extract_grade)
            hm_gd = hm_gd[hm_gd['年级分类'] != '未知']
            gd_g = hm_gd.groupby('年级分类')['当日加微人数(包含删除好友)'].sum().sort_values(ascending=False)
            g_rows = []
            for n, v in gd_g.items():
                g_rows.append(f"<tr><td>{n}</td><td class=\"num\">{num(safe_int(v))}</td><td class=\"num\">{pct_f1(safe_int(v), pin_jw)}</td></tr>")
            d_pin['gd_rows'] = '\n'.join(g_rows) if g_rows else ''

        # 学科
        if cfg['need_subject']:
            hm_sj = hm_pin.copy()
            hm_sj['学科分类'] = hm_sj['活码名称'].apply(extract_subject)
            hm_sj = hm_sj[hm_sj['学科分类'] != '未知']
            sj_g = hm_sj.groupby('学科分类')['当日加微人数(包含删除好友)'].sum().sort_values(ascending=False)
            s_rows = []
            for n, v in sj_g.items():
                s_rows.append(f"<tr><td>{n}</td><td class=\"num\">{num(safe_int(v))}</td><td class=\"num\">{pct_f1(safe_int(v), pin_jw)}</td></tr>")
            d_pin['sj_rows'] = '\n'.join(s_rows) if s_rows else ''

    # ========== 渲染HTML ==========
    html = Path(TEMPLATE).read_text(encoding='utf-8')

    # 整体变量
    for k, v in repl.items():
        html = html.replace('{{' + k + '}}', str(v))

    # 对外按品拆分
    is_multi = len(cfg['pins']) > 1
    if is_multi:
        outer_parts = []
        for d in pins_data:
            outer_parts.append(
                f"<div><div class=\"ext-col-title {d['col_title_class']}\">{d['icon']} {d['name']}</div>"
                f"<table class=\"ext-table\"><thead><tr><th>指标</th><th class=\"this-week\">本周</th><th class=\"last-week\">上周</th><th>环比</th><th>累计</th></tr></thead><tbody>"
                + '\n'.join(d['outer_rows']) +
                f"</tbody></table></div>"
            )
        outer_html = '<div class="table-label">按产品拆分</div><div class="ext-two-col">' + '\n'.join(outer_parts) + '</div>'
    else:
        outer_html = ''
    html = html.replace('{{pin_outer_section}}', outer_html)

    # ========== 内部详析区 ==========
    # 月度数据处理（所有品共享同一个df_month基础）
    def build_month_rows(pin_cfg, ch_key):
        """生成月度明细表行"""
        rows = []
        # 过滤该渠道该品的月度数据
        m_ch = df_month[df_month['渠道'] == ch_key]
        if m_ch.empty:
            return '<tr><td colspan="7">暂无数据</td></tr>'
        # 按月汇总（用leads明细按月聚合更准确）
        df_leads['月份'] = df_leads['leads购课日期'].apply(
            lambda d: pd.Timestamp(d).strftime('%Y-%m') if pd.notna(d) else None
        )
        # 重构月度聚合：按渠道×月×0元/低价
        m_data = []
        for _, row in m_ch.iterrows():
            ym = str(row.get('月份', ''))[:7]
            if not ym or ym == 'nan': continue
            jw = safe_int(row.get('加微信号人数', 0))
            # 月度数据只有总计，用leads明细按月×品补充细分
            ld_hot = safe_int(row.get('leads人次', 0))  # 这里用热点leads做近似
            ld_sf = 0
            # 更准确的做法：从leads明细中按月聚合
            ym_hot = df_leads[(df_leads['leads计划名称'] == cfg['hot_plan']) & (df_leads['月份'] == ym)]
            if 'hot_units' in pin_cfg:
                ym_hot = ym_hot[ym_hot['leads单元名称'].isin(pin_cfg['hot_units'])]
            elif 'hot_unit_include' in pin_cfg:
                m = pd.Series(True, index=ym_hot.index)
                for inc in pin_cfg['hot_unit_include']:
                    m = m | ym_hot['leads单元名称'].str.contains(inc, na=False)
                if 'hot_units_exclude' in cfg:
                    for exc in cfg['hot_units_exclude']:
                        m = m & ~ym_hot['leads单元名称'].str.contains(exc, na=False)
                ym_hot = ym_hot[m]
            sf_spus = pin_cfg.get('sf_spus', [])
            ym_sf = df_leads[(df_leads['月份'] == ym) & (df_leads['leads图书SPU名称'].isin(sf_spus))] if sf_spus else pd.DataFrame()
            ym_hot_m = safe_int(ym_hot['leads人次'].sum()) if not ym_hot.empty else 0
            ld_sf_m = safe_int(ym_sf['leads人次'].sum()) if (not ym_sf.empty and 'leads人次' in ym_sf.columns) else 0
            zz_m = safe_int(ym_hot['转正人次'].sum()) + (safe_int(ym_sf['转正人次'].sum()) if (not ym_sf.empty and '转正人次' in ym_sf.columns) else 0)
            ld_total = ym_hot_m + ld_sf_m
            lk_rate = pct_val(ld_total, jw) if jw else 0
            label = ym[-2:]
            rows.append(
                f"<tr><td>{label}月</td>"
                f"<td class=\"num\">{num(jw)}</td>"
                f"<td class=\"num\">{num(ym_hot_m)}</td>"
                f"<td class=\"num\">{num(ld_sf_m)}</td>"
                f"<td class=\"num\">{num(ld_total)}</td>"
                f"<td class=\"num wd-rate\">{lk_rate:.1f}%</td>"
                f"<td class=\"num wd-zz\">{zz_m}</td></tr>"
            )
        return '\n'.join(rows) if rows else '<tr><td colspan="7">暂无数据</td></tr>'

    # 汉知简同步作文：第2个品（index=1）只显示月度
    is_multi = len(cfg['pins']) > 1
    if is_multi:
        inner_parts = []
        for pi, d in enumerate(pins_data):
            pin_cfg = cfg['pins'][pi]
            # 同步作文降级为月度
            if is_multi and pi == 1 and ch_key == 'hanzhijian':
                m_rows = build_month_rows(pin_cfg, cfg['sheet_key'])
                inner_parts.append(build_pin_section(d, True, show_weekly=False, month_rows=m_rows))
            else:
                inner_parts.append(build_pin_section(d, True, show_weekly=True))
        inner_html = '<div class="int-two-col">' + '\n'.join(inner_parts) + '</div>'
    else:
        inner_html = '<div class="int-section">' + build_pin_section(pins_data[0], False, show_weekly=True) + '</div>'
    html = html.replace('{{pin_inner_section}}', inner_html)

    # 资源位表
    res_tables = []
    for d in pins_data:
        res_tables.append(
            f"<div><div class=\"mini-label\">{d['icon']} {d['name']}（累计 {d['cum_ld']} leads）</div>"
            f"<table class=\"int-table\"><thead><tr><th>资源位</th><th class=\"num\">加微人数</th><th class=\"num\">占比</th></tr></thead><tbody>{d['res_rows']}</tbody></table></div>"
        )
    html = html.replace('{{pin_res_tables}}', '\n'.join(res_tables))

    # 年级表（第一个品的数据展示）
    if cfg['need_grade'] and pins_data[0]['gd_rows']:
        grade_html = f"<div class=\"mini-label\" style=\"margin-top:14px;\">📚 年级拆分（加微维度）</div><div style=\"display:grid;grid-template-columns:{'1fr 1fr' if is_multi else '1fr'};gap:20px;\">"
        for d in pins_data:
            if d['gd_rows']:
                grade_html += f"<div><div style=\"font-size:12px;color:#64748b;margin-bottom:4px;\">{d['name']}</div><table class=\"int-table\"><thead><tr><th>年级</th><th class=\"num\">加微数</th><th class=\"num\">占比</th></tr></thead><tbody>{d['gd_rows']}</tbody></table></div>"
        grade_html += '</div>'
        html = html.replace('{{pin_grade_table}}', grade_html)
    else:
        html = html.replace('{{pin_grade_table}}', '')

    # 学科表
    if cfg['need_subject'] and pins_data[0]['sj_rows']:
        sj_html = f"<div class=\"mini-label\" style=\"margin-top:14px;\">📐 学科拆分（加微维度）</div><div style=\"display:grid;grid-template-columns:{'1fr 1fr' if is_multi else '1fr'};gap:20px;\">"
        for d in pins_data:
            if d['sj_rows']:
                sj_html += f"<div><div style=\"font-size:12px;color:#64748b;margin-bottom:4px;\">{d['name']}</div><table class=\"int-table\"><thead><tr><th>学科</th><th class=\"num\">加微数</th><th class=\"num\">占比</th></tr></thead><tbody>{d['sj_rows']}</tbody></table></div>"
        sj_html += '</div>'
        html = html.replace('{{pin_subject_table}}', sj_html)
    else:
        html = html.replace('{{pin_subject_table}}', '')

    # 月度拆分（整体渠道级别，不按品）
    # 按月从df_month聚合近3月数据
    df_leads['月份'] = df_leads['leads购课日期'].apply(
        lambda d: pd.Timestamp(d).strftime('%Y-%m') if pd.notna(d) else None
    )
    m_ch = df_month[df_month['渠道'] == cfg['sheet_key']].sort_values('月份', ascending=False).head(6)
    month_rows = []
    for _, row in m_ch.iterrows():
        ym = str(row.get('月份', ''))[:7]
        if not ym or ym == 'nan': continue
        jw = safe_int(row.get('加微信号人数', 0))
        ld = safe_int(row.get('leads人次', 0))
        jk = safe_int(row.get('结课leads人次', 0))
        zz = safe_int(row.get('结课转正人次', 0))
        lk_r = pct_val(ld, jw)
        zz_r = pct_val(zz, jk)
        label = ym[-2:]
        month_rows.append(
            f"<tr><td>{label}月</td>"
            f"<td class=\"num\">{num(jw)}</td>"
            f"<td class=\"num\">{num(ld)}</td>"
            f"<td class=\"num wd-rate\">{lk_r:.1f}%</td>"
            f"<td class=\"num\">{num(jk)}</td>"
            f"<td class=\"num wd-zz\">{zz}</td>"
            f"<td class=\"num wd-rate\">{zz_r:.2f}%</td></tr>"
        )
    month_rows_str = '\n'.join(month_rows) if month_rows else '<tr><td colspan="7">暂无数据</td></tr>'
    html = html.replace('{{pin_month_rows}}', month_rows_str)

    # Chart scripts
    html = html.replace('{{chart_scripts}}', '\n'.join(chart_scripts))

    Path(cfg['output_html']).write_text(html, encoding='utf-8')
    print(f"[OK] {cfg['output_html']} 已生成")


def build_pin_section(d, is_multi, show_weekly=True, month_rows=''):
    c = d['cumul']
    # 分隔线规则：group-right 在 0元转正，group-left 在 低价leads
    weekly_section = ''
    if show_weekly:
        weekly_section = f"""
      <div class="mini-label">📅 近{d['week_count']}周周明细（按价位拆分）</div>
      <table class="int-table" style="margin-bottom:14px;">
        <thead class="grouped"><tr>
          <th>周</th>
          <th class="num wd-jw-header">加微</th>
          <th class="num wd-course-header">0元leads</th>
          <th class="num wd-rate-header">0元率</th>
          <th class="num wd-zz group-right">0元转正</th>
          <th class="num wd-course-header group-left">低价leads</th>
          <th class="num wd-rate-header">低价率</th>
          <th class="num wd-zz">低价转正</th>
        </tr></thead>
        <tbody>{d['week_rows']}</tbody>
      </table>
      <div class="chart-wrap-combo"><canvas id="{d['chart_id']}"></canvas></div>
"""
    else:
        weekly_section = f"""
      <div class="mini-label">📅 月度明细（按价位拆分）</div>
      <table class="int-table" style="margin-bottom:8px;">
        <thead><tr><th>月份</th><th class="num wd-jw-header">加微</th><th class="num">0元leads</th><th class="num">低价leads</th><th class="num">合计leads</th><th class="num">leads率</th><th class="num wd-zz">转正</th></tr></thead>
        <tbody>{month_rows}</tbody>
      </table>
"""
    section = f"""
  <div class="pin-section">
    <div class="pin-bar {d['bar_class']}">{d['icon']} {d['name']}<span class="pin-total">累计 {d['cum_ld']} leads ｜ 转正 {d['cum_zz']} 人</span></div>
    <div class="pin-body">
      {weekly_section}

      <div class="cumul-card">
        <div class="cumul-title">📊 累计结构一览（截止 2026/06/30）</div>
        <div class="cumul-row">
          <div class="cumul-item">
            <div class="ci-label">📖 加微 → leads 漏斗</div>
            <div>加微 <span class="ci-val">{c['jw']}</span> 人</div>
            <div>├ 热点析出 <span class="ci-val">{c['hot_ld']}</span> 人 <span class="ci-sub">（转正 {c['hot_zz']} 人 / {c['hot_zzl']}）</span></div>
            <div>└ 索粉析出 <span class="ci-val">{c['sf_ld']}</span> 人 <span class="ci-sub">（转正 {c['sf_zz']} 人 / {c['sf_zzl']}）</span></div>
          </div>
          <div class="cumul-item">
            <div class="ci-label">🔍 来源结构</div>
            <div>热点析出 <span class="ci-val">{c['hot_ld']}</span> <span class="ci-sub">（{c['hot_pct']}，转正 {c['hot_zz']}）</span></div>
            <div>索粉析出 <span class="ci-val">{c['sf_ld']}</span> <span class="ci-sub">（{c['sf_pct']}，转正 {c['sf_zz']}）</span></div>
          </div>
          <div class="cumul-item">
            <div class="ci-label">💡 关键发现</div>
            <div style="font-size:12px;color:#475569;">
              {c['findings']}
            </div>
          </div>
        </div>
      </div>

      <div class="triple-grid">
        <div>
          <div class="mini-label">课型拆分</div>
          <table class="int-table">
            <thead><tr><th>课型</th><th class="num">人次</th><th class="num">占比</th><th class="num wd-zz">转正</th></tr></thead>
            <tbody>
              {''.join(d['course_rows'])}
            </tbody>
          </table>
        </div>
        <div>
          <div class="mini-label">来源拆分</div>
          <table class="int-table">
            <thead><tr><th>来源</th><th class="num">人次</th><th class="num">占比</th><th class="num wd-zz">转正</th></tr></thead>
            <tbody>
              {''.join(d['source_rows'])}
            </tbody>
          </table>
        </div>
      </div>

      <div class="mini-label">TOP 热点来源（{d['hot_top_label']}）</div>
      <table class="int-table">
        <thead><tr><th>#</th><th>{d['hot_top_label']}</th><th class="num">leads</th><th class="num wd-zz">转正</th></tr></thead>
        <tbody>{d['top_hot_rows']}</tbody>
      </table>

      <div class="mini-label" style="margin-top:14px;">TOP 索粉来源（计划名称）</div>
      <table class="int-table">
        <thead><tr><th>#</th><th>计划名称</th><th class="num">leads</th><th class="num wd-zz">转正</th></tr></thead>
        <tbody>{d['top_sf_rows']}</tbody>
      </table>
    </div>
  </div>"""
    return section


if __name__ == '__main__':
    arg = sys.argv[1] if len(sys.argv) > 1 else 'all'
    if arg == 'all':
        for key in CHANNELS:
            generate_channel(key)
    elif arg in CHANNELS:
        generate_channel(arg)
    else:
        print(f"未知渠道: {arg}, 可选: {', '.join(CHANNELS.keys())} / all")
        sys.exit(1)
