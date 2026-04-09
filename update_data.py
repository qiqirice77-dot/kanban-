#!/usr/bin/env python3
"""
update_data.py — 从 Google Sheets 拉取最新数据，注入 index.html
用法：cd ~/Desktop/七组看板 && python3 update_data.py
"""

import json, re, datetime, calendar, os, sys, urllib.request, csv, io

# ══════════════════════════════════════════════════════════════════
#  配置
# ══════════════════════════════════════════════════════════════════

HTML_PATH = os.path.expanduser("~/七组看板/逸品商务BI看板.html")
SHEET_ID  = "1yOzqYjNAEyDmto_PS6Ac3583mcHIQGT6favAWiczAFk"

SHEET_GIDS = {"mar": "905506719", "feb": "611405091", "jan": "1922816921"}

VALID_NAMES = ["卢登", "黄凯悦", "吴挺", "郭永正", "陈亦凡"]

COLORS = {
    "卢登":   "#1E88E5",
    "黄凯悦": "#FF9100",
    "吴挺":   "#00BCD4",
    "郭永正": "#7C4DFF",
    "陈亦凡": "#00E676",
}

# 个人 KPI 目标（元）
KPI = {
    "卢登":   {"total": 6936050, "new": 3000000, "renew": 3936050},
    "黄凯悦": {"total": 4639500, "new": 2600000, "renew": 2039500},
    "吴挺":   {"total": 3463460, "new": 1500000, "renew": 1963460},
    "郭永正": {"total": 3320000, "new": 1600000, "renew": 1720000},
    "陈亦凡": {"total": 1500000, "new": 1500000, "renew":       0},
}

# 团队总目标（元）
TOTAL_TARGET = 19859010

# 列索引（0-indexed）
COL_SALES  = 6   # 列7  销售
COL_IND    = 10  # 列11 广告类型（行业）
COL_MODE   = 11  # 列12 合作模式
COL_EXPIRE = 13  # 列14 到期时间
COL_AMT    = 15  # 列16 实际金额
COL_NOTE   = 18  # 列19 备注（新增/续费）
COL_PAY    = 20  # 列21 付款时间
COL_CLIENT = 21  # 列22 广告主联系方式
COL_GROUP  = 22  # 列23 合作群名

# 客户分层阈值（累计实收，元）
TIER_THRESHOLDS = [
    ("战略客户", 10_000_000),
    ("核心客户",  5_000_000),
    ("重点客户",  1_000_000),
    ("成长客户",    500_000),
    ("常规客户",    200_000),
    ("培育客户",    100_000),
    ("观察客户",          0),
]

# 单笔触发 5 天预警的最低金额（50万）
EXPIRE_SOON_MIN_AMT = 500_000

# 周报月报提交记录 GID
REPORT_GID = "848744054"

# 周报截止：每周一 15:30 / 月报截止：每月1号 15:30
REPORT_WEEKLY_HOUR, REPORT_WEEKLY_MIN   = 15, 30
REPORT_MONTHLY_HOUR, REPORT_MONTHLY_MIN = 15, 30

# Telegram 机器人配置
TG_TOKEN   = "8597350679:AAF411QUfaoi-ebxRoi1G4gMnNcMuZqQZTY"
TG_CHAT_ID = "5582821848"

MODE_META = {
    "CPT": {"label": "按时长计费", "icon": "⏱️", "color": "#1E88E5"},
    "CPA": {"label": "按行为计费", "icon": "🖱️", "color": "#7C4DFF"},
    "CPS": {"label": "按销售提成", "icon": "💰", "color": "#00E676"},
}

TODAY = datetime.date.today()

# ══════════════════════════════════════════════════════════════════
#  工具函数
# ══════════════════════════════════════════════════════════════════

def tg_send(msg):
    """发送 Telegram 消息，失败不中断主流程"""
    try:
        url  = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
        data = json.dumps({"chat_id": TG_CHAT_ID, "text": msg, "parse_mode": "HTML"}).encode()
        req  = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=10)
    except Exception as e:
        print(f"  ⚠️  Telegram 发送失败: {e}")

def fetch_csv(gid):
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    raw = urllib.request.urlopen(req, timeout=30).read().decode("utf-8")
    return list(csv.reader(io.StringIO(raw)))

def g(row, idx):
    return row[idx].strip() if idx < len(row) else ""

def parse_amt(s):
    try:
        return float(str(s).replace(",", "").replace("，", "").strip())
    except:
        return 0.0

def parse_date(s):
    s = str(s).strip()
    m = re.match(r"(\d{4})[-./年](\d{1,2})[-./月](\d{1,2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except:
            pass
    m = re.match(r"(\d{1,2})[-./](\d{1,2})", s)
    if m:
        try:
            return datetime.date(TODAY.year, int(m.group(1)), int(m.group(2)))
        except:
            pass
    return None

def parse_datetime(s):
    """解析 '2026-03-30 14:22' 或 '2026/03/30 14:22' 等格式"""
    s = str(s).strip()
    for pat in [
        r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s+(\d{1,2}):(\d{1,2})",
        r"(\d{4})年(\d{1,2})月(\d{1,2})日\s*(\d{1,2}):(\d{1,2})",
    ]:
        m = re.match(pat, s)
        if m:
            try:
                return datetime.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                                         int(m.group(4)), int(m.group(5)))
            except:
                pass
    return None

def sum_col_amt(raw_rows):
    """总实收：rows[2:] 列16 无条件求和，跳过含"合计"/"总计"的汇总行"""
    total = 0.0
    for r in raw_rows[2:]:
        row_text = " ".join(r)
        if "合计" in row_text or "总计" in row_text:
            continue
        total += parse_amt(g(r, COL_AMT))
    return total

def parse_detail_rows(raw_rows):
    """明细行解析：rows[2:] 列16 > 0 的行，跳过合计/汇总行"""
    result = []
    for r in raw_rows[2:]:
        row_text = " ".join(r)
        if "合计" in row_text or "总计" in row_text:
            continue
        amt = parse_amt(g(r, COL_AMT))
        if amt <= 0:
            continue
        result.append({
            "销售":  g(r, COL_SALES),
            "行业":  g(r, COL_IND),
            "模式":  g(r, COL_MODE).upper(),
            "到期":  g(r, COL_EXPIRE),
            "金额":  amt,
            "备注":  g(r, COL_NOTE),
            "付款":  g(r, COL_PAY),
            "客户":  g(r, COL_CLIENT),
            "群名":  g(r, COL_GROUP),
        })
    return result

# ══════════════════════════════════════════════════════════════════
#  build_month_data — 为单个月份构建完整数据集
# ══════════════════════════════════════════════════════════════════

def build_month_data(raw_rows, year, month, ref_date):
    """
    raw_rows  : fetch_csv 返回的原始行
    year/month: 该月年份和月份
    ref_date  : "今日"基准（当月用 TODAY，历史月用该月最后一天）
    """
    rows          = parse_detail_rows(raw_rows)
    days_in_month = calendar.monthrange(year, month)[1]

    # ── salesList + renewData ──
    new_t  = {n: 0.0 for n in VALID_NAMES}
    renew_t= {n: 0.0 for n in VALID_NAMES}
    cnt    = {n: {"new": 0, "old": 0} for n in VALID_NAMES}

    for r in rows:
        name = r["销售"]
        if name not in VALID_NAMES:
            continue
        if "续费" in r["备注"]:
            renew_t[name] += r["金额"]; cnt[name]["old"] += 1
        else:
            new_t[name]   += r["金额"]; cnt[name]["new"] += 1

    sales_list = [
        {"name": n,
         "actual":     int(new_t[n] + renew_t[n]),
         "target":     KPI[n]["total"],
         "newOrders":  cnt[n]["new"],
         "oldClients": cnt[n]["old"],
         "color":      COLORS[n]}
        for n in VALID_NAMES
    ]
    renew_data = [
        {"name":        n,
         "newActual":   int(new_t[n]),
         "newTarget":   KPI[n]["new"],
         "renewActual": int(renew_t[n]),
         "renewTarget": KPI[n]["renew"]}
        for n in VALID_NAMES
    ]

    # ── adData.modes ──
    mode_totals = {"CPT": 0.0, "CPA": 0.0, "CPS": 0.0}
    for r in rows:
        m = r["模式"]
        if m in mode_totals:
            mode_totals[m] += r["金额"]

    ad_modes = [
        {"mode":   m,
         "label":  MODE_META[m]["label"],
         "actual": int(mode_totals[m]),
         "target": int(mode_totals[m] * 1.1),
         "mom":    0,
         "icon":   MODE_META[m]["icon"],
         "color":  MODE_META[m]["color"]}
        for m in ["CPT", "CPA", "CPS"]
    ]

    # ── adData.industries ──
    ind_totals: dict = {}
    for r in rows:
        ind = r["行业"].strip() or "其他"
        if ind not in ind_totals:
            ind_totals[ind] = {"actual": 0.0, "clients": set()}
        ind_totals[ind]["actual"]  += r["金额"]
        ind_totals[ind]["clients"].add(r["客户"])

    ad_industries = sorted([
        {"name":    k,
         "actual":  int(v["actual"]),
         "target":  int(v["actual"] * 1.1),
         "mom":     0,
         "clients": len(v["clients"])}
        for k, v in ind_totals.items() if v["actual"] > 0
    ], key=lambda x: -x["actual"])

    # ── personData（只看5人，只用当月数据）──
    person_data = {}
    for name in VALID_NAMES:
        p_rows = [r for r in rows if r["销售"] == name]

        mode_detail: dict = {}
        for r in p_rows:
            m = r["模式"]
            if not m:
                continue
            md = mode_detail.setdefault(m, {"mode": m, "amount": 0.0, "orders": 0})
            md["amount"] += r["金额"]; md["orders"] += 1
        mode_details = sorted(
            [{"mode": v["mode"], "amount": int(v["amount"]), "orders": v["orders"]}
             for v in mode_detail.values() if v["amount"] > 0],
            key=lambda x: -x["amount"],
        )

        ind_detail: dict = {}
        for r in p_rows:
            ind = r["行业"].strip() or "其他"
            id_ = ind_detail.setdefault(ind, {"industry": ind, "amount": 0.0, "orders": 0})
            id_["amount"] += r["金额"]; id_["orders"] += 1
        ind_details = sorted(
            [{"industry": v["industry"], "amount": int(v["amount"]), "orders": v["orders"]}
             for v in ind_detail.values() if v["amount"] > 0],
            key=lambda x: -x["amount"],
        )

        person_data[name] = {
            "industryDetails": ind_details or [{"industry": "未分类", "amount": 0, "orders": 0}],
            "modeDetails":     mode_details or [{"mode": "CPT", "amount": 0, "orders": 0}],
            "clients":         [],
        }

    # ── adTable（到期预警用）──
    ad_table = []
    for r in rows:
        exp_date = parse_date(r["到期"])
        if not exp_date:
            continue
        ad_table.append({
            "client":      r["客户"][:20] or "未知客户",
            "group":       r["群名"][:30] if r["群名"] else "",
            "displayName": (r["群名"] or r["客户"])[:30] or "未知客户",
            "sales":       r["销售"],
            "amount":      int(r["金额"]),
            "type":        "普通",
            "expireDate":  exp_date.strftime("%Y-%m-%d"),
        })

    expire_today = [
        {**x, "renewProb": "中", "contact": x["client"][:15]}
        for x in ad_table if parse_date(x["expireDate"]) == ref_date
    ]
    expire_soon = sorted([
        {**x,
         "daysLeft": (parse_date(x["expireDate"]) - ref_date).days,
         "contact":  x["client"][:15]}
        for x in ad_table
        if 1 <= (parse_date(x["expireDate"]) - ref_date).days <= 5
    ], key=lambda x: x["daysLeft"])

    # ── dailyByDate（每天明细，供单日视角使用）──
    daily_by_date = {}
    for day in range(1, days_in_month + 1):
        date_str = f"{year}-{month:02d}-{day:02d}"
        d_new    = {n: 0.0 for n in VALID_NAMES}
        d_renew  = {n: 0.0 for n in VALID_NAMES}
        d_cnt    = {n: {"new": 0, "old": 0} for n in VALID_NAMES}
        d_total  = 0.0

        for r in rows:
            if date_str not in r["付款"]:
                continue
            d_total += r["金额"]
            name = r["销售"]
            if name not in VALID_NAMES:
                continue
            if "续费" in r["备注"]:
                d_renew[name] += r["金额"]; d_cnt[name]["old"] += 1
            else:
                d_new[name]   += r["金额"]; d_cnt[name]["new"] += 1

        daily_by_date[date_str] = {
            "label": f"{month}月{day}日",
            "salesList": [
                {"name":       n,
                 "actual":     int(d_new[n] + d_renew[n]),
                 "target":     KPI[n]["total"] // days_in_month,
                 "newOrders":  d_cnt[n]["new"],
                 "oldClients": d_cnt[n]["old"],
                 "color":      COLORS[n]}
                for n in VALID_NAMES
            ],
            "renewData": [
                {"name":        n,
                 "newActual":   int(d_new[n]),
                 "newTarget":   KPI[n]["new"] // days_in_month,
                 "renewActual": int(d_renew[n]),
                 "renewTarget": KPI[n]["renew"] // days_in_month}
                for n in VALID_NAMES
            ],
            "totalActual": int(d_total),
            "totalTarget": TOTAL_TARGET // days_in_month,
            "adData":      {"modes": ad_modes, "industries": ad_industries},
        }

    return {
        "year":        year,
        "month":       month,
        "daysInMonth": days_in_month,
        "label":       f"{year}年{month}月",
        "salesList":   sales_list,
        "renewData":   renew_data,
        "adData":      {"modes": ad_modes, "industries": ad_industries},
        "personData":  person_data,
        "totalActual": int(sum_col_amt(raw_rows)),
        "totalTarget": TOTAL_TARGET,
        "expireToday": expire_today,
        "expireSoon":  expire_soon,
        "adTable":     ad_table,
        "dailyByDate": daily_by_date,
    }

# ══════════════════════════════════════════════════════════════════
#  build_report_data — 周报月报提交记录
# ══════════════════════════════════════════════════════════════════

def build_report_data(raw_rows):
    """
    raw_rows: 提交记录 Sheet，表头行1（跳过），数据列：
      0=姓名  1=周期  2=提交时间  3=文件链接（可选）
    """
    # ── 截止时间 ──
    days_since_monday = TODAY.weekday()           # 0=周一
    this_monday = TODAY - datetime.timedelta(days=days_since_monday)
    weekly_ddl = datetime.datetime(
        this_monday.year, this_monday.month, this_monday.day,
        REPORT_WEEKLY_HOUR, REPORT_WEEKLY_MIN)

    if TODAY.month == 12:
        next_m1 = datetime.date(TODAY.year + 1, 1, 1)
    else:
        next_m1 = datetime.date(TODAY.year, TODAY.month + 1, 1)
    monthly_ddl = datetime.datetime(
        next_m1.year, next_m1.month, next_m1.day,
        REPORT_MONTHLY_HOUR, REPORT_MONTHLY_MIN)

    week_start  = datetime.datetime(this_monday.year, this_monday.month, this_monday.day)
    month_start = datetime.datetime(TODAY.year, TODAY.month, 1)

    # ── ���析提交记录 ──
    submissions = []
    for row in raw_rows[1:]:         # 跳过表头
        if not row or len(row) < 3:
            continue
        name   = g(row, 0)
        period = g(row, 1)
        st_str = g(row, 2)
        link   = g(row, 3) if len(row) > 3 else ""
        if name not in VALID_NAMES or not st_str:
            continue
        st_dt = parse_datetime(st_str)
        submissions.append({
            "name": name, "period": period,
            "submitTime": st_str, "submit_dt": st_dt, "link": link,
            "is_weekly":  "周报" in period,
            "is_monthly": "月报" in period,
        })

    # ── 为每人找本周周报 / 本月月报 ──
    result = []
    for name in VALID_NAMES:
        psubs = [s for s in submissions if s["name"] == name]

        w_subs = [s for s in psubs
                  if s["is_weekly"] and s["submit_dt"] and s["submit_dt"] >= week_start]
        w = w_subs[-1] if w_subs else None

        m_subs = [s for s in psubs
                  if s["is_monthly"] and s["submit_dt"] and s["submit_dt"] >= month_start]
        mo = m_subs[-1] if m_subs else None

        result.append({
            "name": name,
            "weekly": {
                "submitted":  w is not None,
                "submitTime": w["submitTime"] if w else "",
                "period":     w["period"]     if w else "",
                "onTime":     bool(w and w["submit_dt"] and w["submit_dt"] <= weekly_ddl),
            },
            "monthly": {
                "submitted":  mo is not None,
                "submitTime": mo["submitTime"] if mo else "",
                "period":     mo["period"]     if mo else "",
                "onTime":     bool(mo and mo["submit_dt"] and mo["submit_dt"] <= monthly_ddl),
            },
        })

    return {
        "list":            result,
        "weeklyDeadline":  weekly_ddl.strftime("%m月%d日 %H:%M"),
        "monthlyDeadline": monthly_ddl.strftime("%m月%d日 %H:%M"),
    }

# ══════════════════════════════════════════════════════════════════
#  ① 拉取数据
# ══════════════════════════════════════════════════════════════════

print("📡 正在拉取数据...")
mar_raw    = fetch_csv(SHEET_GIDS["mar"])
feb_raw    = fetch_csv(SHEET_GIDS["feb"])
jan_raw    = fetch_csv(SHEET_GIDS["jan"])
try:
    report_raw = fetch_csv(REPORT_GID)
    print(f"  ✓ 周报月报提交记录 {len(report_raw)} 行")
except Exception as e:
    print(f"  ⚠️  周报记录拉取失败（{e}），使用空数据")
    report_raw = []
print(f"  ✓ 3月 {len(mar_raw)} 行 | 2月 {len(feb_raw)} 行 | 1月 {len(jan_raw)} 行（含表头/汇总）")

# ══════════════════════════════════════════════════════════════════
#  ② 构建三个月数据
# ══════════════════════════════════════════════════════════════════

mar_data = build_month_data(mar_raw, 2026, 3, TODAY)
feb_data = build_month_data(feb_raw, 2026, 2, datetime.date(2026, 2, 28))
jan_data = build_month_data(jan_raw, 2026, 1, datetime.date(2026, 1, 31))

mar_rows = parse_detail_rows(mar_raw)
feb_rows = parse_detail_rows(feb_raw)
jan_rows = parse_detail_rows(jan_raw)
print(f"  ✓ 3月明细 {len(mar_rows)} 条 | 2月 {len(feb_rows)} 条 | 1月 {len(jan_rows)} 条")

# ══════════════════════════════════════════════════════════════════
#  ③ 计算环比（MOM）
# ══════════════════════════════════════════════════════════════════

def apply_mom(curr_list, prev_list, key="mode"):
    prev = {x[key]: x["actual"] for x in prev_list}
    for x in curr_list:
        p = prev.get(x[key], 0)
        x["mom"] = round((x["actual"] - p) / p * 100, 1) if p else 0

apply_mom(mar_data["adData"]["modes"],      feb_data["adData"]["modes"])
apply_mom(feb_data["adData"]["modes"],      jan_data["adData"]["modes"])
apply_mom(mar_data["adData"]["industries"], feb_data["adData"]["industries"], key="name")
apply_mom(feb_data["adData"]["industries"], jan_data["adData"]["industries"], key="name")

# ══════════════════════════════════════════════════════════════════
#  ④ 验证输出
# ══════════════════════════════════════════════════════════════════

print(f"\n📊 三月验证：")
print(f"  全表总实收: {mar_data['totalActual']:,}")
team_total = 0
for s in mar_data["salesList"]:
    team_total += s["actual"]
    kpi = KPI[s["name"]]["total"]
    pct = round(s["actual"] / kpi * 100, 1)
    rd  = next(r for r in mar_data["renewData"] if r["name"] == s["name"])
    print(f"  {s['name']:4s}: 实收={s['actual']:>10,}  目标={kpi:>10,}  {pct:5.1f}%"
          f"  (新增={rd['newActual']:>9,}  续费={rd['renewActual']:>9,}  {s['newOrders']+s['oldClients']}笔)")
print(f"\n  5人合计 : {team_total:>12,}")
print(f"  全表合计: {mar_data['totalActual']:>12,}  差值={mar_data['totalActual']-team_total:,}（非团队成员）\n")

# ══════════════════════════════════════════════════════════════════
#  ⑤ 客户分层 + 三月到期预警（二月+三月合并）
# ══════════════════════════════════════════════════════════════════

from collections import defaultdict

# 计算所有客户三个月累计金额 → 分层
def _tier_key(r):
    return r.get("群名", "").strip() or r.get("客户", "").strip() or ""

client_totals: dict = defaultdict(float)
for r in jan_rows + feb_rows + mar_rows:
    k = _tier_key(r)
    if k:
        client_totals[k] += r["金额"]

def get_client_tier(r):
    k = _tier_key(r) if isinstance(r, dict) else str(r)
    total = client_totals.get(k, 0)
    for tier, thresh in TIER_THRESHOLDS:
        if total >= thresh:
            return tier
    return "观察客户"

# 合并二月+三月 adTable，加上来源和客户层级
combined_feb_mar = [
    {**x, "source": "二月", "tier": get_client_tier(x)}
    for x in feb_data["adTable"]
] + [
    {**x, "source": "三月", "tier": get_client_tier(x)}
    for x in mar_data["adTable"]
]

# expireToday：今日到期（全部显示，不限金额）
expire_today_combined = sorted([
    {**x, "daysLeft": 0}
    for x in combined_feb_mar
    if parse_date(x["expireDate"]) == TODAY
], key=lambda x: -x["amount"])

# expireSoon：单笔 ≥ 50万，5天内到期
expire_soon_combined = sorted([
    {**x, "daysLeft": (parse_date(x["expireDate"]) - TODAY).days}
    for x in combined_feb_mar
    if 1 <= (parse_date(x["expireDate"]) - TODAY).days <= 5
    and x["amount"] >= EXPIRE_SOON_MIN_AMT
], key=lambda x: (x["daysLeft"], -x["amount"]))

# 覆盖三月 expireToday / expireSoon
mar_data["expireToday"] = expire_today_combined
mar_data["expireSoon"]  = expire_soon_combined

print(f"  🔴 今日到期  : {len(expire_today_combined)} 单")
print(f"  🟠 5天预警   : {len(expire_soon_combined)} 单（≥50万）\n")

# ══════════════════════════════════════════════════════════════════
#  ⑥ 构建其余数据
# ══════════════════════════════════════════════════════════════════

# ── trendData（三月走势）──
def name_month_total(rows, name):
    return int(sum(r["金额"] for r in rows if r["销售"] == name))

trend_data = {
    "months": ["1月", "2月", "3月（本月）"],
    "series": [
        {"name":  n,
         "data":  [name_month_total(jan_rows, n),
                   name_month_total(feb_rows, n),
                   name_month_total(mar_rows, n)],
         "color": COLORS[n]}
        for n in VALID_NAMES
    ],
}

# ── weeklyReport（从 Sheets 提交记录构建）──
weekly_report = build_report_data(report_raw)
print(f"  ✓ 周报提交：{sum(1 for r in weekly_report['list'] if r['weekly']['submitted'])}/5 人"
      f" | 月报提交：{sum(1 for r in weekly_report['list'] if r['monthly']['submitted'])}/5 人")

# ══════════════════════════════════════════════════════════════════
#  ⑥ 组装并注入 MOCK_DATA
# ══════════════════════════════════════════════════════════════════

mock_data = {
    "monthlyData": {
        "mar": mar_data,
        "feb": feb_data,
        "jan": jan_data,
    },
    "weeklyReport": weekly_report,
}

def replace_mock_data(content, new_data):
    marker = "const MOCK_DATA = {"
    start  = content.find(marker)
    if start == -1:
        return None
    depth = 0
    i = start + len(marker) - 1
    while i < len(content):
        if content[i] == "{":
            depth += 1
        elif content[i] == "}":
            depth -= 1
            if depth == 0:
                j = i + 1
                while j < len(content) and content[j] in " \t\r\n": j += 1
                if j < len(content) and content[j] == ";":
                    new_block = "const MOCK_DATA = " + json.dumps(new_data, ensure_ascii=False, indent=2) + ";"
                    return content[:start] + new_block + content[j+1:]
                break
        i += 1
    return None

if not os.path.exists(HTML_PATH):
    print(f"❌ 找不到 HTML 文件: {HTML_PATH}")
    sys.exit(1)

with open(HTML_PATH, "r", encoding="utf-8") as f:
    html = f.read()

updated = replace_mock_data(html, mock_data)
if updated is None:
    print("❌ 未找到 MOCK_DATA 替换点")
    sys.exit(1)

# 注入 trendData（独立块）
trend_block = "MOCK_DATA.trendData = " + json.dumps(trend_data, ensure_ascii=False, indent=2) + ";"
updated = re.sub(r"MOCK_DATA\.trendData\s*=\s*\{.*?\};", trend_block, updated, flags=re.DOTALL)

with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(updated)

print("✅ 数据注入完成！")
print(f"   👥 销售人数  : {len(VALID_NAMES)}")
print(f"   📊 三月合同  : {len(mar_rows)} 条")
print(f"   📊 二月合同  : {len(feb_rows)} 条")
print(f"   📊 一月合同  : {len(jan_rows)} 条")
print(f"   💰 总实收    : {mar_data['totalActual']:,}")
print(f"   🔴 今日到期  : {len(mar_data['expireToday'])} 单")
print(f"   🟠 5天内预警 : {len(mar_data['expireSoon'])} 单")
print(f"   📁 已更新    : {HTML_PATH}")
print()
print("👉 双击 index.html 打开看板查看最新数据")

# ══════════════════════════════════════════════════════════════════
#  ⑦ Telegram 通知
# ══════════════════════════════════════════════════════════════════

NOW = datetime.datetime.now()

# ── 每日：今日到期 + 5天预警 ──
def notify_expiry():
    today_list = mar_data["expireToday"]
    soon_list  = mar_data["expireSoon"]
    if not today_list and not soon_list:
        print("  📭 无到期提醒，跳过 Telegram 推送")
        return

    lines = [f"📅 <b>逸品销售 · {TODAY.month}/{TODAY.day} 到期提醒</b>\n"]

    if today_list:
        lines.append(f"🔴 <b>今日到期 {len(today_list)} 单</b>")
        for x in today_list:
            lines.append(f"• {x['displayName']} · {x['sales']} · ¥{x['amount']:,} [{x.get('tier','—')}]")
        lines.append("")

    if soon_list:
        lines.append(f"🟠 <b>5天内到期 {len(soon_list)} 单（≥50万）</b>")
        for x in soon_list:
            lines.append(f"• {x['displayName']} · 还剩{x.get('daysLeft','?')}天 · {x['sales']} · ¥{x['amount']:,}")

    tg_send("\n".join(lines))
    print("  📤 已推送到期提醒至 Telegram")

# ── 周一15:30后：周报未提交提醒 ──
def notify_weekly_report():
    wr_list       = weekly_report["list"]
    not_submitted = [r["name"] for r in wr_list if not r["weekly"]["submitted"]]
    submitted     = [r["name"] for r in wr_list if r["weekly"]["submitted"]]

    lines = [
        f"📋 <b>逸品销售 · 本周周报截止提醒</b>",
        f"截止时间：{TODAY.month}/{TODAY.day} 15:30\n",
    ]
    if not_submitted:
        lines.append(f"⚠️ <b>未提交（{len(not_submitted)} 人）：</b>")
        for name in not_submitted:
            lines.append(f"• {name}")
    else:
        lines.append("✅ 全员已按时提交！")
    if submitted:
        lines.append(f"\n✅ 已提交：{'、'.join(submitted)}")

    tg_send("\n".join(lines))
    print("  📤 已推送周报提醒至 Telegram")

notify_expiry()

# 仅周一 15:30 之后触发周报提醒
if NOW.weekday() == 0 and (NOW.hour > 15 or (NOW.hour == 15 and NOW.minute >= 30)):
    notify_weekly_report()
