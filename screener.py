# -*- coding: utf-8 -*-
"""
可转债双低策略筛选 · 网页自动版（GitHub Actions 每日跑）
功能：
  1) 抓取东财可转债行情 + 基本信息（requests 直连，失败回退本地缓存）
  2) 计算：剩余本息锚/买入上限、评级≥AA-、规模2~30亿、估值水位总开关、强赎关注区
  3) 输出：Excel / CSV / Apple风格 index.html / 追加 history.csv
  4) 可选：SERVERCHAN_KEY 环境变量存在时，推送微信摘要
数据日期：运行当天（收盘后）
"""
import os
import sys
import json
import glob
import re
import datetime as dt
import pandas as pd

try:
    import requests
except ImportError:
    requests = None

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "output")
HISTORY = os.path.join(BASE, "history.csv")
INDEX = os.path.join(BASE, "index.html")
os.makedirs(OUT, exist_ok=True)

TODAY = dt.datetime.now().strftime("%Y%m%d")
NOW = pd.Timestamp.now()

# ---------- 策略参数 ----------
BUFFER = 0.05            # 买入上限 = 锚 × (1+BUFFER)
MIN_RATING_RANK = 16     # AA- 及以上
SCALE_MIN, SCALE_MAX = 2.0, 30.0
EMPTY_SIGNAL_N = 15      # 双低<130 合格标的 < 此数 → 空仓信号

RATING_RANK = {
    'C': 1, 'CC': 2, 'CCC': 3, 'B-': 4, 'B': 5, 'B+': 6,
    'BB-': 7, 'BB': 8, 'BB+': 9, 'BBB-': 10, 'BBB': 11, 'BBB+': 12,
    'A-': 13, 'A': 14, 'A+': 15, 'AA-': 16, 'AA': 17, 'AA+': 18, 'AAA': 19,
}
CN_NUM = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
          '七': 7, '八': 8, '九': 9}

QUOTE_HOSTS = [
    "https://push2delay.eastmoney.com/api/qt/clist/get",
    "https://push2.eastmoney.com/api/qt/clist/get",
    "https://16.push2.eastmoney.com/api/qt/clist/get",
]
QUOTE_FIELDS = "f2,f3,f12,f14,f227,f235,f236,f237,f240"
INFO_HOSTS = [
    "https://datacenter-web.eastmoney.com/api/data/v1/get",
    "https://datacenter.eastmoney.com/api/data/v1/get",
]


# ================= 抓取层 =================
def _get_json(url, params, timeout=20):
    if requests is None:
        return None
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    try:
        r = requests.get(url, params=params, headers=headers, timeout=timeout)
        if r.status_code == 200:
            return r.json()
    except Exception:
        return None
    return None


def fetch_quotes():
    """返回 list[dict]（diff）或 None（失败回退缓存）"""
    for host in QUOTE_HOSTS:
        rows = []
        for p in range(1, 5):
            d = _get_json(host, {
                "pn": p, "pz": 100, "po": 1, "np": 1,
                "ut": "bd1d9ddb04089700cf9c27f6f7426281", "fltt": 2,
                "invt": 2, "fid": "f243", "fs": "b:MK0354",
                "fields": QUOTE_FIELDS,
            })
            if not d:
                break
            diff = d.get("data", {}).get("diff", [])
            if not diff:
                break
            rows += diff
            if len(diff) < 100:
                break
        if rows:
            # 落本地缓存，便于回退与调试
            with open(os.path.join(OUT, f"quote_cache_{TODAY}.json"), "w", encoding="utf-8") as f:
                json.dump(rows, f, ensure_ascii=False)
            return rows
    # 回退：读本地已有缓存
    cache = sorted(glob.glob(os.path.join(OUT, "quote_cache_*.json")))
    if cache:
        print("[warn] 行情实时抓取失败，回退本地缓存:", cache[-1])
        return json.load(open(cache[-1], encoding="utf-8"))
    # 再回退：旧 q2_p*.json
    rows = []
    for f in sorted(glob.glob(os.path.join(OUT, "q2_p*.json"))):
        rows += json.load(open(f, encoding="utf-8")).get("data", {}).get("diff", [])
    return rows or None


def fetch_info():
    """返回 list[dict] 或 None"""
    for host in INFO_HOSTS:
        frames = []
        for p in range(1, 5):
            d = _get_json(host, {
                "reportName": "RPT_BOND_CB_LIST", "columns": "ALL",
                "pageNumber": p, "pageSize": 600,
                "sortColumns": "LISTING_DATE", "sortTypes": "-1",
                "source": "WEB", "client": "WEB",
            })
            if not d:
                break
            rows = d.get("result", {}).get("data", [])
            if not rows:
                break
            frames.append(rows)
            if len(rows) < 600:
                break
        if frames:
            allrows = [r for fr in frames for r in fr]
            with open(os.path.join(OUT, f"info_cache_{TODAY}.json"), "w", encoding="utf-8") as f:
                json.dump(allrows, f, ensure_ascii=False)
            return allrows
    cache = sorted(glob.glob(os.path.join(OUT, "info_cache_*.json")))
    if cache:
        print("[warn] 基本信息实时抓取失败，回退本地缓存:", cache[-1])
        return json.load(open(cache[-1], encoding="utf-8"))
    frames = []
    for f in sorted(glob.glob(os.path.join(OUT, "kzz_info*.json"))):
        frames.append(pd.read_json(f))
    if frames:
        return pd.concat(frames, ignore_index=True).to_dict("records")
    return None


# ================= 计算层 =================
def parse_redeem_price(row):
    for col in ["INTEREST_RATE_EXPLAIN", "REDEEM_CLAUSE"]:
        txt = str(row.get(col, "") or "")
        m = re.search(r"到期赎回价为?\s*([\d.]+)\s*元", txt)
        if m:
            return float(m.group(1))
    txt = str(row.get("REDEEM_CLAUSE", "") or "")
    m = re.search(r"面值的\s*([\d.]+)\s*%", txt)
    if m:
        return float(m.group(1))
    return None


def parse_coupon_rates(row):
    txt = str(row.get("INTEREST_RATE_EXPLAIN", "") or "")
    out = []
    for m in re.finditer(r"第([一二三四五六七八九])年\s*([\d.]+)\s*%", txt):
        yr = CN_NUM.get(m.group(1))
        if yr:
            out.append((yr, float(m.group(2)) / 100.0))
    out.sort()
    return out


def compute_anchor(row):
    rp = parse_redeem_price(row)
    rates = parse_coupon_rates(row)
    start = row.get("BOND_START_DATE")
    n = len(rates)
    future_sum = 0.0
    if rp is not None and n > 0 and pd.notna(start):
        for yr, rate in rates:
            if yr >= n:
                continue
            pay_date = start + pd.DateOffset(years=yr - 1)
            if pay_date > NOW:
                future_sum += rate * 100
        anchor = rp + future_sum
        src = "条款解析"
    else:
        bv = row.get("bond_value")
        anchor = bv if pd.notna(bv) else None
        src = "纯债价值回退" if anchor is not None else "无数据"
    if anchor is None:
        return None, None, None
    return anchor, anchor * (1 + BUFFER), src


def build_df():
    qrows = fetch_quotes()
    irows = fetch_info()
    if not qrows:
        print("[error] 无行情数据"); sys.exit(1)
    if not irows:
        print("[error] 无基本信息数据"); sys.exit(1)

    q = pd.DataFrame(qrows).rename(columns={
        "f2": "price", "f3": "pchange", "f12": "code", "f14": "name",
        "f227": "bond_value", "f235": "stock_price", "f236": "convert_value",
        "f237": "premium", "f240": "redeem_trig",
    })
    for c in ["price", "pchange", "bond_value", "stock_price",
              "convert_value", "premium", "redeem_trig"]:
        q[c] = pd.to_numeric(q[c], errors="coerce")
    q["code"] = q["code"].astype(str).str.zfill(6)

    info = pd.DataFrame(irows)
    info["code"] = info["SECURITY_CODE"].astype(str).str.zfill(6)
    info["存续"] = info["DELIST_DATE"].isna()
    if "ACTUAL_ISSUE_SCALE" in info:
        s = pd.to_numeric(info["ACTUAL_ISSUE_SCALE"], errors="coerce")
        med = s.median()
        if pd.notna(med) and med > 1000:
            s = s / 1e8
        elif pd.notna(med) and med < 5:
            s = s / 1e4
        info["规模_亿"] = s
    for col in ["EXPIRE_DATE", "BOND_START_DATE"]:
        if col in info:
            info[col] = pd.to_datetime(info[col], errors="coerce")

    cols = ["code", "SECURITY_NAME_ABBR", "RATING", "规模_亿", "BOND_EXPIRE",
            "EXPIRE_DATE", "BOND_START_DATE", "INTEREST_RATE_EXPLAIN",
            "REDEEM_CLAUSE", "存续"]
    cols = [c for c in cols if c in info.columns]
    df = q.merge(info[cols], on="code", how="left", suffixes=("", "_info"))
    df["double_low"] = df["price"] + df["premium"]

    anc = df.apply(compute_anchor, axis=1, result_type="expand")
    df["剩余本息锚"], df["买入上限"], df["锚来源"] = anc[0], anc[1], anc[2]
    df["偏贵"] = df.apply(
        lambda r: (pd.notna(r["买入上限"]) and pd.notna(r["price"]) and r["price"] > r["买入上限"]),
        axis=1)
    df["评级等级"] = df["RATING"].astype(str).str.strip().map(RATING_RANK)
    if "EXPIRE_DATE" in df:
        df["剩余年限"] = (df["EXPIRE_DATE"] - NOW).dt.days / 365.25
    return df


def analyze(df):
    alive = df[df["price"].notna() & (df["price"] > 0) & df["存续"]].copy()
    n_total = len(alive)
    med_price = alive["price"].median()
    med_dl = alive["double_low"].median()
    n_dl130 = int((alive["double_low"] < 130).sum())
    if med_price < 110:
        verdict = "估值偏低/舒适区 —— 可积极建仓"
    elif med_price < 125:
        verdict = "中性区间 —— 正常轮动"
    elif med_price < 135:
        verdict = "估值偏高 —— 谨慎、控仓位"
    else:
        verdict = "估值高位 —— 双低已失效，空仓观望"
    empty = n_dl130 < EMPTY_SIGNAL_N

    stats = {
        "数据日期": TODAY,
        "可交易转债数": n_total,
        "价格中位数": round(med_price, 2),
        "价格均值": round(alive["price"].mean(), 2),
        "价格<115数量": int((alive["price"] < 115).sum()),
        "价格>=130(强赎关注)数量": int((alive["price"] >= 130).sum()),
        "溢价率中位数": round(alive["premium"].median(), 2),
        "双低值中位数": round(med_dl, 2),
        "双低值<130数量": n_dl130,
        "双低值<140数量": int((alive["double_low"] < 140).sum()),
        "估值水位结论": verdict,
        "双低失效空仓信号": "触发" if empty else "未触发",
    }

    rated = alive[alive["评级等级"].notna() & (alive["评级等级"] >= MIN_RATING_RANK)].copy()
    if "规模_亿" in rated:
        rated = rated[rated["规模_亿"].fillna(0).between(SCALE_MIN, SCALE_MAX)]
    cand = rated[rated["price"] <= 130].copy()
    cand = cand[cand["premium"].fillna(0) >= 0]
    cand = cand.sort_values("double_low", ascending=True, na_position="last").reset_index(drop=True)
    redeem_watch = alive[alive["price"] >= 130].copy().sort_values("price", ascending=False)
    return alive, cand, redeem_watch, stats, empty


# ================= 输出层 =================
OUT_COLS = ["code", "name", "price", "pchange", "double_low", "premium",
            "convert_value", "stock_price", "bond_value", "RATING", "规模_亿",
            "剩余年限", "剩余本息锚", "买入上限", "偏贵", "锚来源",
            "EXPIRE_DATE", "存续"]
REN = {"code": "转债代码", "name": "转债名称", "price": "转债价格",
       "pchange": "转债涨跌%", "double_low": "双低值", "premium": "转股溢价率%",
       "convert_value": "转股价值", "stock_price": "正股价", "bond_value": "纯债价值",
       "RATING": "评级", "规模_亿": "发行规模(亿)", "剩余年限": "剩余年限",
       "剩余本息锚": "剩余本息锚", "买入上限": "买入上限", "偏贵": "价格超买入上限",
       "锚来源": "锚来源", "EXPIRE_DATE": "到期日", "存续": "存续"}


def export_excel(alive, cand, redeem_watch, stats):
    path = os.path.join(OUT, f"可转债双低筛选_{TODAY}.xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        alive.sort_values("double_low")[OUT_COLS].rename(columns=REN).to_excel(
            w, sheet_name="全市场按双低排序", index=False)
        cand[OUT_COLS].rename(columns=REN).to_excel(w, sheet_name="安全双低候选(AA-+规模)", index=False)
        redeem_watch[OUT_COLS].rename(columns=REN).to_excel(w, sheet_name="强赎关注区(价>=130)", index=False)
        pd.DataFrame([stats]).to_excel(w, sheet_name="市场统计", index=False)
    print(f"[done] Excel: {path}")


def append_history(stats, empty):
    # 同日已存在则跳过（避免重复追加）
    if os.path.exists(HISTORY):
        try:
            old = pd.read_csv(HISTORY)
            if TODAY in set(old["date"].astype(str).tolist()):
                print(f"[skip] history.csv 已有 {TODAY} 记录，跳过追加")
                return
        except Exception:
            pass
    row = {
        "date": TODAY,
        "price_median": stats["价格中位数"],
        "price_mean": stats["价格均值"],
        "premium_median": stats["溢价率中位数"],
        "dl_median": stats["双低值中位数"],
        "dl130_count": stats["双低值<130数量"],
        "dl140_count": stats["双低值<140数量"],
        "redeem_watch": stats["价格>=130(强赎关注)数量"],
        "verdict": stats["估值水位结论"],
        "empty_signal": stats["双低失效空仓信号"],
    }
    write_header = not os.path.exists(HISTORY)
    df = pd.DataFrame([row])
    df.to_csv(HISTORY, mode="a", index=False, header=write_header, encoding="utf-8-sig")
    print(f"[done] history.csv: {HISTORY}")


def sparkline(values, w=300, h=60, color="#0a84ff"):
    if len(values) < 2:
        return ""
    mn, mx = min(values), max(values)
    if mx == mn:
        mx = mn + 1
    pts = []
    for i, v in enumerate(values):
        x = 8 + i * (w - 16) / (len(values) - 1)
        y = h - 8 - (v - mn) / (mx - mn) * (h - 16)
        pts.append(f"{x:.1f},{y:.1f}")
    return (f'<svg viewBox="0 0 {w} {h}" width="100%" height="{h}" '
            f'preserveAspectRatio="none"><polyline fill="none" stroke="{color}" '
            f'stroke-width="2" points="{" ".join(pts)}"/></svg>')


def build_html(stats, cand, redeem_watch, empty):
    # 读历史做趋势
    hist_svg_pm, hist_svg_dl = "", ""
    if os.path.exists(HISTORY):
        h = pd.read_csv(HISTORY)
        if len(h) >= 2:
            hist_svg_pm = sparkline(h["price_median"].tolist(), color="#ff9f0a")
            hist_svg_dl = sparkline(h["dl130_count"].tolist(), color="#0a84ff")

    badge = ("risk" if empty else "ok")
    badge_text = ("⚠️ 空仓观望" if empty else "✅ 可关注")
    action_cls = ("risk" if empty else "ok")
    action_text = (
        "当前双低&lt;130 的合格标的不足 15 只，说明市场整体偏贵。<strong>操作建议：空仓观望，不新建仓</strong>，等该数量回到 15 只以上再出手。"
        if empty else
        "市场处于可操作区间，可结合仓位从「安全双低候选」中挑选标的，但仍需单只独立判断风险。"
    )

    cand_rows = ""
    show = cand.head(20)
    for _, r in show.iterrows():
        over = "偏贵" if r.get("偏贵") else "—"
        over_cls = "tag-warn" if r.get("偏贵") else "tag-dim"
        cand_rows += (
            f"<tr><td>{r['code']}</td><td>{r['name']}</td>"
            f"<td class='num'>{r['price']:.2f}</td>"
            f"<td class='num'>{r['double_low']:.1f}</td>"
            f"<td class='num'>{r['premium']:.1f}%</td>"
            f"<td>{r['RATING']}</td>"
            f"<td class='num'>{r['规模_亿']:.1f}</td>"
            f"<td><span class='{over_cls}'>{over}</span></td></tr>")

    red_rows = ""
    for _, r in redeem_watch.head(10).iterrows():
        red_rows += (f"<tr><td>{r['code']}</td><td>{r['name']}</td>"
                     f"<td class='num'>{r['price']:.2f}</td>"
                     f"<td class='num'>{r['premium']:.1f}%</td>"
                     f"<td class='num'>{r['double_low']:.1f}</td>"
                     f"<td>{r['RATING']}</td></tr>")

    stats_cards = "".join(
        f"<div class='card'><div class='k'>{k}</div><div class='v'>{v}</div></div>"
        for k, v in [
            ("可交易转债", stats["可交易转债数"]),
            ("价格中位数", stats["价格中位数"]),
            ("双低值中位数", stats["双低值中位数"]),
            ("双低<130数量", stats["双低值<130数量"]),
            ("强赎关注区", stats["价格>=130(强赎关注)数量"]),
            ("空仓信号", stats["双低失效空仓信号"]),
        ])

    html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>可转债双低筛选 · {TODAY}</title>
<style>
:root{{--bg:#f2f2f7;--card:#fff;--fg:#1c1c1e;--dim:#8e8e93;--line:#e5e5ea;
--ok:#34c759;--warn:#ff9f0a;--risk:#ff3b30;--blue:#0a84ff;}}
@media(prefers-color-scheme:dark){{:root{{--bg:#000;--card:#1c1c1e;--fg:#f2f2f7;--dim:#8e8e93;--line:#2c2c2e;}}}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,"SF Pro","PingFang SC","Microsoft YaHei",sans-serif;
background:var(--bg);color:var(--fg);padding:16px;max-width:720px;margin:0 auto;-webkit-font-smoothing:antialiased}}
header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}}
h1{{font-size:20px;font-weight:700}}
h2{{font-size:15px;margin-bottom:10px;font-weight:700}}
h3{{font-size:13px;font-weight:700;margin:14px 0 6px;color:var(--fg)}}
.date{{color:var(--dim);font-size:13px}}
.badge{{padding:6px 12px;border-radius:999px;font-size:13px;font-weight:600}}
.badge.ok{{background:rgba(52,199,89,.15);color:var(--ok)}}
.badge.risk{{background:rgba(255,59,48,.15);color:var(--risk)}}
.verdict{{background:var(--card);border-radius:14px;padding:14px 16px;margin-bottom:14px;
font-size:15px;font-weight:600;border:1px solid var(--line)}}
.grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:18px}}
.card{{background:var(--card);border-radius:14px;padding:12px;text-align:center;border:1px solid var(--line)}}
.card .k{{font-size:12px;color:var(--dim);margin-bottom:6px}}
.card .v{{font-size:19px;font-weight:700}}
.section{{background:var(--card);border-radius:14px;padding:14px 16px;margin-bottom:16px;border:1px solid var(--line)}}
.guide{{font-size:13px;line-height:1.7;color:var(--fg)}}
.guide ul{{margin:8px 0 0 18px;color:var(--dim)}}
.guide li{{margin-bottom:5px}}
.guide strong{{color:var(--fg)}}
.action{{border-radius:12px;padding:12px 14px;margin-top:10px;font-size:13px;line-height:1.6}}
.action.ok{{background:rgba(52,199,89,.08)}}
.action.risk{{background:rgba(255,59,48,.08)}}
.action strong{{color:var(--fg)}}
.metric{{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px dashed var(--line);font-size:13px}}
.metric .mk{{color:var(--dim)}}
.metric .mv{{color:var(--fg);font-weight:600}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th,td{{padding:8px 6px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}}
th{{color:var(--dim);font-weight:600;font-size:12px}}
.th-num{{text-align:right}}
td.num{{text-align:right;font-variant-numeric:tabular-nums}}
.tag-warn{{color:var(--warn);font-weight:600}}
.tag-dim{{color:var(--dim)}}
.spark{{margin-top:8px}}
.spark .lab{{font-size:11px;color:var(--dim);margin:8px 0 2px}}
.note{{font-size:12px;color:var(--dim);margin:8px 0 12px;line-height:1.5}}
.sub{{font-size:12px;color:var(--dim);margin:-6px 0 10px;line-height:1.5}}
footer{{color:var(--dim);font-size:11px;text-align:center;margin-top:20px;line-height:1.6}}
</style></head><body>
<header><h1>可转债双低筛选</h1><span class="date">{TODAY}</span></header>
<div class="badge {badge}">{badge_text}</div>
<div class="verdict" style="margin-top:12px">{stats['估值水位结论']}</div>
<div class="grid">{stats_cards}</div>

<div class="section"><h2>今日结论与操作建议</h2>
<p class="note">本页基于「双低 + 安全债底 + 估值水位」的量化模型自动生成，每日北京 22:00 更新。</p>
<div class="action {action_cls}">{action_text}</div>
</div>

<div class="section"><h2>筛选规则与指标说明</h2>
<div class="guide">
<h3>一、两个核心区域</h3>
<ul>
<li><strong>安全双低候选</strong>：评级 ≥ AA-、发行规模 2~30 亿、价格 ≤ 130 元，再按双低值由小到大排序。这是本策略的潜在买入观察池。</li>
<li><strong>强赎关注区</strong>：价格 ≥ 130 元的高价债集合。它们股性强、波动大、触发强赎概率高，<strong>不是买入清单</strong>，仅作风险警示。</li>
</ul>
<h3>二、常用指标</h3>
<div class="metric"><span class="mk">价格</span><span class="mv">可转债的市场价格</span></div>
<div class="metric"><span class="mk">溢价率</span><span class="mv">转股溢价率 = (债价/转股价值 - 1) × 100%，越低股性越强</span></div>
<div class="metric"><span class="mk">双低值</span><span class="mv">价格 + 溢价率 × 100，越小越符合「债底安全 + 股性不差」</span></div>
<div class="metric"><span class="mk">评级</span><span class="mv">发债主体信用评级，AAA 最安全，A 及以下风险上升</span></div>
<div class="metric"><span class="mk">规模(亿)</span><span class="mv">实际发行规模；过小易被炒作，过大弹性弱</span></div>
<div class="metric"><span class="mk">债底 / 偏贵</span><span class="mv">债底=剩余本息锚；价格超过「买入上限（债底×1.05）」标为偏贵</span></div>
<h3>三、空仓信号</h3>
<ul>
<li>当「双低值 &lt; 130」的安全候选少于 {EMPTY_SIGNAL_N} 只时，触发<strong>空仓观望</strong>。</li>
<li>含义：市场整体的债价和溢价率都很高，双低策略暂时失效，不硬买。</li>
</ul>
</div></div>

<div class="section"><h2>估值水位趋势</h2>
<div class="spark"><div class="lab">价格中位数</div>{hist_svg_pm}</div>
<div class="spark"><div class="lab">双低&lt;130 标的数量</div>{hist_svg_dl}</div>
<p class="note">说明：横轴为历史运行日期。价格中位数反映市场整体贵贱；双低&lt;130 标的数量反映可选机会多少。数量越多，市场越便宜。</p>
</div>

<div class="section"><h2>安全双低候选（AA-及以上 · 规模2~30亿 · 价≤130）Top 20</h2>
<p class="sub">入选条件：评级 ≥ AA-、规模 2~30 亿、价格 ≤ 130 元，按双低值由低到高排。「债底」为偏贵表示当前价已高于剩余本息锚的 105%，安全边际不足。</p>
<table><thead><tr><th>代码</th><th>名称</th><th class="th-num">价格</th><th class="th-num">双低值</th><th class="th-num">溢价率</th><th>评级</th><th class="th-num">规模(亿)</th><th>债底</th></tr></thead>
<tbody>{cand_rows}</tbody></table></div>

<div class="section"><h2>强赎关注区（价≥130）Top 10</h2>
<p class="note">⚠️ 这里是「高价高溢价的妖债/强赎博弈区」：价格远超 130、股性极强，正股回调时跌幅可能很大；<strong>非买入清单</strong>，仅作风险警示与观察。</p>
<table><thead><tr><th>代码</th><th>名称</th><th class="th-num">价格</th><th class="th-num">溢价率</th><th class="th-num">双低值</th><th>评级</th></tr></thead>
<tbody>{red_rows}</tbody></table></div>

<footer>数据来源：东方财富 · 仅为量化筛选，不构成投资建议<br>
策略：双低+安全债底+评级/规模过滤+估值水位总开关 · 自用研究</footer>
</body></html>"""
    with open(INDEX, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[done] index.html: {INDEX}")


def send_serverchan(key, stats, empty):
    if not key or requests is None:
        return
    title = f"可转债双低 {TODAY} · {'空仓观望' if empty else '可关注'}"
    desp = (f"估值水位：{stats['估值水位结论']}\n"
            f"可交易：{stats['可交易转债数']}只\n"
            f"价格中位数：{stats['价格中位数']}\n"
            f"双低值中位数：{stats['双低值中位数']}\n"
            f"双低<130数量：{stats['双低值<130数量']}\n"
            f"强赎关注区：{stats['价格>=130(强赎关注)数量']}只\n"
            f"空仓信号：{stats['双低失效空仓信号']}")
    try:
        r = requests.post(f"https://sctapi.ftqq.com/{key}.send",
                         data={"title": title, "desp": desp}, timeout=10)
        print("[info] Server酱推送:", "成功" if r.status_code == 200 else f"失败{r.status_code}")
    except Exception as e:
        print("[warn] Server酱推送异常:", e)


def main():
    df = build_df()
    alive, cand, redeem_watch, stats, empty = analyze(df)
    export_excel(alive, cand, redeem_watch, stats)
    cand[OUT_COLS].rename(columns=REN).to_csv(
        os.path.join(OUT, f"候选清单_{TODAY}.csv"), index=False, encoding="utf-8-sig")
    append_history(stats, empty)
    build_html(stats, cand, redeem_watch, empty)

    print("\n=== 市场统计 / 估值水位 ===")
    for k, v in stats.items():
        print(f"  {k}: {v}")
    if empty:
        print(f"\n⚠️ 双低<130合格标的不足 {EMPTY_SIGNAL_N} 只 → 触发【空仓观望】信号")

    send_serverchan(os.environ.get("SERVERCHAN_KEY"), stats, empty)


if __name__ == "__main__":
    main()
