#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
动态电价下的微网调度与储能响应
=====================================================================

定位
----
Q4-2：严格继承问题2最终V2“历史预测 + EWMA残差偏差校正 + MAD稳健风险缓冲 + 10分钟实时补救”框架，
      仅把固定电价扩展为动态电价信息。
Q4-3：严格继承问题3“0/6/12/18 多时刻PV预报 + 有偿合同调整”框架，仅增加动态电价信息。

价格信息层级（所有方案最终均按附件4真实动态价格结算，保证信息价值比较公平）
--------------------------------------------------------------------------
P1_FIXED    : 决策阶段仍使用附件1固定价格，仅作为低价格信息基准；最终按附件4结算。
P2_FORECAST : 正式可实施方案。0点日前合同仅使用历史价格预测；进入每个10分钟区间时，
              该区间报价视为已公布，可用于本区间实时控制/合同调整；未来区间仍沿用0点预测；最终按附件4结算。
P3_ORACLE   : 价格完全信息离线基准，决策时提前知道当天真实未来价格；仅用于价格完全信息Oracle对照，
              绝不能作为正式可实施策略，也不直接称为“全局理论下界”。

P2 默认价格预测：过去7日同一10分钟时段均值
    p_hat[d,t] = mean(p[d-1,t],...,p[d-7,t])
并提供 week / weekmean 两个简单可解释对照方法。

信息边界
--------
1. 负荷：Q4-2严格沿用Q2最终V2历史预测，并使用过去残差EWMA偏差校正与MAD风险缓冲；
   Q4-3沿用Q3历史负荷预测。当天未来真实负荷不会进入日前/合同调整优化。
2. 光伏：Q4-2沿用Q2历史预测；Q4-3使用附件3在0/6/12/18已发布的官方预报。
3. 价格：P2的0点日前申报规定在首区间报价揭示前完成，因此0点全天计划不读取当天任何真实价格；
   进入某10分钟区间时，假设该区间交易报价已公布，可用于本区间实时控制；Q4-3在6/12/18更新节点
   也只把“当前区间报价”作为新增价格信息，未来区间仍沿用0点历史预测，不滚动修正整条未来价格曲线。
4. 实际负荷/PV只在当前10分钟执行阶段进入实时补救；已执行时段不可回改。
5. SOC从2025-01-01 0:00的6000 kWh开始跨日连续，不每日重置。
6. P1/P2/P3受控比较时固定非价格条件：Q4-2的P1/P3默认复用P2逐日Policy；进入评价期后SOC由各方案独立执行递推。
   Q4-3价格信息比较固定同一更新时间组合、末端库存口径和同目标解选择规则。
7. 正式评价期为2025-02-01~12-31。为避免1月预热路径混入受控比较，P1/P2/P3及Q4-3各D方案在2月1日使用统一评价期初始SOC；
   Q4-3沿用Q3思想，以P2-B1在1月的共同预热终态作为统一起点。

结算
----
Q4-2：C = sum(p_real * q_plan) + 5*sum(p_real * e_emergency)
Q4-3：
  C_contract = sum[p_real*x + 1.5*p_real*(y-x)^+ - 0.5*p_real*(x-y)^+]
  C_total    = C_contract + 5*sum(p_real*e)
其中 x 为0点计划合同，y为最终有效合同。

输出
----
- 正式P2全年Q4-2：result4-2.xlsx
- 正式P2全年Q4-3：result4-3.xlsx
- daily_summary.csv / dispatch_detail.csv / run_summary.json
- 价格预测诊断、费用分解、储能价格响应指标、价格加权净负荷低估/PV高估事后暴露指标
- 若P1/P2/P3均已运行，可用 --mode compare 输出价格信息价值比较

推荐运行顺序
------------
1) python dynamic_price_dispatch.py --mode audit
2) python dynamic_price_dispatch.py --mode demo --branch q42 --price-info P2
3) python dynamic_price_dispatch.py --mode demo --branch q43 --price-info P2 --q43-scenarios core
4) python dynamic_price_dispatch.py --mode annual --branch q42 --price-info P2 --resume
5) python dynamic_price_dispatch.py --mode annual --branch q43 --price-info P2 --q43-scenarios all-d --resume
6) 如需论文信息价值对照，再分别跑P1、P3，最后：
   python dynamic_price_dispatch.py --mode compare --branch q42
   python dynamic_price_dispatch.py --mode compare --branch q43 --compare-q43-scenario auto

依赖
----
python -m pip install numpy scipy openpyxl matplotlib
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import re
import shutil
import time
import warnings
from dataclasses import dataclass, asdict
from datetime import date, datetime, time as dtime, timedelta
from functools import lru_cache
from pathlib import Path

import numpy as np
from openpyxl import load_workbook
from scipy.optimize import milp, Bounds, LinearConstraint
from scipy.sparse import lil_matrix, hstack, vstack, csr_matrix

# =====================================================================
# 0. 常量：与Q2/Q3保持一致
# =====================================================================
T = 144
DT = 1 / 6
ETA = 0.9
EMIN, EMAX, EINITIAL = 1200.0, 10800.0, 6000.0
B = 5000.0 * DT                      # 每10分钟最大充/放电量 kWh
TOL = 2e-4
ISSUE_HOURS = (0, 6, 12, 18)
LEX_COST_EPS = 1e-6
EVAL_START = date(2025, 2, 1)
EVAL_END = date(2025, 12, 31)
ALGORITHM_VERSION = "q4-dynamic-price-v3-q42-v2-inherited"
Q43_PROTOCOL_VERSION = "q43-common-evaluation-warmup-v1"

P1 = "P1"
P2 = "P2"
P3 = "P3"
PRICE_INFOS = (P1, P2, P3)

# =====================================================================
# 1. 数据结构
# =====================================================================
@dataclass(frozen=True)
class Policy:
    load_mode: str = "week_mean"
    pv_days: int = 3
    reserve: float = 0.0
    mu_factor: float = 0.0


@dataclass
class Data:
    dates: list[date]
    load: np.ndarray                 # kWh, (365,144)
    pv: np.ndarray                   # kWh, (365,144)
    fixed_price: np.ndarray          # 元/kWh, (144,)
    price_real: np.ndarray           # 元/kWh, (365,144)
    forecasts: dict[datetime, np.ndarray]  # 附件3，24个整点功率(kW)
    template42: Path | None
    template43: Path | None
    template42_labels: list[str] | None
    template43_labels: list[str] | None
    files: dict[str, Path]


# =====================================================================
# 2. 通用工具
# =====================================================================
def require(condition, message):
    if not condition:
        raise ValueError(message)


def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip().split()[0].replace("/", "-").replace(".", "-")
    parts = s.split("-")
    if len(parts) == 3:
        s = "-".join([parts[0], parts[1].zfill(2), parts[2].zfill(2)])
    return datetime.strptime(s, "%Y-%m-%d").date()


def time_minute(value):
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, dtime):
        return value.hour * 60 + value.minute
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return round(float(value) * 1440)
    s = str(value).strip().replace("：", ":")
    next_day = "+1" in s
    s = s.replace("+1", "")
    h, m = s.split(":")[:2]
    return int(h) * 60 + int(m) + (1440 if next_day else 0)


def clock(minutes):
    # 论文物理区间保留24:00写法；超过24h时才回绕。
    if minutes == 1440:
        return "24:00"
    if minutes > 1440 or minutes < 0:
        minutes %= 1440
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def raw_sample_label(t):
    end = (t + 1) * 10
    return "0:00+1" if end == 1440 else clock(end)


def physical_label(t):
    return f"{clock(t * 10)}-{clock((t + 1) * 10)}"


def write_csv(path: Path, rows, fieldnames=None):
    rows = list(rows)
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = list(fieldnames or rows[0].keys())
    seen = set(fields)
    for row in rows:
        for k in row:
            if k not in seen:
                fields.append(k)
                seen.add(k)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def write_json(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")


def file_sha(path: Path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def resolve_xlsx(folder: Path, stem: str, optional=False, include_attachment5=True):
    roots = [folder]
    if include_attachment5:
        roots += [folder / "附件5", folder / "附件"]
        roots += [folder / "附件" / "附件5"]
    pattern = re.compile(re.escape(stem) + r"(?:\(\d+\)|（\d+）)*\.xlsx", re.I)
    found = []
    for root in roots:
        if root.exists():
            found += [p for p in root.iterdir() if p.is_file() and pattern.fullmatch(p.name)]
    found = sorted(set(found))
    if not found and optional:
        return None
    require(len(found) == 1, f"{stem}.xlsx 应有且仅有一个可用版本，找到：{found}")
    return found[0]


# =====================================================================
# 3. 读取附件1/2/3/4与官方模板
# =====================================================================
def _actual_sheet(ws):
    rows = [r for r in ws.values if any(v is not None for v in r)]
    require([time_minute(v) for v in rows[0][1:145]] == list(range(10, 1441, 10)),
            f"{ws.title}表头不是00:10~24:00右端点序列")
    dates = [as_date(r[0]) for r in rows[1:]]
    expected = [date(2025, 1, 1) + timedelta(days=i) for i in range(365)]
    require(dates == expected, f"{ws.title}日期缺失/重复/乱序")
    a = np.asarray([r[1:145] for r in rows[1:]], dtype=float)
    require(a.shape == (365, T) and np.isfinite(a).all() and (a >= 0).all(),
            f"{ws.title}存在缺失/负值/异常形状 {a.shape}")
    return dates, a * DT


def read_pv_forecasts(path: Path, dates):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb.worksheets[0].values)
    finally:
        wb.close()
    expected_header = ["日期", "预报时刻"] + [f"预报{h}小时" for h in range(1, 25)]
    require(list(rows[0][:26]) == expected_header, "附件3表头异常")
    forecasts = {}
    current_date = None
    for row_index, row in enumerate(rows[1:], 2):
        if all(v is None or v == "" for v in row):
            continue
        if row[0] is not None and str(row[0]).strip():
            current_date = as_date(row[0])
        require(current_date is not None, f"附件3第{row_index}行无法继承日期")
        minute = time_minute(row[1])
        require(minute in (0, 360, 720, 1080), f"附件3第{row_index}行预报时刻异常")
        issue = datetime.combine(current_date, dtime()) + timedelta(minutes=minute)
        vals = np.asarray(row[2:26], dtype=float)
        require(vals.shape == (24,) and np.isfinite(vals).all() and (vals >= 0).all(),
                f"附件3第{row_index}行预测值异常")
        require(issue not in forecasts, f"附件3重复发布时间 {issue}")
        forecasts[issue] = vals
    expected = {datetime.combine(d, dtime(h)) for d in dates for h in ISSUE_HOURS}
    require(set(forecasts) == expected, "附件3应包含365×4条完整发布记录")
    return forecasts


def read_dynamic_price(path: Path, dates):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = [r for r in wb.worksheets[0].values if any(v is not None for v in r)]
    finally:
        wb.close()
    require(len(rows) == 366 and len(rows[0]) >= 145, "附件4形状异常")
    require([time_minute(v) for v in rows[0][1:145]] == list(range(10, 1441, 10)),
            "附件4时间表头不是00:10~24:00")
    ds = [as_date(r[0]) for r in rows[1:]]
    require(ds == dates, "附件4日期与附件2不一致")
    p = np.asarray([r[1:145] for r in rows[1:]], dtype=float)
    require(p.shape == (365, T) and np.isfinite(p).all() and (p > 0).all(),
            "附件4存在非正/缺失/异常电价")
    return p


def read_fixed_price(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = [r for r in wb.worksheets[0].values if r[0] is not None]
    finally:
        wb.close()
    require(len(rows) == 145, "附件1应有1行表头+144行")
    require([time_minute(r[0]) for r in rows[1:]] == list(range(10, 1441, 10)), "附件1时间异常")
    p = np.asarray([r[1] for r in rows[1:]], dtype=float)
    require(p.shape == (T,) and np.isfinite(p).all() and (p > 0).all(), "附件1固定价格异常")
    return p


def _validate_template(path: Path, q43: bool):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        expected_sheets = ["计划购电量", "调整购电量", "充放电量", "紧急购电量"] if q43 \
                          else ["计划购电量", "充放电量", "紧急购电量"]
        require(wb.sheetnames == expected_sheets, f"{path.name}工作表异常：{wb.sheetnames}")
        ws = wb["计划购电量"]
        labels = [str(v) for v in next(ws.values)[1:145]]
        require(len(labels) == T, f"{path.name}计划购电量时间列不是144个")
        dates = [as_date(r[0]) for r in list(ws.values)[1:] if r[0] not in (None, "")]
        require(dates == [EVAL_START + timedelta(days=i) for i in range(334)],
                f"{path.name}日期不是2025-02-01~12-31")
        if q43:
            other = [str(v) for v in next(wb["调整购电量"].values)[1:145]]
            require(other == labels, "result4-3计划/调整表头不一致")
        return labels
    finally:
        wb.close()


def read_data(folder: Path):
    folder = Path(folder)
    require(folder.exists(), f"数据目录不存在：{folder}")
    files = {}
    for stem in ("附件1", "附件2", "附件3", "附件4"):
        files[stem] = resolve_xlsx(folder, stem, include_attachment5=False)

    wb = load_workbook(files["附件2"], read_only=True, data_only=True)
    try:
        ds, load = _actual_sheet(wb["小区负载"])
        ds2, pv = _actual_sheet(wb["光伏发电实际功率"])
    finally:
        wb.close()
    require(ds == ds2, "附件2负荷和PV日期不一致")

    fixed_price = read_fixed_price(files["附件1"])
    forecasts = read_pv_forecasts(files["附件3"], ds)
    price_real = read_dynamic_price(files["附件4"], ds)

    t42 = resolve_xlsx(folder, "result4-2", optional=True)
    t43 = resolve_xlsx(folder, "result4-3", optional=True)
    labels42 = _validate_template(t42, False) if t42 else None
    labels43 = _validate_template(t43, True) if t43 else None
    if t42: files["result4-2"] = t42
    if t43: files["result4-3"] = t43

    return Data(ds, load, pv, fixed_price, price_real, forecasts,
                t42, t43, labels42, labels43, files)


# =====================================================================
# 4. 负荷/PV预测：与Q2/Q3保持一致
# =====================================================================
def forecast(history_l, history_v, policy: Policy):
    """只允许传入目标日前历史数据。"""
    n = len(history_l)
    if n == 0:
        raise ValueError("Bootstrap day must not request a forecast")
    if n >= 14 and policy.load_mode == "week_mean":
        l = (history_l[-7] + history_l[-14]) / 2
    elif n >= 7:
        l = history_l[-7].copy()
    else:
        l = history_l.mean(axis=0)
    v = history_v[-min(n, policy.pv_days):].mean(axis=0)
    return l.copy(), v.copy()


def known_anchor(data: Data, issue: datetime):
    idx = (issue.date() - data.dates[0]).days
    require(0 <= idx < 365, "PV锚点日期越界")
    if issue.hour:
        return float(data.pv[idx, issue.hour * 6 - 1] / DT)
    if idx > 0:
        return float(data.pv[idx - 1, -1] / DT)
    raise ValueError("1月1日0点没有前一日PV观测锚点")


def get_pv_forecast(issue_time: datetime, decision_time: datetime, issued_values, anchor_kw):
    require(issue_time == decision_time, "PV预报只允许在其正式发布时间启用")
    require(issue_time.hour in ISSUE_HOURS and issue_time.minute == 0, "无效PV预报发布时间")
    values = np.asarray(issued_values, dtype=float)
    require(values.shape == (24,) and np.isfinite(values).all(), "PV预报必须为24个整点")
    steps = (24 - issue_time.hour) * 6
    powers = np.interp(np.arange(1, steps + 1) / 6, np.arange(25), np.r_[anchor_kw, values])
    return powers * DT


# =====================================================================
# 5. 价格预测与信息层级
# =====================================================================
def historical_price_forecast(data: Data, day_idx: int, method="mean7"):
    """目标日day_idx的144点未来价格预测，只使用day_idx之前真实价格。"""
    require(day_idx > 0, "首日没有历史价格，不能做P2预测")
    hist = data.price_real[:day_idx]
    if method == "mean7":
        return hist[-min(7, len(hist)):].mean(axis=0).copy()
    if method == "week":
        if day_idx >= 7:
            return data.price_real[day_idx - 7].copy()
        return hist.mean(axis=0).copy()
    if method == "weekmean":
        if day_idx >= 14:
            return ((data.price_real[day_idx - 7] + data.price_real[day_idx - 14]) / 2).copy()
        if day_idx >= 7:
            return data.price_real[day_idx - 7].copy()
        return hist.mean(axis=0).copy()
    raise ValueError(f"未知价格预测方法：{method}")


def base_decision_price(data: Data, day_idx: int, price_info: str, method="mean7"):
    """返回0点可用于未来整日优化的价格路径。P3仅供离线完全信息基准。"""
    require(price_info in PRICE_INFOS, f"未知价格信息层级 {price_info}")
    if price_info == P1:
        return data.fixed_price.copy()
    if price_info == P2:
        return historical_price_forecast(data, day_idx, method)
    return data.price_real[day_idx].copy()  # P3 Oracle


def decision_price_horizon(data: Data, day_idx: int, start: int, price_info: str,
                           method="mean7", observe_current=True):
    """
    从start开始的决策价格路径。
    P2：未来仍用0点历史预测；若observe_current=True，仅将当前第一个区间替换为已观测真实交易价。
    P1：完全忽略动态价格信息，始终用固定价。
    P3：使用真实未来价格（Oracle）。
    """
    p = base_decision_price(data, day_idx, price_info, method)[start:].copy()
    if price_info == P2 and observe_current and len(p):
        p[0] = data.price_real[day_idx, start]
    return p


def price_forecast_diagnostics(data: Data, method: str, output: Path | None = None):
    rows = []
    errors = []
    for i, ds in enumerate(data.dates):
        if ds < EVAL_START:
            continue
        pred = historical_price_forecast(data, i, method)
        actual = data.price_real[i]
        err = pred - actual
        errors.append(err)
        rows.append(dict(
            date=str(ds),
            mae=float(np.mean(np.abs(err))),
            rmse=float(np.sqrt(np.mean(err ** 2))),
            bias=float(np.mean(err)),
            actual_mean=float(np.mean(actual)),
            actual_min=float(np.min(actual)),
            actual_max=float(np.max(actual)),
            actual_peak_valley=float(np.max(actual) - np.min(actual)),
        ))
    e = np.vstack(errors)
    summary = dict(
        method=method,
        days=len(rows),
        mae=float(np.mean(np.abs(e))),
        rmse=float(np.sqrt(np.mean(e ** 2))),
        bias=float(np.mean(e)),
    )
    if output:
        write_csv(output / "price_forecast_daily.csv", rows)
        write_json(output / "price_forecast_summary.json", summary)
    return summary


# =====================================================================
# 6. MILP核心：继承Q2/Q3
# =====================================================================
@lru_cache(maxsize=144)
def skeleton(n):
    # blocks: purchase/emergency, charge, discharge, unused, S_end, binary, reserve_shortfall; terminal shortfall
    nv = 7 * n + 1
    a = lil_matrix((5 * n + 1, nv), dtype=float)
    for t in range(n):
        a[t, t] = 1
        a[t, n + t] = -1
        a[t, 2 * n + t] = 1
        a[t, 3 * n + t] = -1
        a[n + t, 4 * n + t] = 1
        if t:
            a[n + t, 4 * n + t - 1] = -1
        a[n + t, n + t] = -ETA
        a[n + t, 2 * n + t] = 1 / ETA
        a[2 * n + t, n + t] = 1
        a[2 * n + t, 5 * n + t] = -B
        a[3 * n + t, 2 * n + t] = 1
        a[3 * n + t, 5 * n + t] = B
        a[4 * n + t, 4 * n + t] = -1
        a[4 * n + t, 6 * n + t] = -1
    a[5 * n, 5 * n - 1] = -1
    a[5 * n, -1] = -1
    lo = np.zeros(nv)
    hi = np.full(nv, np.inf)
    hi[n:3 * n] = B
    lo[4 * n:5 * n] = EMIN
    hi[4 * n:5 * n] = EMAX
    hi[5 * n:6 * n] = 1
    hi[6 * n:7 * n] = EMAX
    hi[-1] = EINITIAL
    integrality = np.zeros(nv)
    integrality[5 * n:6 * n] = 1
    return a.tocsr(), Bounds(lo, hi), integrality


def optimize(net, price, s0, policy: Policy, terminal_lambda, fixed_q=None,
             time_limit=30.0, lexicographic=True):
    net = np.asarray(net, dtype=float)
    price = np.asarray(price, dtype=float)
    n = len(net)
    require(price.shape == (n,), "price/net长度不一致")
    a, base_bounds, integer = skeleton(n)
    lb = np.array(base_bounds.lb, dtype=float, copy=True)
    ub = np.array(base_bounds.ub, dtype=float, copy=True)
    if fixed_q is not None:
        fixed_q = np.asarray(fixed_q, dtype=float)
        require(fixed_q.shape == (n,), "fixed_q长度不一致")
        # 紧急电只允许弥补计划购电后的正净缺口，不允许拿5倍紧急电给储能充电
        ub[:n] = np.maximum(net - fixed_q, 0.0)
    bounds = Bounds(lb, ub)
    rhs = np.zeros(5 * n + 1)
    rhs[:n] = net if fixed_q is None else net - fixed_q
    rhs[n] = s0
    rhs[3 * n:4 * n] = B
    rhs[4 * n:5 * n] = -(EMIN + policy.reserve)
    rhs[-1] = -EINITIAL
    lower = np.full_like(rhs, -np.inf)
    lower[:2 * n] = rhs[:2 * n]
    obj = np.zeros(7 * n + 1)
    obj[:n] = price * (1 if fixed_q is None else 5)
    obj[6 * n:7 * n] = policy.mu_factor * terminal_lambda * DT
    obj[-1] = terminal_lambda
    opts = {"time_limit": time_limit, "mip_rel_gap": 1e-7, "threads": 1, "mip_feasibility_tolerance": 1e-9}
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Unrecognized options detected.*")
        res = milp(obj, integrality=integer, bounds=bounds,
                   constraints=LinearConstraint(a, lower, rhs), options=opts)
    if not res.success or res.x is None:
        raise RuntimeError(f"MILP failed: {res.message}; horizon={n}, SOC={s0}")
    x = res.x
    primary = x.copy()
    primary_value = float(obj @ x)

    # 与Q3一致：在不恶化主经济目标的极小容差内，优先减少当前未利用电量，避免多解的物理怪象
    if fixed_q is not None and lexicographic and primary[3 * n] > TOL:
        second_obj = np.zeros_like(obj)
        second_obj[3 * n] = 1.0
        cost_limit = LinearConstraint(csr_matrix(obj.reshape(1, -1)), -np.inf, primary_value + LEX_COST_EPS)
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message="Unrecognized options detected.*")
            second = milp(second_obj, integrality=integer, bounds=bounds,
                          constraints=[LinearConstraint(a, lower, rhs), cost_limit],
                          options={"time_limit": time_limit, "mip_rel_gap": 1e-9, "threads": 1,
                                   "mip_feasibility_tolerance": 1e-9})
        if second.success and second.x is not None:
            cand = second.x
            if float(obj @ cand) <= primary_value + LEX_COST_EPS + 1e-5 and cand[3 * n] < primary[3 * n] - 1e-8:
                x = cand

    values = a @ x
    if np.max(values - rhs) > TOL or np.max(np.abs(values[:2 * n] - rhs[:2 * n])) > TOL:
        raise RuntimeError("MILP feasibility audit failed")
    if np.max(lb - x) > TOL or np.max(x - ub) > TOL:
        raise RuntimeError("MILP variable bound audit failed")
    return tuple(x[i * n:(i + 1) * n].copy() for i in range(5))


def current_action(lhat, vhat, current_l, current_v, q, t, s0, policy, price_horizon, lam, limit):
    net = (lhat[t:] - vhat[t:]).copy()
    net[0] = current_l - current_v
    arrays = optimize(net, price_horizon, s0, policy, lam, fixed_q=q[t:], time_limit=limit)
    return np.array([x[0] for x in arrays])


def audit_trace(x, l, v):
    q, e, c, d, w, s0, s1 = x.T
    emergency_cap = np.maximum(l - v - q, 0.0)
    checks = [
        np.max(np.abs(q + e + v + d - l - c - w)),
        np.max(np.abs(s1 - s0 - ETA * c + d / ETA)),
        np.max(np.abs(s1[:-1] - s0[1:])) if len(x) > 1 else 0.0,
        max(0.0, EMIN - min(s0.min(), s1.min())),
        max(0.0, max(s0.max(), s1.max()) - EMAX),
        max(0.0, c.max() - B, d.max() - B),
        max(0.0, -x[:, :5].min()),
        np.maximum(np.minimum(c, d), 0).max(),
        max(0.0, float(np.max(e - emergency_cap))),
    ]
    if max(checks) > TOL or not np.isfinite(x).all():
        raise AssertionError(f"Physical audit failed: {checks}")


def bootstrap(l, v, s0):
    out = np.zeros((T, 7))
    out[:, 1] = np.maximum(l - v, 0.0)
    out[:, 4] = np.maximum(v - l, 0.0)
    out[:, 5:] = s0
    audit_trace(out, l, v)
    return out


# =====================================================================
# 7. Q4-3合同调整与结算
# =====================================================================
def settlement_bill(original, final, emergency, price_real, rule="refund"):
    up = np.maximum(final - original, 0.0)
    down = np.maximum(original - final, 0.0)
    cp = float(price_real @ original)
    ca_inc = float(price_real @ (1.5 * up))
    ca_dec = float(price_real @ ((-0.5 if rule == "refund" else 0.5) * down))
    ce = float(5 * price_real @ emergency)
    return dict(
        plan_cost=cp,
        adjustment_increase_cost=ca_inc,
        adjustment_decrease_credit=ca_dec,  # refund口径下为负数
        adjustment_net_cost=ca_inc + ca_dec,
        contract_cost=cp + ca_inc + ca_dec,
        emergency_cost=ce,
        total_actual_cost=cp + ca_inc + ca_dec + ce,
        increase_kwh=float(up.sum()),
        decrease_kwh=float(down.sum()),
    )


def adjustment_problem(net, decision_price, soc, original, policy, lam, rule, limit, fixed=None):
    """Q3正式调整MILP；decision_price仅用于当前节点对未来的预测性成本比较。"""
    net = np.asarray(net, dtype=float)
    decision_price = np.asarray(decision_price, dtype=float)
    original = np.asarray(original, dtype=float)
    n = len(net)
    require(decision_price.shape == (n,) and original.shape == (n,), "调整问题维度不一致")
    base, bounds, integer = skeleton(n)
    nb = 7 * n + 1
    extra = lil_matrix((5 * n + 1, 3 * n))
    extra[:n, :n] = np.eye(n)  # emergency
    a = hstack([base, extra.tocsr()], format="csr")
    rows = lil_matrix((n, nb + 3 * n))
    for t in range(n):
        rows[t, t] = 1
        rows[t, nb + n + t] = -1       # increase
        rows[t, nb + 2 * n + t] = 1   # decrease
    a = vstack([a, rows.tocsr()], format="csr")
    rhs = np.zeros(6 * n + 1)
    rhs[:n] = net
    rhs[n] = soc
    rhs[3 * n:4 * n] = B
    rhs[4 * n:5 * n] = -(EMIN + policy.reserve)
    rhs[5 * n] = -EINITIAL
    rhs[5 * n + 1:] = original
    lower = np.full_like(rhs, -np.inf)
    lower[:2 * n] = rhs[:2 * n]
    lower[5 * n + 1:] = original
    lb = np.r_[bounds.lb, np.zeros(3 * n)]
    ub = np.r_[bounds.ub, np.full(3 * n, np.inf)]
    if fixed is not None:
        fixed = np.asarray(fixed, dtype=float)
        lb[:n] = fixed
        ub[:n] = fixed
    objective = np.zeros(nb + 3 * n)
    objective[6 * n:7 * n] = policy.mu_factor * lam * DT
    objective[nb - 1] = lam
    objective[nb:nb + n] = 5 * decision_price
    objective[nb + n:nb + 2 * n] = 1.5 * decision_price
    objective[nb + 2 * n:] = (-0.5 if rule == "refund" else 0.5) * decision_price
    integrality = np.r_[integer, np.zeros(3 * n)]
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Unrecognized options detected.*")
        result = milp(objective, integrality=integrality, bounds=Bounds(lb, ub),
                      constraints=LinearConstraint(a, lower, rhs),
                      options={"time_limit": limit, "mip_rel_gap": 1e-7, "threads": 1})
    if not result.success or result.x is None:
        raise RuntimeError(f"调整MILP失败：{result.message}; horizon={n}")
    x = result.x
    ax = a @ x
    finite = np.isfinite(lower)
    violation = max(float(np.max(ax - rhs)), float(np.max(lower[finite] - ax[finite])),
                    float(np.max(lb - x)), float(np.max(x - ub)))
    require(violation < TOL, "调整MILP可行性审计失败")
    q = np.maximum(x[:n], 0.0)
    # 原合同基础费用 + 调整/紧急/末端项，供保持 vs 调整公平比较
    value = float(result.fun + decision_price @ original)
    return q, value


def decide_revision(net, decision_price, soc, original, current, policy, lam, rule, limit):
    q, new_value = adjustment_problem(net, decision_price, soc, original, policy, lam, rule, limit)
    _, keep_value = adjustment_problem(net, decision_price, soc, original, policy, lam, rule, limit, fixed=current)
    improvement = keep_value - new_value
    tol = max(1e-5, 1e-7 * max(abs(keep_value), abs(new_value), 1.0))
    require(improvement >= -2 * tol, "允许调整解反而劣于保持解，检查求解精度")
    chosen = q if improvement > tol else current.copy()
    return chosen, dict(
        predicted_keep_objective=keep_value,
        predicted_adjust_objective=new_value,
        predicted_net_improvement=max(0.0, improvement),
        changed=bool(np.max(np.abs(chosen - current)) > TOL),
        numerical_tolerance=tol,
    )


# =====================================================================
# 8. 价格-储能行为指标
# =====================================================================
def response_metrics(trace, actual_price):
    c = trace[:, 2]
    d = trace[:, 3]
    low = float(np.quantile(actual_price, 0.25))
    high = float(np.quantile(actual_price, 0.75))
    csum = float(c.sum())
    dsum = float(d.sum())
    low_charge = float(c[actual_price <= low].sum())
    high_discharge = float(d[actual_price >= high].sum())
    avg_c = float(actual_price @ c / csum) if csum > 1e-9 else None
    avg_d = float(actual_price @ d / dsum) if dsum > 1e-9 else None
    return dict(
        price_q25=low,
        price_q75=high,
        charge_kwh=csum,
        discharge_kwh=dsum,
        low_price_charge_kwh=low_charge,
        high_price_discharge_kwh=high_discharge,
        low_price_charge_ratio=(low_charge / csum if csum > 1e-9 else None),
        high_price_discharge_ratio=(high_discharge / dsum if dsum > 1e-9 else None),
        avg_charge_price=avg_c,
        avg_discharge_price=avg_d,
        # ETA是单向效率，因此效率修正套利判据使用往返效率 ETA^2
        efficiency_adjusted_arbitrage_margin=(avg_d - avg_c / (ETA * ETA)
                                               if avg_c is not None and avg_d is not None else None),
    )


def price_weighted_exposure_indicators(actual_l, actual_pv, forecast_l, forecast_pv, actual_price):
    """事后误差暴露指标，不是概率风险度量，也不等同于实际经济损失。

    所有负荷/PV输入在本程序中均为10分钟区间电量(kWh)，因此这里不再乘DT。
    - 净负荷低估：actual_net - forecast_net > 0
    - PV高估：forecast_pv - actual_pv > 0
    两者乘真实价格后单位为“元等价暴露量”，仅用于解释高价时段误差暴露。
    """
    actual_net = np.asarray(actual_l) - np.asarray(actual_pv)
    pred_net = np.asarray(forecast_l) - np.asarray(forecast_pv)
    net_under = np.maximum(actual_net - pred_net, 0.0)
    pv_over = np.maximum(np.asarray(forecast_pv) - np.asarray(actual_pv), 0.0)
    price = np.asarray(actual_price)
    return dict(
        price_weighted_net_under_exposure=float(price @ net_under),
        net_under_kwh=float(net_under.sum()),
        price_weighted_pv_over_exposure=float(price @ pv_over),
        pv_over_kwh=float(pv_over.sum()),
    )


# =====================================================================
# 8.1 Q2最终V2预测/风险模块：Q4-2原样继承非价格逻辑
# =====================================================================
def historical_net_residual_matrix_q42_v2(data: Data, day_idx: int, policy: Policy,
                                           window: int = 28):
    """仅用目标日前历史数据构造净负荷预测残差(actual-forecast)，严格past-only。"""
    window = int(window)
    require(window >= 1, "V2 residual window必须为正")
    if day_idx <= 1:
        return np.zeros((0, T), dtype=float), np.zeros(0, dtype=int)
    start = max(1, day_idx - window)
    residuals, indices = [], []
    for j in range(start, day_idx):
        lh, vh = forecast(data.load[:j], data.pv[:j], policy)
        residuals.append((data.load[j] - data.pv[j]) - (lh - vh))
        indices.append(j)
    if not residuals:
        return np.zeros((0, T), dtype=float), np.zeros(0, dtype=int)
    return np.vstack(residuals), np.asarray(indices, dtype=int)


def q42_v2_net_profiles(data: Data, day_idx: int, policy: Policy, args):
    """严格继承Q2最终V2：基础预测 + EWMA残差偏差校正 + MAD稳健风险缓冲。

    日前计划使用 corrected_net + risk_buffer；
    实时未来使用 corrected_net，避免把风险缓冲重复计入实时控制。
    所有输入/输出均为10分钟区间电量(kWh)。
    """
    lhat, vhat = forecast(data.load[:day_idx], data.pv[:day_idx], policy)
    base = lhat - vhat
    residuals, idx = historical_net_residual_matrix_q42_v2(
        data, day_idx, policy, args.v2_residual_window)
    if len(residuals) == 0:
        z = np.zeros(T, dtype=float)
        return lhat.copy(), vhat.copy(), base.copy(), base.copy(), z, {
            "bias": z.copy(), "sigma": z.copy(), "tail_cap": z.copy(), "history_days": 0
        }

    ages = (day_idx - 1) - idx
    half = float(args.v2_bias_half_life)
    weights = np.power(0.5, ages / half)
    weights = weights / weights.sum()
    bias_raw = np.sum(residuals * weights[:, None], axis=0)
    bias = float(args.v2_bias_shrink) * bias_raw

    centered = residuals - bias[None, :]
    med = np.median(centered, axis=0)
    mad = 1.4826 * np.median(np.abs(centered - med[None, :]), axis=0)
    std = np.std(centered, axis=0)
    # 与Q2最终V2一致：MAD为主，0.35*std避免小样本/量化数据出现零风险估计。
    sigma = np.maximum(mad, 0.35 * std)
    raw_buffer = float(args.v2_risk_k) * sigma
    tail = np.quantile(centered, float(args.v2_tail_quantile), axis=0)
    tail_cap = np.maximum(tail, 0.0)
    buffer = np.minimum(np.maximum(raw_buffer, 0.0), tail_cap)

    corrected = base + bias
    return lhat.copy(), vhat.copy(), base.copy(), corrected.copy(), buffer.copy(), {
        "bias": bias.copy(), "sigma": sigma.copy(), "tail_cap": tail_cap.copy(),
        "history_days": int(len(residuals)),
    }


def optimize_q42_v2(net, price, s0, policy: Policy, terminal_lambda,
                    fixed_q=None, time_limit=30.0):
    """Q4-2专用求解器：保持Q2最终V2的主MILP与同目标解规则，不引入Q3二级词典序。"""
    net = np.asarray(net, dtype=float)
    price = np.asarray(price, dtype=float)
    n = len(net)
    require(price.shape == (n,), "Q4-2 V2 price/net长度不一致")
    a, base_bounds, integer = skeleton(n)
    lb = np.array(base_bounds.lb, dtype=float, copy=True)
    ub = np.array(base_bounds.ub, dtype=float, copy=True)
    if fixed_q is not None:
        fixed_q = np.asarray(fixed_q, dtype=float)
        require(fixed_q.shape == (n,), "Q4-2 V2 fixed_q长度不一致")
        ub[:n] = np.maximum(net - fixed_q, 0.0)
    bounds = Bounds(lb, ub)
    rhs = np.zeros(5 * n + 1)
    rhs[:n] = net if fixed_q is None else net - fixed_q
    rhs[n] = s0
    rhs[3 * n:4 * n] = B
    rhs[4 * n:5 * n] = -(EMIN + policy.reserve)
    rhs[-1] = -EINITIAL
    lower = np.full_like(rhs, -np.inf)
    lower[:2 * n] = rhs[:2 * n]
    obj = np.zeros(7 * n + 1)
    obj[:n] = price * (1 if fixed_q is None else 5)
    obj[6 * n:7 * n] = policy.mu_factor * terminal_lambda * DT
    obj[-1] = terminal_lambda
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Unrecognized options detected.*")
        res = milp(
            obj, integrality=integer, bounds=bounds,
            constraints=LinearConstraint(a, lower, rhs),
            options={"time_limit": time_limit, "mip_rel_gap": 1e-7, "threads": 1},
        )
    if not res.success or res.x is None:
        raise RuntimeError(f"Q4-2 V2 MILP failed: {res.message}; horizon={n}, SOC={s0}")
    x = res.x
    values = a @ x
    if np.max(values - rhs) > TOL or np.max(np.abs(values[:2 * n] - rhs[:2 * n])) > TOL:
        raise RuntimeError("Q4-2 V2 MILP feasibility audit failed")
    return tuple(x[i * n:(i + 1) * n].copy() for i in range(5))


def current_action_net_q42_v2(realtime_net, current_net, q, t, s0, policy,
                               price_horizon, lam, limit):
    """Q2-V2实时动作：当前净负荷用真实值，未来仍用V2校正预测；价格规则由Q4传入。"""
    net = np.asarray(realtime_net[t:], dtype=float).copy()
    require(len(net) == len(price_horizon), "Q4-2 V2实时净负荷与价格预测长度不一致")
    net[0] = float(current_net)
    arrays = optimize_q42_v2(net, price_horizon, s0, policy, lam,
                             fixed_q=q[t:], time_limit=limit)
    return np.array([x[0] for x in arrays])


def q42_v2_exposure_indicators(actual_l, actual_pv, realtime_net, base_pv_forecast,
                                actual_price):
    """Q4-2 V2事后误差暴露。

    净负荷低估以V2实时未来使用的corrected_net为预测基准；
    PV高估仍使用Q2基础PV预测，因为V2偏差校正作用于净负荷而非单独PV分量。
    """
    actual_net = np.asarray(actual_l) - np.asarray(actual_pv)
    net_under = np.maximum(actual_net - np.asarray(realtime_net), 0.0)
    pv_over = np.maximum(np.asarray(base_pv_forecast) - np.asarray(actual_pv), 0.0)
    price = np.asarray(actual_price)
    return dict(
        price_weighted_net_under_exposure=float(price @ net_under),
        net_under_kwh=float(net_under.sum()),
        price_weighted_pv_over_exposure=float(price @ pv_over),
        pv_over_kwh=float(pv_over.sum()),
    )


# =====================================================================
# 9. Q4-2：严格继承Q2最终V2日前—实时框架
# =====================================================================
def q2_candidates():
    return [Policy(l, k, r, m)
            for l in ["week", "week_mean"]
            for k in [3, 7]
            for r, m in [(0.0, 0.0), (600.0, 0.10)]]


def simulate_q42_day(data: Data, day_idx: int, s0: float, policy: Policy,
                     price_info: str, price_method: str, lam: float, limit: float, args):
    """Q4-2：完整继承Q2最终V2的非价格逻辑，仅替换/扩展价格信息层。

    - 日前净负荷：corrected_net + risk_buffer
    - 实时未来净负荷：corrected_net
    - 当前净负荷：当前真实负荷-PV
    - 价格：严格沿用P1/P2/P3既定规则，不在此函数内改写
    """
    ds = data.dates[day_idx]
    require(day_idx > 0, "Q4-2正式预测从第2天起")
    lhat, vhat, base_net, corrected_net, risk_buffer, meta = q42_v2_net_profiles(
        data, day_idx, policy, args)
    plan_net = corrected_net + risk_buffer

    # 日前申报在首区间报价揭示前完成：0点计划不得读取当天第一个真实报价。
    p0 = decision_price_horizon(
        data, day_idx, 0, price_info, price_method, observe_current=False)
    q = optimize_q42_v2(plan_net, p0, s0, policy, lam, time_limit=limit)[0]

    out = np.zeros((T, 7))
    s = s0
    for t in range(T):
        # P2仅当前区间替换为真实报价，未来仍沿用0点价格预测；P1/P3规则保持原定义。
        ph = decision_price_horizon(
            data, day_idx, t, price_info, price_method, observe_current=True)
        current_net = float(data.load[day_idx, t] - data.pv[day_idx, t])
        e, c, d, w, snext = current_action_net_q42_v2(
            corrected_net, current_net, q, t, s, policy, ph, lam, limit)
        out[t] = q[t], e, c, d, w, s, snext
        s = snext

    audit_trace(out, data.load[day_idx], data.pv[day_idx])
    exposure = q42_v2_exposure_indicators(
        data.load[day_idx], data.pv[day_idx], corrected_net, vhat,
        data.price_real[day_idx])
    return dict(
        trace=out,
        load_forecast=lhat,
        pv_forecast=vhat,
        base_net_forecast=base_net,
        corrected_net_forecast=corrected_net,
        risk_buffer=risk_buffer,
        plan_net_forecast=plan_net,
        v2_bias=meta["bias"],
        v2_sigma=meta["sigma"],
        v2_tail_cap=meta["tail_cap"],
        v2_history_days=int(meta["history_days"]),
        price_forecast=base_decision_price(data, day_idx, price_info, price_method),
        **exposure,
        date=str(ds), policy=policy,
    )

def q42_day_summary(data: Data, day_idx: int, result, price_info):
    x = result["trace"]
    p = data.price_real[day_idx]
    normal = float(p @ x[:, 0])
    emergency = float(5 * p @ x[:, 1])
    trajectory = np.r_[x[0, 5], x[:, 6]]
    m = response_metrics(x, p)

    actual_net = data.load[day_idx] - data.pv[day_idx]
    base_net = np.asarray(result["base_net_forecast"], dtype=float)
    corrected_net = np.asarray(result["corrected_net_forecast"], dtype=float)
    buffer = np.asarray(result["risk_buffer"], dtype=float)
    base_err = base_net - actual_net
    corr_err = corrected_net - actual_net

    return dict(
        date=str(data.dates[day_idx]),
        price_info=price_info,
        normal_cost=normal,
        plan_cost=normal,
        adjustment_net_cost=0.0,
        emergency_cost=emergency,
        total_actual_cost=normal + emergency,
        plan_kwh=float(x[:, 0].sum()),
        emergency_kwh=float(x[:, 1].sum()),
        emergency_slots=int((x[:, 1] > 1e-6).sum()),
        soc_start=float(x[0, 5]), soc_end=float(x[-1, 6]),
        soc_min_kwh=float(trajectory.min()), soc_max_kwh=float(trajectory.max()),
        price_weighted_net_under_exposure=float(result["price_weighted_net_under_exposure"]),
        net_under_kwh=float(result["net_under_kwh"]),
        price_weighted_pv_over_exposure=float(result["price_weighted_pv_over_exposure"]),
        pv_over_kwh=float(result["pv_over_kwh"]),
        # Q2 V2诊断：与最终Q2输出口径一致，kWh/DT还原为kW。
        base_net_mae_kw=float(np.mean(np.abs(base_err)) / DT),
        corrected_net_mae_kw=float(np.mean(np.abs(corr_err)) / DT),
        base_net_bias_kw=float(np.mean(base_err) / DT),
        corrected_net_bias_kw=float(np.mean(corr_err) / DT),
        mean_abs_bias_correction_kw=float(np.mean(np.abs(result["v2_bias"])) / DT),
        mean_risk_buffer_kw=float(np.mean(buffer) / DT),
        max_risk_buffer_kw=float(np.max(buffer) / DT),
        v2_history_days=int(result["v2_history_days"]),
        **{k: v for k, v in m.items() if v is not None},
        **asdict(result["policy"]),
    )


def tune_q42(data: Data, day_idx: int, snapshots: dict[int, float], args, lam, price_info):
    """按Q2最终V2口径滚动选择Policy；价格信息层只由price_info控制。"""
    start = max(14, day_idx - args.validation_days)
    if start >= day_idx:
        return Policy()
    scored = []
    for pol in q2_candidates():
        state = snapshots[start]
        total = 0.0
        for j in range(start, day_idx):
            rr = simulate_q42_day(
                data, j, state, pol, price_info, args.price_method,
                lam, args.time_limit, args)
            x = rr["trace"]
            p = data.price_real[j]
            total += float(p @ x[:, 0] + 5 * p @ x[:, 1])
            state = float(x[-1, 6])
        # 与Q2最终V2一致：所有候选使用同一线性末端库存修正，避免验证窗末端放空偏好。
        correction = lam * max(EINITIAL - state, 0.0)
        scored.append((total + correction, pol))
    return min(scored, key=lambda z: z[0])[1]


# =====================================================================
# 10. Q4-3：继承Q3多时刻合同调整
# =====================================================================
def scenario_config(name):
    if name == "B1":
        return True, (), False
    require(name.startswith("D"), f"Q4-3仅支持B1/D系列，未知方案 {name}")
    hours = tuple(int(x) for x in name[1:].split("_"))
    require(hours == tuple(sorted(set(hours))) and all(h in (6, 12, 18) for h in hours),
            f"非法Q4-3更新时刻：{name}")
    return True, hours, True


def midnight_plan_q43(data: Data, day_idx: int, soc: float, name: str,
                      price_info: str, price_method: str, lam: float, limit: float):
    ds = data.dates[day_idx]
    policy = Policy()
    lh, _ = forecast(data.load[:day_idx], data.pv[:day_idx], policy)
    issue = datetime.combine(ds, dtime())
    vh = get_pv_forecast(issue, issue, data.forecasts[issue], known_anchor(data, issue))
    # 0点原始合同在首区间报价揭示前申报；当天真实价格不进入日前合同。
    p = decision_price_horizon(data, day_idx, 0, price_info, price_method, observe_current=False)
    q = optimize(lh - vh, p, soc, policy, lam, time_limit=limit)[0]
    return np.maximum(q, 0.0), lh, vh, p


def simulate_q43_day(data: Data, day_idx: int, soc: float, name: str,
                     price_info: str, price_method: str, lam: float, limit: float,
                     settlement_rule="refund"):
    ds = data.dates[day_idx]
    original, lh, vh, p0 = midnight_plan_q43(data, day_idx, soc, name, price_info, price_method, lam, limit)
    current = original.copy()
    _, hours, allow = scenario_config(name)
    policy = Policy()
    trace = np.zeros((T, 7))
    used_pv = np.zeros(T)
    issue_at = np.zeros(T, dtype=int)
    update_logs = []
    revisions = []
    s = soc
    last_hour = 0

    for t in range(T):
        h = t // 6
        if t % 36 == 0 and h in hours:
            issue = datetime.combine(ds, dtime(h))
            vh[t:] = get_pv_forecast(issue, issue, data.forecasts[issue], known_anchor(data, issue))
            last_hour = h
            old = current[t:].copy()
            log = dict(date=str(ds), scenario=name, issue_hour=h, soc_at_update=float(s),
                       price_info=price_info)
            if allow:
                ph = decision_price_horizon(data, day_idx, t, price_info, price_method, observe_current=True)
                chosen, diag = decide_revision(lh[t:] - vh[t:], ph, s, original[t:], old,
                                               policy, lam, settlement_rule, limit)
                current[t:] = chosen
                log.update(diag)
                log["revision_l1_kwh"] = float(np.abs(chosen - old).sum())
            else:
                log.update(changed=False, revision_l1_kwh=0.0,
                           predicted_keep_objective=None, predicted_adjust_objective=None,
                           predicted_net_improvement=None, numerical_tolerance=None)
            update_logs.append(log)
            for j in range(t, T):
                revisions.append(dict(
                    date=str(ds), scenario=name, issue_hour=h, slot=j,
                    physical_interval=physical_label(j),
                    original_plan_kwh=float(original[j]),
                    previous_contract_kwh=float(old[j - t]),
                    new_contract_kwh=float(current[j]),
                    pv_forecast_kwh=float(vh[j]),
                ))

        used_pv[t] = vh[t]
        issue_at[t] = last_hour
        ph = decision_price_horizon(data, day_idx, t, price_info, price_method, observe_current=True)
        e, c, d, w, snext = current_action(
            lh, vh, float(data.load[day_idx, t]), float(data.pv[day_idx, t]),
            current, t, s, policy, ph, lam, limit)
        trace[t] = current[t], e, c, d, w, s, snext
        s = snext

    audit_trace(trace, data.load[day_idx], data.pv[day_idx])
    exposure = price_weighted_exposure_indicators(
        data.load[day_idx], data.pv[day_idx], lh, used_pv, data.price_real[day_idx])
    return dict(
        trace=trace, original=original, load_forecast=lh, pv_forecast=used_pv,
        issue_hour=issue_at, logs=update_logs, revisions=revisions,
        dayahead_price_forecast=p0, **exposure,
    )


def q43_day_summary(data: Data, day_idx: int, result, name: str, price_info: str, rule="refund"):
    x = result["trace"]
    p = data.price_real[day_idx]
    settlement = settlement_bill(result["original"], x[:, 0], x[:, 1], p, rule)
    trajectory = np.r_[x[0, 5], x[:, 6]]
    m = response_metrics(x, p)
    return dict(
        date=str(data.dates[day_idx]), scenario=name, price_info=price_info,
        normal_cost=settlement["plan_cost"],
        **settlement,
        plan_kwh=float(result["original"].sum()),
        executed_contract_kwh=float(x[:, 0].sum()),
        emergency_kwh=float(x[:, 1].sum()),
        emergency_slots=int((x[:, 1] > 1e-6).sum()),
        unused_kwh=float(x[:, 4].sum()),
        soc_start=float(x[0, 5]), soc_end=float(x[-1, 6]),
        soc_min_kwh=float(trajectory.min()), soc_max_kwh=float(trajectory.max()),
        adjustment_nodes=sum(bool(r.get("changed")) for r in result["logs"]),
        adjustment_abs_kwh=float(settlement["increase_kwh"] + settlement["decrease_kwh"]),
        price_weighted_net_under_exposure=float(result["price_weighted_net_under_exposure"]),
        net_under_kwh=float(result["net_under_kwh"]),
        price_weighted_pv_over_exposure=float(result["price_weighted_pv_over_exposure"]),
        pv_over_kwh=float(result["pv_over_kwh"]),
        **{k: v for k, v in m.items() if v is not None},
    )


# =====================================================================
# 11. 运行、断点与导出
# =====================================================================
def _manifest(folder: Path, data: Data, args, branch, price_info, scenario=None):
    m = dict(
        version=ALGORITHM_VERSION,
        branch=branch, price_info=price_info, scenario=scenario,
        price_method=args.price_method,
        settlement=args.settlement,
        validation_days=args.validation_days, tune_every=args.tune_every,
        q42_policy_source=getattr(args, "q42_policy_source", "auto"),
        terminal_weight=args.terminal_weight,
        files={k: file_sha(v) for k, v in data.files.items() if v.exists()},
    )
    if branch == "q43":
        m["q43_common_warmup"] = Q43_PROTOCOL_VERSION
    if branch == "q42":
        # Q4-2严格继承Q2最终V2，V2参数属于正式模型配置，必须进入manifest防止断点混跑。
        m.update(
            q42_prediction_model="Q2_V2_EWMA_MAD",
            v2_residual_window=int(args.v2_residual_window),
            v2_bias_half_life=float(args.v2_bias_half_life),
            v2_bias_shrink=float(args.v2_bias_shrink),
            v2_risk_k=float(args.v2_risk_k),
            v2_tail_quantile=float(args.v2_tail_quantile),
        )
        effective = _effective_q42_policy_source(args, price_info)
        m["q42_policy_source_effective"] = effective
        if effective == "p2":
            p2_policy_file = args.output_root / "q42" / P2 / "daily_summary.csv"
            if not p2_policy_file.exists():
                raise FileNotFoundError(
                    f"P1/P3受控比较要求先完成Q4-2/P2，缺少 {p2_policy_file}")
            m["p2_policy_schedule_sha256"] = file_sha(p2_policy_file)
    return m


def _load_checkpoint(path: Path, q43=False):
    with np.load(path, allow_pickle=False) as f:
        if q43:
            r = dict(
                trace=f["trace"].copy(), original=f["original"].copy(),
                load_forecast=f["load_forecast"].copy(), pv_forecast=f["pv_forecast"].copy(),
                issue_hour=f["issue_hour"].copy(),
                dayahead_price_forecast=f["dayahead_price_forecast"].copy(),
                logs=json.loads(str(f["logs"])), revisions=json.loads(str(f["revisions"])),
                price_weighted_net_under_exposure=float(f["price_weighted_net_under_exposure"]),
                net_under_kwh=float(f["net_under_kwh"]),
                price_weighted_pv_over_exposure=float(f["price_weighted_pv_over_exposure"]),
                pv_over_kwh=float(f["pv_over_kwh"]),
            )
        else:
            pd = json.loads(str(f["policy"]))
            r = dict(
                trace=f["trace"].copy(), load_forecast=f["load_forecast"].copy(),
                pv_forecast=f["pv_forecast"].copy(), price_forecast=f["price_forecast"].copy(),
                base_net_forecast=f["base_net_forecast"].copy(),
                corrected_net_forecast=f["corrected_net_forecast"].copy(),
                risk_buffer=f["risk_buffer"].copy(),
                plan_net_forecast=f["plan_net_forecast"].copy(),
                v2_bias=f["v2_bias"].copy(), v2_sigma=f["v2_sigma"].copy(),
                v2_tail_cap=f["v2_tail_cap"].copy(),
                v2_history_days=int(f["v2_history_days"]),
                price_weighted_net_under_exposure=float(f["price_weighted_net_under_exposure"]),
                net_under_kwh=float(f["net_under_kwh"]),
                price_weighted_pv_over_exposure=float(f["price_weighted_pv_over_exposure"]),
                pv_over_kwh=float(f["pv_over_kwh"]), policy=Policy(**pd),
            )
        return r


def _save_checkpoint(path: Path, result, q43=False):
    tmp = path.with_suffix(".tmp")
    with tmp.open("wb") as f:
        if q43:
            np.savez_compressed(
                f, trace=result["trace"], original=result["original"],
                load_forecast=result["load_forecast"], pv_forecast=result["pv_forecast"],
                issue_hour=result["issue_hour"], dayahead_price_forecast=result["dayahead_price_forecast"],
                logs=json.dumps(result["logs"], ensure_ascii=False),
                revisions=json.dumps(result["revisions"], ensure_ascii=False),
                price_weighted_net_under_exposure=result["price_weighted_net_under_exposure"],
                net_under_kwh=result["net_under_kwh"],
                price_weighted_pv_over_exposure=result["price_weighted_pv_over_exposure"],
                pv_over_kwh=result["pv_over_kwh"],
            )
        else:
            np.savez_compressed(
                f, trace=result["trace"], load_forecast=result["load_forecast"],
                pv_forecast=result["pv_forecast"], price_forecast=result["price_forecast"],
                base_net_forecast=result["base_net_forecast"],
                corrected_net_forecast=result["corrected_net_forecast"],
                risk_buffer=result["risk_buffer"], plan_net_forecast=result["plan_net_forecast"],
                v2_bias=result["v2_bias"], v2_sigma=result["v2_sigma"],
                v2_tail_cap=result["v2_tail_cap"], v2_history_days=int(result["v2_history_days"]),
                price_weighted_net_under_exposure=result["price_weighted_net_under_exposure"],
                net_under_kwh=result["net_under_kwh"],
                price_weighted_pv_over_exposure=result["price_weighted_pv_over_exposure"],
                pv_over_kwh=result["pv_over_kwh"],
                policy=json.dumps(asdict(result["policy"]), ensure_ascii=False),
            )
    tmp.replace(path)


def _prepare_run_folder(folder: Path, manifest: dict, resume: bool):
    folder.mkdir(parents=True, exist_ok=True)
    cp = folder / "checkpoints"
    cp.mkdir(exist_ok=True)
    mf = folder / "manifest.json"
    if mf.exists():
        old = json.loads(mf.read_text(encoding="utf-8"))
        require(old == manifest, f"{folder}存在不同模型配置的旧结果，请换输出目录或删除旧结果")
    else:
        if list(cp.glob("*.npz")) and not resume:
            raise ValueError(f"{folder}已有断点但无manifest")
        write_json(mf, manifest)
    return cp


def _annual_aggregate(summaries):
    ev = [r for r in summaries if date.fromisoformat(r["date"]) >= EVAL_START]
    require(ev, "没有评价期结果")
    keys_sum = [
        "normal_cost", "plan_cost", "adjustment_net_cost", "emergency_cost", "total_actual_cost",
        "plan_kwh", "emergency_kwh", "charge_kwh", "discharge_kwh",
        "price_weighted_net_under_exposure", "net_under_kwh",
        "price_weighted_pv_over_exposure", "pv_over_kwh",
    ]
    report = {k: float(sum(float(r.get(k, 0.0)) for r in ev)) for k in keys_sum}
    report["evaluation_days"] = len(ev)
    report["emergency_days"] = sum(float(r.get("emergency_kwh", 0)) > 1e-6 for r in ev)
    report["emergency_slots"] = sum(int(r.get("emergency_slots", 0)) for r in ev)
    report["terminal_soc"] = float(ev[-1]["soc_end"])

    total_c = sum(float(r.get("charge_kwh", 0)) for r in ev)
    total_d = sum(float(r.get("discharge_kwh", 0)) for r in ev)
    low_c = sum(float(r.get("low_price_charge_kwh", 0)) for r in ev)
    high_d = sum(float(r.get("high_price_discharge_kwh", 0)) for r in ev)
    report["low_price_charge_ratio"] = float(low_c / total_c) if total_c > 1e-9 else None
    report["high_price_discharge_ratio"] = float(high_d / total_d) if total_d > 1e-9 else None

    # 电量加权平均充/放电价格（由每日平均价×当日电量回算）
    cnum = sum(float(r.get("avg_charge_price", 0)) * float(r.get("charge_kwh", 0))
               for r in ev if r.get("avg_charge_price") not in (None, ""))
    dnum = sum(float(r.get("avg_discharge_price", 0)) * float(r.get("discharge_kwh", 0))
               for r in ev if r.get("avg_discharge_price") not in (None, ""))
    report["avg_charge_price"] = float(cnum / total_c) if total_c > 1e-9 else None
    report["avg_discharge_price"] = float(dnum / total_d) if total_d > 1e-9 else None
    if report["avg_charge_price"] is not None and report["avg_discharge_price"] is not None:
        report["efficiency_adjusted_arbitrage_margin"] = (
            report["avg_discharge_price"] - report["avg_charge_price"] / (ETA * ETA)
        )

    net_exp = np.array([float(r.get("price_weighted_net_under_exposure", 0)) for r in ev])
    pv_exp = np.array([float(r.get("price_weighted_pv_over_exposure", 0)) for r in ev])
    emg_cost = np.array([float(r.get("emergency_cost", 0)) for r in ev])
    if len(net_exp) >= 2 and np.std(net_exp) > 1e-12 and np.std(emg_cost) > 1e-12:
        report["corr_net_under_exposure_vs_emergency_cost"] = float(np.corrcoef(net_exp, emg_cost)[0, 1])
    else:
        report["corr_net_under_exposure_vs_emergency_cost"] = None
    if len(pv_exp) >= 2 and np.std(pv_exp) > 1e-12 and np.std(emg_cost) > 1e-12:
        report["corr_pv_over_exposure_vs_emergency_cost"] = float(np.corrcoef(pv_exp, emg_cost)[0, 1])
    else:
        report["corr_pv_over_exposure_vs_emergency_cost"] = None
    report["exposure_note"] = (
        "价格加权净负荷低估/PV高估均为事后误差暴露指标，不是概率风险度量或实际损失；"
        "与紧急购电费的相关性仅用于解释，不作因果证明。")
    return report


def export_q42(data: Data, results: dict[date, dict], summaries, folder: Path, args, price_info: str, full: bool):
    folder.mkdir(parents=True, exist_ok=True)
    write_csv(folder / "daily_summary.csv", summaries)
    details = []
    for ds, r in results.items():
        i = (ds - data.dates[0]).days
        if ds < EVAL_START:
            continue
        for t, x in enumerate(r["trace"]):
            details.append(dict(
                date=str(ds), slot=t, physical_interval=physical_label(t),
                actual_price=float(data.price_real[i, t]),
                decision_price_midnight=float(r["price_forecast"][t]),
                plan_kwh=float(x[0]), emergency_kwh=float(x[1]),
                charge_kwh=float(x[2]), discharge_kwh=float(x[3]), unused_kwh=float(x[4]),
                soc_start=float(x[5]), soc_end=float(x[6]),
                load_forecast_kwh=float(r["load_forecast"][t]), pv_forecast_kwh=float(r["pv_forecast"][t]),
                base_net_forecast_kwh=float(r["base_net_forecast"][t]),
                corrected_net_forecast_kwh=float(r["corrected_net_forecast"][t]),
                risk_buffer_kwh=float(r["risk_buffer"][t]),
                plan_net_forecast_kwh=float(r["plan_net_forecast"][t]),
                v2_bias_kwh=float(r["v2_bias"][t]), v2_sigma_kwh=float(r["v2_sigma"][t]),
                actual_load_kwh=float(data.load[i, t]), actual_pv_kwh=float(data.pv[i, t]),
            ))
    write_csv(folder / "dispatch_detail.csv", details)

    target = None
    if data.template42 is not None:
        wb = load_workbook(data.template42)
        plan, bat, emg = wb["计划购电量"], wb["充放电量"], wb["紧急购电量"]
        plan_rows = {as_date(plan.cell(r, 1).value): r for r in range(2, plan.max_row + 1)
                     if plan.cell(r, 1).value not in (None, "")}
        for ws in (bat, emg):
            for mr in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(mr))
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)
        for ds, r in results.items():
            if ds < EVAL_START:
                continue
            require(ds in plan_rows, f"result4-2模板无日期{ds}")
            row = plan_rows[ds]
            x = r["trace"]
            summary = next(s for s in summaries if s["date"] == str(ds))
            for t in range(T):
                plan.cell(row, 2 + t).value = float(x[t, 0])
            plan.cell(row, 146).value = float(x[:, 0].sum())
            plan.cell(row, 147).value = float(summary["total_actual_cost"])
            for b in range(6):
                z = x[b * 24:(b + 1) * 24]
                bat.append([ds if b == 0 else None, f"{clock(b*240)}-{clock((b+1)*240)}",
                            float(z[:, 2].sum()), float(z[:, 3].sum()),
                            "0:00" if b == 0 else ("24:00" if b == 1 else None),
                            float(x[0, 5]) if b == 0 else (float(x[-1, 6]) if b == 1 else None)])
            for j, (label, amount) in enumerate(emergency_intervals(x) or [("无", 0.0)]):
                emg.append([ds if j == 0 else None, label, amount])
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, date): cell.number_format = "yyyy/mm/dd"
                    elif isinstance(cell.value, (float, int)): cell.number_format = "0.0000"
        filename = "result4-2.xlsx" if full and price_info == P2 else f"result4-2_{price_info}_diagnostic.xlsx"
        target = folder / filename
        wb.save(target)
        wb.close()

    report = _annual_aggregate(summaries)
    ev = [r for r in summaries if date.fromisoformat(r["date"]) >= EVAL_START]
    report.update(
        branch="Q4-2", price_info=price_info, price_method=args.price_method,
        q42_prediction_model="Q2_V2_EWMA_MAD",
        q42_prediction_description=(
            "严格继承Q2最终V2：基础历史预测 + 过去残差EWMA偏差校正 + "
            "MAD稳健风险缓冲；日前使用corrected+buffer，实时未来仅用corrected"),
        v2_residual_window=int(args.v2_residual_window),
        v2_bias_half_life=float(args.v2_bias_half_life),
        v2_bias_shrink=float(args.v2_bias_shrink),
        v2_risk_k=float(args.v2_risk_k),
        v2_tail_quantile=float(args.v2_tail_quantile),
        mean_base_net_mae_kw=(float(np.mean([float(r["base_net_mae_kw"]) for r in ev])) if ev else None),
        mean_corrected_net_mae_kw=(float(np.mean([float(r["corrected_net_mae_kw"]) for r in ev])) if ev else None),
        mean_risk_buffer_kw=(float(np.mean([float(r["mean_risk_buffer_kw"]) for r in ev])) if ev else None),
        complete_evaluation_period=full, workbook=str(target) if target else None,
    )
    write_json(folder / "run_summary.json", report)
    return report


def emergency_intervals(x):
    out = []
    t = 0
    while t < T:
        if x[t, 1] <= 1e-6:
            t += 1
            continue
        e = t + 1
        while e < T and x[e, 1] > 1e-6:
            e += 1
        out.append((f"{clock(t*10)}-{clock(e*10)}", float(x[t:e, 1].sum())))
        t = e
    return out


def export_q43(data: Data, results: dict[date, dict], summaries, folder: Path, args,
               price_info: str, scenario: str, full: bool):
    folder.mkdir(parents=True, exist_ok=True)
    write_csv(folder / "daily_summary.csv", summaries)
    details = []
    logs = []
    revisions = []
    for ds, r in results.items():
        i = (ds - data.dates[0]).days
        if ds >= EVAL_START:
            for t, x in enumerate(r["trace"]):
                details.append(dict(
                    date=str(ds), scenario=scenario, price_info=price_info, slot=t,
                    physical_interval=physical_label(t), actual_price=float(data.price_real[i, t]),
                    original_plan_kwh=float(r["original"][t]), executed_contract_kwh=float(x[0]),
                    load_forecast_kwh=float(r["load_forecast"][t]), pv_forecast_kwh=float(r["pv_forecast"][t]),
                    latest_issue_hour=int(r["issue_hour"][t]), actual_load_kwh=float(data.load[i, t]),
                    actual_pv_kwh=float(data.pv[i, t]), emergency_kwh=float(x[1]),
                    charge_kwh=float(x[2]), discharge_kwh=float(x[3]), unused_kwh=float(x[4]),
                    soc_start_kwh=float(x[5]), soc_end_kwh=float(x[6]),
                ))
            logs.extend(r["logs"])
            revisions.extend(r["revisions"])
    write_csv(folder / "dispatch_detail.csv", details)
    if logs: write_csv(folder / "update_log.csv", logs)
    if revisions: write_csv(folder / "contract_versions.csv", revisions)

    target = None
    if data.template43 is not None:
        wb = load_workbook(data.template43)
        plan, adj = wb["计划购电量"], wb["调整购电量"]
        bat, emg = wb["充放电量"], wb["紧急购电量"]
        plan_rows = {as_date(plan.cell(r, 1).value): r for r in range(2, plan.max_row + 1)
                     if plan.cell(r, 1).value not in (None, "")}
        adj_rows = {as_date(adj.cell(r, 1).value): r for r in range(2, adj.max_row + 1)
                    if adj.cell(r, 1).value not in (None, "")}
        for ws in (bat, emg):
            for mr in list(ws.merged_cells.ranges): ws.unmerge_cells(str(mr))
            if ws.max_row > 1: ws.delete_rows(2, ws.max_row - 1)
        for ds, r in results.items():
            if ds < EVAL_START:
                continue
            require(ds in plan_rows and ds in adj_rows, f"result4-3模板无日期{ds}")
            rp, ra = plan_rows[ds], adj_rows[ds]
            x = r["trace"]
            summary = next(s for s in summaries if s["date"] == str(ds))
            for t in range(T):
                plan.cell(rp, 2 + t).value = float(r["original"][t])
                adj.cell(ra, 2 + t).value = float(x[t, 0])
            plan.cell(rp, 146).value = float(r["original"].sum())
            plan.cell(rp, 147).value = float(summary["plan_cost"])
            adj.cell(ra, 146).value = float(x[:, 0].sum())
            adj.cell(ra, 147).value = float(summary["total_actual_cost"])
            for b in range(6):
                z = x[b * 24:(b + 1) * 24]
                bat.append([ds if b == 0 else None, f"{clock(b*240)}-{clock((b+1)*240)}",
                            float(z[:, 2].sum()), float(z[:, 3].sum()),
                            "0:00" if b == 0 else ("24:00" if b == 1 else None),
                            float(x[0, 5]) if b == 0 else (float(x[-1, 6]) if b == 1 else None)])
            for j, (label, amount) in enumerate(emergency_intervals(x) or [("无", 0.0)]):
                emg.append([ds if j == 0 else None, label, amount])
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, date): cell.number_format = "yyyy/mm/dd"
                    elif isinstance(cell.value, (float, int)): cell.number_format = "0.0000"
        filename = "result4-3.xlsx" if full and price_info == P2 else f"result4-3_{price_info}_{scenario}_diagnostic.xlsx"
        target = folder / filename
        wb.save(target)
        wb.close()

    report = _annual_aggregate(summaries)
    report.update(branch="Q4-3", scenario=scenario, price_info=price_info,
                  price_method=args.price_method, settlement=args.settlement,
                  adjustment_nodes=sum(int(s.get("adjustment_nodes", 0)) for s in summaries if s["date"] >= str(EVAL_START)),
                  total_adjustment_abs_kwh=sum(float(s.get("adjustment_abs_kwh", 0)) for s in summaries if s["date"] >= str(EVAL_START)),
                  complete_evaluation_period=full, workbook=str(target) if target else None)
    write_json(folder / "run_summary.json", report)
    return report


def _effective_q42_policy_source(args, price_info: str) -> str:
    """P2自行按既有Q2逻辑选Policy；P1/P3默认复用P2逐日Policy做受控价格信息比较。"""
    requested = getattr(args, "q42_policy_source", "auto")
    if requested == "auto":
        return "self" if price_info == P2 else "p2"
    if price_info == P2 and requested == "p2":
        return "self"
    return requested


def _load_p2_policy_schedule(args) -> dict[str, Policy]:
    path = args.output_root / "q42" / P2 / "daily_summary.csv"
    if not path.exists():
        raise FileNotFoundError(
            f"受控比较需要先完成Q4-2/P2全年运行，缺少 {path}。"
            "请先运行：python dynamic_price_dispatch.py --mode annual --branch q42 --price-info P2 --resume")
    rows = list(csv.DictReader(path.open(encoding="utf-8-sig")))
    schedule = {}
    for row in rows:
        ds = row.get("date")
        if not ds:
            continue
        schedule[ds] = Policy(
            load_mode=row.get("load_mode", "week_mean"),
            pv_days=int(float(row.get("pv_days", 3))),
            reserve=float(row.get("reserve", 0.0)),
            mu_factor=float(row.get("mu_factor", 0.0)),
        )
    if not schedule:
        raise ValueError(f"{path}中未读取到逐日Policy")
    return schedule


def _read_p2_evaluation_start_soc(args) -> float:
    """受控P1/P3比较统一使用P2在2月1日的评价期起始SOC。"""
    path = args.output_root / "q42" / P2 / "daily_summary.csv"
    if not path.exists():
        raise FileNotFoundError(f"缺少P2正式结果 {path}，无法统一评价期初始SOC")
    for row in csv.DictReader(path.open(encoding="utf-8-sig")):
        if row.get("date") == str(EVAL_START):
            return float(row["soc_start"])
    raise ValueError(f"{path}缺少{EVAL_START}的SOC记录")


def _q43_common_warmup_soc(data: Data, args) -> float:
    """Q4-3正式年度比较的共同预热：仅用P2-B1运行1月，返回2月1日统一SOC。

    这与Q3最终版“B1共同预热后各方案从同一SOC进入评价期”的实验公平性保持一致。
    P1/P3价格信息对照也使用同一评价期起点，避免1月未计费差异污染2~12月比较。
    """
    folder = args.output_root / "q43" / P2 / "common_warmup"
    manifest = _manifest(folder, data, args, "q43_common_warmup", P2, "B1")
    cpdir = _prepare_run_folder(folder, manifest, args.resume)
    summary_path = folder / "warmup_summary.json"
    if args.resume and summary_path.exists():
        d = read_json(summary_path)
        return float(d["evaluation_start_soc_kwh"])

    lam = float(data.fixed_price.min() / ETA) * args.terminal_weight
    state = EINITIAL
    warm_summaries = []
    for j in range((EVAL_START - data.dates[0]).days):
        ds = data.dates[j]
        cp = cpdir / f"{ds}.npz"
        if args.resume and cp.exists():
            r = _load_checkpoint(cp, q43=True)
        elif j == 0:
            x = bootstrap(data.load[j], data.pv[j], state)
            r = dict(
                trace=x, original=np.zeros(T), load_forecast=np.zeros(T), pv_forecast=np.zeros(T),
                issue_hour=np.zeros(T, dtype=int), logs=[], revisions=[],
                dayahead_price_forecast=data.fixed_price.copy(),
                price_weighted_net_under_exposure=0.0, net_under_kwh=0.0,
                price_weighted_pv_over_exposure=0.0, pv_over_kwh=0.0)
            _save_checkpoint(cp, r, q43=True)
        else:
            r = simulate_q43_day(data, j, state, "B1", P2, args.price_method,
                                 lam, args.time_limit, args.settlement)
            _save_checkpoint(cp, r, q43=True)
        state = float(r["trace"][-1, 6])
        warm_summaries.append(dict(date=str(ds), soc_end=state))

    write_json(summary_path, dict(
        role="Q4-3 common warmup", price_info=P2, scenario="B1",
        warmup_start=str(data.dates[0]), warmup_end=str(EVAL_START - timedelta(days=1)),
        evaluation_start=str(EVAL_START), evaluation_start_soc_kwh=state,
        note="1月不计入正式评价；P2-B1共同预热终态用于所有Q4-3方案及P1/P2/P3受控价格信息比较。"))
    write_csv(folder / "warmup_daily_soc.csv", warm_summaries)
    return state


def run_q42(data: Data, args, price_info: str, start_day=date(2025, 1, 1), end_day=EVAL_END):
    folder = args.output_root / "q42" / price_info
    manifest = _manifest(folder, data, args, "q42", price_info)
    checkpoints = _prepare_run_folder(folder, manifest, args.resume)
    lam = float(data.fixed_price.min() / ETA) * args.terminal_weight
    results, summaries = {}, []
    snapshots = {}
    state = EINITIAL
    pol = Policy()
    policy_source = _effective_q42_policy_source(args, price_info)
    controlled_schedule = _load_p2_policy_schedule(args) if policy_source == "p2" else None
    begin = time.perf_counter()

    # P2正式方案仍从1月1日按Q2-V2连续预热；P1/P3受控比较从相同的P2评价期SOC进入2月1日。
    if args.mode == "annual" and controlled_schedule is not None:
        state = _read_p2_evaluation_start_soc(args)
        indices = range((EVAL_START - data.dates[0]).days, (end_day - data.dates[0]).days + 1)
    elif args.mode == "annual":
        indices = range(0, (end_day - data.dates[0]).days + 1)
    else:
        indices = range((start_day - data.dates[0]).days, (end_day - data.dates[0]).days + 1)

    for j in indices:
        ds = data.dates[j]
        snapshots[j] = state
        cp = checkpoints / f"{ds}.npz"
        cached = args.resume and cp.exists()
        if cached:
            r = _load_checkpoint(cp, q43=False)
            # 断点续跑必须同步恢复当日Policy，否则下一未缓存日可能错误回到默认Policy。
            pol = r["policy"]
        elif j == 0:
            x = bootstrap(data.load[j], data.pv[j], state)
            z = np.zeros(T, dtype=float)
            r = dict(
                trace=x, load_forecast=z.copy(), pv_forecast=z.copy(),
                base_net_forecast=z.copy(), corrected_net_forecast=z.copy(),
                risk_buffer=z.copy(), plan_net_forecast=z.copy(),
                v2_bias=z.copy(), v2_sigma=z.copy(), v2_tail_cap=z.copy(), v2_history_days=0,
                price_forecast=data.fixed_price.copy(),
                price_weighted_net_under_exposure=0.0, net_under_kwh=0.0,
                price_weighted_pv_over_exposure=0.0, pv_over_kwh=0.0, policy=pol)
            _save_checkpoint(cp, r, q43=False)
        else:
            if controlled_schedule is not None:
                # 受控价格信息实验：仅复用P2的非价格Policy；SOC仍由当前方案自身独立递推。
                try:
                    pol = controlled_schedule[str(ds)]
                except KeyError as exc:
                    raise KeyError(f"P2 Policy日程缺少 {ds}，无法进行P1/P3受控比较") from exc
            elif ds >= EVAL_START and not args.no_tune_q42 and (j - 31) % args.tune_every == 0:
                # P2主模型（或显式self对照）仍按既有Q2逻辑滚动选择Policy。
                if args.mode == "annual":
                    pol = tune_q42(data, j, snapshots, args, lam, price_info)
            r = simulate_q42_day(data, j, state, pol, price_info, args.price_method,
                                 lam, args.time_limit, args)
            _save_checkpoint(cp, r, q43=False)
        state = float(r["trace"][-1, 6])
        results[ds] = r
        s = q42_day_summary(data, j, r, price_info)
        summaries.append(s)
        write_csv(folder / "daily_summary_progress.csv", summaries)
        print(f"[Q4-2/{price_info}] {ds} cost={s['total_actual_cost']:.2f} "
              f"emergency={s['emergency_kwh']:.2f} SOC={state:.2f} "
              f"{'cached' if cached else f'{time.perf_counter()-begin:.1f}s'}", flush=True)

    full = args.mode == "annual" and end_day == EVAL_END
    report = export_q42(data, results, summaries, folder, args, price_info, full)
    report["q42_policy_source"] = policy_source
    report["controlled_price_information_comparison"] = bool(policy_source == "p2" or price_info == P2)
    eval_start_soc = float(next(s["soc_start"] for s in summaries if s["date"] >= str(EVAL_START)))
    report.update(algorithm_version=ALGORITHM_VERSION, initial_soc_kwh=EINITIAL, evaluation_start_soc_kwh=eval_start_soc,
                  evaluation_start_soc_source=("P2_own_January_warmup" if price_info == P2 else "P2_common_evaluation_start"),
                  terminal_weight=args.terminal_weight, terminal_lambda=lam,
                  single_way_efficiency=ETA, realtime_tiebreak="q2_v2_primary_milp_only")
    write_json(folder / "run_summary.json", report)
    if price_info == P2:
        price_forecast_diagnostics(data, args.price_method, folder)
    return report


def run_q43_scenario(data: Data, args, price_info: str, scenario: str,
                     start_day=date(2025, 1, 1), end_day=EVAL_END):
    folder = args.output_root / "q43" / price_info / scenario
    manifest = _manifest(folder, data, args, "q43", price_info, scenario)
    checkpoints = _prepare_run_folder(folder, manifest, args.resume)
    lam = float(data.fixed_price.min() / ETA) * args.terminal_weight
    results, summaries = {}, []
    state = EINITIAL
    begin = time.perf_counter()

    if args.mode == "annual":
        state = _q43_common_warmup_soc(data, args)
        indices = range((EVAL_START - data.dates[0]).days, (end_day - data.dates[0]).days + 1)
    else:
        indices = range((start_day - data.dates[0]).days, (end_day - data.dates[0]).days + 1)

    for j in indices:
        ds = data.dates[j]
        cp = checkpoints / f"{ds}.npz"
        cached = args.resume and cp.exists()
        if cached:
            r = _load_checkpoint(cp, q43=True)
        elif j == 0:
            x = bootstrap(data.load[j], data.pv[j], state)
            r = dict(trace=x, original=np.zeros(T), load_forecast=np.zeros(T), pv_forecast=np.zeros(T),
                     issue_hour=np.zeros(T, dtype=int), logs=[], revisions=[],
                     dayahead_price_forecast=data.fixed_price.copy(),
                     price_weighted_net_under_exposure=0.0, net_under_kwh=0.0,
                     price_weighted_pv_over_exposure=0.0, pv_over_kwh=0.0)
            _save_checkpoint(cp, r, q43=True)
        else:
            r = simulate_q43_day(data, j, state, scenario, price_info, args.price_method,
                                 lam, args.time_limit, args.settlement)
            _save_checkpoint(cp, r, q43=True)
        state = float(r["trace"][-1, 6])
        results[ds] = r
        s = q43_day_summary(data, j, r, scenario, price_info, args.settlement)
        summaries.append(s)
        write_csv(folder / "daily_summary_progress.csv", summaries)
        print(f"[Q4-3/{price_info}/{scenario}] {ds} cost={s['total_actual_cost']:.2f} "
              f"emergency={s['emergency_kwh']:.2f} adjust_nodes={s['adjustment_nodes']} SOC={state:.2f} "
              f"{'cached' if cached else f'{time.perf_counter()-begin:.1f}s'}", flush=True)

    full = args.mode == "annual" and end_day == EVAL_END
    report = export_q43(data, results, summaries, folder, args, price_info, scenario, full)
    eval_start_soc = float(summaries[0]["soc_start"])
    report.update(algorithm_version=ALGORITHM_VERSION, initial_soc_kwh=EINITIAL, evaluation_start_soc_kwh=eval_start_soc,
                  evaluation_start_soc_source="P2_B1_common_January_warmup",
                  terminal_weight=args.terminal_weight, terminal_lambda=lam,
                  single_way_efficiency=ETA, realtime_tiebreak="lexicographic_same_primary_cost",
                  current_quote_rule="interval_start_quote_known; future_prices_use_midnight_forecast_for_P2")
    write_json(folder / "run_summary.json", report)
    return report


# =====================================================================
# 12. 价格信息价值比较与Q4-3方案选择
# =====================================================================
def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_compare_q43_scenario(args):
    """价格信息比较必须固定同一个Q4-3场景；auto时采用P2全年最优方案。"""
    if args.compare_q43_scenario != "auto":
        return args.compare_q43_scenario
    best_path = args.output_root / "q43" / P2 / "best_main_scenario.json"
    if not best_path.exists():
        raise FileNotFoundError(
            f"缺少 {best_path}；请先完整运行 Q4-3/P2/all-d，再做P1/P2/P3公平比较"
        )
    return read_json(best_path)["scenario"]


def compare_price_information(args, branch: str):
    root = args.output_root / branch
    scenario = resolve_compare_q43_scenario(args) if branch == "q43" else None
    reports = {}
    detail_paths = {}
    for pinfo in PRICE_INFOS:
        folder = root / pinfo / scenario if branch == "q43" else root / pinfo
        rp = folder / "run_summary.json"
        if not rp.exists():
            raise FileNotFoundError(f"缺少 {rp}；请先完成{branch}/{pinfo}运行")
        reports[pinfo] = read_json(rp)
        detail_paths[pinfo] = folder / "dispatch_detail.csv"

    # 公平性硬检查：评价期、初始SOC、末端库存权重/系数必须一致。
    for key in ["evaluation_days", "initial_soc_kwh", "evaluation_start_soc_kwh", "terminal_weight", "terminal_lambda",
                "single_way_efficiency", "realtime_tiebreak"]:
        vals = [reports[p].get(key) for p in PRICE_INFOS]
        if any(v is None for v in vals) or any(str(v) != str(vals[0]) for v in vals[1:]):
            raise ValueError(f"P1/P2/P3受控比较条件不一致：{key}={vals}")
    if branch == "q42":
        if reports[P1].get("q42_policy_source") != "p2" or reports[P3].get("q42_policy_source") != "p2":
            raise ValueError("Q4-2价格信息比较要求P1/P3复用P2逐日Policy；请使用默认--q42-policy-source auto重跑")
        # 除价格信息外，Q2-V2预测/风险参数也必须完全固定。
        for key in ["q42_prediction_model", "v2_residual_window", "v2_bias_half_life",
                    "v2_bias_shrink", "v2_risk_k", "v2_tail_quantile"]:
            vals = [reports[p].get(key) for p in PRICE_INFOS]
            if any(v is None for v in vals) or any(str(v) != str(vals[0]) for v in vals[1:]):
                raise ValueError(f"Q4-2 P1/P2/P3的Q2-V2参数不一致：{key}={vals}")

    rows = []
    for pinfo in PRICE_INFOS:
        r = reports[pinfo]
        rows.append(dict(
            price_info=pinfo,
            total_cost=r["total_actual_cost"],
            normal_cost=r.get("normal_cost", r.get("plan_cost")),
            adjustment_cost=r.get("adjustment_net_cost", 0.0),
            emergency_cost=r.get("emergency_cost", 0.0),
            emergency_kwh=r.get("emergency_kwh", 0.0),
            low_price_charge_ratio=r.get("low_price_charge_ratio"),
            high_price_discharge_ratio=r.get("high_price_discharge_ratio"),
            avg_charge_price=r.get("avg_charge_price"),
            avg_discharge_price=r.get("avg_discharge_price"),
        ))
    c1, c2, c3 = [float(reports[k]["total_actual_cost"]) for k in PRICE_INFOS]
    value = c1 - c2
    loss = c2 - c3
    potential = c1 - c3
    order_ok = (c3 <= c2 + 1e-6) and (c2 <= c1 + 1e-6)
    # 分母过小时比值会数值失稳：至少要求1元，且大于费用量级的1e-6。
    ratio_den_tol = max(1.0, 1e-6 * max(abs(c1), abs(c3), 1.0))
    raw_ratio = value / potential if abs(potential) > ratio_den_tol else None
    interpretable = bool(order_ok and potential > ratio_den_tol)
    summary = dict(
        branch=branch, scenario=scenario,
        C_P1=c1, C_P2=c2, C_P3=c3,
        price_information_improvement_P1_minus_P2=value,
        price_forecast_gap_to_oracle_P2_minus_P3=loss,
        P1_minus_P3_reference_gap=potential,
        relative_oracle_improvement_ratio_raw=raw_ratio,
        relative_oracle_improvement_ratio=(raw_ratio if interpretable else None),
        information_order_P3_le_P2_le_P1=order_ok,
        ratio_denominator_tolerance=ratio_den_tol,
        relative_ratio_interpretable=interpretable,
        note=("P3为价格完全信息Oracle对照，不是严格成本下界。P1/P2/P3均按附件4真实价格结算；"
              "Q4-2受控比较默认固定P2逐日非价格Policy，并统一2月1日评价期起始SOC；进入评价期后SOC由各方案独立递推。"
              "仅当C_P3<=C_P2<=C_P1且P1-P3差值显著大于数值阈值时，才报告‘相对价格完全信息对照的改善比例’。"),
    )

    # P1 vs P2 决策变化量（若明细存在）
    try:
        d1 = list(csv.DictReader(detail_paths[P1].open(encoding="utf-8-sig")))
        d2 = list(csv.DictReader(detail_paths[P2].open(encoding="utf-8-sig")))
        if len(d1) == len(d2):
            if branch == "q42":
                plan_key, ch_key, dis_key, soc_key = "plan_kwh", "charge_kwh", "discharge_kwh", "soc_end"
            else:
                plan_key, ch_key, dis_key, soc_key = "executed_contract_kwh", "charge_kwh", "discharge_kwh", "soc_end_kwh"
            a = lambda rr, k: np.asarray([float(x[k]) for x in rr], dtype=float)
            summary.update(
                decision_L1_purchase_kwh=float(np.abs(a(d2, plan_key) - a(d1, plan_key)).sum()),
                decision_L1_charge_kwh=float(np.abs(a(d2, ch_key) - a(d1, ch_key)).sum()),
                decision_L1_discharge_kwh=float(np.abs(a(d2, dis_key) - a(d1, dis_key)).sum()),
                soc_rmse_kwh=float(np.sqrt(np.mean((a(d2, soc_key) - a(d1, soc_key)) ** 2))),
            )
    except Exception as exc:
        summary["decision_change_note"] = f"未计算P1/P2动作变化：{exc}"

    out = args.output_root / "analysis"
    out.mkdir(parents=True, exist_ok=True)
    suffix = f"_{scenario}" if scenario else ""
    write_csv(out / f"price_information_comparison_{branch}{suffix}.csv", rows)
    write_json(out / f"price_information_value_{branch}{suffix}.json", summary)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return summary


def select_best_q43(args, price_info=P2):
    root = args.output_root / "q43" / price_info
    rows = []
    for d in sorted(root.iterdir()) if root.exists() else []:
        rp = d / "run_summary.json"
        if rp.exists():
            r = read_json(rp)
            rows.append(dict(scenario=d.name,
                             total_actual_cost=float(r["total_actual_cost"]),
                             emergency_cost=float(r.get("emergency_cost", 0)),
                             adjustment_net_cost=float(r.get("adjustment_net_cost", 0)),
                             adjustment_nodes=int(r.get("adjustment_nodes", 0))))
    require(rows, f"{root}没有已完成的Q4-3方案")
    # 所有候选方案必须来自同一算法版本和同一评价期初始SOC，避免不同预热路径混入方案优选。
    starts = []
    versions = []
    for d in sorted(root.iterdir()) if root.exists() else []:
        rp = d / "run_summary.json"
        if rp.exists():
            rr = read_json(rp)
            if rr.get("scenario") in {r["scenario"] for r in rows}:
                starts.append(float(rr.get("evaluation_start_soc_kwh")))
                versions.append(rr.get("algorithm_version"))
    require(starts and max(starts) - min(starts) <= TOL, f"Q4-3候选方案评价期初始SOC不一致：{starts}")
    require(all(v == ALGORITHM_VERSION for v in versions), f"Q4-3候选方案混入旧算法版本：{versions}")
    rows.sort(key=lambda x: x["total_actual_cost"])
    write_csv(root / "comparison_main_D.csv", rows)
    best = rows[0]
    b1 = next((r for r in rows if r["scenario"] == "B1"), None)
    if b1 is not None:
        flex_value = float(b1["total_actual_cost"] - best["total_actual_cost"])
        best["B1_total_actual_cost"] = float(b1["total_actual_cost"])
        best["contract_flexibility_value_B1_minus_best"] = flex_value
        best["contract_flexibility_saving_rate"] = (
            flex_value / float(b1["total_actual_cost"])
            if abs(float(b1["total_actual_cost"])) > 1e-9 else None
        )
    src = root / best["scenario"] / ("result4-3.xlsx" if price_info == P2 else f"result4-3_{price_info}_{best['scenario']}_diagnostic.xlsx")
    if price_info == P2 and src.exists():
        shutil.copy2(src, root / "result4-3.xlsx")
    write_json(root / "best_main_scenario.json", best)
    return best


# =====================================================================
# 13. 可视化：只读结果，不影响优化
# =====================================================================
def make_plots(data: Data, args):
    try:
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
        from matplotlib.colors import LinearSegmentedColormap
    except ImportError as exc:
        raise RuntimeError("绘图需要 matplotlib") from exc

    # -------------------------------------------------------------
    # 论文图统一样式：仅影响可视化，不改变任何模型、数据和结果。
    # 色卡按最终科研配色方案统一。
    # -------------------------------------------------------------
    colors = {
        "blue_dark": "#104E8B",
        "blue": "#376B9E",
        "blue_mid": "#5F89B1",
        "blue_light": "#AFC3D8",
        "mint": "#C5E9E3",
        "blue_pale": "#D7E1EB",
        "red_pale": "#F2DADA",
        "red_light": "#E5B5B5",
        "red_mid": "#D89090",
        "red_dark": "#B22222",
        "text": "#222222",
        "grid": "#D7E1EB",
    }

    # Windows优先使用微软雅黑；兼容其他常见中文字体。
    try:
        installed = {f.name for f in font_manager.fontManager.ttflist}
        for name in ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC",
                     "Source Han Sans SC", "PingFang SC", "Arial Unicode MS"]:
            if name in installed:
                plt.rcParams["font.sans-serif"] = [name]
                break
    except Exception:
        pass
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["font.size"] = 11
    plt.rcParams["axes.titlesize"] = 14
    plt.rcParams["axes.labelsize"] = 11
    plt.rcParams["legend.fontsize"] = 10

    def style_axis(ax, grid_axis="y"):
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#666666")
        ax.spines["bottom"].set_color("#666666")
        ax.tick_params(colors=colors["text"])
        if grid_axis:
            ax.grid(axis=grid_axis, color=colors["grid"], linewidth=0.8, alpha=0.75)
            ax.set_axisbelow(True)

    out = args.output_root / "figures"
    out.mkdir(parents=True, exist_ok=True)

    # 图1：全年动态价格热力图
    # 保持原数据不变，仅替换为论文色卡并将坐标改为真实时刻/月度标签。
    price_cmap = LinearSegmentedColormap.from_list(
        "q4_research_price",
        [colors["blue_dark"], colors["blue_mid"], colors["blue_pale"],
         colors["red_pale"], colors["red_mid"], colors["red_dark"]],
        N=256,
    )
    fig, ax = plt.subplots(figsize=(12.4, 6.0))
    im = ax.imshow(data.price_real, aspect="auto", origin="lower",
                   interpolation="nearest", cmap=price_cmap)
    hour_ticks = np.arange(0, 25, 4)
    slot_ticks = np.minimum(hour_ticks * 6, T - 1)
    ax.set_xticks(slot_ticks)
    ax.set_xticklabels([f"{h}:00" if h < 24 else "24:00" for h in hour_ticks])
    ax.set_xlabel("时刻")

    month_positions = []
    month_labels = []
    for m in range(1, 13):
        ds = date(2025, m, 1)
        if data.dates[0] <= ds <= data.dates[-1]:
            month_positions.append((ds - data.dates[0]).days)
            month_labels.append(f"{m}月")
    ax.set_yticks(month_positions)
    ax.set_yticklabels(month_labels)
    ax.set_ylabel("日期")
    ax.set_title("2025年全年动态电价热力图", pad=12, weight="bold")
    cbar = fig.colorbar(im, ax=ax, pad=0.018)
    cbar.set_label("实际电价（元/kWh）")
    cbar.outline.set_edgecolor("#999999")
    fig.tight_layout()
    fig.savefig(out / "Fig_Q4_01_全年动态电价热力图.svg", format="svg", bbox_inches="tight")
    plt.close(fig)

    # 图2、图3：P2正式Q4-3典型日电价/充放电/SOC（优先最佳方案）
    # 典型日选择规则保持原代码不变：选择价格峰谷差最接近全年中位数的日期。
    best_path = args.output_root / "q43" / P2 / "best_main_scenario.json"
    scenario = args.compare_q43_scenario
    if best_path.exists():
        scenario = read_json(best_path)["scenario"]
    detail_path = args.output_root / "q43" / P2 / scenario / "dispatch_detail.csv"
    if detail_path.exists():
        rows = list(csv.DictReader(detail_path.open(encoding="utf-8-sig")))
        dates = sorted({r["date"] for r in rows})
        if dates:
            spreads = []
            for ds in dates:
                i = (date.fromisoformat(ds) - data.dates[0]).days
                spreads.append((float(np.ptp(data.price_real[i])), ds))
            med = float(np.median([x[0] for x in spreads]))
            typical = min(spreads, key=lambda x: abs(x[0] - med))[1]
            rr = [r for r in rows if r["date"] == typical]
            rr.sort(key=lambda r: int(r["slot"]))
            i = (date.fromisoformat(typical) - data.dates[0]).days
            t = np.arange(T) / 6
            charge = np.asarray([float(r["charge_kwh"]) / DT for r in rr])
            discharge = np.asarray([float(r["discharge_kwh"]) / DT for r in rr])
            battery = discharge - charge
            soc = np.asarray([float(r["soc_end_kwh"]) for r in rr])
            price = data.price_real[i]

            # 图2：储能净功率填充 + 动态电价折线。
            fig, ax1 = plt.subplots(figsize=(12.4, 5.6))
            ax1.axhline(0, color="#777777", linewidth=0.8)
            ax1.fill_between(t, 0, np.maximum(battery, 0), step="post",
                             color=colors["blue_dark"], alpha=0.92, label="放电功率")
            ax1.fill_between(t, 0, np.minimum(battery, 0), step="post",
                             color=colors["blue_light"], alpha=0.95, label="充电功率")
            ax1.step(t, battery, where="post", color=colors["blue"], linewidth=1.15)
            ax1.set_xlim(0, 24)
            ax1.set_xticks(np.arange(0, 25, 3))
            ax1.set_xlabel("时刻 / h")
            ax1.set_ylabel("储能功率 / kW（放电为正，充电为负）")
            style_axis(ax1, "y")
            for h in (6, 12, 18):
                ax1.axvline(h, color=colors["blue_pale"], linestyle="--",
                            linewidth=1.1, zorder=0)

            ax2 = ax1.twinx()
            ax2.step(t, price, where="post", color=colors["red_dark"],
                     linewidth=2.0, label="实际电价")
            ax2.set_ylabel("实际电价 /（元/kWh）", color=colors["red_dark"])
            ax2.tick_params(axis="y", colors=colors["red_dark"])
            ax2.spines["top"].set_visible(False)
            ax2.spines["right"].set_color(colors["red_mid"])
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines1 + lines2, labels1 + labels2,
                       loc="upper left", frameon=False, ncol=3)
            ax1.set_title("典型日电价与储能充放电响应",
                          pad=12, weight="bold")
            fig.tight_layout()
            fig.savefig(out / "Fig_Q4_02_典型日电价充放电响应.svg",
                        format="svg", bbox_inches="tight")
            plt.close(fig)

            # 图3：SOC安全区间 + 动态电价联合曲线。
            fig, ax1 = plt.subplots(figsize=(12.4, 5.6))
            ax1.axhspan(EMIN, EMAX, color=colors["blue_pale"], alpha=0.30, zorder=0)
            ax1.axhline(EMIN, color=colors["blue_mid"], linestyle=":",
                        linewidth=1.2, label="SOC下限")
            ax1.axhline(EMAX, color=colors["blue_mid"], linestyle=":",
                        linewidth=1.2, label="SOC上限")
            ax1.step(t, soc, where="post", color=colors["blue_dark"],
                     linewidth=2.1, label="SOC")
            ax1.set_xlim(0, 24)
            ax1.set_xticks(np.arange(0, 25, 3))
            ax1.set_xlabel("时刻 / h")
            ax1.set_ylabel("储能电量 SOC / kWh")
            style_axis(ax1, "y")
            for h in (6, 12, 18):
                ax1.axvline(h, color=colors["blue_pale"], linestyle="--",
                            linewidth=1.1, zorder=0)

            ax2 = ax1.twinx()
            ax2.step(t, price, where="post", color=colors["red_dark"],
                     linewidth=1.8, label="实际电价")
            ax2.set_ylabel("实际电价 /（元/kWh）", color=colors["red_dark"])
            ax2.tick_params(axis="y", colors=colors["red_dark"])
            ax2.spines["top"].set_visible(False)
            ax2.spines["right"].set_color(colors["red_mid"])
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines1 + lines2, labels1 + labels2,
                       loc="upper left", frameon=False, ncol=4)
            ax1.set_title("典型日电价与储能SOC响应",
                          pad=12, weight="bold")
            fig.tight_layout()
            fig.savefig(out / "Fig_Q4_03_典型日电价SOC联合曲线.svg",
                        format="svg", bbox_inches="tight")
            plt.close(fig)

    # 图4：P1/P2/P3价格信息层级下Q4-2与Q4-3全年实际结算费用。
    # 只读取compare已经生成的受控比较CSV，不重新计算优化结果。
    compare_scenario = resolve_compare_q43_scenario(args)
    comp42 = args.output_root / "analysis" / "price_information_comparison_q42.csv"
    comp43 = args.output_root / "analysis" / f"price_information_comparison_q43_{compare_scenario}.csv"
    if comp42.exists() and comp43.exists():
        rows42 = list(csv.DictReader(comp42.open(encoding="utf-8-sig")))
        rows43 = list(csv.DictReader(comp43.open(encoding="utf-8-sig")))
        map42 = {r["price_info"]: float(r["total_cost"]) / 1e4 for r in rows42}
        map43 = {r["price_info"]: float(r["total_cost"]) / 1e4 for r in rows43}
        labels = [P1, P2, P3]
        q42 = np.asarray([map42[x] for x in labels], dtype=float)
        q43 = np.asarray([map43[x] for x in labels], dtype=float)
        x = np.arange(len(labels))

        fig, ax = plt.subplots(figsize=(10.8, 5.7))
        ax.plot(x, q42, color=colors["blue_dark"], marker="o", markersize=8,
                linewidth=2.4, label="Q4-2")
        ax.plot(x, q43, color=colors["red_dark"], marker="s", markersize=7.5,
                linewidth=2.4, label="Q4-3")
        for xi, yi in zip(x, q42):
            ax.text(xi, yi + 0.8, f"{yi:.2f}", ha="center", va="bottom",
                    fontsize=9.5, color=colors["blue_dark"])
        for xi, yi in zip(x, q43):
            ax.text(xi, yi + 0.8, f"{yi:.2f}", ha="center", va="bottom",
                    fontsize=9.5, color=colors["red_dark"])

        ax.set_xticks(x)
        ax.set_xticklabels(["P1\n固定价格认知", "P2\n因果价格预测",
                            "P3\n全日价格完全信息参考"])
        ax.set_ylabel("全年实际结算费用 / 万元")
        ax.set_title("不同价格信息层级下的全年实际结算费用",
                     pad=12, weight="bold")
        all_vals = np.r_[q42, q43]
        margin = max(5.0, float(np.ptp(all_vals)) * 0.18)
        ax.set_ylim(float(np.min(all_vals) - margin), float(np.max(all_vals) + margin))
        style_axis(ax, "y")
        ax.legend(frameon=False, loc="upper right", ncol=2)
        fig.tight_layout()
        fig.savefig(out / "Fig_Q4_04_价格信息层级费用分解.svg",
                    format="svg", bbox_inches="tight")
        plt.close(fig)

    # 图5：Q4-3 D系列方案“经济性—紧急购电风险”权衡图。
    # 横轴越小表示紧急购电越少，纵轴越低表示全年实际费用越低；
    # 气泡面积仅表示全年实际调整节点数。只读P2各方案run_summary，不重新求解。
    q43_root = args.output_root / "q43" / P2
    if q43_root.exists():
        scenario_rows = []
        for folder in sorted(q43_root.iterdir()):
            if not folder.is_dir():
                continue
            rp = folder / "run_summary.json"
            if not rp.exists():
                continue
            r = read_json(rp)
            sc = str(r.get("scenario", folder.name))
            if sc != "B1" and not sc.startswith("D"):
                continue
            if r.get("total_actual_cost") is None or r.get("emergency_kwh") is None:
                continue
            scenario_rows.append(dict(
                scenario=sc,
                total_cost=float(r["total_actual_cost"]) / 1e4,
                emergency_mwh=float(r["emergency_kwh"]) / 1000.0,
                adjustment_nodes=int(r.get("adjustment_nodes", 0)),
            ))

        if scenario_rows:
            best_info = read_json(best_path) if best_path.exists() else None
            best_sc = str(best_info.get("scenario")) if best_info else min(
                scenario_rows, key=lambda z: z["total_cost"])["scenario"]
            max_nodes = max([r["adjustment_nodes"] for r in scenario_rows] + [1])

            fig, ax = plt.subplots(figsize=(10.8, 6.2))
            for r in scenario_rows:
                if r["scenario"] == best_sc:
                    face, edge, zorder = colors["blue_dark"], colors["blue_dark"], 4
                elif r["scenario"] == "B1":
                    face, edge, zorder = colors["red_mid"], colors["red_dark"], 3
                else:
                    face, edge, zorder = colors["blue_light"], colors["blue_mid"], 2
                size = 90 + 430 * r["adjustment_nodes"] / max_nodes
                ax.scatter(
                    r["emergency_mwh"], r["total_cost"],
                    s=size, color=face, edgecolor=edge, linewidth=1.1,
                    alpha=0.92, zorder=zorder
                )
                ax.annotate(
                    r["scenario"],
                    (r["emergency_mwh"], r["total_cost"]),
                    xytext=(6, 6), textcoords="offset points",
                    fontsize=9.0,
                    color=colors["blue_dark"] if r["scenario"] == best_sc else colors["text"],
                    weight="bold" if r["scenario"] == best_sc else "normal",
                )

            ax.set_xlabel("全年紧急购电量 / MWh")
            ax.set_ylabel("全年实际结算费用 / 万元")
            ax.set_title("D系列方案的经济性—紧急购电风险权衡",
                         pad=12, weight="bold")
            style_axis(ax, "both")
            ax.text(
                0.02, 0.97, "左下方向更优；气泡面积表示全年合同调整节点数",
                transform=ax.transAxes, ha="left", va="top",
                fontsize=9.2, color="#555555"
            )
            fig.tight_layout()
            fig.savefig(out / "Fig_Q4_05_D系列经济性风险权衡.svg",
                        format="svg", bbox_inches="tight")
            plt.close(fig)

    # 图6：Q4-3从B1到最优D方案的费用变化瀑布图。
    # 严格按真实结算恒等式拆分：
    # total_actual_cost = plan_cost + adjustment_net_cost + emergency_cost。
    if best_path.exists():
        best_info = read_json(best_path)
        best_sc = str(best_info["scenario"])
        b1_path = args.output_root / "q43" / P2 / "B1" / "run_summary.json"
        best_run_path = args.output_root / "q43" / P2 / best_sc / "run_summary.json"
        if b1_path.exists() and best_run_path.exists():
            b1 = read_json(b1_path)
            br = read_json(best_run_path)

            start = float(b1["total_actual_cost"]) / 1e4
            delta_plan = (float(br.get("plan_cost", 0.0)) -
                          float(b1.get("plan_cost", 0.0))) / 1e4
            delta_adjust = (float(br.get("adjustment_net_cost", 0.0)) -
                            float(b1.get("adjustment_net_cost", 0.0))) / 1e4
            delta_emergency = (float(br.get("emergency_cost", 0.0)) -
                               float(b1.get("emergency_cost", 0.0))) / 1e4
            final = float(br["total_actual_cost"]) / 1e4

            deltas = [delta_plan, delta_adjust, delta_emergency]
            labels_wf = [
                "B1\n总费用",
                "0点合同\n费用变化",
                "合同调整\n净费用变化",
                "紧急购电\n费用变化",
                f"{best_sc}\n总费用",
            ]

            fig, ax = plt.subplots(figsize=(11.2, 6.0))
            xw = np.arange(5)
            width = 0.64
            ax.bar(xw[0], start, width=width, color=colors["blue_mid"],
                   edgecolor=colors["blue_dark"], linewidth=1.0)

            prev = start
            for j, delta in enumerate(deltas, start=1):
                new = prev + delta
                bottom = min(prev, new)
                height = abs(delta)
                bar_color = colors["red_mid"] if delta > 0 else colors["blue_dark"]
                ax.bar(xw[j], height, bottom=bottom, width=width,
                       color=bar_color, edgecolor=bar_color, alpha=0.92)
                ax.plot([xw[j-1] + width/2, xw[j] - width/2],
                        [prev, prev], color=colors["blue_pale"],
                        linewidth=1.2, linestyle="--")
                ax.text(
                    xw[j], bottom + height / 2, f"{delta:+.2f}",
                    ha="center", va="center", fontsize=9.6,
                    color=colors["text"],
                    weight="bold"
                )
                prev = new

            ax.bar(xw[4], final, width=width, color=colors["blue_dark"],
                   edgecolor=colors["blue_dark"], linewidth=1.0)
            ax.plot([xw[3] + width/2, xw[4] - width/2],
                    [prev, prev], color=colors["blue_pale"],
                    linewidth=1.2, linestyle="--")
            ax.text(xw[0], start + 1.0, f"{start:.2f}",
                    ha="center", va="bottom", fontsize=9.8, color=colors["blue_dark"])
            ax.text(xw[4], final + 1.0, f"{final:.2f}",
                    ha="center", va="bottom", fontsize=9.8, color=colors["blue_dark"])

            ax.set_xticks(xw)
            ax.set_xticklabels(labels_wf)
            ax.set_ylabel("全年实际结算费用 / 万元")
            ax.set_title("日内预测更新的费用改善来源", pad=12, weight="bold")
            style_axis(ax, "y")
            fig.tight_layout()
            fig.savefig(out / "Fig_Q4_06_Q43费用改善瀑布图.svg",
                        format="svg", bbox_inches="tight")
            plt.close(fig)

    # 图7：实际价格分位与储能充放电响应。
    # 使用Q4-2/P2正式逐时结果，将334天评价期实际电价按排序等频分为5组，
    # 统计每组平均充/放电功率。该图仅作事后行为解释，不改变调度策略。
    q42_detail = args.output_root / "q42" / P2 / "dispatch_detail.csv"
    if q42_detail.exists():
        rows = list(csv.DictReader(q42_detail.open(encoding="utf-8-sig")))
        if rows:
            actual_price = np.asarray([float(r["actual_price"]) for r in rows], dtype=float)
            charge_kw = np.asarray([float(r["charge_kwh"]) / DT for r in rows], dtype=float)
            discharge_kw = np.asarray([float(r["discharge_kwh"]) / DT for r in rows], dtype=float)

            n = len(rows)
            order = np.argsort(actual_price, kind="mergesort")
            group_id = np.empty(n, dtype=int)
            for g, ids in enumerate(np.array_split(order, 5)):
                group_id[ids] = g

            avg_charge = np.asarray(
                [charge_kw[group_id == g].mean() for g in range(5)], dtype=float)
            avg_discharge = np.asarray(
                [discharge_kw[group_id == g].mean() for g in range(5)], dtype=float)
            avg_price = np.asarray(
                [actual_price[group_id == g].mean() for g in range(5)], dtype=float)

            y = np.arange(5)
            fig, ax = plt.subplots(figsize=(10.8, 6.0))
            ax.barh(y, -avg_charge, height=0.58,
                    color=colors["blue_light"], edgecolor=colors["blue_mid"],
                    linewidth=0.9, label="平均充电功率")
            ax.barh(y, avg_discharge, height=0.58,
                    color=colors["blue_dark"], edgecolor=colors["blue_dark"],
                    linewidth=0.9, label="平均放电功率")
            ax.axvline(0, color="#666666", linewidth=0.9)

            qlabels = [
                f"Q{g+1} {'低价' if g == 0 else ('高价' if g == 4 else '')}\n"
                f"均价 {avg_price[g]:.3f}"
                for g in range(5)
            ]
            ax.set_yticks(y)
            ax.set_yticklabels(qlabels)
            ax.set_xlabel("平均储能功率 / kW（左：充电；右：放电）")
            ax.set_ylabel("实际电价等频分位组")
            ax.set_title("不同电价分位下的储能充放电响应",
                         pad=12, weight="bold")
            style_axis(ax, "x")
            ax.legend(frameon=False, loc="lower right")
            fig.tight_layout()
            fig.savefig(out / "Fig_Q4_07_价格分位储能响应.svg",
                        format="svg", bbox_inches="tight")
            plt.close(fig)


    # -----------------------------------------------------------------
    # 图8：P2因果价格预测误差的“月份 × 时刻”热力图
    # 价值：不是只给一个全年MAE，而是展示误差在季节和日内时段上的异质性。
    # 数据：Q4-2/P2正式dispatch_detail中的0点价格预测与附件4实际价格。
    # -----------------------------------------------------------------
    q42_detail = args.output_root / "q42" / P2 / "dispatch_detail.csv"
    if q42_detail.exists():
        rows8 = list(csv.DictReader(q42_detail.open(encoding="utf-8-sig")))
        rows8 = [r for r in rows8 if date.fromisoformat(r["date"]) >= EVAL_START]
        if rows8:
            mae_sum = np.zeros((11, 24), dtype=float)
            mae_cnt = np.zeros((11, 24), dtype=int)
            for r in rows8:
                ds = date.fromisoformat(r["date"])
                m = ds.month - 2
                h = int(r["slot"]) // 6
                if 0 <= m < 11 and 0 <= h < 24:
                    err = abs(float(r["decision_price_midnight"]) -
                              float(r["actual_price"]))
                    mae_sum[m, h] += err
                    mae_cnt[m, h] += 1
            mae_mat = np.divide(
                mae_sum, mae_cnt,
                out=np.full_like(mae_sum, np.nan),
                where=mae_cnt > 0
            )

            mae_cmap = LinearSegmentedColormap.from_list(
                "q4_price_mae",
                [colors["blue_pale"], colors["red_pale"],
                 colors["red_mid"], colors["red_dark"]],
                N=256,
            )
            fig, ax = plt.subplots(figsize=(12.4, 5.8))
            im = ax.imshow(mae_mat, aspect="auto", origin="lower",
                           interpolation="nearest", cmap=mae_cmap)
            ax.set_xticks(np.arange(0, 24, 2))
            ax.set_xticklabels([f"{h:02d}:00" for h in range(0, 24, 2)])
            ax.set_yticks(np.arange(11))
            ax.set_yticklabels([f"{m}月" for m in range(2, 13)])
            ax.set_xlabel("时刻")
            ax.set_ylabel("月份")
            ax.set_title("P2因果价格预测的月—时段平均绝对误差",
                         pad=12, weight="bold")
            cbar = fig.colorbar(im, ax=ax, pad=0.018)
            cbar.set_label("平均绝对误差 /（元/kWh）")
            cbar.outline.set_edgecolor("#999999")
            fig.tight_layout()
            fig.savefig(out / "Fig_Q4_08_P2价格预测误差热力图.svg",
                        format="svg", bbox_inches="tight")
            plt.close(fig)

    # -----------------------------------------------------------------
    # 图9：价格加权净负荷低估暴露 vs 紧急购电费用
    # 价值：验证“高价时段净负荷低估”与5倍紧急购电损失的关联，
    #       为Q4-2继续继承Q2-V2风险机制提供解释证据。
    # 注意：暴露指标仅为解释指标，不等同于真实账单。
    # -----------------------------------------------------------------
    q42_daily = args.output_root / "q42" / P2 / "daily_summary.csv"
    if q42_daily.exists():
        rows9 = list(csv.DictReader(q42_daily.open(encoding="utf-8-sig")))
        rows9 = [r for r in rows9 if date.fromisoformat(r["date"]) >= EVAL_START]
        if len(rows9) >= 3:
            x9 = np.asarray(
                [float(r["price_weighted_net_under_exposure"]) / 1e4 for r in rows9],
                dtype=float
            )
            y9 = np.asarray(
                [float(r["emergency_cost"]) / 1e4 for r in rows9],
                dtype=float
            )
            finite = np.isfinite(x9) & np.isfinite(y9)
            x9, y9 = x9[finite], y9[finite]
            if len(x9) >= 3:
                fig, ax = plt.subplots(figsize=(9.6, 6.1))
                ax.scatter(
                    x9, y9, s=30, color=colors["blue_mid"],
                    edgecolor=colors["blue_dark"], linewidth=0.45,
                    alpha=0.58, label="评价日"
                )
                corr = None
                if np.ptp(x9) > 1e-12:
                    coef = np.polyfit(x9, y9, 1)
                    xx = np.linspace(float(x9.min()), float(x9.max()), 160)
                    yy = coef[0] * xx + coef[1]
                    ax.plot(xx, yy, color=colors["red_dark"],
                            linewidth=2.0, label="线性趋势")
                    corr = float(np.corrcoef(x9, y9)[0, 1])
                ax.set_xlabel("价格加权净负荷低估暴露 / 万元（解释指标）")
                ax.set_ylabel("当日紧急购电费用 / 万元")
                ax.set_title("风险暴露与紧急购电费用的关联",
                             pad=12, weight="bold")
                style_axis(ax, "both")
                ax.legend(frameon=False, loc="upper left")
                if corr is not None and np.isfinite(corr):
                    ax.text(
                        0.98, 0.04, f"Pearson r = {corr:.3f}",
                        transform=ax.transAxes, ha="right", va="bottom",
                        fontsize=9.6, color=colors["text"]
                    )
                fig.tight_layout()
                fig.savefig(out / "Fig_Q4_09_风险暴露与紧急购电关联.svg",
                            format="svg", bbox_inches="tight")
                plt.close(fig)

    # -----------------------------------------------------------------
    # 图10：6/12/18时刻“信息利用率—合同调整强度”气泡图
    # 价值：直接回答“有新预测是否每次都值得改合同”。
    # y轴为触发调整的比例；气泡大小为触发时平均L1合同调整量。
    # contract_versions包含每个更新节点的旧合同与新合同，可严格回算。
    # -----------------------------------------------------------------
    if best_path.exists():
        best_info10 = read_json(best_path)
        best_sc10 = str(best_info10["scenario"])
        cv_path = args.output_root / "q43" / P2 / best_sc10 / "contract_versions.csv"
        if cv_path.exists():
            cv_rows = list(csv.DictReader(cv_path.open(encoding="utf-8-sig")))
            grouped = {}
            for r in cv_rows:
                h = int(r["issue_hour"])
                if h not in (6, 12, 18):
                    continue
                key = (r["date"], h)
                grouped.setdefault(key, 0.0)
                grouped[key] += abs(
                    float(r["new_contract_kwh"]) -
                    float(r["previous_contract_kwh"])
                )

            hours10 = [6, 12, 18]
            rates10, avg_mwh10, opps10 = [], [], []
            for h in hours10:
                vals = [v for (ds, hh), v in grouped.items() if hh == h]
                opps10.append(len(vals))
                if vals:
                    arr = np.asarray(vals, dtype=float)
                    trig = arr > 1e-6
                    rates10.append(100.0 * float(np.mean(trig)))
                    avg_mwh10.append(
                        float(arr[trig].mean() / 1000.0) if np.any(trig) else 0.0
                    )
                else:
                    rates10.append(0.0)
                    avg_mwh10.append(0.0)

            if any(opps10):
                max_size_base = max(avg_mwh10 + [1e-9])
                sizes10 = [150 + 950 * v / max_size_base for v in avg_mwh10]

                fig, ax = plt.subplots(figsize=(9.6, 5.9))
                ax.vlines(hours10, 0, rates10,
                          color=colors["blue_light"], linewidth=4.0, alpha=0.9)
                ax.scatter(
                    hours10, rates10, s=sizes10,
                    color=[colors["blue"], colors["blue_dark"], colors["blue_mid"]],
                    edgecolor=colors["blue_dark"], linewidth=1.2,
                    alpha=0.92, zorder=3
                )
                for h, rate, mag in zip(hours10, rates10, avg_mwh10):
                    # 百分比标签与气泡顶部留出足够间距，避免100%标签压在大气泡上。
                    label_y = min(rate + 10.0, 111.0)
                    ax.text(h, label_y, f"{rate:.1f}%",
                            ha="center", va="bottom",
                            fontsize=10, color=colors["blue_dark"], weight="bold")
                    # 调整强度放在气泡下方，同样预留间距。
                    mag_y = max(rate - 12.0, 4.0)
                    ax.text(h, mag_y, f"{mag:.2f} MWh/次",
                            ha="center", va="top",
                            fontsize=9.0, color=colors["text"])
                ax.set_xlim(4.5, 19.5)
                ax.set_ylim(0, 116)
                ax.set_xticks(hours10)
                ax.set_xticklabels(["6:00", "12:00", "18:00"])
                ax.set_xlabel("预测更新时间")
                ax.set_ylabel("合同调整触发率 / %")
                ax.set_title("日内预测更新的信息利用率与合同调整强度",
                             pad=12, weight="bold")
                style_axis(ax, "y")
                fig.tight_layout()
                fig.savefig(out / "Fig_Q4_10_预测更新利用率与调整强度.svg",
                            format="svg", bbox_inches="tight")
                plt.close(fig)

    # -----------------------------------------------------------------
    # 图11：Q4-3价格信息价值的月度分布
    # 价值：避免只报告一个全年总数，检验价格信息价值是否具有季节异质性。
    # 比较固定同一个best D场景的P1/P2/P3，仍遵守受控比较条件。
    # -----------------------------------------------------------------
    if best_path.exists():
        best_info11 = read_json(best_path)
        best_sc11 = str(best_info11["scenario"])
        monthly_cost = {}
        all_exist = True
        for pinfo in PRICE_INFOS:
            p = args.output_root / "q43" / pinfo / best_sc11 / "daily_summary.csv"
            if not p.exists():
                all_exist = False
                break
            rr = list(csv.DictReader(p.open(encoding="utf-8-sig")))
            sums = {m: 0.0 for m in range(2, 13)}
            for r in rr:
                ds = date.fromisoformat(r["date"])
                if EVAL_START <= ds <= EVAL_END:
                    sums[ds.month] += float(r["total_actual_cost"])
            monthly_cost[pinfo] = sums

        if all_exist:
            months11 = np.arange(2, 13)
            save_p2 = np.asarray([
                (monthly_cost[P1][m] - monthly_cost[P2][m]) / 1e4
                for m in months11
            ])
            save_p3 = np.asarray([
                (monthly_cost[P1][m] - monthly_cost[P3][m]) / 1e4
                for m in months11
            ])
            x11 = np.arange(len(months11))
            w11 = 0.36

            fig, ax = plt.subplots(figsize=(12.0, 5.9))
            ax.bar(
                x11 - w11/2, save_p2, width=w11,
                color=colors["blue_light"], edgecolor=colors["blue_mid"],
                linewidth=0.9, label="P1→P2 节省"
            )
            ax.bar(
                x11 + w11/2, save_p3, width=w11,
                color=colors["blue_dark"], edgecolor=colors["blue_dark"],
                linewidth=0.9, label="P1→P3 节省"
            )
            ax.axhline(0, color="#666666", linewidth=0.9)
            ax.set_xticks(x11)
            ax.set_xticklabels([f"{m}月" for m in months11])
            ax.set_xlabel("月份")
            ax.set_ylabel("相对P1的月度费用节省 / 万元")
            ax.set_title("Q4-3价格信息价值的月度分布",
                         pad=12, weight="bold")
            style_axis(ax, "y")
            ax.legend(frameon=False, loc="best", ncol=2)
            fig.tight_layout()
            fig.savefig(out / "Fig_Q4_11_价格信息价值月度分布.svg",
                        format="svg", bbox_inches="tight")
            plt.close(fig)

    print(f"图件输出：{out.resolve()}")


# =====================================================================
# 14. 自检
# =====================================================================
def audit_all(data: Data, args):
    require(data.load.shape == (365, T) and data.pv.shape == (365, T), "附件2维度异常")
    require(data.price_real.shape == (365, T), "附件4维度异常")
    require(raw_sample_label(0) == "00:10" and physical_label(0) == "00:00-00:10", "时间映射错误")
    require(raw_sample_label(T - 1) == "0:00+1" and physical_label(T - 1) == "23:50-24:00", "日末映射错误")

    # 价格预测因果性：修改目标日及未来价格不能改变目标日P2预测
    idx = (date(2025, 3, 20) - data.dates[0]).days
    p1 = historical_price_forecast(data, idx, args.price_method)
    altered = data.price_real.copy()
    altered[idx:] *= 9
    fake = Data(data.dates, data.load, data.pv, data.fixed_price, altered, data.forecasts,
                data.template42, data.template43, data.template42_labels, data.template43_labels, data.files)
    p2 = historical_price_forecast(fake, idx, args.price_method)
    require(np.allclose(p1, p2), "P2价格预测读取了当天/未来价格")
    # 时间顺序：日前价格路径不得使用首时段真实报价；进入当前区间后只替换当前报价。
    p0 = decision_price_horizon(data, idx, 0, P2, args.price_method, observe_current=False)
    require(np.allclose(p0, historical_price_forecast(data, idx, args.price_method)),
            "P2日前计划错误读取了当天真实报价")
    pt = decision_price_horizon(data, idx, 36, P2, args.price_method, observe_current=True)
    require(abs(pt[0] - data.price_real[idx, 36]) < 1e-12, "当前区间真实报价未进入实时价格路径")
    require(np.allclose(pt[1:], p0[37:]), "P2未来价格被日内真实报价错误滚动修正")

    # Q4-2 V2因果性：修改目标日及未来负荷/PV，不得改变目标日0点V2预测、EWMA偏差或MAD缓冲。
    base_v2 = q42_v2_net_profiles(data, idx, Policy(), args)
    alt_load = data.load.copy(); alt_pv = data.pv.copy()
    alt_load[idx:] *= 3.0
    alt_pv[idx:] = 0.0
    fake_v2 = Data(data.dates, alt_load, alt_pv, data.fixed_price, data.price_real, data.forecasts,
                   data.template42, data.template43, data.template42_labels, data.template43_labels, data.files)
    alt_v2 = q42_v2_net_profiles(fake_v2, idx, Policy(), args)
    for a, b, name in zip(base_v2[:5], alt_v2[:5],
                          ["load_forecast", "pv_forecast", "base_net", "corrected_net", "risk_buffer"]):
        require(np.allclose(a, b), f"Q4-2 V2 {name}读取了目标日/未来真实负荷或PV")
    require(np.allclose(base_v2[5]["bias"], alt_v2[5]["bias"]) and
            np.allclose(base_v2[5]["sigma"], alt_v2[5]["sigma"]),
            "Q4-2 V2 EWMA/MAD统计读取了目标日/未来真实数据")

    # 单区间调整结算硬测试
    s = settlement_bill(np.array([100.0]), np.array([80.0]), np.array([0.0]), np.array([1.0]), "refund")
    require(abs(s["total_actual_cost"] - 90.0) < 1e-9, "50%调减退款结算错误")
    s = settlement_bill(np.array([100.0]), np.array([120.0]), np.array([0.0]), np.array([1.0]), "refund")
    require(abs(s["total_actual_cost"] - 130.0) < 1e-9, "1.5倍调增结算错误")

    # MILP物理小样本
    parts = optimize(np.array([-100.0, 200.0]), np.array([1.0, 2.0]), EMIN,
                     Policy(), 0.0, fixed_q=np.array([0.0, 200.0]), time_limit=args.time_limit)
    require(parts[3][0] < TOL and abs(parts[1][0] - 100.0) < TOL, "富余电优先吸收自检失败")

    # 真实数据1天Q4-2/Q4-3快速因果/可行性测试
    lam = float(data.fixed_price.min() / ETA) * args.terminal_weight
    rr42 = simulate_q42_day(data, idx, EINITIAL, Policy(), P2, args.price_method,
                             lam, args.time_limit, args)
    audit_trace(rr42["trace"], data.load[idx], data.pv[idx])
    rr43 = simulate_q43_day(data, idx, EINITIAL, "D6_12_18", P2, args.price_method,
                            lam, args.time_limit, args.settlement)
    audit_trace(rr43["trace"], data.load[idx], data.pv[idx])
    print("Q4自检通过：数据、时间、价格预测因果性、合同结算、储能物理约束、Q4-2/Q4-3单日运行均正常。")


# =====================================================================
# 15. CLI
# =====================================================================
def parse_scenarios(value):
    if value == "core":
        return ["B1", "D6", "D6_12", "D6_12_18"]
    if value == "all-d":
        return ["B1", "D6", "D12", "D18", "D6_12", "D6_18", "D12_18", "D6_12_18"]
    return [x.strip() for x in value.split(",") if x.strip()]


def main():
    root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--mode", choices=["audit", "demo", "annual", "compare", "figures"], default="audit")
    parser.add_argument("--branch", choices=["q42", "q43", "both"], default="both")
    parser.add_argument("--data-dir", type=Path, default=root / "data")
    parser.add_argument("--output-root", type=Path, default=root / "q4_output")
    parser.add_argument("--price-info", choices=[P1, P2, P3, "ALL"], default=P2)
    parser.add_argument("--price-method", choices=["mean7", "week", "weekmean"], default="mean7")
    parser.add_argument("--q43-scenarios", default="core", help="core / all-d / auto(P2全年最优D*) / 自定义逗号列表")
    parser.add_argument("--compare-q43-scenario", default="auto", help="Q4-3价格信息比较场景；auto=读取P2全年最优方案")
    parser.add_argument("--settlement", choices=["refund", "no_refund"], default="refund")
    parser.add_argument("--time-limit", type=float, default=30.0)
    parser.add_argument("--terminal-weight", type=float, default=1.0)
    parser.add_argument("--validation-days", type=int, default=7)
    parser.add_argument("--tune-every", type=int, default=28)
    parser.add_argument("--no-tune-q42", action="store_true")
    parser.add_argument("--q42-policy-source", choices=["auto", "self", "p2"], default="auto",
                        help="Q4-2非价格Policy来源：auto=P2自行选择、P1/P3复用P2；self=各自调参；p2=强制复用P2")
    # Q4-2严格继承Q2最终V2参数；默认值与Q2正式交付完全一致。
    parser.add_argument("--v2-residual-window", type=int, default=28,
                        help="Q2 V2历史净负荷残差窗口（日）")
    parser.add_argument("--v2-bias-half-life", type=float, default=5.0,
                        help="Q2 V2 EWMA残差偏差半衰期（日）")
    parser.add_argument("--v2-bias-shrink", type=float, default=0.75,
                        help="Q2 V2偏差收缩系数")
    parser.add_argument("--v2-risk-k", type=float, default=0.50,
                        help="Q2 V2 MAD风险缓冲倍数")
    parser.add_argument("--v2-tail-quantile", type=float, default=0.90,
                        help="Q2 V2正残差经验尾部分位数上限")
    parser.add_argument("--resume", action="store_true")
    parser.add_argument("--demo-date", default="2025-03-20")
    parser.add_argument("--demo-days", type=int, default=1)
    parser.add_argument("--end-date", default="2025-12-31")
    args = parser.parse_args()

    require(args.time_limit > 0 and args.terminal_weight >= 0, "时间上限/末端权重非法")
    require(args.validation_days >= 1 and args.tune_every >= 1, "验证窗口/调参间隔必须为正")
    require(args.v2_residual_window >= 2, "v2-residual-window必须>=2")
    require(math.isfinite(args.v2_bias_half_life) and args.v2_bias_half_life > 0,
            "v2-bias-half-life必须为正有限数")
    require(math.isfinite(args.v2_bias_shrink) and 0 <= args.v2_bias_shrink <= 1,
            "v2-bias-shrink必须在[0,1]")
    require(math.isfinite(args.v2_risk_k) and args.v2_risk_k >= 0,
            "v2-risk-k必须为非负有限数")
    require(math.isfinite(args.v2_tail_quantile) and 0.5 < args.v2_tail_quantile < 1,
            "v2-tail-quantile必须在(0.5,1)")
    data = read_data(args.data_dir)

    if args.mode == "audit":
        audit_all(data, args)
        return
    if args.mode == "figures":
        make_plots(data, args)
        return
    if args.mode == "compare":
        if args.branch in ("q42", "both"):
            compare_price_information(args, "q42")
        if args.branch in ("q43", "both"):
            compare_price_information(args, "q43")
        return

    # ALL时必须先跑P2，P1/P3的auto受控比较才能读取P2逐日Policy。
    price_infos = [P2, P1, P3] if args.price_info == "ALL" else [args.price_info]
    if args.q43_scenarios == "auto":
        scenarios = [resolve_compare_q43_scenario(args)]
    else:
        scenarios = parse_scenarios(args.q43_scenarios)
    for s in scenarios:
        scenario_config(s)  # validate

    if args.mode == "demo":
        start = date.fromisoformat(args.demo_date)
        end = start + timedelta(days=args.demo_days - 1)
        require(data.dates[0] < start <= EVAL_END and start <= end <= EVAL_END, "demo日期范围异常")
    else:
        start = data.dates[0]
        end = date.fromisoformat(args.end_date)
        require(EVAL_START <= end <= EVAL_END, "annual end-date必须在2025-02-01~12-31")

    for pinfo in price_infos:
        if args.branch in ("q42", "both"):
            run_q42(data, args, pinfo, start, end)
        if args.branch in ("q43", "both"):
            for scenario in scenarios:
                run_q43_scenario(data, args, pinfo, scenario, start, end)
            # P2正式分支自动选择已跑D/B1方案中全年成本最低者，并复制官方result4-3
            if args.mode == "annual" and end == EVAL_END:
                best = select_best_q43(args, pinfo)
                print(f"[Q4-3/{pinfo}] 当前已跑方案最优：{best}", flush=True)

    # 如果一次性ALL跑完，自动生成信息价值比较（Q43按用户指定对比场景）
    if args.price_info == "ALL" and args.mode == "annual" and end == EVAL_END:
        if args.branch in ("q42", "both"):
            compare_price_information(args, "q42")
        if args.branch in ("q43", "both"):
            compare_price_information(args, "q43")


if __name__ == "__main__":
    main()

