"""
日前计划与实时滚动调度
======================

主模型：
    基于历史预测、软风险备用与10分钟实时滚动修正的日前—实时微网调度。

信息边界：
    第d日0:00只使用d日前历史数据制定并锁定计划购电q；
    运行中仅当前10分钟时段真实负荷/PV可见，未来仍使用0点预测。

正式评价：
    2025-02-01~2025-12-31；2025年1月用于SOC连续运行与历史预热。

本文件保留并统一了此前已经完成的所有Q2有效工作：
    1) 正式主模型全年运行；
    2) 公平“无软备用”B0消融对照（仅R=0, mu=0，预测仍滚动选择）；
    3) 固定R敏感性实验入口（不代表所有R都已经实际跑完）；
    4) 预测误差诊断；
    5) SOC / 紧急购电诊断图与数值汇总；
    6) 主模型 vs 无备用消融对比表与图；
    7) 因果性、SOC、功率、紧急购电用途等自检。
    8) V2实验：EWMA净负荷偏差校正 + MAD稳健风险缓冲（独立输出，不覆盖既有正式结果）。

默认直接运行仍然等价于此前最终 problem2_final.py 的主模型：
    python rolling_dispatch.py

常用模式：
    python rolling_dispatch.py --mode no-reserve
    python rolling_dispatch.py --mode diagnostics
    python rolling_dispatch.py --mode postprocess
    python rolling_dispatch.py --mode r-sensitivity --reserve 900

注意：
    - 默认主模型的数学求解核心未改写，目的是保持与既有最终结果对应。
    - diagnostics / postprocess 只读取已有CSV，不重新求MILP。
    - R敏感性入口仅用于复现实验；没有真实跑出的R点不得写成论文结果。
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import time
import warnings
import subprocess
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time as dtime, timedelta
from functools import lru_cache
from pathlib import Path

import numpy as np
from openpyxl import load_workbook, Workbook
from scipy.optimize import milp, Bounds, LinearConstraint
from scipy.sparse import lil_matrix

T = 144
DT = 1 / 6
ETA = 0.9
EMIN, EMAX, EINITIAL = 1200., 10800., 6000.
B = 5000 * DT
TOL = 2e-4

# ---------------------------------------------------------------------
# 实验模式控制
# ---------------------------------------------------------------------
# 这些全局量只控制“候选策略集合”，不修改 optimize/current_action/
# simulate_day 等已经产生最终结果的数学求解核心。
EXPERIMENT_MODE = 'main'
FIXED_RESERVE = None
FIXED_MU_FACTOR = None

# 已经真实跑完并核验过的全年结果。仅用于 --verify-reference 对照，
# 不参与任何优化，也不会“把答案写死”进求解过程。
REFERENCE_RESULTS = {
    'main': {
        'evaluation_days': 334,
        'planned_cost': 12219255.813539516,
        'emergency_cost': 2934024.7184292674,
        'total_cost': 15153280.531968784,
        'planned_kwh': 20119733.674096465,
        'emergency_kwh': 631023.6000027342,
        'emergency_days': 328,
        'emergency_slots_10min': 9976,
        'terminal_soc': 1760.7052664814817,
    },
    'no-reserve': {
        'evaluation_days': 334,
        'planned_cost': 12217818.134978352,
        'emergency_cost': 3143971.9284786936,
        'total_cost': 15361790.063457046,
        'planned_kwh': 20098297.432582278,
        'emergency_kwh': 677611.7143531463,
        'emergency_days': 331,
        'emergency_slots_10min': 10403,
        'terminal_soc': 1216.7421579629633,
    },
}



@dataclass(frozen=True)
class Policy:
    load_mode: str = 'week_mean'
    pv_days: int = 3
    reserve: float = 0.
    mu_factor: float = 0.


@dataclass
class Data:
    dates: list[date]
    load: np.ndarray
    pv: np.ndarray
    price: np.ndarray
    labels: list[str]


def time_minute(value):
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, dtime):
        return value.hour * 60 + value.minute
    if isinstance(value, (int, float)):
        return round(float(value) * 1440)
    s = str(value).strip().replace('：', ':')
    next_day = '+1' in s
    s = s.replace('+1', '')
    h, m = s.split(':')[:2]
    return int(h) * 60 + int(m) + (1440 if next_day else 0)


def as_date(x):
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    return datetime.strptime(str(x).split()[0].replace('/', '-'), '%Y-%m-%d').date()


def physical_label(t):
    def fmt(m):
        return f'{m // 60:02d}:{m % 60:02d}'
    return f'{fmt(t * 10)}-{fmt((t + 1) * 10)}'


def attachment_time_label(t):
    """附件144点为区间终点：00:10,...,23:50,0:00+1。"""
    m = (t + 1) * 10
    if m == 1440:
        return '0:00+1'
    return f'{m // 60:02d}:{m % 60:02d}'


def read_data(folder):
    def matrix(ws):
        rows = list(ws.values)
        assert [time_minute(x) for x in rows[0][1:145]] == list(range(10, 1441, 10)), 'Input time sequence is not 00:10..24:00'
        rows = [r for r in rows[1:] if r[0] is not None]
        ds = [as_date(r[0]) for r in rows]
        a = np.array([r[1:145] for r in rows], dtype=float)
        assert a.shape == (365, T) and np.isfinite(a).all() and (a >= 0).all(), 'Missing, negative or nonfinite actual data'
        assert ds == [date(2025, 1, 1) + timedelta(days=i) for i in range(365)], 'Dates must be unique, chronological and complete'
        return ds, a
    wb = load_workbook(folder / '附件2.xlsx', read_only=True, data_only=True)
    ds, l = matrix(wb['小区负载'])
    ds2, v = matrix(wb['光伏发电实际功率'])
    wb.close()
    assert ds == ds2
    wb = load_workbook(folder / '附件1.xlsx', read_only=True, data_only=True)
    rr = [r for r in list(wb.worksheets[0].values)[1:] if r[0] is not None]
    assert len(rr) == T and [time_minute(r[0]) for r in rr] == list(range(10, 1441, 10))
    p = np.array([r[1] for r in rr], float)
    wb.close()
    assert np.isfinite(p).all() and (p > 0).all(), 'Q2 expects positive prices'
    wb = load_workbook(folder / 'result2.xlsx', read_only=True, data_only=True)
    labels = [str(x) for x in next(wb['计划购电量'].values)[1:145]]
    wb.close()
    assert len(labels) == T
    return Data(ds, l * DT, v * DT, p, labels)


def forecast(history_l, history_v, policy):
    """Only past rows enter this function; never receives target-day actuals."""
    n = len(history_l)
    if n == 0:
        raise ValueError('Bootstrap day must not request a forecast')
    if n >= 14 and policy.load_mode == 'week_mean':
        l = (history_l[-7] + history_l[-14]) / 2
    elif n >= 7:
        l = history_l[-7].copy()
    else:
        l = history_l.mean(axis=0)
    v = history_v[-min(n, policy.pv_days):].mean(axis=0)
    return l.copy(), v.copy()


def historical_risk_buffer(data, day, policy, alpha=0.80, window=28):
    """仅用day之前的历史预测残差，计算144时段的上行净负荷风险修正量(kWh)。

    residual = 实际净负荷 - 当时可得到的预测净负荷。
    对每个10分钟时段取历史残差的alpha分位数，并截断为非负：
        B_t = max(0, Q_alpha(residual_t))

    该修正只用于0:00日前计划购电；实时滚动未来负荷/PV仍使用原点预测，
    因而不会把“风险修正”重复计入实时预测。
    """
    alpha=float(alpha); window=int(window)
    if not (0.5 <= alpha < 1.0):
        raise ValueError('risk alpha必须满足0.5 <= alpha < 1')
    if window < 1:
        raise ValueError('risk window必须为正整数')
    if day <= 1:
        return np.zeros(T, dtype=float)

    start=max(1, day-window)
    residuals=[]
    for j in range(start, day):
        lh,vh=forecast(data.load[:j], data.pv[:j], policy)
        predicted_net=lh-vh
        actual_net=data.load[j]-data.pv[j]
        residuals.append(actual_net-predicted_net)

    if not residuals:
        return np.zeros(T, dtype=float)
    q=np.quantile(np.vstack(residuals), alpha, axis=0)
    return np.maximum(q, 0.0)



def historical_net_residual_matrix(data, day, policy, window=28):
    """Return past net-load forecast residuals (actual - forecast), using past-only forecasts."""
    window=int(window)
    if window < 1:
        raise ValueError('V2 residual window must be positive')
    if day <= 1:
        return np.zeros((0,T),dtype=float), np.zeros(0,dtype=int)
    start=max(1,day-window)
    residuals=[]; indices=[]
    for j in range(start,day):
        lh,vh=forecast(data.load[:j],data.pv[:j],policy)
        residuals.append((data.load[j]-data.pv[j])-(lh-vh))
        indices.append(j)
    return np.vstack(residuals) if residuals else np.zeros((0,T),dtype=float), np.asarray(indices,dtype=int)


def v2_net_profiles(data, day, policy, args):
    """V2 past-only net-load forecast for day.

    Components:
      1) base forecast = existing Q2 historical predictor;
      2) EWMA residual bias = tracks persistent seasonal/regime under/over prediction;
      3) robust uncertainty buffer = k * MAD-scale, capped by an empirical positive tail.

    Day-ahead planning uses corrected_net + risk_buffer.
    Realtime MPC uses corrected_net only, so the risk allowance is not counted twice.
    """
    lh,vh=forecast(data.load[:day],data.pv[:day],policy)
    base=lh-vh
    residuals,idx=historical_net_residual_matrix(data,day,policy,args.v2_residual_window)
    if len(residuals)==0:
        z=np.zeros(T,dtype=float)
        return base.copy(),base.copy(),z,{
            'bias':z,'sigma':z,'tail_cap':z,'history_days':0
        }

    ages=(day-1)-idx
    half=float(args.v2_bias_half_life)
    weights=np.power(0.5,ages/half)
    weights=weights/weights.sum()
    bias_raw=np.sum(residuals*weights[:,None],axis=0)
    bias=float(args.v2_bias_shrink)*bias_raw

    centered=residuals-bias[None,:]
    med=np.median(centered,axis=0)
    mad=1.4826*np.median(np.abs(centered-med[None,:]),axis=0)
    std=np.std(centered,axis=0)
    # MAD is robust; 0.35*std avoids a zero-risk estimate in very small/quantized samples.
    sigma=np.maximum(mad,0.35*std)
    raw_buffer=float(args.v2_risk_k)*sigma
    tail=np.quantile(centered,float(args.v2_tail_quantile),axis=0)
    tail_cap=np.maximum(tail,0.0)
    buffer=np.minimum(np.maximum(raw_buffer,0.0),tail_cap)

    corrected=base+bias
    return base.copy(),corrected,buffer,{
        'bias':bias,'sigma':sigma,'tail_cap':tail_cap,'history_days':len(residuals)
    }


def current_action_net(net_hat, current_net, q, t, s0, policy, price, lam, limit):
    """Realtime action when future information is represented directly as net-load forecast."""
    net=np.asarray(net_hat[t:],dtype=float).copy()
    net[0]=float(current_net)
    arrays=optimize(net,price[t:],s0,policy,lam,q[t:],limit)
    return np.array([x[0] for x in arrays])


def simulate_day_v2(lactual,vactual,plan_net,realtime_net,price,s0,policy,lam,limit):
    """V2 day simulation: risk-aware day-ahead plan + bias-corrected realtime horizon."""
    plan_net=np.asarray(plan_net,dtype=float)
    realtime_net=np.asarray(realtime_net,dtype=float)
    if plan_net.shape!=(T,) or realtime_net.shape!=(T,):
        raise ValueError('V2 net profiles must be 144-dimensional')
    if not np.isfinite(plan_net).all() or not np.isfinite(realtime_net).all():
        raise ValueError('V2 net profiles contain non-finite values')
    q=optimize(plan_net,price,s0,policy,lam,time_limit=limit)[0]
    out=np.zeros((T,7))
    s=s0
    for t in range(T):
        current_net=float(lactual[t]-vactual[t])
        e,c,a,w,snext=current_action_net(realtime_net,current_net,q,t,s,policy,price,lam,limit)
        out[t]=q[t],e,c,a,w,s,snext
        s=snext
    audit(out,lactual,vactual)
    return out


@lru_cache(maxsize=144)
def skeleton(n):
    # Blocks: purchase(q or emergency), charge, discharge, unused, S_end,
    # binary_charge, reserve_shortfall; final variable: terminal_shortfall.
    nv = 7 * n + 1
    a = lil_matrix((5 * n + 1, nv), dtype=float)
    for t in range(n):
        a[t, t] = 1
        a[t, n+t] = -1
        a[t, 2*n+t] = 1
        a[t, 3*n+t] = -1
        a[n+t, 4*n+t] = 1
        if t:
            a[n+t, 4*n+t-1] = -1
        a[n+t, n+t] = -ETA
        a[n+t, 2*n+t] = 1/ETA
        a[2*n+t, n+t] = 1
        a[2*n+t, 5*n+t] = -B
        a[3*n+t, 2*n+t] = 1
        a[3*n+t, 5*n+t] = B
        a[4*n+t, 4*n+t] = -1
        a[4*n+t, 6*n+t] = -1
    a[5*n, 5*n-1] = -1
    a[5*n, -1] = -1
    lo = np.zeros(nv)
    hi = np.full(nv, np.inf)
    hi[n:3*n] = B
    lo[4*n:5*n] = EMIN
    hi[4*n:5*n] = EMAX
    hi[5*n:6*n] = 1
    hi[6*n:7*n] = EMAX
    hi[-1] = EINITIAL
    integrality = np.zeros(nv)
    integrality[5*n:6*n] = 1
    return a.tocsr(), Bounds(lo, hi), integrality


def optimize(net, price, s0, policy, terminal_lambda, fixed_q=None, time_limit=30.):
    n = len(net)
    a, base_bounds, integer = skeleton(n)
    lb = np.array(base_bounds.lb, dtype=float, copy=True)
    ub = np.array(base_bounds.ub, dtype=float, copy=True)
    # 实时阶段第1组变量为紧急购电e。其上界为剩余净负荷缺口，
    # 防止5倍紧急电用于给储能充电。
    if fixed_q is not None:
        ub[:n] = np.maximum(np.asarray(net) - np.asarray(fixed_q), 0.0)
    bounds = Bounds(lb, ub)
    rhs = np.zeros(5*n+1)
    rhs[:n] = net if fixed_q is None else net-fixed_q
    rhs[n] = s0
    rhs[3*n:4*n] = B
    rhs[4*n:5*n] = -(EMIN + policy.reserve)
    rhs[-1] = -EINITIAL
    lower = np.full_like(rhs, -np.inf)
    lower[:2*n] = rhs[:2*n]
    obj = np.zeros(7*n+1)
    obj[:n] = price * (1 if fixed_q is None else 5)
    obj[6*n:7*n] = policy.mu_factor * terminal_lambda * DT
    obj[-1] = terminal_lambda
    # Small repeated MILPs are faster without starting many solver threads.
    # scipy forwards the native HiGHS option, while warning it is not a scipy option.
    with warnings.catch_warnings():
        warnings.filterwarnings('ignore', message='Unrecognized options detected.*')
        res = milp(obj, integrality=integer, bounds=bounds,
                   constraints=LinearConstraint(a, lower, rhs),
                   options={'time_limit': time_limit, 'mip_rel_gap': 1e-7, 'threads': 1})
    # Fail explicitly rather than silently publish a nonoptimal/time-limited plan.
    if not res.success or res.x is None:
        raise RuntimeError(f'MILP failed: {res.message}; horizon={n}, initial={s0}')
    x = res.x
    values = a @ x
    if np.max(values-rhs) > TOL or np.max(abs(values[:2*n]-rhs[:2*n])) > TOL:
        raise RuntimeError('Solver feasibility audit failed')
    return tuple(x[i*n:(i+1)*n].copy() for i in range(5))


def current_action(lhat, vhat, current_l, current_v, q, t, s0, policy, price, lam, limit):
    """Receives only a scalar current observation and future FORECASTS."""
    net = (lhat[t:] - vhat[t:]).copy()
    net[0] = current_l-current_v
    arrays = optimize(net, price[t:], s0, policy, lam, q[t:], limit)
    return np.array([x[0] for x in arrays])


def simulate_day(lactual, vactual, lhat, vhat, price, s0, policy, lam, limit):
    q = optimize(lhat-vhat, price, s0, policy, lam, time_limit=limit)[0]
    out = np.zeros((T, 7))  # q,e,c,a,w,S_start,S_end
    s = s0
    for t in range(T):
        e,c,a,w,snext = current_action(lhat, vhat, float(lactual[t]), float(vactual[t]),
                                      q, t, s, policy, price, lam, limit)
        out[t] = q[t],e,c,a,w,s,snext
        s = snext
    audit(out, lactual, vactual)
    return out


def simulate_day_risk_buffer(lactual, vactual, lhat, vhat, price, s0, policy,
                             lam, limit, buffer_kwh):
    """风险修正版：只在0:00日前购电计划中加入buffer；实时预测仍保持原lhat/vhat。"""
    buffer_kwh=np.asarray(buffer_kwh,dtype=float)
    if buffer_kwh.shape != (T,) or not np.isfinite(buffer_kwh).all() or np.min(buffer_kwh) < -TOL:
        raise ValueError('risk buffer必须是144维非负有限数组')
    q = optimize((lhat-vhat)+buffer_kwh, price, s0, policy, lam, time_limit=limit)[0]
    out = np.zeros((T, 7))  # q,e,c,a,w,S_start,S_end
    s = s0
    for t in range(T):
        e,c,a,w,snext = current_action(lhat, vhat, float(lactual[t]), float(vactual[t]),
                                      q, t, s, policy, price, lam, limit)
        out[t] = q[t],e,c,a,w,s,snext
        s = snext
    audit(out, lactual, vactual)
    return out


def bootstrap(l, v, s0):
    out = np.zeros((T,7))
    out[:,1] = np.maximum(l-v,0)
    out[:,4] = np.maximum(v-l,0)
    out[:,5:] = s0
    audit(out,l,v)
    return out


def audit(x, l, v):
    q,e,c,a,w,s0,s1 = x.T
    emergency_cap = np.maximum(l - v - q, 0.0)
    checks = [np.max(abs(q+e+v+a-l-c-w)),
              np.max(abs(s1-s0-ETA*c+a/ETA)),
              np.max(abs(s1[:-1]-s0[1:])),
              max(0., EMIN-min(s0.min(),s1.min())),
              max(0., max(s0.max(),s1.max())-EMAX),
              max(0., c.max()-B,a.max()-B),
              max(0., -x[:,:5].min()),
              np.maximum(np.minimum(c,a),0).max(),
              max(0., float(np.max(e-emergency_cap)))]
    if max(checks) > TOL or not np.isfinite(x).all():
        raise AssertionError(f'Physical audit failed: {checks}')


def bill(x, price):
    return float(price @ x[:,0]), float(5*price @ x[:,1])


def candidates():
    """返回当前实验模式的预声明候选集合。

    main:
        原最终主模型的8个候选，完全保持旧版顺序；
    no-reserve:
        公平B0，只固定 reserve=0, mu=0，预测方式仍由过去窗口选择；
    r-sensitivity:
        固定某个R，只在4种预测方式之间滚动选择。
    """
    if EXPERIMENT_MODE in {'main','risk-buffer','v2'}:
        return [Policy(l,k,r,m) for l in ['week','week_mean'] for k in [3,7]
                for r,m in [(0.,0.),(600.,0.10)]]
    if EXPERIMENT_MODE == 'no-reserve':
        return [Policy(l,k,0.0,0.0)
                for l in ['week','week_mean'] for k in [3,7]]
    if EXPERIMENT_MODE == 'r-sensitivity':
        if FIXED_RESERVE is None or FIXED_MU_FACTOR is None:
            raise RuntimeError('r-sensitivity模式未设置FIXED_RESERVE/FIXED_MU_FACTOR')
        return [Policy(l,k,float(FIXED_RESERVE),float(FIXED_MU_FACTOR))
                for l in ['week','week_mean'] for k in [3,7]]
    raise RuntimeError(f'未知EXPERIMENT_MODE={EXPERIMENT_MODE!r}')



def candidate_description():
    if EXPERIMENT_MODE == 'main':
        return "负荷week/week_mean × PV 3/7天 × (R,mu_factor)=(0,0)/(600,0.10)"
    if EXPERIMENT_MODE == 'risk-buffer':
        return ("风险修正实验：保持main的预测/软备用候选集合；"
                "仅在0:00日前计划净负荷中加入历史预测误差分位数buffer")
    if EXPERIMENT_MODE == 'v2':
        return ("V2：原预测/软备用候选集合 + 过去残差EWMA偏差校正 + "
                "MAD稳健风险缓冲；日前计入风险缓冲，实时未来仅用偏差校正预测")
    if EXPERIMENT_MODE == 'risk-screen':
        return ("快速筛选：直接复用既有主模型每日已选Policy，不重新tune；"
                "仅用于筛选risk-alpha，不作为正式全年结果")
    if EXPERIMENT_MODE == 'no-reserve':
        return "公平B0：负荷week/week_mean × PV 3/7天；固定(R,mu_factor)=(0,0)"
    if EXPERIMENT_MODE == 'r-sensitivity':
        return (f"固定R={float(FIXED_RESERVE):.0f} kWh, "
                f"mu_factor={float(FIXED_MU_FACTOR):.2f}；"
                "负荷week/week_mean × PV 3/7天")
    return EXPERIMENT_MODE


def tune(data, day, snapshots, args, lam, log):
    start = max(14, day-args.validation_days)
    if start >= day:
        return Policy()
    scored = []
    for pol in candidates():
        state = snapshots[start]
        total = 0.
        for j in range(start,day):
            lh,vh = forecast(data.load[:j],data.pv[:j],pol)
            if EXPERIMENT_MODE == 'risk-buffer':
                buffer=historical_risk_buffer(data,j,pol,args.risk_alpha,args.risk_window)
                x=simulate_day_risk_buffer(data.load[j],data.pv[j],lh,vh,data.price,state,
                                           pol,lam,args.time_limit,buffer)
            elif EXPERIMENT_MODE == 'v2':
                _,corrected,buffer,_=v2_net_profiles(data,j,pol,args)
                x=simulate_day_v2(data.load[j],data.pv[j],corrected+buffer,corrected,
                                  data.price,state,pol,lam,args.time_limit)
            else:
                x = simulate_day(data.load[j],data.pv[j],lh,vh,data.price,state,pol,lam,args.time_limit)
            total += sum(bill(x,data.price))
            state = x[-1,6]
        # Same linear inventory valuation for ALL candidates, including surplus.
        correction = lam * max(EINITIAL-state, 0.0)
        score = total+correction
        record = {'selection_date':str(data.dates[day]),'validation_start':str(data.dates[start]),
                  'validation_end':str(data.dates[day-1]), **pol.__dict__,
                  'actual_bill':total,'end_soc':state,'common_terminal_correction':correction,'score':score}
        log.append(record)
        scored.append((score,pol))
        print(f'  validation {pol}: bill={total:.2f}, score={score:.2f}',flush=True)
    return min(scored,key=lambda z:z[0])[1]


def csv_write(path, rows, headers=None):
    with path.open('w',encoding='utf-8-sig',newline='') as f:
        writer = csv.DictWriter(f,fieldnames=headers or list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def export_forecast_diagnostics(data, policies, results, output):
    """复现每天0点实际采用的预测规则并输出误差诊断；不重新求解MILP。"""
    daily_rows = []
    all_el, all_ev, all_en = [], [], []
    under_daily, emg_daily, emg_cost_daily = [], [], []
    month_groups = {}
    for j, ds in enumerate(data.dates):
        if ds < date(2025, 2, 1) or j not in results:
            continue
        pol = policies[j]
        lhat, vhat = forecast(data.load[:j], data.pv[:j], pol)
        la, va = data.load[j], data.pv[j]
        el_kw = (lhat - la) / DT
        ev_kw = (vhat - va) / DT
        en_kw = ((lhat-vhat) - (la-va)) / DT
        under_kwh = float(np.maximum((la-va)-(lhat-vhat), 0.0).sum())
        over_kwh = float(np.maximum((lhat-vhat)-(la-va), 0.0).sum())
        x = results[j]
        _, ce = bill(x, data.price)
        row = {
            'date': str(ds), 'load_mode': pol.load_mode, 'pv_days': pol.pv_days,
            'load_mae_kw': float(np.mean(np.abs(el_kw))),
            'load_rmse_kw': float(np.sqrt(np.mean(el_kw**2))),
            'pv_mae_kw': float(np.mean(np.abs(ev_kw))),
            'pv_rmse_kw': float(np.sqrt(np.mean(ev_kw**2))),
            'net_mae_kw': float(np.mean(np.abs(en_kw))),
            'net_rmse_kw': float(np.sqrt(np.mean(en_kw**2))),
            'net_under_kwh': under_kwh, 'net_over_kwh': over_kwh,
            'emergency_kwh': float(x[:,1].sum()), 'emergency_cost': ce,
        }
        daily_rows.append(row)
        all_el.append(el_kw); all_ev.append(ev_kw); all_en.append(en_kw)
        under_daily.append(under_kwh); emg_daily.append(row['emergency_kwh']); emg_cost_daily.append(ce)
        month_groups.setdefault(str(ds)[:7], []).append(row)
    if not daily_rows:
        return
    def corr(a, b):
        a=np.asarray(a,float); b=np.asarray(b,float)
        if len(a)<2 or np.std(a)<1e-12 or np.std(b)<1e-12: return float('nan')
        return float(np.corrcoef(a,b)[0,1])
    el=np.concatenate(all_el); ev=np.concatenate(all_ev); en=np.concatenate(all_en)
    summary = {
        'evaluation_days': len(daily_rows),
        'load_mae_kw': float(np.mean(np.abs(el))),
        'load_rmse_kw': float(np.sqrt(np.mean(el**2))),
        'pv_mae_kw': float(np.mean(np.abs(ev))),
        'pv_rmse_kw': float(np.sqrt(np.mean(ev**2))),
        'net_load_mae_kw': float(np.mean(np.abs(en))),
        'net_load_rmse_kw': float(np.sqrt(np.mean(en**2))),
        'net_under_total_kwh': float(sum(r['net_under_kwh'] for r in daily_rows)),
        'net_over_total_kwh': float(sum(r['net_over_kwh'] for r in daily_rows)),
        'corr_daily_under_vs_emergency_kwh': corr(under_daily, emg_daily),
        'corr_daily_under_vs_emergency_cost': corr(under_daily, emg_cost_daily),
    }
    monthly=[]
    for m in sorted(month_groups):
        rr=month_groups[m]
        monthly.append({
            'month':m,'days':len(rr),
            'mean_load_mae_kw':float(np.mean([r['load_mae_kw'] for r in rr])),
            'mean_pv_mae_kw':float(np.mean([r['pv_mae_kw'] for r in rr])),
            'mean_net_mae_kw':float(np.mean([r['net_mae_kw'] for r in rr])),
            'net_under_kwh':float(sum(r['net_under_kwh'] for r in rr)),
            'emergency_kwh':float(sum(r['emergency_kwh'] for r in rr)),
            'emergency_cost':float(sum(r['emergency_cost'] for r in rr)),
        })
    csv_write(output/'forecast_error_daily.csv', daily_rows)
    csv_write(output/'forecast_error_monthly.csv', monthly)
    (output/'forecast_error_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2), encoding='utf-8')


def export(data, results, policies, output, template, args, lam, tuning_log):
    """输出官方result2（仅3个sheet）及论文/审计辅助CSV。"""
    wb = load_workbook(template)
    official = {'计划购电量','充放电量','紧急购电量'}
    for ws in list(wb.worksheets):
        if ws.title not in official:
            wb.remove(ws)
    if set(wb.sheetnames) != official:
        raise AssertionError(f'result2模板工作表异常: {wb.sheetnames}')
    plan = wb['计划购电量']
    plan.delete_rows(2, max(0, plan.max_row-1))
    bat = wb['充放电量']; em = wb['紧急购电量']
    for ws in [bat, em]:
        for merged in list(ws.merged_cells.ranges): ws.unmerge_cells(str(merged))
        ws.delete_rows(2, max(0, ws.max_row-1))
    daily=[]; detail=[]; selected=[]; emergency_rows=[]
    specified_dates = {date(2025,3,20),date(2025,6,21),date(2025,9,23),date(2025,12,21)}
    specified_slots = {60,72,84,96,108,120}
    for j,x in results.items():
        ds = data.dates[j]; cp,ce = bill(x,data.price)
        daily.append({'date':str(ds),'evaluation':ds>=date(2025,2,1),
                      'planned_kwh':float(x[:,0].sum()),'emergency_kwh':float(x[:,1].sum()),
                      'planned_cost':cp,'emergency_cost':ce,'total_cost':cp+ce,
                      'soc_start':float(x[0,5]),'soc_end':float(x[-1,6]),
                      'emergency_intervals_10min':int(np.sum(x[:,1]>1e-6)),**policies[j].__dict__})
        if ds < date(2025,2,1): continue
        plan.append([ds]+x[:,0].tolist()+[float(x[:,0].sum()),cp+ce])
        for block in range(6):
            y=x[24*block:24*(block+1)]
            bat.append([ds if block==0 else None,f'{block*4}:00-{(block+1)*4}:00',
                        float(y[:,2].sum()),float(y[:,3].sum()),
                        '0:00' if block==0 else ('24:00' if block==1 else None),
                        float(x[0,5]) if block==0 else (float(x[-1,6]) if block==1 else None)])
        intervals=[]; t=0
        while t<T:
            if x[t,1]<=1e-6: t+=1; continue
            end=t+1
            while end<T and x[end,1]>1e-6: end+=1
            label=physical_label(t).split('-')[0]+'-'+physical_label(end-1).split('-')[1]
            amount=float(x[t:end,1].sum())
            intervals.append((label,amount,t,end))
            emergency_rows.append({'date':str(ds),'emergency_interval':label,'emergency_kwh':amount,
                                   'start_slot':t,'end_slot_exclusive':end})
            t=end
        for k,(label,amount,_,__) in enumerate(intervals): em.append([ds if k==0 else None,label,amount])
        if not intervals: em.append([ds,'无',0.])
        for t,row in enumerate(x):
            record={'date':str(ds),'slot_index':t,'attachment_time':attachment_time_label(t),
                    'physical_interval':physical_label(t),'template_label':data.labels[t],
                    'price':float(data.price[t]),'load_kwh':float(data.load[j,t]),'pv_kwh':float(data.pv[j,t])}
            record.update(dict(zip(['plan_kwh','emergency_kwh','charge_kwh','discharge_kwh','unused_kwh','soc_start','soc_end'],map(float,row))))
            detail.append(record)
            if ds in specified_dates and t in specified_slots: selected.append(record)
    for ws in [plan,bat,em]:
        ws.freeze_panes='B2'
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value,date): cell.number_format='yyyy/mm/dd'
                elif isinstance(cell.value,(float,int)): cell.number_format='0.0000'
    eval_indices=[j for j,ds in enumerate(data.dates) if ds>=date(2025,2,1)]
    complete = (len(eval_indices)==334 and all(j in results for j in eval_indices))
    target = output/('result2.xlsx' if complete else 'result2_partial.xlsx')
    wb.save(target)
    csv_write(output/'daily_summary.csv',daily)
    if detail:
        csv_write(output/'dispatch_detail_physical.csv',detail)
        export_daily_soc(output,output)
    if selected: csv_write(output/'specified_dates_table1_physical.csv',selected)
    if emergency_rows: csv_write(output/'emergency_intervals_physical.csv',emergency_rows)
    if tuning_log: csv_write(output/'parameter_validation.csv',tuning_log)
    mapping=[{'slot_index':t,'attachment_time':attachment_time_label(t),
              'physical_interval':physical_label(t),'template_label':data.labels[t]} for t in range(T)]
    csv_write(output/'time_mapping.csv',mapping)
    risk_note = ''
    if EXPERIMENT_MODE == 'risk-buffer':
        risk_note = (f"\nrisk_alpha={args.risk_alpha}\n"
                     f"risk_window_days={args.risk_window}\n"
                     "risk_buffer说明=仅以过去预测残差分位数上调0:00日前净负荷；实时未来预测不加buffer。\n")
    elif EXPERIMENT_MODE == 'v2':
        risk_note = (f"\nV2_residual_window={args.v2_residual_window}\n"
                     f"V2_bias_half_life={args.v2_bias_half_life}\n"
                     f"V2_bias_shrink={args.v2_bias_shrink}\n"
                     f"V2_risk_k={args.v2_risk_k}\n"
                     f"V2_tail_quantile={args.v2_tail_quantile}\n"
                     "V2说明=过去残差EWMA校正季节性偏差；MAD估计不确定度；日前计入风险buffer，实时未来只用校正预测。\n")
    notes = f'''Q2最终模型实现口径

1. 日前：每天0:00只用历史信息预测并锁定常规购电q。
2. 实时：只读取当前实际负荷/PV；未来仍使用0点预测；每10分钟滚动MILP。
3. 紧急购电：5倍交易电价，只补负荷缺口，不允许用于储能充电。
4. 储能：SOC 1200~10800 kWh；最大充/放电功率5000 kW；效率0.9/0.9；充放互斥。
5. 软备用：R与mu为软风险惩罚，不是硬SOC下限。
6. 日末6000 kWh仅为软参考，用于缓解滚动时域末端短视，不计入实际账单。
7. 跨日：2025-01-01 0:00为6000 kWh，之后实际SOC连续传递。
8. 评价：2025-02-01~2025-12-31；1月用于历史积累与运行预热。
9. 时间：附件第i条数据对应模型第i个10分钟物理区间；计划q不循环移位。
10. result2.xlsx只保留官方三个sheet。

lambda={lam:.10f}
lambda_factor={args.lambda_factor}
说明：按原公式，lambda倍数同时影响终端缺口惩罚、软备用惩罚及验证期终端修正。
validation_days={args.validation_days}
tune_every={args.tune_every}
实验模式={EXPERIMENT_MODE}
候选集合={candidate_description()}
{risk_note}'''
    (output/'model_notes.txt').write_text(notes,encoding='utf-8')
    eval_rows=[r for r in daily if r['evaluation']]
    planned_cost=sum(r['planned_cost'] for r in eval_rows); emergency_cost=sum(r['emergency_cost'] for r in eval_rows)
    planned_kwh=sum(r['planned_kwh'] for r in eval_rows); emergency_kwh=sum(r['emergency_kwh'] for r in eval_rows)
    simulated_days=(365 if complete and getattr(args,'resume_from',None) else len(results))
    report={'complete_year':complete,'simulated_days':simulated_days,'evaluation_days':len(eval_rows),
            'planned_kwh':planned_kwh,'emergency_kwh':emergency_kwh,'total_purchase_kwh':planned_kwh+emergency_kwh,
            'planned_cost':planned_cost,'emergency_cost':emergency_cost,'total_cost':planned_cost+emergency_cost,
            'emergency_days':sum(r['emergency_kwh']>1e-6 for r in eval_rows),
            'emergency_slots_10min':sum(r['emergency_intervals_10min'] for r in eval_rows),
            'terminal_soc':daily[-1]['soc_end'],'lambda':lam,
            'experiment_mode':EXPERIMENT_MODE,
            'candidate_description':candidate_description(),
            'arguments':{k:str(v) if isinstance(v,Path) else v for k,v in vars(args).items()}}
    (output/'run_summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    if EXPERIMENT_MODE == 'v2':
        export_v2_diagnostics(data,policies,results,output,args)
    else:
        export_forecast_diagnostics(data,policies,results,output)
    check=load_workbook(target,read_only=True,data_only=True)
    assert check.sheetnames == ['计划购电量','充放电量','紧急购电量']
    rows=list(check['计划购电量'].values)[1:]; expected=eval_rows
    assert len(rows)==len(expected)
    if complete:
        assert len(expected)==334
    for row,r in zip(rows,expected):
        assert abs(sum(row[1:145])-r['planned_kwh'])<1e-5
        assert abs(row[146]-r['total_cost'])<1e-5
    check.close()
    print(json.dumps(report,ensure_ascii=False,indent=2),flush=True)
    print(f'Output: {target}',flush=True)
    return report



# =====================================================================
# 已有结果后处理：不重新求MILP
# =====================================================================

def read_csv_rows(path):
    path = Path(path)
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def first_existing(folder, names):
    folder = Path(folder)
    for name in names:
        p = folder / name
        if p.exists():
            return p
    raise FileNotFoundError(
        f"{folder} 中找不到以下任一文件：" + ", ".join(names)
    )


def _evaluation_flag(v):
    return str(v).strip().lower() in {'true', '1', 'yes', 'y'}


def _configure_plot_font(plt):
    try:
        from matplotlib import font_manager
        installed = {f.name for f in font_manager.fontManager.ttflist}
        for name in ['Microsoft YaHei','SimHei','Noto Sans CJK SC',
                     'Source Han Sans SC','PingFang SC','Arial Unicode MS']:
            if name in installed:
                plt.rcParams['font.sans-serif'] = [name]
                break
    except Exception:
        pass
    plt.rcParams['axes.unicode_minus'] = False


def _load_matplotlib(no_plots=False):
    if no_plots:
        return None, None
    try:
        import matplotlib.pyplot as plt
        import matplotlib.dates as mdates
        _configure_plot_font(plt)
        return plt, mdates
    except Exception as exc:
        print(f'[warning] matplotlib不可用，仅输出数值诊断：{exc}', flush=True)
        return None, None


def aggregate_daily_output(output_dir):
    """兼容新旧Q2输出目录，汇总正式评价期334天。"""
    daily_path = first_existing(output_dir, [
        'daily_summary.csv',
        'daily_summary_main.csv',
        'daily_summary_no_reserve.csv',
    ])
    rows = read_csv_rows(daily_path)
    ev = [r for r in rows if _evaluation_flag(r.get('evaluation', 'false'))]
    if not ev:
        # 极老版本若没有evaluation字段，则按日期>=2025-02-01。
        ev = [r for r in rows if r.get('date','') >= '2025-02-01']

    slot_key = ('emergency_intervals_10min'
                if ev and 'emergency_intervals_10min' in ev[0]
                else 'emergency_intervals')

    out = {
        'evaluation_days': len(ev),
        'planned_kwh': sum(float(r['planned_kwh']) for r in ev),
        'emergency_kwh': sum(float(r['emergency_kwh']) for r in ev),
        'planned_cost': sum(float(r['planned_cost']) for r in ev),
        'emergency_cost': sum(float(r['emergency_cost']) for r in ev),
        'total_cost': sum(float(r['total_cost']) for r in ev),
        'emergency_days': sum(float(r['emergency_kwh']) > 1e-6 for r in ev),
        'emergency_slots_10min': sum(int(float(r.get(slot_key, 0) or 0)) for r in ev),
        'terminal_soc': float(rows[-1]['soc_end']),
        'daily_path': str(daily_path),
    }
    return out, rows, ev



def verify_no_reserve_output(output_dir):
    """核验公平B0：评价期和参数验证记录中的R、mu必须全部为0。"""
    output_dir=Path(output_dir)
    _, rows, ev = aggregate_daily_output(output_dir)
    bad=[
        r for r in ev
        if abs(float(r.get('reserve',0)))>1e-12
        or abs(float(r.get('mu_factor',0)))>1e-12
    ]
    if bad:
        raise AssertionError('公平无备用结果中出现非0 reserve/mu')

    pv=output_dir/'parameter_validation.csv'
    if pv.exists():
        rr=read_csv_rows(pv)
        for r in rr:
            if abs(float(r.get('reserve',0)))>1e-12 or abs(float(r.get('mu_factor',0)))>1e-12:
                raise AssertionError('无备用parameter_validation中出现非0 reserve/mu')
    print('[B0 verify] PASS：仅取消软备用，预测滚动选择仍保留。',flush=True)


def verify_r_sensitivity_output(output_dir, reserve):
    """保留原q2_r_sensitivity.py的实验核验与metadata。"""
    output_dir=Path(output_dir)
    reserve=float(reserve)
    pv=output_dir/'parameter_validation.csv'
    if pv.exists():
        rr=read_csv_rows(pv)
        if rr:
            vals=sorted({round(float(r['reserve']),9) for r in rr})
            if vals != [round(reserve,9)]:
                raise AssertionError(f'R敏感性参数验证出现非目标R：{vals}，目标={reserve}')

    summary_path=output_dir/'run_summary.json'
    if not summary_path.exists():
        raise FileNotFoundError(summary_path)
    summary=json.loads(summary_path.read_text(encoding='utf-8'))
    if summary.get('complete_year',False):
        if int(summary.get('evaluation_days',-1)) != 334:
            raise AssertionError('完整R敏感性实验的正式评价天数不是334')

    meta={
        'experiment':'fixed_soft_reserve_sensitivity',
        'reserve_kwh':reserve,
        'mu_factor':0.0 if reserve==0 else 0.10,
        'forecast_candidates':[
            {'load_mode':'week','pv_days':3},
            {'load_mode':'week','pv_days':7},
            {'load_mode':'week_mean','pv_days':3},
            {'load_mode':'week_mean','pv_days':7},
        ],
        'note':(
            'R固定，仅预测方法仍按rolling_dispatch.py的历史滚动验证选择；'
            '该实验用于离线敏感性分析，不用于回填过去决策。'
        ),
    }
    (output_dir/'sensitivity_experiment_meta.json').write_text(
        json.dumps(meta,ensure_ascii=False,indent=2),encoding='utf-8'
    )
    print(f'[R sensitivity verify] PASS：R={reserve:.0f} kWh',flush=True)


def verify_reference_summary(summary, mode, strict=False):
    """与已经真实跑完的最终结果核对；不参与优化。"""
    if mode not in REFERENCE_RESULTS:
        print(f'[reference] {mode} 无内置参考值，跳过。', flush=True)
        return True

    ref = REFERENCE_RESULTS[mode]
    checks = {}
    # 允许跨机器/HiGHS版本的极小数值差异，但不能掩盖策略改变。
    tolerances = {
        'evaluation_days': 0,
        'planned_cost': 20.0,
        'emergency_cost': 20.0,
        'total_cost': 20.0,
        'planned_kwh': 5.0,
        'emergency_kwh': 5.0,
        'emergency_days': 0,
        'emergency_slots_10min': 2,
        'terminal_soc': 2.0,
    }
    ok = True
    for k, rv in ref.items():
        av = summary.get(k)
        if av is None:
            checks[k] = {'actual': None, 'reference': rv, 'ok': False}
            ok = False
            continue
        diff = abs(float(av)-float(rv))
        tol = tolerances[k]
        this_ok = diff <= tol
        checks[k] = {'actual': av, 'reference': rv, 'abs_diff': diff, 'tolerance': tol, 'ok': this_ok}
        ok &= this_ok

    print('[reference] ' + ('PASS' if ok else 'FAIL') + f' mode={mode}', flush=True)
    if not ok:
        for k,v in checks.items():
            if not v['ok']:
                print('  mismatch', k, v, flush=True)
        if strict:
            raise AssertionError(f'{mode} 与既有最终结果不对应')
    return ok


def soc_emergency_diagnostics(main_output, diagnostic_output, no_plots=False):
    """只读取已有CSV，输出SOC/紧急购电诊断与论文图。"""
    main_output = Path(main_output)
    diagnostic_output = Path(diagnostic_output)
    diagnostic_output.mkdir(parents=True, exist_ok=True)

    detail_path = first_existing(main_output, [
        'dispatch_detail_physical.csv',
        'dispatch_detail.csv',
    ])
    daily_path = first_existing(main_output, [
        'daily_summary.csv',
        'daily_summary_main.csv',
    ])

    export_daily_soc(main_output,diagnostic_output)
    detail = read_csv_rows(detail_path)
    daily = read_csv_rows(daily_path)
    daily_eval = [r for r in daily if _evaluation_flag(r.get('evaluation','false'))]
    if not daily_eval:
        daily_eval = [r for r in daily if r.get('date','') >= '2025-02-01']

    # detail最终版只保存评价期；若某旧版含1月，显式过滤。
    detail = [r for r in detail if r.get('date','') >= '2025-02-01']
    dates = sorted({r['date'] for r in detail})
    if len(dates) != 334:
        raise AssertionError(f'SOC诊断应有334个评价日，实际{len(dates)}')
    if len(detail) != 334*T:
        raise AssertionError(f'dispatch明细应有{334*T}行，实际{len(detail)}')

    by_date = defaultdict(list)
    for r in detail:
        by_date[r['date']].append(r)
    for d in dates:
        by_date[d].sort(key=lambda r: int(r['slot_index']))
        if len(by_date[d]) != T:
            raise AssertionError(f'{d} 不是144个10分钟时段')

    soc_mat = np.array([
        [float(r['soc_end']) for r in by_date[d]] for d in dates
    ], dtype=float)
    emg_mat = np.array([
        [float(r['emergency_kwh']) for r in by_date[d]] for d in dates
    ], dtype=float)
    charge_mat = np.array([
        [float(r['charge_kwh']) for r in by_date[d]] for d in dates
    ], dtype=float)
    discharge_mat = np.array([
        [float(r['discharge_kwh']) for r in by_date[d]] for d in dates
    ], dtype=float)

    day_end = soc_mat[:, -1]
    positive_days = int(np.sum(emg_mat.sum(axis=1) > 1e-6))
    positive_slots = int(np.sum(emg_mat > 1e-6))
    touch_low = int(np.sum(np.min(soc_mat, axis=1) <= EMIN + 1e-5))
    touch_high = int(np.sum(np.max(soc_mat, axis=1) >= EMAX - 1e-5))

    summary = {
        'evaluation_days': 334,
        'detail_rows': len(detail),
        'soc_all_mean_kwh': float(soc_mat.mean()),
        'soc_all_median_kwh': float(np.median(soc_mat)),
        'soc_day_end_mean_kwh': float(day_end.mean()),
        'soc_day_end_median_kwh': float(np.median(day_end)),
        'soc_min_kwh': float(soc_mat.min()),
        'soc_max_kwh': float(soc_mat.max()),
        'days_touch_soc_min': touch_low,
        'days_touch_soc_max': touch_high,
        'emergency_positive_days': positive_days,
        'emergency_positive_slots_10min': positive_slots,
        'emergency_slot_share_pct': float(100*positive_slots/(334*T)),
        'emergency_kwh': float(emg_mat.sum()),
        'emergency_cost_yuan': float(sum(float(r['emergency_cost']) for r in daily_eval)),
        'charge_total_kwh': float(charge_mat.sum()),
        'discharge_total_kwh': float(discharge_mat.sum()),
        'year_end_soc_kwh': float(day_end[-1]),
    }

    # 月度汇总
    monthly = defaultdict(lambda: {'emergency_kwh':0.0,'emergency_cost':0.0,'days':0,'slots':0})
    for r in daily_eval:
        m = r['date'][:7]
        e = float(r['emergency_kwh'])
        monthly[m]['emergency_kwh'] += e
        monthly[m]['emergency_cost'] += float(r['emergency_cost'])
        monthly[m]['days'] += int(e > 1e-6)
        slot_key = 'emergency_intervals_10min' if 'emergency_intervals_10min' in r else 'emergency_intervals'
        monthly[m]['slots'] += int(float(r.get(slot_key,0) or 0))
    monthly_rows = [
        {'month':m, **monthly[m]} for m in sorted(monthly)
    ]

    csv_write(diagnostic_output/'soc_emergency_monthly.csv', monthly_rows)
    (diagnostic_output/'soc_emergency_summary.json').write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8'
    )
    with (diagnostic_output/'soc_emergency_summary.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f); w.writerow(['指标','数值'])
        for k,v in summary.items(): w.writerow([k,v])
    (diagnostic_output/'soc_emergency_summary.txt').write_text(
        '\n'.join(f'{k}: {v}' for k,v in summary.items()), encoding='utf-8'
    )

    plt, mdates = _load_matplotlib(no_plots)
    if plt is not None:
        # 1) 评价期SOC：每日最小-最大区间 + 日末SOC轨迹
        # 原始334天×144时段轨迹过于密集。这里不改变任何模型结果，只将
        # 每天的144个10分钟SOC压缩成“日内范围 + 日末状态”，提高论文可读性。
        from datetime import datetime as _dt
        day_x=[]; day_min=[]; day_max=[]; day_end=[]
        for d in dates:
            rr=by_date[d]
            x=_dt.strptime(d,'%Y-%m-%d')
            soc_path=[float(rr[0]['soc_start'])] + [float(r['soc_end']) for r in rr]
            day_x.append(x)
            day_min.append(min(soc_path))
            day_max.append(max(soc_path))
            day_end.append(float(rr[-1]['soc_end']))

        fig,ax=plt.subplots(figsize=(13.8,5.4))
        # 日内SOC波动范围：浅蓝带
        ax.fill_between(day_x,day_min,day_max,
                        color='#AFC3D8',alpha=.35,linewidth=0,
                        label='日内SOC范围（最小–最大）')
        # 日末SOC：深蓝主线
        ax.plot(day_x,day_end,linewidth=2.1,color='#104E8B',
                label='日末SOC')

        # 物理边界与参考库存
        ax.axhline(EMIN,linestyle='--',linewidth=1.25,color='#B22222')
        ax.axhline(EMAX,linestyle='--',linewidth=1.25,color='#B22222')
        ax.axhline(EINITIAL,linestyle=':',linewidth=1.25,color='#5F89B1')

        ax.set_ylim(0,12000)
        ax.set_xlabel('日期')
        ax.set_ylabel('SOC（kWh）')
        ax.set_title('评价期储能SOC日内波动区间及日末状态')
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%m月'))
        ax.grid(axis='y',color='#D7E1EB',alpha=.80,linewidth=.9)
        ax.set_axisbelow(True)
        for side in ['top','right']:
            ax.spines[side].set_visible(False)
        ax.spines['left'].set_color('#AFC3D8')
        ax.spines['bottom'].set_color('#AFC3D8')

        # 图例放在图外左下方，避免遮挡数据。
        ax.legend(loc='upper left',bbox_to_anchor=(0.0,-0.13),
                  ncol=2,frameon=False,borderaxespad=0.0)

        # 三条关键参考线仅在右侧标注，不再添加“SOC轨迹”长箭头。
        guide=dict(arrowstyle='-',linewidth=.9,shrinkA=2,shrinkB=2)
        ax.annotate('上限 10800 kWh',xy=(day_x[-1],EMAX),xycoords='data',
                    xytext=(1.015,EMAX),textcoords=ax.get_yaxis_transform(),
                    ha='left',va='center',color='#B22222',
                    arrowprops={**guide,'color':'#B22222'},annotation_clip=False)
        ax.annotate('参考值 6000 kWh',xy=(day_x[-1],EINITIAL),xycoords='data',
                    xytext=(1.015,EINITIAL),textcoords=ax.get_yaxis_transform(),
                    ha='left',va='center',color='#5F89B1',
                    arrowprops={**guide,'color':'#5F89B1'},annotation_clip=False)
        ax.annotate('下限 1200 kWh',xy=(day_x[-1],EMIN),xycoords='data',
                    xytext=(1.015,EMIN),textcoords=ax.get_yaxis_transform(),
                    ha='left',va='center',color='#B22222',
                    arrowprops={**guide,'color':'#B22222'},annotation_clip=False)

        fig.subplots_adjust(right=.80,bottom=.22)
        fig.savefig(diagnostic_output/'Fig_SOC全年轨迹.svg',format='svg',bbox_inches='tight')
        plt.close(fig)

        # 2) 日内平均SOC（之前最需要补的诊断）
        hours=np.arange(1,T+1)/6.0
        mean_soc=soc_mat.mean(axis=0)
        q25=np.quantile(soc_mat,.25,axis=0); q75=np.quantile(soc_mat,.75,axis=0)
        fig,ax=plt.subplots(figsize=(10,5.2))
        ax.plot(hours,mean_soc,color='#104E8B',label='平均SOC')
        ax.fill_between(hours,q25,q75,color='#C5E9E3',alpha=.65,label='25%–75%区间')
        ax.axhline(EMIN,linestyle='--',linewidth=1,color='#B22222')
        ax.axhline(EMAX,linestyle='--',linewidth=1,color='#B22222')
        ax.set_xlim(0,24); ax.set_xticks([0,4,8,12,16,20,24])
        ax.set_xlabel('时刻（h）'); ax.set_ylabel('SOC（kWh）')
        ax.set_title('储能SOC日内平均演化')
        ax.grid(color='#D7E1EB',alpha=.8); ax.legend(frameon=False)
        fig.tight_layout()
        fig.savefig(diagnostic_output/'Fig_SOC日内平均.svg',format='svg',bbox_inches='tight')
        plt.close(fig)

        # 3) 紧急购电热力图
        fig,ax=plt.subplots(figsize=(12,7))
        from matplotlib.colors import LinearSegmentedColormap
        _emg_cmap=LinearSegmentedColormap.from_list('research_blue',['#D7E1EB','#C5E9E3','#AFC3D8','#5F89B1','#376B9E','#104E8B'])
        im=ax.imshow(emg_mat,aspect='auto',origin='upper',cmap=_emg_cmap)
        cb=fig.colorbar(im,ax=ax); cb.set_label('紧急购电量（kWh / 10min）')
        ax.set_xticks([0,24,48,72,96,120,143])
        ax.set_xticklabels(['0','4','8','12','16','20','24'])
        month_pos=[]; month_lab=[]; last=None
        for i,d in enumerate(dates):
            m=d[:7]
            if m!=last:
                month_pos.append(i); month_lab.append(d[5:7]+'月'); last=m
        ax.set_yticks(month_pos); ax.set_yticklabels(month_lab)
        ax.set_xlabel('日内时刻（h）'); ax.set_ylabel('日期')
        ax.set_title('全年紧急购电时空分布')
        fig.tight_layout()
        fig.savefig(diagnostic_output/'Fig_紧急购电热力图.svg',format='svg',bbox_inches='tight')
        plt.close(fig)

        # 4) 月度紧急购电
        months=[r['month'] for r in monthly_rows]
        energy=np.array([r['emergency_kwh'] for r in monthly_rows])/1000
        cost=np.array([r['emergency_cost'] for r in monthly_rows])/10000
        xx=np.arange(len(months))
        fig,ax1=plt.subplots(figsize=(11,5.2))
        ax1.bar(xx,energy,color='#5F89B1',label='紧急购电量')
        ax1.set_xticks(xx); ax1.set_xticklabels([m[-2:]+'月' for m in months])
        ax1.set_xlabel('月份'); ax1.set_ylabel('紧急购电量（MWh）')
        ax1.grid(axis='y',color='#D7E1EB',alpha=.8)
        ax2=ax1.twinx(); ax2.plot(xx,cost,marker='o',color='#B22222',label='紧急购电费')
        ax2.set_ylabel('紧急购电费（万元）')
        h1,l1=ax1.get_legend_handles_labels(); h2,l2=ax2.get_legend_handles_labels()
        ax1.legend(h1+h2,l1+l2,frameon=False)
        ax1.set_title('月度紧急购电规模与费用')
        fig.tight_layout()
        fig.savefig(diagnostic_output/'Fig_月度紧急购电.svg',format='svg',bbox_inches='tight')
        plt.close(fig)

    print(json.dumps(summary,ensure_ascii=False,indent=2),flush=True)
    print(f'SOC/紧急购电诊断输出：{diagnostic_output}',flush=True)
    return summary


def ablation_compare(main_output, no_reserve_output, diagnostic_output, no_plots=False):
    """主模型 vs 公平无备用B0。"""
    diagnostic_output = Path(diagnostic_output)
    diagnostic_output.mkdir(parents=True, exist_ok=True)
    main_sum, main_rows, main_ev = aggregate_daily_output(main_output)
    b0_sum, b0_rows, b0_ev = aggregate_daily_output(no_reserve_output)

    # 公平性硬检查：B0评价期必须全部R=0,mu=0。
    bad = [
        r for r in b0_ev
        if abs(float(r.get('reserve',0))) > 1e-12
        or abs(float(r.get('mu_factor',0))) > 1e-12
    ]
    if bad:
        raise AssertionError('无备用输出中发现 reserve/mu 非0，不是公平B0')

    main_dates=[r['date'] for r in main_ev]; b0_dates=[r['date'] for r in b0_ev]
    if main_dates != b0_dates:
        raise AssertionError('主模型和B0评价日期集合不一致')

    savings = b0_sum['total_cost'] - main_sum['total_cost']
    out = {
        'main': main_sum,
        'no_reserve': b0_sum,
        'main_saves_yuan_vs_no_reserve': savings,
        'main_saves_pct_vs_no_reserve': 100*savings/b0_sum['total_cost'],
        'emergency_cost_reduction_pct': 100*(b0_sum['emergency_cost']-main_sum['emergency_cost'])/b0_sum['emergency_cost'],
        'emergency_kwh_reduction_pct': 100*(b0_sum['emergency_kwh']-main_sum['emergency_kwh'])/b0_sum['emergency_kwh'],
    }

    rows = [
        {'方案':'主模型_软备用', **{k:v for k,v in main_sum.items() if k!='daily_path'}},
        {'方案':'B0_无备用', **{k:v for k,v in b0_sum.items() if k!='daily_path'}},
    ]
    csv_write(diagnostic_output/'ablation_main_vs_no_reserve.csv', rows)
    (diagnostic_output/'ablation_summary.json').write_text(
        json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8'
    )

    plt,_ = _load_matplotlib(no_plots)
    if plt is not None:
        names=['主模型\n软备用','B0\n无备用']
        plan=np.array([main_sum['planned_cost'],b0_sum['planned_cost']])/10000
        emg=np.array([main_sum['emergency_cost'],b0_sum['emergency_cost']])/10000
        ekwh=np.array([main_sum['emergency_kwh'],b0_sum['emergency_kwh']])/10000
        x=np.arange(2)
        fig,ax1=plt.subplots(figsize=(9.6,5.2))
        ax1.bar(x,plan,color=['#376B9E','#5F89B1'],label='计划购电费')
        ax1.bar(x,emg,bottom=plan,color='#E5B5B5',label='紧急购电费')
        ax1.set_xticks(x); ax1.set_xticklabels(names)
        ax1.set_ylabel('全年费用（万元）'); ax1.grid(axis='y',color='#D7E1EB',alpha=.8)
        ax2=ax1.twinx(); ax2.plot(x,ekwh,marker='o',linestyle='--',color='#B22222',label='紧急购电量')
        ax2.set_ylabel('紧急购电量（万kWh）')
        h1,l1=ax1.get_legend_handles_labels(); h2,l2=ax2.get_legend_handles_labels()
        ax1.legend(h1+h2,l1+l2,frameon=False,loc='upper left',bbox_to_anchor=(1.13,1.0),borderaxespad=0.)
        ax1.set_title('软备用机制公平消融：主模型 vs 无备用B0')
        fig.subplots_adjust(right=.78)
        fig.savefig(diagnostic_output/'Fig_软备用公平消融.svg',format='svg',bbox_inches='tight')
        plt.close(fig)

    print(json.dumps(out,ensure_ascii=False,indent=2),flush=True)
    print(f'消融对比输出：{diagnostic_output}',flush=True)
    return out


def recompute_forecast_diagnostics(data_dir, main_output, diagnostic_output):
    """保留旧 q2_forecast_diagnostics.py 的能力，不重新求MILP。"""
    data_dir=Path(data_dir); main_output=Path(main_output); diagnostic_output=Path(diagnostic_output)
    diagnostic_output.mkdir(parents=True,exist_ok=True)

    data=read_data(data_dir)
    daily_path=first_existing(main_output,['daily_summary.csv','daily_summary_main.csv'])
    daily_rows=read_csv_rows(daily_path)
    daily_map={r['date']:r for r in daily_rows}

    out=[]; all_el=[]; all_ev=[]; all_en=[]
    under=[]; emg=[]; emg_cost=[]; months=defaultdict(list)
    for j,ds in enumerate(data.dates):
        if ds < date(2025,2,1): continue
        r=daily_map.get(str(ds))
        if r is None: raise AssertionError(f'daily_summary缺少{ds}')
        pol=Policy(r['load_mode'],int(float(r['pv_days'])),
                   float(r.get('reserve',0)),float(r.get('mu_factor',0)))
        lh,vh=forecast(data.load[:j],data.pv[:j],pol)
        la,va=data.load[j],data.pv[j]
        el=(lh-la)/DT; epv=(vh-va)/DT; en=((lh-vh)-(la-va))/DT
        uk=float(np.maximum((la-va)-(lh-vh),0).sum())
        ok=float(np.maximum((lh-vh)-(la-va),0).sum())
        row={
            'date':str(ds),'load_mode':pol.load_mode,'pv_days':pol.pv_days,
            'load_mae_kw':float(np.mean(np.abs(el))),
            'load_rmse_kw':float(np.sqrt(np.mean(el**2))),
            'pv_mae_kw':float(np.mean(np.abs(epv))),
            'pv_rmse_kw':float(np.sqrt(np.mean(epv**2))),
            'net_mae_kw':float(np.mean(np.abs(en))),
            'net_rmse_kw':float(np.sqrt(np.mean(en**2))),
            'net_under_kwh':uk,'net_over_kwh':ok,
            'emergency_kwh':float(r['emergency_kwh']),
            'emergency_cost':float(r['emergency_cost']),
        }
        out.append(row); all_el.append(el); all_ev.append(epv); all_en.append(en)
        under.append(uk); emg.append(row['emergency_kwh']); emg_cost.append(row['emergency_cost'])
        months[str(ds)[:7]].append(row)

    def corr(a,b):
        a=np.asarray(a,float); b=np.asarray(b,float)
        if len(a)<2 or np.std(a)<1e-12 or np.std(b)<1e-12:return float('nan')
        return float(np.corrcoef(a,b)[0,1])

    el=np.concatenate(all_el); epv=np.concatenate(all_ev); en=np.concatenate(all_en)
    summary={
        'evaluation_days':len(out),
        'load_mae_kw':float(np.mean(np.abs(el))),
        'load_rmse_kw':float(np.sqrt(np.mean(el**2))),
        'pv_mae_kw':float(np.mean(np.abs(epv))),
        'pv_rmse_kw':float(np.sqrt(np.mean(epv**2))),
        'net_load_mae_kw':float(np.mean(np.abs(en))),
        'net_load_rmse_kw':float(np.sqrt(np.mean(en**2))),
        'net_under_total_kwh':float(sum(r['net_under_kwh'] for r in out)),
        'net_over_total_kwh':float(sum(r['net_over_kwh'] for r in out)),
        'corr_daily_under_vs_emergency_kwh':corr(under,emg),
        'corr_daily_under_vs_emergency_cost':corr(under,emg_cost),
    }
    monthly=[]
    for m in sorted(months):
        rr=months[m]
        monthly.append({
            'month':m,'days':len(rr),
            'mean_load_mae_kw':float(np.mean([r['load_mae_kw'] for r in rr])),
            'mean_pv_mae_kw':float(np.mean([r['pv_mae_kw'] for r in rr])),
            'mean_net_mae_kw':float(np.mean([r['net_mae_kw'] for r in rr])),
            'net_under_kwh':float(sum(r['net_under_kwh'] for r in rr)),
            'emergency_kwh':float(sum(r['emergency_kwh'] for r in rr)),
            'emergency_cost':float(sum(r['emergency_cost'] for r in rr)),
        })
    csv_write(diagnostic_output/'forecast_error_daily.csv',out)
    csv_write(diagnostic_output/'forecast_error_monthly.csv',monthly)
    (diagnostic_output/'forecast_error_summary.json').write_text(
        json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8'
    )
    print(json.dumps(summary,ensure_ascii=False,indent=2),flush=True)
    return summary


def forecast_method_benchmark(data_dir, main_output, diagnostic_output):
    """比较三种日前预测方法的MAE/RMSE；只读历史数据和已有主模型策略，不求解MILP。"""
    data_dir=Path(data_dir); main_output=Path(main_output); diagnostic_output=Path(diagnostic_output)
    diagnostic_output.mkdir(parents=True,exist_ok=True)

    data=read_data(data_dir)
    daily_path=first_existing(main_output,['daily_summary.csv','daily_summary_main.csv'])
    daily_rows=read_csv_rows(daily_path)
    daily_map={r['date']:r for r in daily_rows}

    methods={
        '昨日持续法': {'load':[], 'pv':[], 'net':[]},
        '最近7日均值': {'load':[], 'pv':[], 'net':[]},
        '本文滚动选择': {'load':[], 'pv':[], 'net':[]},
    }

    evaluation_days=0
    for j,ds in enumerate(data.dates):
        if ds < date(2025,2,1):
            continue
        evaluation_days += 1
        if j < 7:
            raise AssertionError('评价期历史长度不足7天')

        la=data.load[j]/DT
        va=data.pv[j]/DT

        # A. 昨日持续法：前一天同一时段作为今日预测
        lhat=data.load[j-1]/DT
        vhat=data.pv[j-1]/DT
        methods['昨日持续法']['load'].append(lhat-la)
        methods['昨日持续法']['pv'].append(vhat-va)
        methods['昨日持续法']['net'].append((lhat-vhat)-(la-va))

        # B. 最近7日均值：最近7天同一时段简单平均
        lhat=data.load[j-7:j].mean(axis=0)/DT
        vhat=data.pv[j-7:j].mean(axis=0)/DT
        methods['最近7日均值']['load'].append(lhat-la)
        methods['最近7日均值']['pv'].append(vhat-va)
        methods['最近7日均值']['net'].append((lhat-vhat)-(la-va))

        # C. 本文方法：严格复现主模型当天实际采用的滚动选择预测规则
        r=daily_map.get(str(ds))
        if r is None:
            raise AssertionError(f'daily_summary缺少{ds}')
        pol=Policy(r['load_mode'],int(float(r['pv_days'])),
                   float(r.get('reserve',0)),float(r.get('mu_factor',0)))
        lh,vh=forecast(data.load[:j],data.pv[:j],pol)
        lhat=lh/DT; vhat=vh/DT
        methods['本文滚动选择']['load'].append(lhat-la)
        methods['本文滚动选择']['pv'].append(vhat-va)
        methods['本文滚动选择']['net'].append((lhat-vhat)-(la-va))

    if evaluation_days != 334:
        raise AssertionError(f'预测方法比较应有334个评价日，实际{evaluation_days}')

    rows=[]
    for name,err in methods.items():
        el=np.concatenate(err['load'])
        epv=np.concatenate(err['pv'])
        en=np.concatenate(err['net'])
        rows.append({
            'method':name,
            'evaluation_days':evaluation_days,
            'load_mae_kw':float(np.mean(np.abs(el))),
            'load_rmse_kw':float(np.sqrt(np.mean(el**2))),
            'pv_mae_kw':float(np.mean(np.abs(epv))),
            'pv_rmse_kw':float(np.sqrt(np.mean(epv**2))),
            'net_load_mae_kw':float(np.mean(np.abs(en))),
            'net_load_rmse_kw':float(np.sqrt(np.mean(en**2))),
        })

    csv_write(diagnostic_output/'forecast_method_comparison.csv',rows)

    print('预测方法MAE/RMSE比较（不求解MILP）：',flush=True)
    for r in rows:
        print(
            f"{r['method']}: "
            f"Load MAE={r['load_mae_kw']:.3f}, RMSE={r['load_rmse_kw']:.3f}; "
            f"PV MAE={r['pv_mae_kw']:.3f}, RMSE={r['pv_rmse_kw']:.3f}; "
            f"Net MAE={r['net_load_mae_kw']:.3f}, RMSE={r['net_load_rmse_kw']:.3f}",
            flush=True
        )
    print(f"输出：{diagnostic_output/'forecast_method_comparison.csv'}",flush=True)
    return rows



def export_v2_diagnostics(data,policies,results,output,args):
    """Export V2 forecast correction/risk diagnostics without re-solving MILPs."""
    output=Path(output)
    rows=[]; base_err=[]; corr_err=[]
    for j,ds in enumerate(data.dates):
        if ds<date(2025,2,1) or j not in results:
            continue
        pol=policies[j]
        base,corrected,buffer,meta=v2_net_profiles(data,j,pol,args)
        actual=data.load[j]-data.pv[j]
        be=base-actual; ce=corrected-actual
        x=results[j];cp,emc=bill(x,data.price)
        rows.append({
            'date':str(ds),'load_mode':pol.load_mode,'pv_days':pol.pv_days,
            'reserve':pol.reserve,'mu_factor':pol.mu_factor,
            'history_days':meta['history_days'],
            'base_net_mae_kw':float(np.mean(np.abs(be))/DT),
            'corrected_net_mae_kw':float(np.mean(np.abs(ce))/DT),
            'base_net_bias_kw':float(np.mean(be)/DT),
            'corrected_net_bias_kw':float(np.mean(ce)/DT),
            'mean_abs_bias_correction_kw':float(np.mean(np.abs(meta['bias']))/DT),
            'mean_risk_buffer_kw':float(np.mean(buffer)/DT),
            'max_risk_buffer_kw':float(np.max(buffer)/DT),
            'planned_cost':cp,'emergency_cost':emc,
            'emergency_kwh':float(x[:,1].sum()),'total_cost':cp+emc,
        })
        base_err.append(be);corr_err.append(ce)
    if not rows:return
    csv_write(output/'v2_daily_diagnostics.csv',rows)
    b=np.concatenate(base_err);c=np.concatenate(corr_err)
    summary={
        'evaluation_days':len(rows),
        'base_net_mae_kw':float(np.mean(np.abs(b))/DT),
        'corrected_net_mae_kw':float(np.mean(np.abs(c))/DT),
        'base_net_rmse_kw':float(np.sqrt(np.mean(b*b))/DT),
        'corrected_net_rmse_kw':float(np.sqrt(np.mean(c*c))/DT),
        'base_net_bias_kw':float(np.mean(b)/DT),
        'corrected_net_bias_kw':float(np.mean(c)/DT),
        'mean_daily_risk_buffer_kw':float(np.mean([r['mean_risk_buffer_kw'] for r in rows])),
        'v2_residual_window':int(args.v2_residual_window),
        'v2_bias_half_life':float(args.v2_bias_half_life),
        'v2_bias_shrink':float(args.v2_bias_shrink),
        'v2_risk_k':float(args.v2_risk_k),
        'v2_tail_quantile':float(args.v2_tail_quantile),
    }
    (output/'v2_diagnostics_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')


def compare_v2_to_reference(reference_output,v2_output):
    """Compare V2 to any existing Q2 output on exactly the V2 evaluation dates."""
    reference_output=Path(reference_output);v2_output=Path(v2_output)
    try:
        _,_,ref_all=aggregate_daily_output(reference_output)
        _,_,v2_ev=aggregate_daily_output(v2_output)
    except FileNotFoundError:
        return None
    ref_map={r['date']:r for r in ref_all}; dates=[r['date'] for r in v2_ev]
    if any(d not in ref_map for d in dates):
        raise AssertionError('V2包含参考结果中不存在的评价日期')
    ref=[ref_map[d] for d in dates]
    def agg(rows):
        return {
            'planned_kwh':sum(float(r['planned_kwh']) for r in rows),
            'emergency_kwh':sum(float(r['emergency_kwh']) for r in rows),
            'planned_cost':sum(float(r['planned_cost']) for r in rows),
            'emergency_cost':sum(float(r['emergency_cost']) for r in rows),
            'total_cost':sum(float(r['total_cost']) for r in rows),
            'emergency_days':sum(float(r['emergency_kwh'])>1e-6 for r in rows),
            'emergency_slots_10min':sum(int(float(r['emergency_intervals_10min'])) for r in rows),
            'terminal_soc':float(rows[-1]['soc_end']) if rows else float('nan'),
        }
    a=agg(ref);b=agg(v2_ev)
    rows=[]
    for key in ['planned_kwh','emergency_kwh','planned_cost','emergency_cost','total_cost',
                'emergency_days','emergency_slots_10min','terminal_soc']:
        x=float(a[key]);y=float(b[key])
        rows.append({'metric':key,'reference':x,'v2':y,'v2_minus_reference':y-x,
                     'saving_reference_minus_v2':x-y,
                     'improvement_pct':100*(x-y)/x if abs(x)>1e-12 else None})
    csv_write(v2_output/'v2_vs_reference.csv',rows)
    summary={
        'reference_dir':str(reference_output),
        'evaluation_days':len(v2_ev),
        'reference_total_cost':a['total_cost'],'v2_total_cost':b['total_cost'],
        'saving_yuan':a['total_cost']-b['total_cost'],
        'saving_pct':100*(a['total_cost']-b['total_cost'])/a['total_cost'],
        'reference_emergency_cost':a['emergency_cost'],'v2_emergency_cost':b['emergency_cost'],
        'reference_emergency_kwh':a['emergency_kwh'],'v2_emergency_kwh':b['emergency_kwh'],
        'reference_emergency_days':a['emergency_days'],'v2_emergency_days':b['emergency_days'],
    }
    (v2_output/'v2_vs_reference.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print('[V2 vs reference]',json.dumps(summary,ensure_ascii=False),flush=True)
    return summary



def run_v2_local_screen(args):
    """Very fast local V2 screen on a selected date range using source SOC/Policy at range start.

    This deliberately does NOT reproduce the V2 path before local_start, so it is only a
    directional screening tool. It is useful for testing crisis months (e.g. June) cheaply.
    """
    global EXPERIMENT_MODE
    EXPERIMENT_MODE='v2'
    source=Path(args.v2_policy_source or args.main_output)
    schedule,source_rows=_policy_schedule_from_main(source)
    source_map={r['date']:r for r in source_rows}
    start=date.fromisoformat(args.local_start)
    end=date.fromisoformat(args.screen_end)
    if not date(2025,2,1)<=start<=end<=date(2025,12,31):
        raise ValueError('local-start/screen-end must satisfy 2025-02-01 <= start <= end <= 2025-12-31')
    if str(start) not in source_map:
        raise AssertionError(f'policy source missing {start}')
    data=read_data(args.data_dir);lam=float(data.price.min()/ETA)*args.lambda_factor
    state=float(source_map[str(start)]['soc_start'])
    output=Path(args.output_dir or (Path(__file__).resolve().parent/'q2_v2_local_screen'))
    output.mkdir(parents=True,exist_ok=True)
    rows=[];begin=time.perf_counter()
    for j,ds in enumerate(data.dates):
        if ds<start:continue
        if ds>end:break
        if str(ds) not in schedule: raise AssertionError(f'policy source missing {ds}')
        pol=schedule[str(ds)]
        _,corrected,buffer,meta=v2_net_profiles(data,j,pol,args)
        x=simulate_day_v2(data.load[j],data.pv[j],corrected+buffer,corrected,
                          data.price,state,pol,lam,args.time_limit)
        cp,ce=bill(x,data.price)
        rows.append({
            'date':str(ds),'planned_kwh':float(x[:,0].sum()),'emergency_kwh':float(x[:,1].sum()),
            'planned_cost':cp,'emergency_cost':ce,'total_cost':cp+ce,
            'soc_start':float(x[0,5]),'soc_end':float(x[-1,6]),
            'emergency_intervals_10min':int(np.sum(x[:,1]>1e-6)),
            'load_mode':pol.load_mode,'pv_days':pol.pv_days,'reserve':pol.reserve,'mu_factor':pol.mu_factor,
            'mean_bias_kw':float(np.mean(meta['bias'])/DT),'mean_abs_bias_kw':float(np.mean(np.abs(meta['bias']))/DT),
            'mean_buffer_kw':float(np.mean(buffer)/DT),
        })
        state=float(x[-1,6])
        print(f'{ds}: V2-local plan={cp:.2f}, emergency={ce:.2f}, SOC={state:.2f}, elapsed={time.perf_counter()-begin:.1f}s',flush=True)
    if not rows: raise ValueError('empty local screen range')
    csv_write(output/'v2_local_daily.csv',rows)
    ref_rows=[source_map[r['date']] for r in rows]
    def agg(rr):
        return {
            'planned_cost':sum(float(r['planned_cost']) for r in rr),
            'emergency_cost':sum(float(r['emergency_cost']) for r in rr),
            'total_cost':sum(float(r['total_cost']) for r in rr),
            'emergency_kwh':sum(float(r['emergency_kwh']) for r in rr),
            'emergency_days':sum(float(r['emergency_kwh'])>1e-6 for r in rr),
            'emergency_slots_10min':sum(int(float(r['emergency_intervals_10min'])) for r in rr),
        }
    ref=agg(ref_rows);v2=agg(rows)
    summary={
        'screening_only':True,'official_result':False,'local_start':str(start),'local_end':str(end),
        'initial_soc_from_reference':float(source_map[str(start)]['soc_start']),
        'reference_dir':str(source),'reference':ref,'v2_local':v2,
        'saving_yuan':ref['total_cost']-v2['total_cost'],
        'saving_pct':100*(ref['total_cost']-v2['total_cost'])/ref['total_cost'],
        'v2_parameters':{
            'residual_window':args.v2_residual_window,'bias_half_life':args.v2_bias_half_life,
            'bias_shrink':args.v2_bias_shrink,'risk_k':args.v2_risk_k,'tail_quantile':args.v2_tail_quantile,
        },
        'warning':'Local screen starts from the reference SOC rather than a V2-generated prior path; use only for fast directional screening.',
    }
    (output/'v2_local_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print('[V2 local screen]',json.dumps(summary,ensure_ascii=False,indent=2),flush=True)
    return summary

def run_v2_screen(args):
    """Fast V2 screen: reuse a completed/partial policy schedule; no re-tuning."""
    global EXPERIMENT_MODE
    EXPERIMENT_MODE='v2'
    source=Path(args.v2_policy_source or args.main_output)
    schedule,source_rows=_policy_schedule_from_main(source)
    data=read_data(args.data_dir);lam=float(data.price.min()/ETA)*args.lambda_factor
    screen_end=date.fromisoformat(args.screen_end)
    if not date(2025,2,1)<=screen_end<=date(2025,12,31):
        raise ValueError('screen-end must be within evaluation period')
    needed=[str(d) for d in data.dates if d<=screen_end]
    missing=[d for d in needed if d not in schedule]
    if missing: raise AssertionError(f'policy source missing dates, e.g. {missing[:3]}')
    output=Path(args.output_dir or (Path(__file__).resolve().parent/'q2_v2_screen'))
    output.mkdir(parents=True,exist_ok=True)
    state=EINITIAL;results={};policies={};begin=time.perf_counter()
    for j,ds in enumerate(data.dates):
        if ds>screen_end:break
        pol=schedule[str(ds)]
        if j==0:
            x=bootstrap(data.load[j],data.pv[j],state)
        else:
            _,corrected,buffer,_=v2_net_profiles(data,j,pol,args)
            x=simulate_day_v2(data.load[j],data.pv[j],corrected+buffer,corrected,
                              data.price,state,pol,lam,args.time_limit)
        results[j]=x;policies[j]=pol;state=float(x[-1,6])
        cp,ce=bill(x,data.price)
        print(f'{ds}: V2-screen plan={cp:.2f}, emergency={ce:.2f}, SOC={state:.2f}, elapsed={time.perf_counter()-begin:.1f}s',flush=True)
    # use regular export, but mark as screen in a separate file
    report=export(data,results,policies,output,args.data_dir/'result2.xlsx',args,lam,[])
    export_v2_diagnostics(data,policies,results,output,args)
    compare_v2_to_reference(source,output)
    (output/'SCREENING_ONLY.txt').write_text(
        'V2-screen reuses an existing daily Policy schedule and does not re-run rolling policy selection.\n'
        'Use --mode v2 for a formal result.\n',encoding='utf-8')
    return report

def export_risk_buffer_diagnostics(data, policies, results, output, alpha, window):
    """风险修正实验的可审计输出；只重算buffer，不重新求MILP。"""
    output=Path(output)
    rows=[]
    for j,ds in enumerate(data.dates):
        if ds < date(2025,2,1) or j not in results:
            continue
        pol=policies[j]
        b=historical_risk_buffer(data,j,pol,alpha,window)
        x=results[j]
        cp,ce=bill(x,data.price)
        rows.append({
            'date':str(ds),
            'risk_alpha':float(alpha),
            'risk_window_days':int(window),
            'history_days_used':min(max(j-1,0),int(window)),
            'load_mode':pol.load_mode,
            'pv_days':pol.pv_days,
            'reserve':pol.reserve,
            'mu_factor':pol.mu_factor,
            'buffer_signal_total_kwh':float(b.sum()),
            'buffer_mean_kw':float(b.mean()/DT),
            'buffer_max_kw':float(b.max()/DT),
            'buffer_positive_slots':int(np.sum(b>1e-12)),
            'planned_kwh':float(x[:,0].sum()),
            'emergency_kwh':float(x[:,1].sum()),
            'planned_cost':cp,
            'emergency_cost':ce,
            'total_cost':cp+ce,
        })
    if not rows:
        return
    csv_write(output/'risk_buffer_daily.csv',rows)
    summary={
        'evaluation_days':len(rows),
        'risk_alpha':float(alpha),
        'risk_window_days':int(window),
        'mean_daily_buffer_signal_kwh':float(np.mean([r['buffer_signal_total_kwh'] for r in rows])),
        'mean_buffer_kw':float(np.mean([r['buffer_mean_kw'] for r in rows])),
        'max_buffer_kw':float(max(r['buffer_max_kw'] for r in rows)),
        'mean_positive_slots_per_day':float(np.mean([r['buffer_positive_slots'] for r in rows])),
        'note':'buffer_signal是日前净负荷风险修正信号，不等于最终增加的计划购电量。',
    }
    (output/'risk_buffer_summary.json').write_text(
        json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8'
    )


def compare_risk_buffer_to_main(main_output, risk_output):
    """将风险修正实验与主模型按相同评价日期比较；支持部分日期测试。"""
    main_output=Path(main_output);risk_output=Path(risk_output)
    try:
        _,_,base_ev_all=aggregate_daily_output(main_output)
        _,_,risk_ev=aggregate_daily_output(risk_output)
    except FileNotFoundError:
        return None
    risk_dates=[r['date'] for r in risk_ev]
    base_map={r['date']:r for r in base_ev_all}
    if any(d not in base_map for d in risk_dates):
        raise AssertionError('risk-buffer包含主模型中不存在的评价日期')
    base_ev=[base_map[d] for d in risk_dates]
    def _agg(rows):
        return {
            'planned_kwh':sum(float(r['planned_kwh']) for r in rows),
            'emergency_kwh':sum(float(r['emergency_kwh']) for r in rows),
            'planned_cost':sum(float(r['planned_cost']) for r in rows),
            'emergency_cost':sum(float(r['emergency_cost']) for r in rows),
            'total_cost':sum(float(r['total_cost']) for r in rows),
            'emergency_days':sum(float(r['emergency_kwh'])>1e-6 for r in rows),
            'emergency_slots_10min':sum(int(float(r['emergency_intervals_10min'])) for r in rows),
            'terminal_soc':float(rows[-1]['soc_end']) if rows else float('nan'),
        }
    base=_agg(base_ev); risk=_agg(risk_ev)
    metrics=['planned_kwh','emergency_kwh','planned_cost','emergency_cost',
             'total_cost','emergency_days','emergency_slots_10min','terminal_soc']
    rows=[]
    for key in metrics:
        b=float(base[key]);r=float(risk[key])
        rows.append({
            'metric':key,
            'main':b,
            'risk_buffer':r,
            'difference_risk_minus_main':r-b,
            'improvement_main_minus_risk':b-r,
            'improvement_pct':(100*(b-r)/b if abs(b)>1e-12 else None),
        })
    csv_write(risk_output/'risk_buffer_vs_main.csv',rows)
    summary={
        'main_total_cost':base['total_cost'],
        'risk_total_cost':risk['total_cost'],
        'saving_yuan':base['total_cost']-risk['total_cost'],
        'saving_pct':100*(base['total_cost']-risk['total_cost'])/base['total_cost'],
        'main_emergency_cost':base['emergency_cost'],
        'risk_emergency_cost':risk['emergency_cost'],
        'main_emergency_kwh':base['emergency_kwh'],
        'risk_emergency_kwh':risk['emergency_kwh'],
        'main_emergency_days':base['emergency_days'],
        'risk_emergency_days':risk['emergency_days'],
        'main_emergency_slots_10min':base['emergency_slots_10min'],
        'risk_emergency_slots_10min':risk['emergency_slots_10min'],
    }
    (risk_output/'risk_buffer_vs_main.json').write_text(
        json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8'
    )
    print('[risk-buffer vs main]',json.dumps(summary,ensure_ascii=False),flush=True)
    return summary


def _policy_schedule_from_main(main_output):
    """读取既有主模型daily_summary中的每日已选Policy；不重新做历史验证。"""
    main_output=Path(main_output)
    rows=read_csv_rows(first_existing(main_output,['daily_summary.csv','daily_summary_main.csv']))
    schedule={}
    for r in rows:
        schedule[r['date']]=Policy(
            r.get('load_mode','week_mean'),
            int(float(r.get('pv_days',3))),
            float(r.get('reserve',0)),
            float(r.get('mu_factor',0)),
        )
    return schedule,rows


def _summary_from_daily_rows(rows):
    if not rows:
        raise ValueError('没有可汇总的评价期记录')
    return {
        'evaluation_days':len(rows),
        'evaluation_start':rows[0]['date'],
        'evaluation_end':rows[-1]['date'],
        'planned_kwh':sum(float(r['planned_kwh']) for r in rows),
        'emergency_kwh':sum(float(r['emergency_kwh']) for r in rows),
        'planned_cost':sum(float(r['planned_cost']) for r in rows),
        'emergency_cost':sum(float(r['emergency_cost']) for r in rows),
        'total_cost':sum(float(r['total_cost']) for r in rows),
        'emergency_days':sum(float(r['emergency_kwh'])>1e-6 for r in rows),
        'emergency_slots_10min':sum(int(float(r['emergency_intervals_10min'])) for r in rows),
        'terminal_soc':float(rows[-1]['soc_end']),
    }


def run_risk_screen(args):
    """快速筛选risk-alpha：固定复用主模型每日Policy，不重新tune，不生成正式result2。"""
    global EXPERIMENT_MODE
    EXPERIMENT_MODE='risk-screen'

    screen_end=date.fromisoformat(args.screen_end)
    if not date(2025,2,1) <= screen_end <= date(2025,12,31):
        raise ValueError('screen-end必须位于2025-02-01~2025-12-31')

    data=read_data(args.data_dir)
    lam=float(data.price.min()/ETA)*args.lambda_factor
    policy_schedule,main_rows=_policy_schedule_from_main(args.main_output)

    needed=[str(d) for d in data.dates if d<=screen_end]
    missing=[d for d in needed if d not in policy_schedule]
    if missing:
        raise AssertionError(f'主模型daily_summary缺少筛选所需日期，例如：{missing[:3]}')

    if args.output_dir is None:
        args.output_dir=(Path(__file__).resolve().parent/'q2_risk_screen'/
                         f'alpha_{int(round(args.risk_alpha*100)):03d}_to_{screen_end}')
    output=Path(args.output_dir)
    output.mkdir(parents=True,exist_ok=True)

    state=EINITIAL
    daily=[]
    begin=time.perf_counter()

    for j,ds in enumerate(data.dates):
        if ds>screen_end:
            break
        pol=policy_schedule[str(ds)]

        if j==0:
            x=bootstrap(data.load[j],data.pv[j],state)
            buffer=np.zeros(T)
        else:
            lh,vh=forecast(data.load[:j],data.pv[:j],pol)
            buffer=historical_risk_buffer(data,j,pol,args.risk_alpha,args.risk_window)
            x=simulate_day_risk_buffer(
                data.load[j],data.pv[j],lh,vh,data.price,state,
                pol,lam,args.time_limit,buffer
            )

        cp,ce=bill(x,data.price)
        row={
            'date':str(ds),
            'evaluation':ds>=date(2025,2,1),
            'planned_kwh':float(x[:,0].sum()),
            'emergency_kwh':float(x[:,1].sum()),
            'planned_cost':cp,
            'emergency_cost':ce,
            'total_cost':cp+ce,
            'soc_start':float(x[0,5]),
            'soc_end':float(x[-1,6]),
            'emergency_intervals_10min':int(np.sum(x[:,1]>1e-6)),
            'load_mode':pol.load_mode,
            'pv_days':pol.pv_days,
            'reserve':pol.reserve,
            'mu_factor':pol.mu_factor,
            'risk_alpha':float(args.risk_alpha),
            'risk_window_days':int(args.risk_window),
            'buffer_signal_total_kwh':float(buffer.sum()),
            'buffer_mean_kw':float(buffer.mean()/DT),
            'buffer_max_kw':float(buffer.max()/DT),
            'buffer_positive_slots':int(np.sum(buffer>1e-12)),
        }
        daily.append(row)
        state=float(x[-1,6])

        print(
            f'{ds}: screen plan={cp:.2f}, emergency={ce:.2f}, '
            f'end_SOC={state:.2f}, elapsed={time.perf_counter()-begin:.1f}s',
            flush=True
        )

    csv_write(output/'risk_screen_daily.csv',daily)

    risk_eval=[r for r in daily if r['evaluation']]
    risk_summary=_summary_from_daily_rows(risk_eval)

    main_eval=[
        r for r in main_rows
        if _evaluation_flag(r.get('evaluation','false'))
        and date.fromisoformat(r['date'])<=screen_end
    ]
    if not main_eval:
        main_eval=[r for r in main_rows
                   if '2025-02-01'<=r.get('date','')<=str(screen_end)]
    if [r['date'] for r in main_eval] != [r['date'] for r in risk_eval]:
        raise AssertionError('risk-screen与主模型筛选期日期集合不一致')
    main_summary=_summary_from_daily_rows(main_eval)

    comparison={
        'screening_only':True,
        'official_result':False,
        'policy_source':'既有主模型daily_summary.csv中的每日已选Policy；未重新tune',
        'simulation_start':'2025-01-01',
        'evaluation_start':risk_summary['evaluation_start'],
        'screen_end':str(screen_end),
        'risk_alpha':float(args.risk_alpha),
        'risk_window_days':int(args.risk_window),
        'main':main_summary,
        'risk_screen':risk_summary,
        'saving_yuan_vs_main':main_summary['total_cost']-risk_summary['total_cost'],
        'saving_pct_vs_main':100*(main_summary['total_cost']-risk_summary['total_cost'])/main_summary['total_cost'],
        'emergency_cost_reduction_yuan':main_summary['emergency_cost']-risk_summary['emergency_cost'],
        'emergency_kwh_reduction':main_summary['emergency_kwh']-risk_summary['emergency_kwh'],
        'warning':(
            '该结果只用于快速筛选alpha。由于未在风险修正后的SOC路径上重新执行滚动Policy验证，'
            '不能替代risk-buffer完整全年正式实验，也不能直接写成全年最终答案。'
        ),
    }

    (output/'risk_screen_summary.json').write_text(
        json.dumps(comparison,ensure_ascii=False,indent=2),encoding='utf-8'
    )
    csv_write(output/'risk_screen_vs_main.csv',[
        {
            'risk_alpha':float(args.risk_alpha),
            'risk_window_days':int(args.risk_window),
            'evaluation_start':risk_summary['evaluation_start'],
            'evaluation_end':risk_summary['evaluation_end'],
            'main_total_cost':main_summary['total_cost'],
            'risk_total_cost':risk_summary['total_cost'],
            'saving_yuan_vs_main':comparison['saving_yuan_vs_main'],
            'saving_pct_vs_main':comparison['saving_pct_vs_main'],
            'main_emergency_cost':main_summary['emergency_cost'],
            'risk_emergency_cost':risk_summary['emergency_cost'],
            'main_emergency_kwh':main_summary['emergency_kwh'],
            'risk_emergency_kwh':risk_summary['emergency_kwh'],
            'main_emergency_days':main_summary['emergency_days'],
            'risk_emergency_days':risk_summary['emergency_days'],
            'main_emergency_slots_10min':main_summary['emergency_slots_10min'],
            'risk_emergency_slots_10min':risk_summary['emergency_slots_10min'],
        }
    ])

    (output/'SCREENING_ONLY.txt').write_text(
        '本目录仅为risk-alpha快速筛选结果，不是正式Q2答案。\\n'
        'Policy直接复用既有主模型每日已选结果，未在风险修正后的SOC路径上重新tune。\\n'
        '若某alpha表现较好，最终仍需使用--mode risk-buffer完整全年运行后才能替换正式Q2结果。\\n',
        encoding='utf-8'
    )

    print('[risk-screen]',json.dumps(comparison,ensure_ascii=False,indent=2),flush=True)
    print(f'快速筛选输出：{output}',flush=True)
    return comparison


def run_risk_screen_sweep(args):
    """快速筛选0.60/0.70/0.80/0.90；每组固定复用主模型Policy，不重新tune。"""
    if args.self_test:
        raise ValueError('risk-screen-sweep请勿组合--self-test')
    root=Path(args.risk_screen_results)
    root.mkdir(parents=True,exist_ok=True)
    alphas=[0.60,0.70,0.80,0.90]
    rows=[]

    for alpha in alphas:
        out=root/f'alpha_{int(round(alpha*100)):03d}_to_{args.screen_end}'
        if not args.risk_summary_only:
            if out.exists() and any(out.iterdir()):
                raise FileExistsError(
                    f'{out}非空；请先删除该筛选目录或使用--risk-summary-only汇总已有结果'
                )
            command=[
                sys.executable,str(Path(__file__).resolve()),
                '--mode','risk-screen',
                '--data-dir',str(args.data_dir.resolve()),
                '--main-output',str(args.main_output.resolve()),
                '--output-dir',str(out.resolve()),
                '--risk-alpha',str(alpha),
                '--risk-window',str(args.risk_window),
                '--screen-end',str(args.screen_end),
                '--time-limit',str(args.time_limit),
            ]
            print(f'开始快速筛选 alpha={alpha:.2f}, 截止={args.screen_end}',flush=True)
            subprocess.run(command,check=True)

        summary=json.loads((out/'risk_screen_summary.json').read_text(encoding='utf-8'))
        rows.append({
            'risk_alpha':alpha,
            'risk_window_days':int(args.risk_window),
            'evaluation_start':summary['evaluation_start'],
            'evaluation_end':summary['screen_end'],
            'main_total_cost':summary['main']['total_cost'],
            'risk_total_cost':summary['risk_screen']['total_cost'],
            'saving_yuan_vs_main':summary['saving_yuan_vs_main'],
            'saving_pct_vs_main':summary['saving_pct_vs_main'],
            'main_emergency_cost':summary['main']['emergency_cost'],
            'risk_emergency_cost':summary['risk_screen']['emergency_cost'],
            'main_emergency_kwh':summary['main']['emergency_kwh'],
            'risk_emergency_kwh':summary['risk_screen']['emergency_kwh'],
            'main_emergency_days':summary['main']['emergency_days'],
            'risk_emergency_days':summary['risk_screen']['emergency_days'],
        })

    csv_write(root/'risk_screen_sweep_summary.csv',rows)
    best=min(rows,key=lambda r:r['risk_total_cost'])
    (root/'risk_screen_sweep_summary.json').write_text(
        json.dumps({
            'screening_only':True,
            'official_result':False,
            'experiments':rows,
            'best_by_screen_total_cost':best,
            'warning':'筛选最优alpha必须再用--mode risk-buffer完整全年验证后才能成为正式结果。',
        },ensure_ascii=False,indent=2),
        encoding='utf-8'
    )
    print('快速筛选完成，筛选期总费用最低：',json.dumps(best,ensure_ascii=False),flush=True)
    return rows


def run_risk_buffer_sweep(args):
    """独立运行alpha=0.60/0.70/0.80/0.90四组全年实验并汇总。"""
    if args.no_tune or args.self_test or args.lambda_factor!=1.0:
        raise ValueError('risk-buffer-sweep请勿组合--no-tune、--self-test或非1倍lambda')
    if args.end_date!='2025-12-31':
        raise ValueError('risk-buffer-sweep要求完整全年；短期测试请用--mode risk-buffer --end-date')
    root=Path(args.risk_results)
    root.mkdir(parents=True,exist_ok=True)
    alphas=[0.60,0.70,0.80,0.90]

    for alpha in alphas:
        out=root/f'alpha_{int(round(alpha*100)):03d}'
        if not args.risk_summary_only:
            if out.exists() and any(out.iterdir()):
                raise FileExistsError(f'{out}非空；请换目录或加--risk-summary-only汇总已有结果')
            command=[
                sys.executable,str(Path(__file__).resolve()),
                '--mode','risk-buffer',
                '--data-dir',str(args.data_dir.resolve()),
                '--output-dir',str(out.resolve()),
                '--main-output',str(args.main_output.resolve()),
                '--risk-alpha',str(alpha),
                '--risk-window',str(args.risk_window),
                '--validation-days',str(args.validation_days),
                '--tune-every',str(args.tune_every),
                '--time-limit',str(args.time_limit),
            ]
            print(f'开始risk-buffer全年实验 alpha={alpha:.2f}',flush=True)
            subprocess.run(command,check=True)

    base,_,base_ev=aggregate_daily_output(args.main_output)
    rows=[]
    for alpha in alphas:
        out=root/f'alpha_{int(round(alpha*100)):03d}'
        summary,_,ev=aggregate_daily_output(out)
        if [r['date'] for r in ev] != [r['date'] for r in base_ev]:
            raise AssertionError(f'{out}与主模型评价日期不一致')
        rows.append({
            'risk_alpha':alpha,
            'risk_window_days':int(args.risk_window),
            'total_cost':summary['total_cost'],
            'planned_cost':summary['planned_cost'],
            'emergency_cost':summary['emergency_cost'],
            'planned_kwh':summary['planned_kwh'],
            'emergency_kwh':summary['emergency_kwh'],
            'emergency_days':summary['emergency_days'],
            'emergency_slots_10min':summary['emergency_slots_10min'],
            'terminal_soc':summary['terminal_soc'],
            'saving_yuan_vs_main':base['total_cost']-summary['total_cost'],
            'saving_pct_vs_main':100*(base['total_cost']-summary['total_cost'])/base['total_cost'],
        })
    csv_write(root/'risk_buffer_sweep_summary.csv',rows)
    best=min(rows,key=lambda r:r['total_cost'])
    (root/'risk_buffer_sweep_summary.json').write_text(
        json.dumps({'baseline_main':base,'experiments':rows,'best_by_total_cost':best},
                   ensure_ascii=False,indent=2),encoding='utf-8'
    )
    print('risk-buffer sweep完成。最低总费用：',json.dumps(best,ensure_ascii=False),flush=True)
    print(f'汇总：{root/"risk_buffer_sweep_summary.csv"}',flush=True)
    return rows



def self_test(data,args,lam):
    # Independent balance audit and explicit future-mutation invariance.
    day=20; pol=Policy('week_mean',3,600.,.1)
    lh,vh=forecast(data.load[:day],data.pv[:day],pol)
    mutated_l=data.load.copy();mutated_v=data.pv.copy()
    mutated_l[day:]*=3;mutated_v[day:]=0
    lh2,vh2=forecast(mutated_l[:day],mutated_v[:day],pol)
    assert np.array_equal(lh,lh2) and np.array_equal(vh,vh2)
    q=optimize(lh-vh,data.price,EINITIAL,pol,lam,time_limit=args.time_limit)[0]
    q2=optimize(lh2-vh2,data.price,EINITIAL,pol,lam,time_limit=args.time_limit)[0]
    assert np.allclose(q,q2,atol=1e-5)
    act=current_action(lh,vh,data.load[day,0],data.pv[day,0],q,0,EINITIAL,pol,data.price,lam,args.time_limit)
    # Future actuals cannot enter current_action's signature.
    mutated_l[day,0]=data.load[day,0];mutated_v[day,0]=data.pv[day,0]
    act2=current_action(lh,vh,mutated_l[day,0],mutated_v[day,0],q,0,EINITIAL,pol,data.price,lam,args.time_limit)
    assert np.allclose(act,act2,atol=1e-5)
    # Low inventory with reserve must remain feasible (soft, not hard reserve).
    optimize(lh-vh,data.price,EMIN,pol,lam,time_limit=args.time_limit)
    print('Self-test passed: future invariance and low-SOC soft-reserve feasibility.',flush=True)




def export_daily_soc(source, output):
    """SOC极值含日初和144个区间末端；循环指标不是雨流计数。"""
    source,output=Path(source),Path(output)
    rows=read_csv_rows(first_existing(source,['dispatch_detail_physical.csv','dispatch_detail.csv']))
    groups=defaultdict(list)
    for r in rows:
        if r['date']>='2025-02-01': groups[r['date']].append(r)
    if not groups: raise ValueError('没有正式评价期SOC明细')
    out=[]
    for day,rr in sorted(groups.items()):
        rr.sort(key=lambda r:int(r['slot_index']))
        if [int(r['slot_index']) for r in rr]!=list(range(T)):
            raise ValueError(f'{day}时段不完整或重复')
        start=np.array([float(r['soc_start']) for r in rr])
        end=np.array([float(r['soc_end']) for r in rr])
        c=np.array([float(r['charge_kwh']) for r in rr]);d=np.array([float(r['discharge_kwh']) for r in rr])
        if not np.isfinite(np.r_[start,end,c,d]).all(): raise ValueError('非有限SOC数据')
        if np.max(abs(start[1:]-end[:-1]))>TOL or np.max(abs(end-start-ETA*c+d/ETA))>TOL:
            raise ValueError(f'{day}SOC连续性或递推不成立')
        if min(c.min(),d.min()) < -TOL: raise ValueError('充放电量不能为负')
        trajectory=np.r_[start[0],end]
        low,high=float(trajectory.min()),float(trajectory.max())
        if low<EMIN-TOL or high>EMAX+TOL: raise ValueError('SOC越界')
        drawdown=float(np.max(np.maximum.accumulate(trajectory)-trajectory))
        throughput=float(ETA*c.sum()+d.sum()/ETA)
        out.append({'date':day,'soc_start_kwh':float(start[0]),'soc_end_kwh':float(end[-1]),
                    'soc_min_kwh':low,'soc_max_kwh':high,'soc_range_kwh':high-low,
                    'soc_range_pct_nominal':100*(high-low)/12000.,
                    'max_discharge_depth_kwh':drawdown,
                    'max_discharge_depth_pct_nominal':100*drawdown/12000.,
                    'max_discharge_depth_pct_usable':100*drawdown/(EMAX-EMIN),
                    'internal_throughput_kwh':throughput,
                    'equivalent_full_cycles_nominal':throughput/(2*12000.),
                    'equivalent_full_cycles_usable':throughput/(2*(EMAX-EMIN))})
    output.mkdir(parents=True,exist_ok=True)
    csv_write(output/'soc_daily_metrics.csv',out)
    (output/'soc_daily_metrics_notes.txt').write_text(
        '每日SOC极值覆盖0:00初值及144个区间末值。\n'
        'soc_range是日内极差，不等于循环深度。\n'
        'max_discharge_depth是日内先出现的峰值到后续谷值的最大下降量；不跨日追踪循环，非雨流循环计数。\n'
        '深度百分比分别按额定容量12000和可用窗口9600归一化。\n'
        '等效循环次数=储能内部充放电吞吐量/(2*容量)，吞吐量=0.9*充电量+放电量/0.9。\n'
        '日内不闭合循环也按吞吐量计入，不代表完成的物理循环个数。\n',encoding='utf-8')
    return out


def run_lambda_sensitivity(args):
    """三组独立全年主模型；保留原lambda与备用惩罚的耦合关系。"""
    root=args.output_dir or Path(__file__).resolve().parent/'q2_lambda_sensitivity'
    root=Path(root);root.mkdir(parents=True,exist_ok=True)
    factors=[0.5,1.,2.]
    if args.no_tune or args.self_test or args.lambda_factor!=1.:
        raise ValueError('lambda-sensitivity固定运行0.5/1/2倍，请勿组合no-tune、self-test或lambda-factor')
    if args.end_date!='2025-12-31': raise ValueError('lambda敏感性要求完整全年；短期验证请用main --lambda-factor')
    for factor in factors:
        out=root/f'lambda_{factor:g}'
        if not args.lambda_summary_only:
            if out.exists() and any(out.iterdir()):
                raise FileExistsError(f'{out}非空；请换输出目录，或使用--lambda-summary-only汇总已有结果')
            command=[sys.executable,str(Path(__file__).resolve()),'--mode','main',
                     '--data-dir',str(args.data_dir.resolve()),'--output-dir',str(out.resolve()),
                     '--lambda-factor',str(factor),'--validation-days',str(args.validation_days),
                     '--tune-every',str(args.tune_every),'--time-limit',str(args.time_limit)]
            print(f'开始lambda={factor:g}倍全年独立实验',flush=True)
            subprocess.run(command,check=True)
        report=json.loads((out/'run_summary.json').read_text(encoding='utf-8'))
        if not report.get('complete_year') or report['evaluation_days']!=334:
            raise ValueError(f'{out}尚未完整完成')
        if float(report['arguments'].get('lambda_factor',1))!=factor:
            raise ValueError('结果中的lambda倍数与实验目录不符')
    summaries=[]
    for factor in factors:
        out=root/f'lambda_{factor:g}'
        r=json.loads((out/'run_summary.json').read_text(encoding='utf-8'))
        metrics=export_daily_soc(out,out)
        summaries.append({'lambda_factor':factor,'lambda':r['lambda'],
                          **{k:r[k] for k in ['total_cost','planned_cost','emergency_cost','emergency_kwh',
                                              'emergency_days','emergency_slots_10min','terminal_soc']},
                          'mean_daily_max_discharge_depth_pct_nominal':float(np.mean([x['max_discharge_depth_pct_nominal'] for x in metrics])),
                          'equivalent_full_cycles_nominal_total':sum(x['equivalent_full_cycles_nominal'] for x in metrics)})
    base=summaries[1]['total_cost']
    for r in summaries: r['cost_change_vs_1x_pct']=100*(r['total_cost']-base)/base if base else None
    csv_write(root/'lambda_sensitivity_summary.csv',summaries)
    (root/'lambda_sensitivity_notes.txt').write_text(
        '三组从1月1日起独立运行，2月1日至12月31日评价。\n'
        'lambda0=min(price)/ETA。按原目标，倍数同时缩放终端缺口惩罚、软备用惩罚及验证期终端修正。\n'
        '这属于耦合风险权重敏感性，不是只改变终端项而固定备用惩罚。实际费用不包含这些软惩罚。\n'
        '主模型候选仍由各组历史验证选择，策略及SOC路径可能不同。\n',encoding='utf-8')
    plt,_=_load_matplotlib(args.no_plots)
    if plt is not None:
        fig,axes=plt.subplots(1,2,figsize=(11,4))
        for ax,key,label,color in zip(axes,['total_cost','emergency_kwh'],['总费用（元）','紧急购电量（kWh）'],['#104E8B','#B22222']):
            ax.plot(factors,[r[key] for r in summaries],marker='o',color=color);ax.set_xticks(factors)
            ax.set_xlabel('lambda倍数');ax.set_ylabel(label);ax.grid(color='#D7E1EB',alpha=.8)
        fig.tight_layout()
        fig.savefig(root/'Fig_lambda敏感性.svg',format='svg');plt.close(fig)
    print(f'lambda敏感性完成：{root}',flush=True)


# ==================== Q2论文补充图：只读取结果，不求解MILP ====================

PAPER_COLORS = dict(
    blue='#104E8B', medium='#376B9E', light='#5F89B1', pale='#AFC3D8',
    mint='#C5E9E3', grid='#D7E1EB', pink='#F2DADA', rose='#E5B5B5',
    salmon='#D89090', red='#B22222'
)


def _paper_save(fig, output, name, plt):
    fig.savefig(output / (name + '.svg'), format='svg', bbox_inches='tight', facecolor='white')
    plt.close(fig)


def _paper_detail(main_output):
    candidates = [
        main_output / 'dispatch_detail_physical.csv',
        main_output / 'dispatch_detail.csv',
        main_output.parent / 'q2_output_timefixed' / 'dispatch_detail_physical.csv'
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError('缺少主模型逐时明细；已检查主目录及同级q2_output_timefixed')


def _paper_style_axes(ax, c, grid_axis='y'):
    ax.set_facecolor('white')
    if grid_axis in {'x', 'both'}:
        ax.grid(axis='x', color=c['grid'], alpha=.70, linewidth=.9)
    if grid_axis in {'y', 'both'}:
        ax.grid(axis='y', color=c['grid'], alpha=.70, linewidth=.9)
    for side in ['top', 'right']:
        ax.spines[side].set_visible(False)
    ax.spines['left'].set_color(c['pale'])
    ax.spines['bottom'].set_color(c['pale'])


def _paper_add_panel_label(ax, label, c):
    ax.text(0.01, 0.98, label, transform=ax.transAxes,
            ha='left', va='top', fontsize=11, fontweight='bold', color=c['blue'])


def _paper_daily_load_weekday_stats(data):
    rows = []
    for ds, load_day, pv_day in zip(data.dates, data.load, data.pv):
        if ds < date(2025, 2, 1):
            continue
        rows.append({
            'date': str(ds),
            'weekday': ds.weekday(),
            'weekday_name': ['周一', '周二', '周三', '周四', '周五', '周六', '周日'][ds.weekday()],
            'load_mwh': float(load_day.sum() / 1000),
            'pv_mwh': float(pv_day.sum() / 1000),
        })
    return rows


def plot_r_results(args, plt, output):
    records = []
    skipped = []
    sources = [args.no_reserve_output] + sorted(args.r_results.glob('R_*'))
    reference = None
    main_report = args.main_output / 'run_summary.json'
    if main_report.exists():
        m = json.loads(main_report.read_text(encoding='utf-8'))
        reference = (
            m.get('lambda'),
            m.get('arguments', {}).get('validation_days', 7),
            m.get('arguments', {}).get('tune_every', 28)
        )
    for source in sources:
        try:
            report = json.loads((source / 'run_summary.json').read_text(encoding='utf-8'))
            if not report.get('complete_year'):
                raise ValueError('没有全年完成标记')
            summary, _, ev = aggregate_daily_output(source)
            expected = [str(date(2025, 2, 1) + timedelta(days=i)) for i in range(334)]
            if [r['date'] for r in ev] != expected:
                raise ValueError('评价日期不完整')
            reserves = {float(r['reserve']) for r in ev}
            mus = {float(r['mu_factor']) for r in ev}
            if len(reserves) != 1:
                raise ValueError('不是固定R实验')
            reserve = next(iter(reserves))
            expected_mu = 0. if reserve == 0 else .1
            if mus != {expected_mu}:
                raise ValueError('mu口径不一致')
            cfg = (
                report.get('lambda'),
                report.get('arguments', {}).get('validation_days', 7),
                report.get('arguments', {}).get('tune_every', 28)
            )
            if reference is None:
                reference = cfg
            if cfg != reference:
                raise ValueError('lambda或历史验证设置不一致')
            if any(r['reserve_kwh'] == reserve for r in records):
                raise ValueError('重复R点')
            records.append({
                'reserve_kwh': reserve,
                **{k: summary[k] for k in ['total_cost', 'emergency_kwh', 'emergency_days']},
                'source': str(source)
            })
        except (OSError, ValueError, KeyError, AssertionError) as exc:
            skipped.append({'source': str(source), 'reason': str(exc)})
    records.sort(key=lambda r: r['reserve_kwh'])
    if records:
        csv_write(output / 'R_sensitivity_verified_points.csv', records)
    status = {
        'valid_points': records,
        'skipped': skipped,
        'generated': False,
        'note': '只接收固定R全年记录；不预设U型。不自动把动态主模型当成固定R=600。旧记录无输入哈希时仍需保证附件来源一致。'
    }
    if len(records) >= 3 and plt is not None:
        c = PAPER_COLORS
        fig, axes = plt.subplots(1, 2, figsize=(11.5, 4.2), layout='constrained')
        configs = [
            ('total_cost', '评价期总费用（万元）', 10000, c['blue']),
            ('emergency_kwh', '紧急购电量（MWh）', 1000, c['red'])
        ]
        for ax, (key, ylabel, divisor, color) in zip(axes, configs):
            xs = [r['reserve_kwh'] for r in records]
            ys = [r[key] / divisor for r in records]
            ax.plot(xs, ys, marker='o', color=color, linewidth=2.2)
            ax.fill_between(xs, ys, color=color, alpha=.10)
            _paper_style_axes(ax, c)
            ax.set(xlabel='软备用目标 R（kWh）', ylabel=ylabel)
            _paper_add_panel_label(ax, '(%s)' % ('a' if key == 'total_cost' else 'b'), c)
        fig.suptitle('固定软备用水平的全年敏感性分析', fontsize=15, color=c['blue'])
        _paper_save(fig, output, 'Fig_R敏感性', plt)
        status['generated'] = True
    elif len(records) < 3:
        print(f'[待补实验] 固定R全年有效点仅{len(records)}个，至少3点后生成R敏感性图。', flush=True)
        (output / 'Fig_R敏感性.svg').unlink(missing_ok=True)
    (output / 'R_sensitivity_plot_status.json').write_text(
        json.dumps(status, ensure_ascii=False, indent=2), encoding='utf-8'
    )


def paper_figures(args):
    output = args.diagnostic_output
    output.mkdir(parents=True, exist_ok=True)
    plt, _ = _load_matplotlib(args.no_plots)
    if not args.no_plots and plt is None:
        raise RuntimeError('补充绘图需要matplotlib')
    c = PAPER_COLORS

    # ---------- 读入并核验主结果 ----------
    _, daily, ev = aggregate_daily_output(args.main_output)
    if len(ev) != 334:
        raise ValueError('论文图要求完整334天评价结果')
    detail = read_csv_rows(_paper_detail(args.main_output))
    groups = defaultdict(list)
    for r in detail:
        if r['date'] >= '2025-02-01':
            groups[r['date']].append(r)
    if set(groups) != {r['date'] for r in ev}:
        raise ValueError('主模型每日汇总与明细日期不一致')
    for ds, rr in groups.items():
        rr.sort(key=lambda r: int(r['slot_index']))
        if [int(r['slot_index']) for r in rr] != list(range(T)):
            raise ValueError('时段缺失或重复')
        d = next(r for r in ev if r['date'] == ds)
        if abs(sum(float(r['emergency_kwh']) for r in rr) - float(d['emergency_kwh'])) > TOL:
            raise ValueError('主模型明细与每日汇总不匹配')

    data = read_data(args.data_dir)
    eval_mask = np.array([ds >= date(2025, 2, 1) for ds in data.dates])
    eval_dates = [str(ds) for ds in data.dates if ds >= date(2025, 2, 1)]
    load_eval = data.load[eval_mask] / DT
    pv_eval = data.pv[eval_mask] / DT

    # ---------- 月度费用表 / 预测诊断表 ----------
    monthly = []
    for m in sorted({r['date'][:7] for r in ev}):
        rr = [r for r in ev if r['date'].startswith(m)]
        monthly.append({
            'month': m,
            'planned_cost': sum(float(r['planned_cost']) for r in rr),
            'emergency_cost': sum(float(r['emergency_cost']) for r in rr),
            'total_cost': sum(float(r['total_cost']) for r in rr),
            'emergency_kwh': sum(float(r['emergency_kwh']) for r in rr),
        })
    csv_write(output / 'paper_monthly_cost.csv', monthly)

    policies = {r['date']: r for r in daily}
    forecasts = {}
    errors = [[], []]
    stats = []
    forecast_daily_rows = []
    for j, ds in enumerate(data.dates):
        ds_str = str(ds)
        if ds_str not in groups:
            continue
        for key, expected_values in [('load_kwh', data.load[j]), ('pv_kwh', data.pv[j])]:
            if not np.allclose([float(x[key]) for x in groups[ds_str]], expected_values, rtol=0, atol=1e-8):
                raise ValueError('预测附件与既有调度输入不一致')
        r = policies[ds_str]
        pol = Policy(r['load_mode'], int(float(r['pv_days'])), float(r['reserve']), float(r['mu_factor']))
        lh, vh = forecast(data.load[:j], data.pv[:j], pol)
        actual_l = data.load[j]
        actual_v = data.pv[j]
        forecasts[ds_str] = (lh / DT, vh / DT, actual_l / DT, actual_v / DT)
        load_err = (lh - actual_l) / DT
        pv_err = (vh - actual_v) / DT
        errors[0].extend(load_err)
        errors[1].extend(pv_err)
        forecast_daily_rows.append({
            'date': ds_str,
            'load_mae_kw': float(np.abs(load_err).mean()),
            'load_rmse_kw': float(np.sqrt(np.mean(load_err ** 2))),
            'load_bias_kw': float(load_err.mean()),
            'pv_mae_kw': float(np.abs(pv_err).mean()),
            'pv_rmse_kw': float(np.sqrt(np.mean(pv_err ** 2))),
            'pv_bias_kw': float(pv_err.mean()),
        })
    for label, err in zip(['load', 'pv'], errors):
        err = np.array(err)
        stats.append({
            'variable': label,
            'mae_kw': float(np.abs(err).mean()),
            'rmse_kw': float(np.sqrt(np.mean(err ** 2))),
            'bias_kw': float(err.mean())
        })
    csv_write(output / 'paper_forecast_metrics.csv', stats)
    csv_write(output / 'paper_forecast_daily_metrics.csv', forecast_daily_rows)

    # ---------- 衍生每日统计 ----------
    weekday_rows = _paper_daily_load_weekday_stats(data)
    csv_write(output / 'paper_weekday_load_distribution.csv', weekday_rows)

    daily_stats = []
    special_dates = ['2025-03-20', '2025-06-21', '2025-09-23', '2025-12-21']
    special_set = set(special_dates)
    for r in ev:
        ds = r['date']
        j = next(idx for idx, d in enumerate(data.dates) if str(d) == ds)
        load_mwh = float(data.load[j].sum() / 1000)
        pv_mwh = float(data.pv[j].sum() / 1000)
        emergency_kwh = float(r['emergency_kwh'])
        total_cost = float(r['total_cost'])
        emergency_cost = float(r['emergency_cost'])
        daily_stats.append({
            'date': ds,
            'load_mwh': load_mwh,
            'pv_mwh': pv_mwh,
            'planned_cost': float(r['planned_cost']),
            'emergency_cost': emergency_cost,
            'total_cost': total_cost,
            'emergency_kwh': emergency_kwh,
            'emergency_mwh': emergency_kwh / 1000,
            'emergency_cost_share_pct': 100 * emergency_cost / total_cost if total_cost > 0 else 0.,
            'weekday_name': ['周一', '周二', '周三', '周四', '周五', '周六', '周日'][date.fromisoformat(ds).weekday()],
            'is_special_date': int(ds in special_set),
        })
    csv_write(output / 'paper_daily_summary_enhanced.csv', daily_stats)

    # ---------- 绘图 ----------
    if plt is not None:
        from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
        from matplotlib.gridspec import GridSpec

        # 1) 信息边界流程图（保留）
        fig, ax = plt.subplots(figsize=(12.5, 7.2))
        ax.set(xlim=(0, 12), ylim=(0, 7))
        ax.axis('off')
        boxes = [
            (0.4, 4.9, 3.1, 1.05, '历史负荷与光伏\n当日初始SOC、固定电价', c['grid']),
            (4.2, 4.9, 3.1, 1.05, '每日0:00历史预测\n未来真实值不可见', c['mint']),
            (8.0, 4.9, 3.4, 1.05, '日前MILP\n锁定144时段计划购电 q', c['pale']),
            (8.0, 2.2, 3.4, 1.35, '每10分钟滚动修正\n计划 q 不变；未来仍用0点预测\n仅执行当前充放电与紧急购电', c['pink']),
            (4.2, 2.2, 3.1, 1.35, '观察当前时段\n实际负荷与光伏\n不读取未来实际值', c['rose']),
            (0.4, 2.2, 3.1, 1.35, '更新SOC并跨日传递\n实际费用 = 计划费用 + 紧急购电费用', c['grid'])
        ]
        for x, y, w, h, label, color in boxes:
            ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.08',
                                        facecolor=color, edgecolor=c['blue'], linewidth=1.3))
            ax.text(x + w / 2, y + h / 2, label, ha='center', va='center', fontsize=10.5, color=c['blue'])
        for start, end in [
            ((3.58, 5.42), (4.10, 5.42)),
            ((7.38, 5.42), (7.92, 5.42)),
            ((9.70, 4.84), (9.70, 3.70)),
            ((7.38, 2.88), (7.92, 2.88)),
        ]:
            ax.add_patch(FancyArrowPatch(start, end, arrowstyle='-|>', mutation_scale=16, color=c['blue'], linewidth=1.4))
        ax.plot([9.7, 9.7, 1.95], [2.15, 1.58, 1.58], color=c['medium'], linewidth=1.3)
        ax.annotate('', xy=(1.95, 2.15), xytext=(1.95, 1.58), arrowprops=dict(arrowstyle='->', color=c['medium']))
        ax.annotate('', xy=(1.95, 4.82), xytext=(1.95, 3.55), arrowprops=dict(arrowstyle='->', color=c['blue'], linestyle='--'))
        ax.text(6.0, 0.72, '因果边界：日前只用历史；实时只揭示当前观测；储能全年连续运行，不重置SOC。',
                ha='center', fontsize=10.5, color=c['blue'])
        ax.set_title('问题二的信息边界与日前—实时滚动调度闭环', fontsize=15.5, color=c['blue'])
        _paper_save(fig, output, 'Fig_信息边界与调度流程', plt)

        # 2) 负荷与光伏的全年典型曲线 + 四个指定日期曲线
        fig = plt.figure(figsize=(13.5, 11.0), layout='constrained')
        gs = GridSpec(3, 2, figure=fig, height_ratios=[1.35, 1.0, 1.0])
        ax_top = fig.add_subplot(gs[0, :])
        hours = np.arange(T) * DT
        load_mean = load_eval.mean(axis=0)
        load_q10 = np.quantile(load_eval, .10, axis=0)
        load_q90 = np.quantile(load_eval, .90, axis=0)
        pv_mean = pv_eval.mean(axis=0)
        pv_q10 = np.quantile(pv_eval, .10, axis=0)
        pv_q90 = np.quantile(pv_eval, .90, axis=0)
        ax_top.plot(hours, load_mean, color=c['blue'], linewidth=2.5, label='小区负载')
        ax_top.fill_between(hours, load_q10, load_q90, color=c['pale'], alpha=.45, label='负载10–90%范围')
        ax_top.plot(hours, pv_mean, color=c['medium'], linewidth=2.5, label='光伏发电')
        ax_top.fill_between(hours, pv_q10, pv_q90, color=c['mint'], alpha=.42, label='光伏10–90%范围')
        _paper_style_axes(ax_top, c, grid_axis='both')
        ax_top.set(xlim=(0, 24), xlabel='时刻 h', ylabel='功率 kW', title='2025年2月1日—12月31日平均时刻曲线')
        # 只保留一组完整图例，并放在上方主图外左下角。
        ax_top.legend(loc='upper left', bbox_to_anchor=(0.0, -0.13), ncol=4, frameon=False,
                      borderaxespad=0.0, handlelength=2.2, columnspacing=1.4)
        _paper_add_panel_label(ax_top, '(a)', c)

        panel_axes = [fig.add_subplot(gs[1, 0]), fig.add_subplot(gs[1, 1]), fig.add_subplot(gs[2, 0]), fig.add_subplot(gs[2, 1])]
        for idx, (ax, ds) in enumerate(zip(panel_axes, special_dates), start=1):
            rr = groups[ds]
            load_kw = np.array([float(r['load_kwh']) / DT for r in rr])
            pv_kw = np.array([float(r['pv_kwh']) / DT for r in rr])
            ax.plot(hours, load_kw, color=c['blue'], linewidth=2.0, label='小区负载')
            ax.plot(hours, pv_kw, color=c['medium'], linewidth=2.0, label='光伏发电')
            _paper_style_axes(ax, c, grid_axis='both')
            ax.set(xlim=(0, 24), xlabel='时刻 h', ylabel='功率 kW', title=ds)
            _paper_add_panel_label(ax, f'({chr(96 + idx + 1)})', c)
        # 下方四个指定日期与主图使用相同的“负荷/光伏”线型，故不再重复设置全局图例。
        _paper_save(fig, output, 'Fig_负荷光伏全年与指定日期特征', plt)

        # 3) 周内负荷分布箱线图
        fig, ax = plt.subplots(figsize=(10.2, 4.8), layout='constrained')
        weekday_groups = [[] for _ in range(7)]
        for r in weekday_rows:
            weekday_groups[r['weekday']].append(r['load_mwh'])
        bp = ax.boxplot(
            weekday_groups,
            tick_labels=['周一', '周二', '周三', '周四', '周五', '周六', '周日'],
            patch_artist=True,
            widths=.56,
            medianprops=dict(color=c['red'], linewidth=1.6),
            boxprops=dict(edgecolor=c['medium'], linewidth=1.2),
            whiskerprops=dict(color=c['medium'], linewidth=1.2),
            capprops=dict(color=c['medium'], linewidth=1.2)
        )
        fill_colors = [c['pale'], c['pale'], c['pale'], c['pale'], c['mint'], c['mint'], c['pale']]
        for patch, color in zip(bp['boxes'], fill_colors):
            patch.set_facecolor(color)
            patch.set_alpha(.88)
        _paper_style_axes(ax, c)
        ax.set_ylabel('日负载电量（MWh）')
        ax.set_title('小区日负载电量的周内分布')
        ax.text(0.99, -0.18, '说明：工作日整体负荷更高，周五与周六显著偏低。', transform=ax.transAxes,
                ha='right', va='top', fontsize=9.5, color=c['medium'])
        _paper_save(fig, output, 'Fig_周内负荷分布箱线图', plt)

        # 4) 全年经济性与紧急购电分析（参考用户给出的双层图）
        fig, axes = plt.subplots(2, 1, figsize=(13.2, 8.1), sharex=True, layout='constrained')
        xs = [datetime.strptime(r['date'], '%Y-%m-%d') for r in daily_stats]
        emergency_series = np.array([r['emergency_mwh'] for r in daily_stats])
        total_cost_series = np.array([r['total_cost'] for r in daily_stats])
        emergency_cost_series = np.array([r['emergency_cost'] for r in daily_stats])
        axes[0].plot(xs, emergency_series, color=c['salmon'], linewidth=1.6, label='紧急购电量')
        axes[0].fill_between(xs, emergency_series, color=c['pink'], alpha=.30)
        for ds in special_dates:
            idx = eval_dates.index(ds)
            axes[0].scatter(xs[idx], emergency_series[idx], color=c['blue'], s=26, zorder=4)
        _paper_style_axes(axes[0], c)
        axes[0].set_ylabel('紧急购电量（MWh）')
        axes[0].legend(loc='upper right', frameon=False)
        axes[0].set_title('评价期紧急购电与费用变化')
        _paper_add_panel_label(axes[0], '(a)', c)

        axes[1].plot(xs, total_cost_series, color=c['blue'], linewidth=1.8, label='总费用')
        axes[1].plot(xs, emergency_cost_series, color=c['salmon'], linewidth=1.4, label='紧急费用')
        _paper_style_axes(axes[1], c)
        axes[1].set_ylabel('费用（元/日）')
        axes[1].set_xlabel('日期')
        axes[1].legend(loc='upper right', frameon=False)
        _paper_add_panel_label(axes[1], '(b)', c)
        _paper_save(fig, output, 'Fig_全年经济性与紧急购电分析', plt)

        # 5) 费用分布与影响因素（参考用户给出的三联图）
        fig, axes = plt.subplots(1, 3, figsize=(14.8, 4.6), layout='constrained')
        total_cost_mw = total_cost_series / 10000
        emg_share = np.array([r['emergency_cost_share_pct'] for r in daily_stats])
        load_mwh = np.array([r['load_mwh'] for r in daily_stats])
        pv_mwh = np.array([r['pv_mwh'] for r in daily_stats])
        emg_mwh = np.array([r['emergency_mwh'] for r in daily_stats])

        axes[0].hist(total_cost_mw, bins=24, color=c['pale'], edgecolor='white')
        axes[0].axvline(total_cost_mw.mean(), color=c['red'], linestyle='--', linewidth=1.6,
                        label=f'均值 {total_cost_mw.mean():.2f} 万元')
        _paper_style_axes(axes[0], c)
        axes[0].set_title('(a) 日总费用分布')
        axes[0].set_xlabel('日总费用（万元）')
        axes[0].set_ylabel('天数')
        axes[0].legend(frameon=False, fontsize=9)

        from matplotlib.colors import ListedColormap
        _paper_blue_cmap = ListedColormap([c['grid'], c['mint'], c['pale'], c['light'], c['medium'], c['blue']])
        _paper_red_cmap = ListedColormap([c['grid'], c['pink'], c['rose'], c['salmon'], c['red']])
        sc = axes[1].scatter(load_mwh, emg_share, c=pv_mwh, cmap=_paper_blue_cmap, s=20, alpha=.85)
        _paper_style_axes(axes[1], c)
        axes[1].set_title('(b) 负荷与紧急费用占比')
        axes[1].set_xlabel('日负载电量（MWh）')
        axes[1].set_ylabel('紧急费用占比（%）')
        cb1 = fig.colorbar(sc, ax=axes[1])
        cb1.set_label('日光伏电量（MWh）')

        sc2 = axes[2].scatter(pv_mwh, total_cost_mw, c=emg_mwh, cmap=_paper_red_cmap, s=20, alpha=.85)
        _paper_style_axes(axes[2], c)
        axes[2].set_title('(c) 光伏与日总费用')
        axes[2].set_xlabel('日光伏电量（MWh）')
        axes[2].set_ylabel('日总费用（万元）')
        cb2 = fig.colorbar(sc2, ax=axes[2])
        cb2.set_label('紧急购电量（MWh）')
        _paper_save(fig, output, 'Fig_费用分布与驱动因素', plt)

        # 6) 月度费用构成
        fig, ax = plt.subplots(figsize=(11.0, 4.8), layout='constrained')
        p = np.array([r['planned_cost'] for r in monthly]) / 10000
        e = np.array([r['emergency_cost'] for r in monthly]) / 10000
        xx = np.arange(len(monthly))
        ax.bar(xx, p, color=c['blue'], width=.66, label='计划购电费')
        ax.bar(xx, e, bottom=p, color=c['rose'], width=.66, label='紧急购电费')
        _paper_style_axes(ax, c)
        ax.set_xticks(xx, [r['month'][5:] + '月' for r in monthly])
        ax.set_ylabel('费用（万元）')
        ax.set_title('评价期月度费用构成（2025年2—12月）')
        # 图例移到对应图表外左下角，避免遮挡12月柱形。
        ax.legend(frameon=False, loc='upper left', bbox_to_anchor=(0.0, -0.12),
                  ncol=2, borderaxespad=0.0, handlelength=1.8, columnspacing=1.5)
        _paper_save(fig, output, 'Fig_月度费用构成', plt)

        # 7) 四个指定日期的完整调度图
        fig, axes = plt.subplots(2, 2, figsize=(13.4, 8.3), layout='constrained')
        first_right = None
        for i, (ax, ds) in enumerate(zip(axes.flat, special_dates)):
            rr = groups[ds]
            load_curve = [float(r['load_kwh']) / DT for r in rr] + [float(rr[-1]['load_kwh']) / DT]
            pv_curve = [float(r['pv_kwh']) / DT for r in rr] + [float(rr[-1]['pv_kwh']) / DT]
            plan_curve = [float(r['plan_kwh']) / DT for r in rr] + [float(rr[-1]['plan_kwh']) / DT]
            emg_curve = [float(r['emergency_kwh']) / DT for r in rr] + [float(rr[-1]['emergency_kwh']) / DT]
            soc_curve = [float(rr[0]['soc_start'])] + [float(r['soc_end']) for r in rr]
            x_step = np.arange(T + 1) * DT
            ax.step(x_step, load_curve, where='post', color=c['blue'], linewidth=1.8, label='负荷')
            ax.step(x_step, pv_curve, where='post', color=c['medium'], linewidth=1.8, label='光伏')
            ax.step(x_step, plan_curve, where='post', color=c['light'], linewidth=1.4, linestyle='--', label='计划购电')
            ax.step(x_step, emg_curve, where='post', color=c['red'], linewidth=1.2, label='紧急购电')
            ax2 = ax.twinx()
            ax2.plot(x_step, soc_curve, color=c['salmon'], linestyle=':', linewidth=1.7, label='SOC')
            ax2.set_ylabel('SOC（kWh）', color=c['salmon'])
            ax2.tick_params(axis='y', colors=c['salmon'])
            ax2.spines['top'].set_visible(False)
            if i == 0:
                first_right = ax2
            _paper_style_axes(ax, c, grid_axis='both')
            ax.set(xlim=(0, 24), xlabel='时刻 h', ylabel='功率 kW', title=ds)
        hh, ll = axes.flat[0].get_legend_handles_labels()
        h2, l2 = first_right.get_legend_handles_labels() if first_right is not None else ([], [])
        fig.legend(hh + h2, ll + l2, loc='outside upper center', ncol=5, frameon=False)
        _paper_save(fig, output, 'Fig_四个指定日期调度', plt)

        # 8) 负荷/光伏 0点预测 vs 实际
        for typ, indices, actual_color, pred_color, ylabel in [
            ('负荷', (0, 2), c['blue'], c['pale'], '负荷功率（kW）'),
            ('光伏', (1, 3), c['medium'], c['mint'], '光伏功率（kW）')
        ]:
            fig, axes = plt.subplots(2, 2, figsize=(12.8, 7.5), layout='constrained')
            for ax, ds in zip(axes.flat, special_dates):
                f = forecasts[ds]
                ax.plot(hours, f[indices[1]], color=actual_color, linewidth=2.0, label='实际')
                ax.plot(hours, f[indices[0]], color=pred_color, linewidth=1.8, linestyle='--', label='0点预测')
                _paper_style_axes(ax, c, grid_axis='both')
                ax.set(title=ds, xlabel='时刻 h', ylabel=ylabel, xlim=(0, 24))
            handles, labels = axes.flat[0].get_legend_handles_labels()
            fig.legend(handles, labels, loc='outside upper center', ncol=2, frameon=False)
            _paper_save(fig, output, f'Fig_{typ}预测与实际', plt)

        # 9) 预测误差箱线图
        fig, ax = plt.subplots(figsize=(7.2, 4.8), layout='constrained')
        bp = ax.boxplot(errors, tick_labels=['负荷误差', '光伏误差'], showfliers=False, patch_artist=True)
        for patch, color in zip(bp['boxes'], [c['pale'], c['rose']]):
            patch.set_facecolor(color)
            patch.set_alpha(.88)
        for key in ['whiskers', 'caps', 'medians']:
            for line in bp[key]:
                line.set_color(c['blue'])
        _paper_style_axes(ax, c)
        ax.axhline(0, color=c['red'], linestyle='--', linewidth=1.2)
        ax.set_ylabel('预测值 − 实际值（kW）')
        ax.set_title('评价期逐时预测误差（箱线图隐藏离群点）')
        _paper_save(fig, output, 'Fig_预测误差箱线图', plt)


    # ---------- V2专属论文图：基于已完成全年结果与V2诊断，不重新求解MILP ----------
    if plt is not None:
        # A) V2 与 alpha=0.70 风险分位基准的全年对比
        reference_dir = args.v2_reference
        if reference_dir is not None and Path(reference_dir).exists():
            try:
                ref_summary, _, ref_ev = aggregate_daily_output(reference_dir)
                v2_summary, _, v2_ev = aggregate_daily_output(args.main_output)
                if [r['date'] for r in ref_ev] != [r['date'] for r in v2_ev]:
                    raise ValueError('V2与alpha=0.70参考结果评价日期不一致')

                metrics = [
                    ('总费用', ref_summary['total_cost'] / 10000, v2_summary['total_cost'] / 10000, '万元'),
                    ('紧急购电费', ref_summary['emergency_cost'] / 10000, v2_summary['emergency_cost'] / 10000, '万元'),
                    ('紧急购电量', ref_summary['emergency_kwh'] / 1000, v2_summary['emergency_kwh'] / 1000, 'MWh'),
                ]
                fig, axes = plt.subplots(1, 3, figsize=(13.2, 4.4), layout='constrained')
                for ax, (name, ref_val, v2_val, unit) in zip(axes, metrics):
                    vals = [ref_val, v2_val]
                    bars = ax.bar([0, 1], vals, color=[c['pale'], c['blue']], width=.62)
                    _paper_style_axes(ax, c)
                    ax.set_xticks([0, 1], ['α=0.70基准', 'V2'])
                    ax.set_ylabel(unit)
                    ax.set_title(name)
                    ymax = max(vals)
                    for bar, val in zip(bars, vals):
                        ax.text(bar.get_x() + bar.get_width()/2, val + ymax*.025,
                                f'{val:.2f}', ha='center', va='bottom', fontsize=9.5, color=c['blue'])
                    saving = ref_val - v2_val
                    if ref_val > 0:
                        ax.text(.5, .93, f'改善 {100*saving/ref_val:.2f}%', transform=ax.transAxes,
                                ha='center', va='top', fontsize=10, color=c['red'])
                fig.suptitle('V2相对固定 α=0.70 基准的全年经济性改善', fontsize=15, color=c['blue'])
                _paper_save(fig, output, 'Fig_V2与alpha070基准对比', plt)

                # 月度总费用对比
                ref_by_m = defaultdict(float)
                v2_by_m = defaultdict(float)
                for r in ref_ev:
                    ref_by_m[r['date'][:7]] += float(r['total_cost'])
                for r in v2_ev:
                    v2_by_m[r['date'][:7]] += float(r['total_cost'])
                months_cmp = sorted(v2_by_m)
                xx = np.arange(len(months_cmp))
                fig, ax = plt.subplots(figsize=(11.4, 4.8), layout='constrained')
                ref_vals = np.array([ref_by_m[m] for m in months_cmp]) / 10000
                v2_vals = np.array([v2_by_m[m] for m in months_cmp]) / 10000
                width = .36
                ax.bar(xx-width/2, ref_vals, width=width, color=c['pale'], label='α=0.70基准')
                ax.bar(xx+width/2, v2_vals, width=width, color=c['blue'], label='V2')
                _paper_style_axes(ax, c)
                ax.set_xticks(xx, [m[5:]+'月' for m in months_cmp])
                ax.set_ylabel('月总费用（万元）')
                ax.set_title('V2与固定 α=0.70 基准的月度费用对比')
                ax.legend(frameon=False, ncol=2)
                _paper_save(fig, output, 'Fig_V2与alpha070月度费用对比', plt)
            except Exception as exc:
                print(f'[warning] V2 vs alpha=0.70 对比图未生成：{exc}', flush=True)
        else:
            print('[warning] 未提供或找不到 --v2-reference，跳过 V2 vs alpha=0.70 对比图。', flush=True)

        # B) V2预测修正效果：直接使用v2_daily_diagnostics或按V2公式重构
        diag_path = Path(args.main_output) / 'v2_daily_diagnostics.csv'
        if diag_path.exists():
            try:
                diag_rows = read_csv_rows(diag_path)
                if len(diag_rows) != 334:
                    raise ValueError(f'V2诊断应为334天，实际{len(diag_rows)}')
                base_mae = np.array([float(r['base_net_mae_kw']) for r in diag_rows])
                corr_mae = np.array([float(r['corrected_net_mae_kw']) for r in diag_rows])
                bias_abs = np.array([float(r['mean_abs_bias_correction_kw']) for r in diag_rows])
                buffer_mean = np.array([float(r['mean_risk_buffer_kw']) for r in diag_rows])
                emg_cost = np.array([float(r['emergency_cost']) for r in diag_rows]) / 10000
                dates_diag = [datetime.strptime(r['date'], '%Y-%m-%d') for r in diag_rows]

                # 月度均值
                month_keys = sorted({r['date'][:7] for r in diag_rows})
                base_m = []
                corr_m = []
                for m in month_keys:
                    rr = [r for r in diag_rows if r['date'].startswith(m)]
                    base_m.append(np.mean([float(r['base_net_mae_kw']) for r in rr]))
                    corr_m.append(np.mean([float(r['corrected_net_mae_kw']) for r in rr]))

                fig, axes = plt.subplots(1, 2, figsize=(13.2, 4.5), layout='constrained')
                axes[0].bar([0, 1], [base_mae.mean(), corr_mae.mean()],
                            color=[c['pale'], c['blue']], width=.62)
                _paper_style_axes(axes[0], c)
                axes[0].set_xticks([0, 1], ['基础净负荷预测', 'EWMA修正后'])
                axes[0].set_ylabel('日均净负荷 MAE（kW）')
                axes[0].set_title('(a) 全年平均预测误差')
                improve = 100 * (base_mae.mean() - corr_mae.mean()) / base_mae.mean()
                axes[0].text(.5, .93, f'MAE改善 {improve:.2f}%', transform=axes[0].transAxes,
                             ha='center', va='top', color=c['red'])

                xx = np.arange(len(month_keys))
                axes[1].plot(xx, base_m, marker='o', color=c['light'], linewidth=1.8, label='基础预测')
                axes[1].plot(xx, corr_m, marker='o', color=c['blue'], linewidth=2.0, label='EWMA修正后')
                _paper_style_axes(axes[1], c, grid_axis='both')
                axes[1].set_xticks(xx, [m[5:]+'月' for m in month_keys])
                axes[1].set_ylabel('净负荷 MAE（kW）')
                axes[1].set_title('(b) 月度预测误差')
                axes[1].legend(frameon=False)
                fig.suptitle('V2净负荷预测偏差修正效果', fontsize=15, color=c['blue'])
                _paper_save(fig, output, 'Fig_V2预测修正效果', plt)

                fig, ax1 = plt.subplots(figsize=(12.3, 5.0), layout='constrained')
                ax1.plot(dates_diag, bias_abs, color=c['medium'], linewidth=1.3, label='EWMA绝对偏差修正')
                ax1.plot(dates_diag, buffer_mean, color=c['red'], linewidth=1.25, alpha=.85, label='MAD风险缓冲')
                _paper_style_axes(ax1, c)
                ax1.set_ylabel('功率修正量（kW）')
                ax1.set_xlabel('日期')
                ax2 = ax1.twinx()
                ax2.plot(dates_diag, emg_cost, color=c['salmon'], linewidth=1.0, alpha=.60, label='紧急购电费')
                ax2.set_ylabel('紧急购电费（万元/日）', color=c['salmon'])
                ax2.tick_params(axis='y', colors=c['salmon'])
                h1, l1 = ax1.get_legend_handles_labels()
                h2, l2 = ax2.get_legend_handles_labels()
                ax1.legend(h1+h2, l1+l2, frameon=False, loc='upper left', ncol=3)
                ax1.set_title('V2自适应偏差修正、稳健风险缓冲与紧急购电')
                _paper_save(fig, output, 'Fig_V2自适应风险信号', plt)
            except Exception as exc:
                print(f'[warning] V2诊断图未生成：{exc}', flush=True)

        # C) 四个指定日期的净负荷：实际 / 基础预测 / EWMA修正 / 修正+buffer
        try:
            fig, axes = plt.subplots(2, 2, figsize=(13.4, 8.0), layout='constrained')
            h = np.arange(T) * DT
            date_to_idx = {str(ds): j for j, ds in enumerate(data.dates)}
            for ax, ds in zip(axes.flat, special_dates):
                j = date_to_idx[ds]
                r = policies[ds]
                pol = Policy(r['load_mode'], int(float(r['pv_days'])), float(r['reserve']), float(r['mu_factor']))
                base, corrected, buffer, _ = v2_net_profiles(data, j, pol, args)
                actual = data.load[j] - data.pv[j]
                ax.plot(h, actual/DT, color=c['blue'], linewidth=2.0, label='实际净负荷')
                ax.plot(h, base/DT, color=c['light'], linestyle='--', linewidth=1.5, label='基础预测')
                ax.plot(h, corrected/DT, color=c['blue'], linewidth=1.7, label='EWMA修正')
                ax.plot(h, (corrected+buffer)/DT, color=c['red'], linestyle='-.', linewidth=1.4, label='日前风险修正')
                _paper_style_axes(ax, c, grid_axis='both')
                ax.set(xlim=(0, 24), xlabel='时刻 h', ylabel='净负荷功率（kW）', title=ds)
            handles, labels = axes.flat[0].get_legend_handles_labels()
            fig.legend(handles, labels, loc='lower left', bbox_to_anchor=(0.055, -0.035),
                       ncol=4, frameon=False, borderaxespad=0.0)
            fig.suptitle('V2指定日期净负荷预测与风险修正', fontsize=15, color=c['blue'])
            _paper_save(fig, output, 'Fig_V2四日净负荷修正', plt)
        except Exception as exc:
            print(f'[warning] V2四日净负荷修正图未生成：{exc}', flush=True)

    plot_r_results(args, plt, output)
    (output / 'paper_figures_notes.txt').write_text(
        '所有论文图仅输出SVG，不再生成PNG或PDF。\n'
        '配色严格限定为指定科研蓝—青—粉—红色卡；连续/分级色标也仅由该色卡构造。\n'
        '新增图重点包括：全年典型负荷/光伏曲线、周内负荷分布、全年经济性与紧急购电分析、费用分布与驱动因素。\n'
        '预测仅按每日记录的策略、当日之前历史重构，不重新优化。\n'
        '图不能替代题目规定表1、表2、表3；流程图不构成独立的信息隔离证明。\n',
        encoding='utf-8'
    )
    print(f'论文补充图完成：{output}', flush=True)


def load_resume_checkpoint(data, checkpoint, start_date, args):
    """从既有risk-buffer/V2部分结果恢复到start_date前一天，并重建评价期逐时结果。"""
    checkpoint=Path(checkpoint)
    if not checkpoint.exists():
        raise FileNotFoundError(f'断点目录不存在：{checkpoint}')
    if args.mode not in {'risk-buffer','v2'}:
        raise ValueError('--resume-from目前仅支持--mode risk-buffer或v2')
    start_date=date.fromisoformat(str(start_date))
    prev_date=start_date-timedelta(days=1)
    if start_date<=date(2025,2,1):
        raise ValueError('断点续跑起点必须晚于2025-02-01')

    daily_path=first_existing(checkpoint,['daily_summary.csv'])
    daily_rows=read_csv_rows(daily_path)
    daily_map={r['date']:r for r in daily_rows}
    prev=daily_map.get(str(prev_date))
    if prev is None:
        raise AssertionError(f'断点daily_summary缺少前一日{prev_date}')

    # 核对断点参数，避免把不同实验口径的结果拼接在一起。
    summary_path=checkpoint/'run_summary.json'
    if summary_path.exists():
        old_summary=json.loads(summary_path.read_text(encoding='utf-8'))
        old_mode=old_summary.get('experiment_mode')
        if old_mode and old_mode!=args.mode:
            raise AssertionError(f'断点实验模式{old_mode}与当前{args.mode}不一致')
        old_args=old_summary.get('arguments',{})
        if args.mode=='risk-buffer':
            if 'risk_alpha' in old_args and abs(float(old_args['risk_alpha'])-float(args.risk_alpha))>1e-12:
                raise AssertionError('断点risk_alpha与当前命令不一致')
            if 'risk_window' in old_args and int(float(old_args['risk_window']))!=int(args.risk_window):
                raise AssertionError('断点risk_window与当前命令不一致')
        elif args.mode=='v2':
            pairs=[
                ('v2_residual_window',int(args.v2_residual_window),int),
                ('v2_bias_half_life',float(args.v2_bias_half_life),float),
                ('v2_bias_shrink',float(args.v2_bias_shrink),float),
                ('v2_risk_k',float(args.v2_risk_k),float),
                ('v2_tail_quantile',float(args.v2_tail_quantile),float),
            ]
            for key,current,kind in pairs:
                if key not in old_args: continue
                old=kind(old_args[key])
                if (kind is int and old!=current) or (kind is float and abs(old-current)>1e-12):
                    raise AssertionError(f'断点{key}与当前命令不一致')

    detail_path=first_existing(checkpoint,['dispatch_detail_physical.csv'])
    detail=read_csv_rows(detail_path)
    by_date=defaultdict(list)
    for r in detail:
        ds=date.fromisoformat(r['date'])
        if date(2025,2,1)<=ds<=prev_date:
            by_date[r['date']].append(r)

    results={}; policies={}
    date_to_index={str(ds):j for j,ds in enumerate(data.dates)}
    expected_dates=[]
    d=date(2025,2,1)
    while d<=prev_date:
        expected_dates.append(str(d)); d+=timedelta(days=1)
    for ds in expected_dates:
        rows=by_date.get(ds,[])
        if len(rows)!=T:
            raise AssertionError(f'断点{ds}逐时明细应为{T}行，实际{len(rows)}')
        rows.sort(key=lambda r:int(float(r['slot_index'])))
        x=np.zeros((T,7),dtype=float)
        for k,r in enumerate(rows):
            if int(float(r['slot_index']))!=k:
                raise AssertionError(f'断点{ds}时段索引不连续')
            x[k]=[
                float(r['plan_kwh']),float(r['emergency_kwh']),float(r['charge_kwh']),
                float(r['discharge_kwh']),float(r['unused_kwh']),float(r['soc_start']),float(r['soc_end'])
            ]
        j=date_to_index[ds]
        audit(x,data.load[j],data.pv[j])
        results[j]=x
        dr=daily_map.get(ds)
        if dr is None:
            raise AssertionError(f'断点daily_summary缺少{ds}')
        policies[j]=Policy(dr['load_mode'],int(float(dr['pv_days'])),
                           float(dr.get('reserve',0)),float(dr.get('mu_factor',0)))

    state=float(prev['soc_end'])
    pol=Policy(prev['load_mode'],int(float(prev['pv_days'])),
               float(prev.get('reserve',0)),float(prev.get('mu_factor',0)))
    tuning_log=[]
    pv=checkpoint/'parameter_validation.csv'
    if pv.exists():
        tuning_log=read_csv_rows(pv)
    print(f'断点恢复成功：{checkpoint}',flush=True)
    print(f'已复用 2025-02-01~{prev_date} 的正式结果；从 {start_date} 继续。',flush=True)
    print(f'续跑初始SOC={state:.6f} kWh；继承Policy={pol}',flush=True)
    return results,policies,state,pol,tuning_log,daily_rows

def main():
    global EXPERIMENT_MODE, FIXED_RESERVE, FIXED_MU_FACTOR

    root=Path(__file__).resolve().parent
    parser=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--mode',
                        choices=['main','no-reserve','r-sensitivity','risk-buffer','risk-buffer-sweep',
                                 'risk-screen','risk-screen-sweep','v2','v2-screen','v2-local-screen',
                                 'lambda-sensitivity','soc-daily','diagnostics','forecast-diagnostics',
                                 'forecast-benchmark','compare','postprocess','paper-figures'],
                        default='main',
                        help='main=正式模型；no-reserve=公平B0；diagnostics=只读CSV诊断；postprocess=全部后处理')
    parser.add_argument('--lambda-factor',type=float,default=1.0,help='lambda相对min(price)/ETA的倍数；默认1保持原模型')
    parser.add_argument('--lambda-summary-only',action='store_true',help='lambda-sensitivity只汇总已完成的三组，不重新求解')
    parser.add_argument('--data-dir',type=Path,default=root/'data')
    parser.add_argument('--output-dir',type=Path,default=None,
                        help='求解输出目录；main默认q2_output，no-reserve默认q2_output_no_reserve')
    parser.add_argument('--main-output',type=Path,default=root/'q2_output',
                        help='已有主模型结果目录，供diagnostics/compare/postprocess读取')
    parser.add_argument('--no-reserve-output',type=Path,default=root/'q2_output_no_reserve',
                        help='已有公平B0结果目录，供compare/postprocess读取')
    parser.add_argument('--diagnostic-output',type=Path,default=root/'q2_diagnostics',
                        help='诊断与消融图输出目录')
    parser.add_argument('--r-results',type=Path,default=root/'q2_r_sensitivity',help='已有固定R全年实验目录，供论文补图读取')
    parser.add_argument('--risk-results',type=Path,default=root/'q2_risk_buffer',
                        help='risk-buffer-sweep输出根目录')
    parser.add_argument('--risk-alpha',type=float,default=0.80,
                        help='risk-buffer历史净负荷低估残差分位数，默认0.80')
    parser.add_argument('--risk-window',type=int,default=28,
                        help='risk-buffer历史残差滚动窗口天数，默认28')
    parser.add_argument('--risk-summary-only',action='store_true',
                        help='risk-buffer-sweep只汇总已有四组结果，不重新运行')
    parser.add_argument('--screen-end',default='2025-03-31',
                        help='risk-screen快速筛选截止日期，默认2025-03-31')
    parser.add_argument('--local-start',default='2025-03-01',
                        help='v2-local-screen起始日期，默认2025-03-01')
    parser.add_argument('--risk-screen-results',type=Path,default=root/'q2_risk_screen',
                        help='risk-screen-sweep快速筛选输出根目录')
    parser.add_argument('--v2-residual-window',type=int,default=28,
                        help='V2 residual history window in days; default 28')
    parser.add_argument('--v2-bias-half-life',type=float,default=5.0,
                        help='V2 EWMA bias half-life in days; default 5')
    parser.add_argument('--v2-bias-shrink',type=float,default=0.75,
                        help='V2 bias shrinkage factor in [0,1]; default 0.75')
    parser.add_argument('--v2-risk-k',type=float,default=0.50,
                        help='V2 robust uncertainty multiplier; 0.50 is near the Gaussian 70th-percentile z-score')
    parser.add_argument('--v2-tail-quantile',type=float,default=0.90,
                        help='V2 positive-tail cap quantile; default 0.90')
    parser.add_argument('--v2-policy-source',type=Path,default=None,
                        help='v2-screen daily Policy source directory; default uses --main-output')
    parser.add_argument('--v2-reference',type=Path,default=None,
                        help='optional comparison output directory for V2; default uses --main-output')
    parser.add_argument('--reserve',type=float,default=None,
                        choices=[0,300,600,900,1200,1500],
                        help='r-sensitivity模式固定R/kWh')
    parser.add_argument('--end-date',default='2025-12-31')
    parser.add_argument('--resume-from',type=Path,default=None,
                        help='从既有risk-buffer或V2部分结果目录续跑')
    parser.add_argument('--start-date',default=None,
                        help='断点续跑起始日期YYYY-MM-DD；必须是断点最后日期的下一天')
    parser.add_argument('--no-tune',action='store_true',
                        help='跳过滚动策略选择并保持默认Policy；仅诊断用，不等于公平无备用')
    parser.add_argument('--validation-days',type=int,default=7)
    parser.add_argument('--tune-every',type=int,default=28)
    parser.add_argument('--time-limit',type=float,default=30.)
    parser.add_argument('--self-test',action='store_true',
                        help='运行因果性和可行性测试后退出')
    parser.add_argument('--verify-reference',action='store_true',
                        help='完整全年运行后与已经核验的主模型/B0结果数值对照')
    parser.add_argument('--no-plots',action='store_true',
                        help='后处理只输出CSV/JSON/TXT，不画图')
    args=parser.parse_args()

    if not math.isfinite(args.lambda_factor) or args.lambda_factor<=0:
        parser.error('lambda-factor必须为正有限数')
    if not math.isfinite(args.risk_alpha) or not (0.5 <= args.risk_alpha < 1.0):
        parser.error('risk-alpha必须满足0.5 <= alpha < 1')
    if args.risk_window < 1:
        parser.error('risk-window必须为正整数')
    if args.v2_residual_window < 2:
        parser.error('v2-residual-window必须>=2')
    if not math.isfinite(args.v2_bias_half_life) or args.v2_bias_half_life<=0:
        parser.error('v2-bias-half-life必须为正有限数')
    if not math.isfinite(args.v2_bias_shrink) or not (0<=args.v2_bias_shrink<=1):
        parser.error('v2-bias-shrink必须在[0,1]')
    if not math.isfinite(args.v2_risk_k) or args.v2_risk_k<0:
        parser.error('v2-risk-k必须为非负有限数')
    if not math.isfinite(args.v2_tail_quantile) or not (0.5<args.v2_tail_quantile<1):
        parser.error('v2-tail-quantile必须在(0.5,1)')
    if args.start_date is not None:
        try:
            _resume_start_check=date.fromisoformat(args.start_date)
        except ValueError:
            parser.error('start-date必须是YYYY-MM-DD')
        if args.resume_from is None:
            parser.error('--start-date必须与--resume-from一起使用')
    if args.resume_from is not None and args.mode not in {'risk-buffer','v2'}:
        parser.error('--resume-from目前仅支持--mode risk-buffer或v2')
    try:
        _screen_end_check=date.fromisoformat(args.screen_end)
    except ValueError:
        parser.error('screen-end必须是YYYY-MM-DD')
    try:
        _local_start_check=date.fromisoformat(args.local_start)
    except ValueError:
        parser.error('local-start必须是YYYY-MM-DD')
    if args.mode=='risk-screen':
        run_risk_screen(args)
        return
    if args.mode=='v2-screen':
        run_v2_screen(args)
        return
    if args.mode=='v2-local-screen':
        run_v2_local_screen(args)
        return
    if args.mode=='risk-screen-sweep':
        run_risk_screen_sweep(args)
        return
    if args.mode=='risk-buffer-sweep':
        run_risk_buffer_sweep(args)
        return
    if args.mode=='paper-figures':
        paper_figures(args)
        return
    if args.mode=='soc-daily':
        export_daily_soc(args.main_output,args.diagnostic_output)
        return
    if args.mode=='lambda-sensitivity':
        run_lambda_sensitivity(args)
        return

    # ----------------------- 纯后处理模式 -----------------------
    if args.mode == 'diagnostics':
        soc_emergency_diagnostics(args.main_output,args.diagnostic_output,args.no_plots)
        return
    if args.mode == 'forecast-diagnostics':
        recompute_forecast_diagnostics(args.data_dir,args.main_output,args.diagnostic_output)
        return
    if args.mode == 'forecast-benchmark':
        forecast_method_benchmark(args.data_dir,args.main_output,args.diagnostic_output)
        return
    if args.mode == 'compare':
        ablation_compare(args.main_output,args.no_reserve_output,args.diagnostic_output,args.no_plots)
        return
    if args.mode == 'postprocess':
        soc_emergency_diagnostics(args.main_output,args.diagnostic_output,args.no_plots)
        try:
            recompute_forecast_diagnostics(args.data_dir,args.main_output,args.diagnostic_output)
        except Exception as exc:
            print(f'[warning] 预测误差诊断未完成：{exc}',flush=True)
        try:
            ablation_compare(args.main_output,args.no_reserve_output,args.diagnostic_output,args.no_plots)
        except Exception as exc:
            print(f'[warning] 消融对比未完成：{exc}',flush=True)
        paper_figures(args)
        return

    # ----------------------- 求解模式 ---------------------------
    if args.validation_days<1 or args.tune_every<1:
        parser.error('Window/cadence must be positive')
    end=date.fromisoformat(args.end_date)
    if not date(2025,1,1)<=end<=date(2025,12,31):
        parser.error('end-date outside 2025')

    if args.mode == 'main':
        EXPERIMENT_MODE='main'
        if args.output_dir is None:
            args.output_dir=root/'q2_output'
    elif args.mode == 'no-reserve':
        if args.no_tune:
            parser.error('公平无备用必须保留历史滚动预测选择，请不要加 --no-tune')
        EXPERIMENT_MODE='no-reserve'
        if args.output_dir is None:
            args.output_dir=root/'q2_output_no_reserve'
        print('='*72,flush=True)
        print('Q2 公平无软备用B0：只取消R和mu，其他机制全部保持',flush=True)
        print('='*72,flush=True)
    elif args.mode == 'r-sensitivity':
        if args.reserve is None:
            parser.error('r-sensitivity必须提供 --reserve')
        if args.no_tune:
            parser.error('R敏感性应保留历史滚动预测选择，请不要加 --no-tune')
        EXPERIMENT_MODE='r-sensitivity'
        FIXED_RESERVE=float(args.reserve)
        FIXED_MU_FACTOR=0.0 if FIXED_RESERVE==0 else 0.10
        if args.output_dir is None:
            args.output_dir=root/'q2_r_sensitivity'/f'R_{int(FIXED_RESERVE):04d}'
        print('='*72,flush=True)
        print(f'Q2 固定R敏感性：R={FIXED_RESERVE:.0f} kWh, mu_factor={FIXED_MU_FACTOR:.2f}',flush=True)
        print('仅固定备用参数；预测方法仍按过去窗口滚动选择',flush=True)
        print('='*72,flush=True)
    elif args.mode == 'risk-buffer':
        if args.no_tune:
            parser.error('risk-buffer应保留历史滚动预测/备用选择，请不要加 --no-tune')
        EXPERIMENT_MODE='risk-buffer'
        if args.output_dir is None:
            args.output_dir=root/'q2_risk_buffer'/f'alpha_{int(round(args.risk_alpha*100)):03d}'
        print('='*72,flush=True)
        print(f'Q2 历史预测误差风险修正：alpha={args.risk_alpha:.2f}, '
              f'window={args.risk_window}天',flush=True)
        print('只修正0:00日前计划净负荷；实时未来预测仍使用原0点预测。',flush=True)
        print('='*72,flush=True)
    elif args.mode == 'v2':
        if args.no_tune:
            parser.error('V2正式模式必须保留滚动策略选择，请不要加 --no-tune')
        EXPERIMENT_MODE='v2'
        if args.output_dir is None:
            args.output_dir=root/'q2_v2'
        print('='*72,flush=True)
        print('Q2 V2：EWMA净负荷偏差校正 + MAD稳健风险缓冲 + 原价格感知MILP',flush=True)
        print(f'window={args.v2_residual_window}, half_life={args.v2_bias_half_life}, '
              f'bias_shrink={args.v2_bias_shrink}, risk_k={args.v2_risk_k}, '
              f'tail_q={args.v2_tail_quantile}',flush=True)
        print('V2全程仅使用目标日前历史信息；实时仅当前时段实际值可见。',flush=True)
        print('='*72,flush=True)
    else:
        parser.error(f'未知求解模式：{args.mode}')

    data=read_data(args.data_dir)
    lam=float(data.price.min()/ETA)*args.lambda_factor
    if args.self_test:
        self_test(data,args,lam)
        return

    args.output_dir.mkdir(parents=True,exist_ok=True)
    snapshots={};begin=time.perf_counter();resume_prefix_daily=[]
    if args.resume_from is not None:
        if args.start_date is None:
            daily_rows=read_csv_rows(first_existing(args.resume_from,['daily_summary.csv']))
            if not daily_rows: raise ValueError('断点daily_summary为空')
            last=date.fromisoformat(daily_rows[-1]['date'])
            start_date=last+timedelta(days=1)
            args.start_date=str(start_date)
        else:
            start_date=date.fromisoformat(args.start_date)
        if args.output_dir.resolve()==args.resume_from.resolve():
            parser.error('output-dir不能与resume-from相同，避免覆盖断点')
        results,policies,state,pol,tuning_log,resume_prefix_daily=load_resume_checkpoint(
            data,args.resume_from,start_date,args)
        start_index=(start_date-date(2025,1,1)).days
    else:
        results={};policies={};tuning_log=[]
        state=EINITIAL;pol=Policy();start_index=0

    # 正常从1月1日运行；断点模式则从start_index开始，但预测仍可访问此前真实历史数据。
    for j,ds in enumerate(data.dates):
        if ds>end:break
        if j<start_index: continue
        snapshots[j]=state
        if j>=31 and not args.no_tune and (j-31)%args.tune_every==0:
            print(f'{ds}: selecting policy using past data only',flush=True)
            pol=tune(data,j,snapshots,args,lam,tuning_log)
        if j==0:
            x=bootstrap(data.load[j],data.pv[j],state)
        else:
            lh,vh=forecast(data.load[:j],data.pv[:j],pol)
            if args.mode=='risk-buffer':
                buffer=historical_risk_buffer(data,j,pol,args.risk_alpha,args.risk_window)
                x=simulate_day_risk_buffer(data.load[j],data.pv[j],lh,vh,data.price,state,
                                           pol,lam,args.time_limit,buffer)
            elif args.mode=='v2':
                _,corrected,buffer,_=v2_net_profiles(data,j,pol,args)
                x=simulate_day_v2(data.load[j],data.pv[j],corrected+buffer,corrected,
                                  data.price,state,pol,lam,args.time_limit)
            else:
                x=simulate_day(data.load[j],data.pv[j],lh,vh,data.price,state,pol,lam,args.time_limit)
        if (j-1) in results:
            assert abs(x[0,5]-results[j-1][-1,6])<=TOL
        results[j]=x;policies[j]=pol;state=float(x[-1,6])
        cp,ce=bill(x,data.price)
        print(f'{ds}: plan={cp:.2f}, emergency={ce:.2f}, end_SOC={state:.2f}, '
              f'elapsed={time.perf_counter()-begin:.1f}s',flush=True)

    report=export(data,results,policies,args.output_dir,args.data_dir/'result2.xlsx',
                  args,lam,tuning_log)

    if args.resume_from is not None and resume_prefix_daily:
        jan=[r for r in resume_prefix_daily if r.get('date','')<'2025-02-01']
        current=read_csv_rows(args.output_dir/'daily_summary.csv')
        if jan:
            csv_write(args.output_dir/'daily_summary.csv',jan+current)

    if args.mode=='risk-buffer':
        export_risk_buffer_diagnostics(data,policies,results,args.output_dir,
                                       args.risk_alpha,args.risk_window)
        compare_risk_buffer_to_main(args.main_output,args.output_dir)

    if args.mode=='v2':
        compare_v2_to_reference(args.v2_reference or args.main_output,args.output_dir)

    if args.mode=='no-reserve':
        verify_no_reserve_output(args.output_dir)
    elif args.mode=='r-sensitivity':
        verify_r_sensitivity_output(args.output_dir,FIXED_RESERVE)

    if args.verify_reference and args.lambda_factor==1.0 and report.get('complete_year') and report.get('evaluation_days')==334:
        ref_mode = 'main' if args.mode=='main' else ('no-reserve' if args.mode=='no-reserve' else None)
        if ref_mode:
            verify_reference_summary(report,ref_mode,strict=True)


if __name__=='__main__':
    main()

