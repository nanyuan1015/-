#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多时点光伏预测信息价值与滚动调度

【主模型定位】
D 系列是第三问主模型：
- B1：只使用 0:00 官方光伏预报，不进行日内合同调整（无日内更新基准）
- D6 / D12 / D18：仅在对应时刻允许使用新预报并调整未来合同
- D6_12 / D6_18 / D12_18：两次更新
- D6_12_18：三次更新

在每个更新节点：
1) 已执行时段锁定；
2) 保留当前真实储能状态；
3) 只用该时刻已经发布的光伏预报重算剩余时段；
4) 分别计算“保持当前合同”和“允许调整合同”的预测成本；
5) 只有允许调整更经济时才修改未来合同。

G 系列仅作为“预测误差风险增强策略/扩展检验”，不混入主模型。
它利用决策日前历史残差，在保持/半幅/全幅之间做经验候选比较。

【结算口径】
正式主口径采用 refund：
- 0点原计划为 x；最终合同为 y；
- 调增部分 (y-x)^+ 按 1.5 倍交易时价计；
- 调减部分 (x-y)^+ 从原计划费用中净退 50%；
合同费用 = sum[p*x + 1.5*p*(y-x)^+ - 0.5*p*(x-y)^+]。
紧急购电按 5 倍交易时价另计。
no_refund 仅保留作解释敏感性，不作为默认主口径。

【物理与信息边界】
- 10分钟一个区间，DT=1/6 h；输入功率 kW 统一乘 DT 转为 kWh 后优化与结算。
- 储能：1200~10800 kWh，最大充放电功率5000 kW，单向效率0.9，充放互斥。
- 2025-01-01 0:00 初始储电量6000 kWh；以后SOC跨日连续，不每日重置。
- 未来真实负荷/PV不得进入计划或合同调整；当前已完成区间观测可用于更新SOC与锚点。
- 0/6/12/18 的24小时光伏预报只使用当日剩余部分；不做跨日滚动前瞻。

【与你当前 C 目录的适配】
默认脚本位于 C/forecast_value.py：
- 数据目录：C/data
- 优先读取 C/data 下原始 附件1.xlsx、附件2.xlsx、附件3.xlsx；
- 若不存在，则读取 C/data/processed 下：附件1_clean.csv、小区负载_clean.csv、光伏实际功率_clean.csv、光伏预测_clean.csv；
- result3 模板优先 C/data/result3.xlsx；若没有，则允许使用 C/data/result4-3.xlsx 作为同四表结构模板副本（原文件绝不覆盖）；
- 输出目录：C/q3_output。

正式 result3 导出保留官方模板的 sheet 和时间表头，只填数值，不再重建/改写官方表头。
time_mapping.csv 单独记录“模型物理区间 ↔ 官方模板列标签”的对应关系。

推荐：
    python forecast_value.py --mode audit
    python forecast_value.py --mode demo --start-date 2025-03-20 --days 7 --scenarios core
    python forecast_value.py --mode annual --scenarios all-d --resume
    python forecast_value.py --mode figures
可选G扩展：
    python forecast_value.py --mode annual --scenarios risk --resume

依赖：python -m pip install numpy scipy openpyxl
"""
from __future__ import annotations
import argparse
import csv
import hashlib
import itertools
import json
import re
import shutil
import sys
import time
import warnings
from dataclasses import asdict, dataclass
from datetime import date, datetime, time as dtime, timedelta
from functools import lru_cache
from pathlib import Path
try:
    import numpy as np
    from openpyxl import load_workbook
    from scipy.optimize import milp, Bounds, LinearConstraint
    from scipy.sparse import lil_matrix, hstack, vstack, csr_matrix
except ImportError as exc:
    raise SystemExit('缺少依赖：python -m pip install numpy scipy openpyxl') from exc

# ================= 用户配置区：VS Code / PyCharm 直接运行 =================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / 'data'
PROCESSED_DIR = DATA_DIR / 'processed'
OUTPUT_DIR = BASE_DIR / 'q3_output'

RUN_MODE = 'demo'               # demo / annual / audit / figures
DEMO_START = '2025-03-20'
DEMO_DAYS = 7
SCENARIOS = 'core'              # core / all-d / risk / all / 逗号分隔自定义
RESUME = True
RUN_SELF_TEST = True

SETTLEMENT = 'refund'           # 正式主口径；no_refund只用于敏感性
TIME_LIMIT = 30.0
USE_LEXICOGRAPHIC = True
LEX_COST_EPS = 1e-6

HISTORY_DAYS = 7                # 仅G风险增强策略使用
MIN_HISTORY_DAYS = 3
TERMINAL_WEIGHT = 1.0           # 日末库存软价值权重；不是“调整启动费”
# ============================================================
T = 144
DT = 1/6
ETA = 0.9
EMIN, EMAX, EINITIAL = 1200.,10800.,6000.
B = 5000*DT
TOL = 2e-4
ISSUE_HOURS = (0,6,12,18)
ALGORITHM_VERSION = 'q3-rolling-v4-D-main-template-preserved'
RESUME_PATCH_TAG = 'output-time-mapping-v1'  # 仅导出/时间标签标准化；模型数学核心与既有断点兼容

@dataclass(frozen=True)
class Policy:
    load_mode: str = 'week_mean'
    pv_days: int = 3  # 仅供保留的Q2函数兼容，Q3基准PV不使用历史均值
    reserve: float = 0.
    mu_factor: float = 0.

@dataclass
class Data:
    dates: list[date]
    load: np.ndarray  # kWh, shape=(365,144)
    pv: np.ndarray    # kWh, shape=(365,144)
    price: np.ndarray
    forecasts: dict[datetime, np.ndarray]  # 24个整点功率，kW
    template_labels: list[str] | None
    files: dict[str, Path]

def require(condition, message):
    """不使用可被python -O关闭的assert检查外部数据。"""
    if not condition:
        raise ValueError(message)

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

def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s=str(value).strip().split()[0].replace('/', '-').replace('.', '-')
    parts=s.split('-')
    if len(parts)==3:
        s='-'.join([parts[0],parts[1].zfill(2),parts[2].zfill(2)])
    return datetime.strptime(s, '%Y-%m-%d').date()

def clock(minutes):
    return f'{minutes // 60:02d}:{minutes % 60:02d}'

def raw_sample_label(t):
    """原始右端点采样标签；最后一个时点按题目记法写作0:00+1。"""
    require(0 <= t < T, f'slot越界：{t}')
    end=(t+1)*10
    return '0:00+1' if end==1440 else clock(end)

def interval_start_label(t):
    require(0 <= t < T, f'slot越界：{t}')
    return clock(t*10)

def interval_end_label(t):
    require(0 <= t < T, f'slot越界：{t}')
    return clock((t+1)*10)

def physical_label(t):
    """统一论文表、图、自生成CSV使用的物理区间标签。"""
    return f'{interval_start_label(t)}-{interval_end_label(t)}'

def resolve_file(folder, stem, optional=False):
    """只在指定目录寻找同名xlsx；支持文件名末尾(数字)版本。"""
    folder = Path(folder)
    if not folder.exists():
        if optional:
            return None
        raise FileNotFoundError(folder)
    pattern = re.compile(re.escape(stem) + r'(?:\(\d+\)|（\d+）)*\.xlsx', re.I)
    found = sorted(p for p in folder.iterdir() if p.is_file() and pattern.fullmatch(p.name))
    if not found and optional:
        return None
    require(len(found) == 1,
            f'{stem}应有且仅有一个版本，找到{len(found)}个：{[p.name for p in found]}；目录：{folder}')
    return found[0]


def numeric(values, name):
    try:
        x = np.asarray(values, dtype=float)
    except (ValueError, TypeError) as exc:
        raise ValueError(f'{name}含非数字内容') from exc
    require(np.isfinite(x).all() and (x >= 0).all(), f'{name}含缺失、非有限值或负值')
    return x


def actual_sheet(ws):
    rows = [r for r in ws.values if any(v is not None for v in r)]
    require([time_minute(v) for v in rows[0][1:145]] == list(range(10,1441,10)),
            f'{ws.title}表头不符合00:10至24:00')
    dates = [as_date(r[0]) for r in rows[1:]]
    expected = [date(2025,1,1)+timedelta(days=i) for i in range(365)]
    require(dates == expected, f'{ws.title}日期缺失、重复、乱序或不在2025全年')
    a = numeric([r[1:145] for r in rows[1:]], ws.title)
    require(a.shape == (365,T), f'{ws.title}形状异常：{a.shape}')
    return dates, a * DT


def read_pv_forecasts(path, dates):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb.worksheets[0].values)
    finally:
        wb.close()
    expected_header = ['日期','预报时刻'] + [f'预报{h}小时' for h in range(1,25)]
    require(list(rows[0][:26]) == expected_header, '附件3表头与预期不符')
    forecasts = {}
    current_date = None
    for row_index, row in enumerate(rows[1:], 2):
        if all(v is None or v == '' for v in row):
            continue
        if row[0] is not None and str(row[0]).strip():
            current_date = as_date(row[0])
        require(current_date is not None, f'附件3第{row_index}行没有可继承日期')
        minute = time_minute(row[1])
        require(minute in (0,360,720,1080), f'附件3第{row_index}行发布时间异常')
        issue = datetime.combine(current_date, dtime()) + timedelta(minutes=minute)
        require(issue not in forecasts, f'附件3重复发布时间：{issue}')
        values = numeric(row[2:26], f'附件3第{row_index}行')
        require(values.shape == (24,), f'附件3第{row_index}行预测数不足24')
        forecasts[issue] = values
    expected = {datetime.combine(d,dtime(h)) for d in dates for h in ISSUE_HOURS}
    require(set(forecasts) == expected, '附件3发布记录不完整或存在范围外日期')
    return forecasts


def _read_csv_rows(path):
    with Path(path).open('r', encoding='utf-8-sig', newline='') as f:
        return list(csv.reader(f))


def _clean_header(value):
    return str(value).strip().replace(' ', '').replace('\ufeff','')


def _find_col(header, keywords, *, required=True):
    norm=[_clean_header(x) for x in header]
    for keyword in keywords:
        for i,name in enumerate(norm):
            if keyword == name or keyword in name:
                return i
    if required:
        raise ValueError(f'CSV表头中找不到字段 {keywords}；实际表头={header}')
    return None


def _read_attachment1_clean(path):
    rows=_read_csv_rows(path)
    require(len(rows)>=145, f'{path.name}行数不足')
    header=rows[0]
    it=_find_col(header,['时间'])
    ip=_find_col(header,['电价'])
    body=[r for r in rows[1:] if any(str(x).strip() for x in r)]
    require(len(body)==T, f'{path.name}应有144个10分钟时点，实际{len(body)}')
    require([time_minute(r[it]) for r in body]==list(range(10,1441,10)), f'{path.name}时间序列异常')
    price=numeric([r[ip] for r in body], '固定电价')
    require((price>0).all(), '固定电价必须为正')
    return price


def _read_actual_clean(path, value_keywords):
    """兼容365×145宽表或 日期/时间/值 长表；返回kWh。"""
    rows=_read_csv_rows(path)
    require(rows, f'{path}为空')
    header=rows[0]
    body=[r for r in rows[1:] if any(str(x).strip() for x in r)]
    expected=[date(2025,1,1)+timedelta(days=i) for i in range(365)]

    # 宽表：日期 + 144个时间列
    if len(header)>=145 and len(body)==365:
        times=[time_minute(x) for x in header[1:145]]
        require(times==list(range(10,1441,10)), f'{path.name}宽表时间列不是00:10~24:00')
        dates=[as_date(r[0]) for r in body]
        require(dates==expected, f'{path.name}日期不完整或乱序')
        arr=numeric([r[1:145] for r in body], path.name)
        require(arr.shape==(365,T), f'{path.name}形状异常：{arr.shape}')
        return dates,arr*DT

    # 长表：日期、时间、数值
    idate=_find_col(header,['日期'])
    itime=_find_col(header,['时间'])
    ivalue=_find_col(header,value_keywords)
    grouped={}
    for r in body:
        ds=as_date(r[idate]); minute=time_minute(r[itime])
        grouped.setdefault(ds,{})[minute]=float(r[ivalue])
    require(sorted(grouped)==expected, f'{path.name}日期不完整')
    arr=[]
    for ds in expected:
        require(sorted(grouped[ds])==list(range(10,1441,10)), f'{path.name} {ds} 时段不完整')
        arr.append([grouped[ds][m] for m in range(10,1441,10)])
    return expected,numeric(arr,path.name)*DT


def _read_forecast_clean(path, dates):
    rows=_read_csv_rows(path)
    require(rows, f'{path}为空')
    header=[_clean_header(x) for x in rows[0]]
    idate=_find_col(header,['日期'])
    iissue=_find_col(header,['预报时刻','发布时间','发布时刻'])
    forecast_cols=[]
    for h in range(1,25):
        idx=_find_col(header,[f'预报{h}小时'],required=False)
        if idx is None:
            # 允许“1小时”“预测1小时”等轻微命名差异
            candidates=[i for i,name in enumerate(header) if str(h) in name and '小时' in name and i not in (idate,iissue)]
            idx=candidates[0] if len(candidates)==1 else None
        require(idx is not None, f'{path.name}缺少预报{h}小时列')
        forecast_cols.append(idx)
    forecasts={};current_date=None
    for row_index,r in enumerate(rows[1:],2):
        if not any(str(x).strip() for x in r):
            continue
        if idate < len(r) and str(r[idate]).strip():
            current_date=as_date(r[idate])
        require(current_date is not None, f'{path.name}第{row_index}行无法确定日期')
        minute=time_minute(r[iissue])
        require(minute in (0,360,720,1080), f'{path.name}第{row_index}行预报时刻异常')
        issue=datetime.combine(current_date,dtime())+timedelta(minutes=minute)
        require(issue not in forecasts, f'{path.name}重复发布时间：{issue}')
        forecasts[issue]=numeric([r[i] for i in forecast_cols], f'{path.name}第{row_index}行')
    expected={datetime.combine(d,dtime(h)) for d in dates for h in ISSUE_HOURS}
    require(set(forecasts)==expected, f'{path.name}应有365×4次预报，当前记录不完整')
    return forecasts


def _resolve_template(folder):
    """优先result3；当前目录截图没有result3时，允许result4-3作为同结构模板副本。"""
    candidates=[]
    for root in [Path(folder), BASE_DIR]:
        if not root.exists():
            continue
        for stem in ('result3','result4-3'):
            p=resolve_file(root,stem,optional=True)
            if p is not None and p not in candidates:
                candidates.append(p)
    if not candidates:
        return None
    result3=[p for p in candidates if p.name.lower().startswith('result3')]
    return result3[0] if result3 else candidates[0]


def _validate_template(template, dates):
    wb=load_workbook(template,read_only=True,data_only=True)
    try:
        require(wb.sheetnames==['计划购电量','调整购电量','充放电量','紧急购电量'],
                f'模板工作表结构异常：{wb.sheetnames}')
        expected_dates=dates[31:]
        for name in ('计划购电量','调整购电量'):
            rows=list(wb[name].values)
            require(len(rows[0])>=147, f'{name}模板列数不足147')
            dates_in_template=[as_date(r[0]) for r in rows[1:] if r[0] not in (None,'')]
            require(dates_in_template==expected_dates, f'{name}模板日期不是2025-02-01~12-31')
        labels=[str(v) for v in next(wb['计划购电量'].values)[1:145]]
        other=[str(v) for v in next(wb['调整购电量'].values)[1:145]]
        require(labels==other and len(labels)==T,'计划/调整模板时间标签不一致')
        return labels
    finally:
        wb.close()


def read_data(folder):
    folder=Path(folder)
    require(folder.is_dir(), f'数据目录不存在：{folder}')
    files={}

    raw=[resolve_file(folder,f'附件{i}',optional=True) for i in (1,2,3)]
    if all(p is not None for p in raw):
        files.update({f'附件{i}':raw[i-1] for i in (1,2,3)})
        wb=load_workbook(files['附件1'],read_only=True,data_only=True)
        try:
            rows=[r for r in wb.worksheets[0].values if r[0] is not None]
        finally:
            wb.close()
        require(list(rows[0][:4])==['时间','电价','小区负载','光伏发电预测功率'],'附件1表头异常')
        require([time_minute(r[0]) for r in rows[1:]]==list(range(10,1441,10)),'附件1时间序列异常')
        price=numeric([r[1] for r in rows[1:]],'电价')
        require((price>0).all(),'当前模型预期正电价')
        wb=load_workbook(files['附件2'],read_only=True,data_only=True)
        try:
            ds,load=actual_sheet(wb['小区负载'])
            ds2,pv=actual_sheet(wb['光伏发电实际功率'])
        finally:
            wb.close()
        require(ds==ds2,'负荷和PV日期不一致')
        forecasts=read_pv_forecasts(files['附件3'],ds)
        source_mode='xlsx原始附件'
    else:
        processed=folder/'processed'
        required={
            '附件1_clean':processed/'附件1_clean.csv',
            '小区负载_clean':processed/'小区负载_clean.csv',
            '光伏实际功率_clean':processed/'光伏实际功率_clean.csv',
            '光伏预测_clean':processed/'光伏预测_clean.csv',
        }
        missing=[str(p) for p in required.values() if not p.exists()]
        require(not missing,
                'data目录未找到完整附件1/2/3.xlsx，processed回退文件也不完整。缺少：'+', '.join(missing))
        files.update(required)
        price=_read_attachment1_clean(required['附件1_clean'])
        ds,load=_read_actual_clean(required['小区负载_clean'],['小区负载','负载'])
        ds2,pv=_read_actual_clean(required['光伏实际功率_clean'],['光伏发电实际功率','光伏实际功率','实际功率'])
        require(ds==ds2,'清洗后的负荷与PV日期不一致')
        forecasts=_read_forecast_clean(required['光伏预测_clean'],ds)
        source_mode='processed清洗CSV'

    template=_resolve_template(folder)
    labels=None
    if template is not None:
        labels=_validate_template(template,ds)
        files['result3_template']=template
    else:
        print('[warning] 未找到result3.xlsx或result4-3.xlsx模板；audit/demo仍可计算，但正式annual导出result3.xlsx前必须补模板。',flush=True)

    print(f'输入模式：{source_mode}',flush=True)
    if template is not None:
        print(f'结果模板：{template}',flush=True)
    return Data(ds,load,pv,price,forecasts,labels,files)


def get_pv_forecast(issue_time, decision_time, issued_values, anchor_kw):
    """只接收一版已发布预报，不接收全年实际矩阵。

    返回从issue_time到当天24点的区间电量。本阶段只允许在发布时刻调用。
    anchor_kw必须由调用者提供当时已知的快照，禁止使用下一时段观测。
    """
    require(issue_time <= decision_time, '禁止使用尚未发布的光伏预报')
    require(issue_time == decision_time, '此接口用于发布节点；实时节点应复用已生成预测')
    require(issue_time.hour in ISSUE_HOURS and issue_time.minute == 0 and issue_time.second == 0,
            '不是有效预报发布时刻')
    values = numeric(issued_values, '预报值')
    require(values.shape == (24,), '必须提供24个整点预报')
    require(np.isfinite(anchor_kw) and anchor_kw >= 0, '插值起点不合法')
    steps = (24-issue_time.hour)*6
    powers = np.interp(np.arange(1,steps+1)/6, np.arange(25), np.r_[anchor_kw,values])
    return powers * DT

def known_anchor(data, issue):
    """按当前端点解释，发布时刻只读取已经结束区间的最后一个样本。"""
    index = (issue.date()-data.dates[0]).days
    require(0 <= index < len(data.dates), '锚点日期越界')
    if issue.hour:
        return float(data.pv[index,issue.hour*6-1]/DT)
    if index > 0:
        return float(data.pv[index-1,-1]/DT)
    raise ValueError('1月1日0点缺少前一日观测；基准诊断请从1月2日或之后开始')

def write_csv(path, fieldnames, rows):
    """稳定写CSV：兼容断点续跑前后字段不完全一致的诊断记录。

    对list/tuple等可重复遍历的记录，自动取所有字典键的并集，
    避免旧checkpoint缺少新诊断字段、而新checkpoint包含新字段时
    csv.DictWriter因'extrasaction=raise'直接中断全年导出。
    对生成器（例如dispatch_detail）保持调用方给定字段，不额外物化。
    """
    fields=list(fieldnames)
    if isinstance(rows,(list,tuple)):
        seen=set(fields)
        for row in rows:
            for key in row:
                if key not in seen:
                    fields.append(key)
                    seen.add(key)
    with path.open('w',encoding='utf-8-sig',newline='') as stream:
        writer = csv.DictWriter(stream,fieldnames=fields,extrasaction='ignore')
        writer.writeheader()
        writer.writerows(rows)

def write_json(path, obj):
    path.write_text(json.dumps(obj,ensure_ascii=False,indent=2,allow_nan=False),encoding='utf-8')

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

def optimize(net, price, s0, policy, terminal_lambda, fixed_q=None, time_limit=30.,
             lexicographic=False, diagnostics=None):
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
                   options={'time_limit': time_limit, 'mip_rel_gap': 1e-7, 'threads': 1, 'mip_feasibility_tolerance': 1e-9})
    # Fail explicitly rather than silently publish a nonoptimal/time-limited plan.
    if not res.success or res.x is None:
        raise RuntimeError(f'MILP failed: {res.message}; horizon={n}, initial={s0}')
    x = res.x
    primary_x=x.copy()
    primary_value=float(obj@x)
    second_attempted=False
    second_solved=False
    second_accepted=False
    second_status='not_needed'
    if fixed_q is not None and lexicographic and primary_x[3*n]>TOL:
        second_attempted=True
        second_obj=np.zeros_like(obj)
        second_obj[3*n]=1.0
        cost_limit=LinearConstraint(csr_matrix(obj.reshape(1,-1)),
                                    -np.inf,primary_value+LEX_COST_EPS)
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore',message='Unrecognized options detected.*')
            second=milp(second_obj,integrality=integer,bounds=bounds,
                        constraints=[LinearConstraint(a,lower,rhs),cost_limit],
                        options={'time_limit':time_limit,'mip_rel_gap':1e-9,'threads':1, 'mip_feasibility_tolerance':1e-9})
        if second.success and second.x is not None:
            second_solved=True
            candidate=second.x
            cost_change=float(obj@candidate)-primary_value
            # 第二级只是同主费用下的数值择优，不是经济模型本体。HiGHS在极少数MILP上
            # 可能返回一个满足费用容差、但w[0]因数值容差反而略大的可行解。
            # 此时保留已经验证可行的一阶段最优解，而不是让全年计算中断。
            if cost_change<=LEX_COST_EPS+1e-5 and candidate[3*n] < primary_x[3*n]-1e-8:
                x=candidate
                second_accepted=True
                second_status='accepted'
            elif cost_change>LEX_COST_EPS+1e-5:
                second_status='rejected_cost_tolerance'
            else:
                second_status='rejected_no_unused_improvement'
        else:
            second_status='solver_failed_fallback_primary'
        if diagnostics is not None and second_status!='accepted':
            diagnostics['lexicographic_note']=second_status
    if diagnostics is not None:
        diagnostics.update(lexicographic_enabled=bool(lexicographic),
                           second_attempted=second_attempted,second_solved=second_solved,
                           second_accepted=second_accepted,second_status=second_status,
                           primary_objective=primary_value,selected_primary_objective=float(obj@x),
                           primary_objective_change=float(obj@x)-primary_value,
                           cost_tolerance_yuan=LEX_COST_EPS,
                           original_current_unused_kwh=float(primary_x[3*n]),
                           selected_current_unused_kwh=float(x[3*n]),
                           original_current_charge_kwh=float(primary_x[n]),
                           selected_current_charge_kwh=float(x[n]))
    values = a @ x
    if np.max(values-rhs) > TOL or np.max(abs(values[:2*n]-rhs[:2*n])) > TOL:
        raise RuntimeError('Solver feasibility audit failed')
    require(np.max(lb-x)<TOL and np.max(x-ub)<TOL,'变量上下界审计失败')
    require(np.max(abs(x[5*n:6*n]-np.round(x[5*n:6*n])))<TOL,'整数状态审计失败')
    return tuple(x[i*n:(i+1)*n].copy() for i in range(5))

def current_action(lhat, vhat, current_l, current_v, q, t, s0, policy, price, lam, limit,
                   lexicographic=False, diagnostics=None):
    """Receives only a scalar current observation and future FORECASTS."""
    net = (lhat[t:] - vhat[t:]).copy()
    net[0] = current_l-current_v
    arrays = optimize(net, price[t:], s0, policy, lam, q[t:], limit,
                      lexicographic=lexicographic,diagnostics=diagnostics)
    return np.array([x[0] for x in arrays])

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


def settlement_bill(original, final, emergency, price, rule):
    """只结算最终有效合同，绝不累加中途优化目标。"""
    up=np.maximum(final-original,0.)
    down=np.maximum(original-final,0.)
    cp=float(price@original)
    ca=float(price@(1.5*up+(-.5 if rule=='refund' else .5)*down))
    ce=float(5*price@emergency)
    return dict(plan_cost=cp,adjustment_net_cost=ca,contract_cost=cp+ca,
                emergency_cost=ce,total_actual_cost=cp+ca+ce,
                increase_kwh=float(up.sum()),decrease_kwh=float(down.sum()))


def adjustment_problem(net, price, soc, original, pol, lam, rule, limit, fixed=None):
    """变量块：[q,c,d,w,S,z,reserve,terminal] + [e,u,v]。
    q-u+v=original；成本original*p+1.5p*u-.5p*v+5p*e（退款口径）。
    fixed用于公平计算保持当前合同的反事实，储能仍允许重新优化。
    """
    n=len(net)
    base,bounds,integer=skeleton(n)
    nb=7*n+1
    extra=lil_matrix((5*n+1,3*n))
    extra[:n,:n]=np.eye(n)
    a=hstack([base,extra.tocsr()],format='csr')
    rows=lil_matrix((n,nb+3*n))
    for t in range(n):
        rows[t,t]=1
        rows[t,nb+n+t]=-1
        rows[t,nb+2*n+t]=1
    a=vstack([a,rows.tocsr()],format='csr')
    rhs=np.zeros(6*n+1)
    rhs[:n]=net
    rhs[n]=soc
    rhs[3*n:4*n]=B
    rhs[4*n:5*n]=-(EMIN+pol.reserve)
    rhs[5*n]=-EINITIAL
    rhs[5*n+1:]=original
    lower=np.full_like(rhs,-np.inf)
    lower[:2*n]=rhs[:2*n]
    lower[5*n+1:]=original
    lb=np.r_[bounds.lb,np.zeros(3*n)]
    ub=np.r_[bounds.ub,np.full(3*n,np.inf)]
    if fixed is not None:
        lb[:n]=fixed
        ub[:n]=fixed
    objective=np.zeros(nb+3*n)
    objective[6*n:7*n]=pol.mu_factor*lam*DT
    objective[nb-1]=lam
    objective[nb:nb+n]=5*price
    objective[nb+n:nb+2*n]=1.5*price
    objective[nb+2*n:]=(-.5 if rule=='refund' else .5)*price
    integrality=np.r_[integer,np.zeros(3*n)]
    with warnings.catch_warnings():
        warnings.filterwarnings('ignore',message='Unrecognized options detected.*')
        result=milp(objective,integrality=integrality,bounds=Bounds(lb,ub),
                    constraints=LinearConstraint(a,lower,rhs),
                    options={'time_limit':limit,'mip_rel_gap':1e-7,'threads':1})
    if not result.success or result.x is None:
        raise RuntimeError(f'调整求解失败：{result.message}，剩余{n}区间')
    x=result.x
    ax=a@x
    finite=np.isfinite(lower)
    violation=max(float(np.max(ax-rhs)),float(np.max(lower[finite]-ax[finite])),
                  float(np.max(lb-x)),float(np.max(x-ub)))
    require(violation<TOL,'调整优化可行性审计失败')
    require(np.max(abs(x[5*n:6*n]-np.round(x[5*n:6*n])))<TOL,'二元状态异常')
    # 仅清除浮点级负合同，保持随后的非负固定边界有效。
    q=np.maximum(x[:n],0.)
    value=float(result.fun+price@original)
    return q,value


def decide_revision(net, price, soc, original, current, pol, lam, args):
    q,new_value=adjustment_problem(net,price,soc,original,pol,lam,args.settlement,args.time_limit)
    _,keep_value=adjustment_problem(net,price,soc,original,pol,lam,args.settlement,args.time_limit,fixed=current)
    improvement=keep_value-new_value
    # 仅为数值容差，不是题目之外的固定调整费或经济阈值。
    tol=max(1e-5,1e-7*max(abs(keep_value),abs(new_value)))
    require(improvement>=-2*tol,'允许调整反而劣于保持：检查求解精度或约束')
    chosen=q if improvement>tol else current.copy()
    return chosen,dict(predicted_keep_objective=keep_value,predicted_adjust_objective=new_value,
                       predicted_net_improvement=max(0.,improvement),
                       changed=bool(np.max(abs(chosen-current))>TOL),
                       numerical_tolerance=tol)


def scenario_config(name):
    if name=='B0':return False,(),False
    if name=='B1':return True,(),False
    if name=='C':return True,(6,12,18),False
    require(name.startswith(('D','G')),'未知方案：'+name)
    hours=tuple(int(x) for x in name[1:].split('_'))
    require(hours==tuple(sorted(set(hours))) and all(h in (6,12,18) for h in hours), '方案时刻非法：'+name)
    return True,hours,True


def historical_net_residuals(data, ds, hour, window):
    """每条轨迹仅来自决策日前已完成的日期。数值单位kWh。
    R=(历史实际负荷-历史实际PV)-(该日0点负荷预测-该日同发布时刻PV预测)。
    禁止随机划分、全年的统计修正、当日未来真值。历史锚点与运行重构相同。
    """
    index=(ds-data.dates[0]).days
    t=hour*6
    residuals=[]
    dates=[]
    # 至少已有14天负荷历史，保证与正式week_mean口径一致。
    for j in range(max(14,index-window),index):
        old_day=data.dates[j]
        require(old_day<ds,'历史残差读取当日或未来')
        old_l,_=forecast(data.load[:j],data.pv[:j],Policy())
        issue=datetime.combine(old_day,dtime(hour))
        old_g=get_pv_forecast(issue,issue,data.forecasts[issue],known_anchor(data,issue))
        residual=(data.load[j,t:]-data.pv[j,t:])-(old_l[t:]-old_g)
        require(np.isfinite(residual).all(),'历史净负荷残差异常')
        residuals.append(residual)
        dates.append(old_day)
    return dates,residuals


def fixed_contract_proxy(net, price, soc, original, contract, pol, lam, args):
    """在一条假设净负荷轨迹上求固定合同下的最小补救费用。
    这不是实际回放：场景内未来轨迹已知，所得值仅为候选比较的近似。
    这里只需最优费用，故不对尚未执行的动作额外做V2第二级求解。
    """
    emergency,charge,discharge,unused,states=optimize(
        net,price,soc,pol,lam,fixed_q=contract,time_limit=args.time_limit)
    contract_only=settlement_bill(original,contract,np.zeros_like(contract),price,args.settlement)['contract_cost']
    reserve=pol.mu_factor*lam*DT*float(np.maximum(EMIN+pol.reserve-states,0.).sum())
    terminal=lam*max(EINITIAL-float(states[-1]),0.)
    return float(contract_only+5*price@emergency+reserve+terminal)


def historical_candidate_check(data, ds, hour, net, price, soc, original, current, proposal, pol, lam, args):
    """对保持/半幅/全幅做相同样本比较；相同成本优先保持或较小调整。"""
    dates,residuals=historical_net_residuals(data,ds,hour,args.history_days)
    info=dict(guard_history_count=len(dates),
              guard_history_start=str(dates[0]) if dates else '',
              guard_history_end=str(dates[-1]) if dates else '',
              guard_selected_fraction=0.,guard_mean_keep=None,guard_mean_selected=None,
              guard_estimated_gain=None,guard_status='insufficient_history')
    if len(dates)<MIN_HISTORY_DAYS:
        return current.copy(),info,[]
    if np.max(abs(proposal-current))<=TOL:
        info['guard_status']='point_proposal_unchanged'
        return current.copy(),info,[]
    fractions=(0.,.5,1.)
    candidates=[np.maximum(current+alpha*(proposal-current),0.) for alpha in fractions]
    values=[]
    records=[]
    for alpha,contract in zip(fractions,candidates):
        costs=[]
        for history_day,residual in zip(dates,residuals):
            assumed_net=net+residual
            # 净负荷可以为负（富余电量）；没有将负值错误裁成0。
            cost=fixed_contract_proxy(assumed_net,price,soc,original,contract,pol,lam,args)
            costs.append(cost)
            records.append(dict(history_date=str(history_day),fraction=alpha,
                                scenario_proxy_cost=cost,selected=False))
        values.append(float(np.mean(costs)))
    best=0
    tolerance=max(1e-5,1e-7*max(abs(x) for x in values))
    for j in (1,2):
        if values[j]<values[best]-tolerance:best=j
    for row in records:row['selected']=row['fraction']==fractions[best]
    info.update(guard_selected_fraction=fractions[best],guard_mean_keep=values[0],
                guard_mean_selected=values[best],guard_estimated_gain=values[0]-values[best],
                guard_status='keep' if best==0 else ('partial' if best==1 else 'full'))
    return candidates[best],info,records


def audit_final_contract(result):
    """从0点合同逐次重建每个交付时段的最终合同，并与实际轨迹对账。"""
    original=result['original']
    require(np.isfinite(original).all() and np.min(original)>=-TOL,'0点合同非法')
    versions=result['revisions']
    for t in range(T):
        current=float(original[t])
        for row in versions:
            if row['slot']!=t:continue
            require(row['issue_hour']*6<=t,'历史执行时段被修改')
            require(abs(row['original_plan_kwh']-original[t])<TOL,'0点计划被覆盖')
            require(abs(row['previous_contract_kwh']-current)<TOL,'合同版本不衔接')
            current=row['new_contract_kwh']
        require(abs(current-result['trace'][t,0])<TOL,'最终合同与实际交付不一致')


def midnight_plan(data, ds, soc, name, args):
    """不接收当天实际曲线，计划仅由历史切片与0点预报决定。"""
    i=(ds-data.dates[0]).days
    pol=Policy()
    lh,vh=forecast(data.load[:i],data.pv[:i],pol)
    official,_,_=scenario_config(name)
    if official:
        issue=datetime.combine(ds,dtime())
        vh=get_pv_forecast(issue,issue,data.forecasts[issue],known_anchor(data,issue))
    lam=float(data.price.min()/ETA)*args.terminal_weight
    q=optimize(lh-vh,data.price,soc,pol,lam,time_limit=args.time_limit)[0]
    return np.maximum(q,0.),lh,vh


def simulate_q3_day(data, ds, soc, name, args):
    i=(ds-data.dates[0]).days
    original,lh,vh=midnight_plan(data,ds,soc,name,args)
    current=original.copy()
    _,hours,allow=scenario_config(name)
    pol=Policy()
    lam=float(data.price.min()/ETA)*args.terminal_weight
    traces=np.zeros((T,7))
    used_pv=np.zeros(T)
    issue_at=np.zeros(T,dtype=int)
    logs=[]
    revisions=[]
    lex_logs=[]
    guard_rows=[]
    last_hour=0
    s=soc
    for t in range(T):
        if t%36==0 and t//6 in hours:
            h=t//6
            issue=datetime.combine(ds,dtime(h))
            # anchor索引t-1；当前t实际量尚不进入合同优化。
            vh[t:]=get_pv_forecast(issue,issue,data.forecasts[issue],known_anchor(data,issue))
            last_hour=h
            old=current[t:].copy()
            log=dict(date=str(ds),scenario=name,issue_hour=h,soc_at_update=s)
            if allow:
                chosen,diagnostic=decide_revision(lh[t:]-vh[t:],data.price[t:],s,
                                                 original[t:],old,pol,lam,args)
                if name.startswith('G'):
                    chosen,guard,records=historical_candidate_check(
                        data,ds,h,lh[t:]-vh[t:],data.price[t:],s,original[t:],old,chosen,pol,lam,args)
                    diagnostic.update(guard)
                    diagnostic['changed']=bool(np.max(abs(chosen-old))>TOL)
                    for record in records:
                        guard_rows.append(dict(date=str(ds),scenario=name,issue_hour=h,**record))
                current[t:]=chosen
                log.update(diagnostic)
                log['revision_l1_kwh']=float(abs(chosen-old).sum())
            else:
                log.update(changed=False,revision_l1_kwh=0.,predicted_keep_objective=None,
                           predicted_adjust_objective=None,predicted_net_improvement=None,numerical_tolerance=None)
            logs.append(log)
            for j in range(t,T):
                revisions.append(dict(date=str(ds),scenario=name,issue_hour=h,slot=j,
                                      raw_sample_time=raw_sample_label(j),
                                      interval_start=interval_start_label(j),
                                      interval_end=interval_end_label(j),
                                      physical_interval=physical_label(j),
                                      original_plan_kwh=float(original[j]),
                                      previous_contract_kwh=float(old[j-t]),new_contract_kwh=float(current[j]),
                                      pv_forecast_kwh=float(vh[j])))
        used_pv[t]=vh[t]
        issue_at[t]=last_hour
        lex_record=dict(date=str(ds),scenario=name,slot=t,
                        raw_sample_time=raw_sample_label(t),
                        interval_start=interval_start_label(t),
                        interval_end=interval_end_label(t),
                        physical_interval=physical_label(t))
        e,c,d,w,snext=current_action(lh,vh,float(data.load[i,t]),float(data.pv[i,t]),
                                    current,t,s,pol,data.price,lam,args.time_limit,
                                    lexicographic=args.lexicographic,diagnostics=lex_record)
        lex_logs.append(lex_record)
        traces[t]=current[t],e,c,d,w,s,snext
        s=snext
    audit(traces,data.load[i],data.pv[i])
    return dict(trace=traces,original=original,load_forecast=lh,pv_forecast=used_pv,
                issue_hour=issue_at,logs=logs,revisions=revisions,lex_logs=lex_logs,guard_rows=guard_rows)


def atomic_json(path, obj):
    temp=path.with_suffix(path.suffix+'.tmp')
    write_json(temp,obj)
    temp.replace(path)


def day_summary(data, ds, result, name, args):
    x=result['trace']
    settlement=settlement_bill(result['original'],x[:,0],x[:,1],data.price,args.settlement)
    trajectory=np.r_[x[0,5],x[:,6]]
    summary=dict(date=str(ds),scenario=name,
                 **settlement,
                 plan_kwh=float(result['original'].sum()),executed_contract_kwh=float(x[:,0].sum()),
                 emergency_kwh=float(x[:,1].sum()),charge_kwh=float(x[:,2].sum()),
                 discharge_kwh=float(x[:,3].sum()),unused_kwh=float(x[:,4].sum()),
                 soc_start=float(x[0,5]),soc_end=float(x[-1,6]),
                 soc_min_kwh=float(trajectory.min()),soc_max_kwh=float(trajectory.max()),
                 soc_range_kwh=float(trajectory.max()-trajectory.min()),
                 emergency_slots=int((x[:,1]>1e-6).sum()),
                 adjustment_nodes=sum(bool(row['changed']) for row in result['logs']),
                 adjustment_abs_kwh=float(settlement['increase_kwh']+settlement['decrease_kwh']),
                 mode=args.mode)
    return summary


def checkpoint_day(path, result):
    temp=path.with_suffix('.tmp')
    with temp.open('wb') as f:
        np.savez_compressed(f,trace=result['trace'],original=result['original'],
                            load_forecast=result['load_forecast'],pv_forecast=result['pv_forecast'],
                            issue_hour=result['issue_hour'],logs=json.dumps(result['logs'],ensure_ascii=False),
                            revisions=json.dumps(result['revisions'],ensure_ascii=False),
                            lex_logs=json.dumps(result.get('lex_logs',[]),ensure_ascii=False),
                            guard_rows=json.dumps(result.get('guard_rows',[]),ensure_ascii=False))
    temp.replace(path)


def load_checkpoint(path):
    with np.load(path,allow_pickle=False) as f:
        result={k:f[k].copy() for k in ['trace','original','load_forecast','pv_forecast','issue_hour']}
        result.update(logs=json.loads(str(f['logs'])),revisions=json.loads(str(f['revisions'])))
        result['lex_logs']=json.loads(str(f['lex_logs']))
        result['guard_rows']=json.loads(str(f['guard_rows']))
    require(result['trace'].shape==(T,7) and result['original'].shape==(T,), '断点形状错误')
    return result


def run_sequence(data, dates, start_soc, name, folder, args):
    folder.mkdir(parents=True,exist_ok=True)
    checkpoints=folder/'checkpoints'
    checkpoints.mkdir(exist_ok=True)
    current_code_sha=hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    manifest=dict(version=ALGORITHM_VERSION,code_sha256=current_code_sha,resume_patch=RESUME_PATCH_TAG,
                  scenario=name,start_date=str(dates[0]),end_date=str(dates[-1]),initial_soc=start_soc,
                  settlement=args.settlement,policy=asdict(Policy()),time_limit=args.time_limit,
                  lexicographic=args.lexicographic,lex_cost_eps=LEX_COST_EPS,
                  history_days=args.history_days,min_history_days=MIN_HISTORY_DAYS,
                  terminal_weight=args.terminal_weight,
                  input_sha256={k:hashlib.sha256(v.read_bytes()).hexdigest() for k,v in data.files.items()})
    mf=folder/'manifest.json'
    if mf.exists():
        previous=json.loads(mf.read_text(encoding='utf-8'))
        if previous!=manifest:
            # 只要模型版本、模型参数和输入文件哈希完全一致，
            # 就允许绘图/CSV/时间标签等输出层源码改动复用既有checkpoint。
            # 若数学模型改变，必须同步提升ALGORITHM_VERSION，此处会自动拒绝旧断点。
            prev_core={k:v for k,v in previous.items() if k not in ('code_sha256','resume_patch')}
            new_core={k:v for k,v in manifest.items() if k not in ('code_sha256','resume_patch')}
            safe_migration=(previous.get('version')==ALGORITHM_VERSION and prev_core==new_core)
            require(safe_migration,
                    f'断点模型配置或输入已变化，请换新的输出目录，避免混用：{folder}')
            print(f'[{name}] 仅检测到源码/导出层变更；模型版本、参数和输入未变，安全复用既有断点。',flush=True)
            atomic_json(mf,manifest)
    else:
        require(not list(checkpoints.glob('*.npz')),'缺少manifest但存在旧断点，请换目录')
        atomic_json(mf,manifest)
    results={}
    summaries=[]
    s=start_soc
    for ds in dates:
        target=checkpoints/f'{ds}.npz'
        begin=time.perf_counter()
        cached=args.resume and target.exists()
        if cached:
            result=load_checkpoint(target)
        elif ds==date(2025,1,1):
            # 首日无历史：保持SOC，实时缺口紧急补购，不用全天真值制定合同。
            x=bootstrap(data.load[0],data.pv[0],s)
            result=dict(trace=x,original=np.zeros(T),load_forecast=np.zeros(T),pv_forecast=np.zeros(T),
                        issue_hour=np.zeros(T,dtype=int),logs=[],revisions=[])
            checkpoint_day(target,result)
        else:
            result=simulate_q3_day(data,ds,s,name,args)
            checkpoint_day(target,result)
        i=(ds-data.dates[0]).days
        audit(result['trace'],data.load[i],data.pv[i])
        audit_final_contract(result)
        require(abs(result['trace'][0,5]-s)<TOL,'断点/跨日初始SOC不一致')
        s=float(result['trace'][-1,6])
        row=day_summary(data,ds,result,name,args)
        summaries.append(row)
        results[ds]=result
        write_csv(folder/'daily_summary.csv',list(row),summaries)
        status='续跑读取' if cached else f'{time.perf_counter()-begin:.1f}s'
        print(f'[{name}] {ds} 总费={row["total_actual_cost"]:.2f} 紧急={row["emergency_kwh"]:.2f}kWh '
              f'调增={row["increase_kwh"]:.2f} 调减={row["decrease_kwh"]:.2f} SOC={s:.2f} ({status})',flush=True)
    return results,summaries


def emergency_intervals(x):
    result=[]
    t=0
    while t<T:
        if x[t,1]<=1e-6:
            t+=1
            continue
        end=t+1
        while end<T and x[end,1]>1e-6:end+=1
        result.append((f'{clock(t*10)}-{clock(end*10)}',float(x[t:end,1].sum())))
        t=end
    return result


def export_run(data, results, summaries, folder, args):
    """导出审计CSV，并在存在官方模板时原样保留模板表头/工作表，仅填数值。"""
    name=summaries[0]['scenario']
    detail_fields=['date','slot','raw_sample_time','interval_start','interval_end','physical_interval',
                   'price','original_plan_kwh','executed_contract_kwh',
                   'load_forecast_kwh','pv_forecast_kwh','latest_issue_hour','actual_load_kwh','actual_pv_kwh',
                   'emergency_kwh','charge_kwh','discharge_kwh','unused_kwh','soc_start_kwh','soc_end_kwh']

    def detail_rows():
        for ds,result in results.items():
            i=(ds-data.dates[0]).days
            for t,x in enumerate(result['trace']):
                row=dict(date=str(ds),slot=t,
                         raw_sample_time=raw_sample_label(t),
                         interval_start=interval_start_label(t),
                         interval_end=interval_end_label(t),
                         physical_interval=physical_label(t),price=float(data.price[t]),
                         original_plan_kwh=float(result['original'][t]),executed_contract_kwh=float(x[0]),
                         load_forecast_kwh=float(result['load_forecast'][t]),pv_forecast_kwh=float(result['pv_forecast'][t]),
                         latest_issue_hour=int(result['issue_hour'][t]),
                         actual_load_kwh=float(data.load[i,t]),actual_pv_kwh=float(data.pv[i,t]))
                row.update(dict(zip(detail_fields[-6:],map(float,x[1:]))))
                yield row

    write_csv(folder/'dispatch_detail.csv',detail_fields,detail_rows())
    logs=[row for result in results.values() for row in result['logs']]
    revisions=[row for result in results.values() for row in result['revisions']]
    if logs: write_csv(folder/'update_log.csv',list(logs[0]),logs)
    if revisions: write_csv(folder/'contract_versions.csv',list(revisions[0]),revisions)
    lex_logs=[row for result in results.values() for row in result.get('lex_logs',[])]
    if lex_logs: write_csv(folder/'lexicographic_audit.csv',list(lex_logs[0]),lex_logs)
    guard_rows=[row for result in results.values() for row in result.get('guard_rows',[])]
    if guard_rows: write_csv(folder/'guard_evaluation.csv',list(guard_rows[0]),guard_rows)
    if name.startswith('G') and logs: write_csv(folder/'decision_history.csv',list(logs[0]),logs)

    # 论文表1/2/3指定日期辅助CSV
    required_dates={date(2025,3,20),date(2025,6,21),date(2025,9,23),date(2025,12,21)}
    table1=[];table2=[];table3=[]
    for ds,result in results.items():
        if ds not in required_dates:
            continue
        x=result['trace']
        for t in (60,72,84,96,108,120):
            table1.append(dict(date=str(ds),interval=physical_label(t),plan_kwh=float(result['original'][t]),
                               final_contract_kwh=float(x[t,0]),emergency_kwh=float(x[t,1]),
                               total_external_kwh=float(x[t,0]+x[t,1])))
        for block in range(6):
            z=x[block*24:(block+1)*24]
            table2.append(dict(date=str(ds),interval=f'{clock(block*240)}-{clock((block+1)*240)}',
                               charge_kwh=float(z[:,2].sum()),discharge_kwh=float(z[:,3].sum()),
                               day_start_soc=float(x[0,5]),day_end_soc=float(x[-1,6])))
        for label,amount in emergency_intervals(x) or [('无',0.)]:
            table3.append(dict(date=str(ds),interval=label,emergency_kwh=amount))
    for stem,rows in [('specified_table1',table1),('specified_table2',table2),('specified_table3',table3)]:
        if rows: write_csv(folder/f'{stem}.csv',list(rows[0]),rows)

    full=(args.mode=='annual' and list(results)==data.dates[31:])
    target=None
    template=data.files.get('result3_template')
    if template is None:
        if full:
            raise FileNotFoundError('正式annual已完成，但未找到result3.xlsx/result4-3.xlsx模板，无法生成官方result3.xlsx')
        print(f'[warning] {name} 无模板：跳过诊断xlsx，仅保留CSV/JSON。',flush=True)
    else:
        wb=load_workbook(template)
        require(wb.sheetnames==['计划购电量','调整购电量','充放电量','紧急购电量'],
                '模板四个工作表结构发生变化')
        plan=wb['计划购电量'];adjust=wb['调整购电量']
        bat=wb['充放电量'];emg=wb['紧急购电量']
        original_plan_header=[c.value for c in plan[1]]
        original_adjust_header=[c.value for c in adjust[1]]

        # 计划/调整表保留官方日期与表头，仅清空数值区域后回填。
        for ws in (plan,adjust):
            for row in ws.iter_rows(min_row=2,min_col=2,max_col=147):
                for cell in row:
                    cell.value=None
        plan_rows={as_date(plan.cell(r,1).value):r for r in range(2,plan.max_row+1) if plan.cell(r,1).value not in (None,'')}
        adjust_rows={as_date(adjust.cell(r,1).value):r for r in range(2,adjust.max_row+1) if adjust.cell(r,1).value not in (None,'')}

        for (ds,result),summary in zip(results.items(),summaries):
            require(ds in plan_rows and ds in adjust_rows,
                    f'模板计划/调整工作表中没有日期 {ds}')
            x=result['trace'];original=result['original']
            rp=plan_rows[ds];ra=adjust_rows[ds]
            for t in range(T):
                # 注意：官方模板时间标签存在偏移歧义；保持官方表头不改，数值按模型slot顺序填入。
                plan.cell(rp,2+t).value=float(original[t])
                adjust.cell(ra,2+t).value=float(x[t,0])
            plan.cell(rp,146).value=float(original.sum())
            plan.cell(rp,147).value=float(summary['plan_cost'])
            adjust.cell(ra,146).value=float(x[:,0].sum())
            adjust.cell(ra,147).value=float(summary['total_actual_cost'])

        # 充放电与紧急购电为变长表：保留官方第一行表头，重写数据区。
        if bat.max_row>1: bat.delete_rows(2,bat.max_row-1)
        if emg.max_row>1: emg.delete_rows(2,emg.max_row-1)
        for (ds,result),summary in zip(results.items(),summaries):
            x=result['trace']
            for b in range(6):
                z=x[b*24:(b+1)*24]
                bat.append([ds if b==0 else None,f'{clock(b*240)}-{clock((b+1)*240)}',
                            float(z[:,2].sum()),float(z[:,3].sum()),
                            '0:00' if b==0 else ('24:00' if b==1 else None),
                            float(x[0,5]) if b==0 else (float(x[-1,6]) if b==1 else None)])
            for j,(label,amount) in enumerate(emergency_intervals(x) or [('无',0.)]):
                emg.append([ds if j==0 else None,label,amount])

        # 不重写官方表头；做硬审计。
        require([c.value for c in plan[1]]==original_plan_header,'计划购电量官方表头被修改')
        require([c.value for c in adjust[1]]==original_adjust_header,'调整购电量官方表头被修改')
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value,date): cell.number_format='yyyy/mm/dd'
                    elif isinstance(cell.value,(float,int)): cell.number_format='0.0000'

        target=folder/('result3.xlsx' if full else 'result3_diagnostic.xlsx')
        tmp=target.with_name(target.stem+'.tmp.xlsx')
        wb.save(tmp);wb.close();tmp.replace(target)

        # 独立回读核对本次实际写入日期、电量和费用。
        check=load_workbook(target,read_only=True,data_only=True)
        try:
            require(check.sheetnames==['计划购电量','调整购电量','充放电量','紧急购电量'],'导出表名错误')
            rows=list(check['调整购电量'].values)
            row_map={as_date(row[0]):row for row in rows[1:] if row[0] not in (None,'')}
            for summ in summaries:
                ds=date.fromisoformat(summ['date']);row=row_map[ds]
                vals=row[1:145]
                require(all(v is not None for v in vals),f'{ds} 调整购电量存在空值')
                require(abs(sum(map(float,vals))-summ['executed_contract_kwh'])<1e-5,'导出电量不一致')
                require(abs(float(row[146])-summ['total_actual_cost'])<1e-5,'导出账单不一致')
        finally:
            check.close()

    report=dict(
        scenario=name,mode=args.mode,complete_evaluation_period=full,
        initial_soc=summaries[0]['soc_start'],terminal_soc=summaries[-1]['soc_end'],
        soc_min_kwh=min(s['soc_min_kwh'] for s in summaries),
        soc_max_kwh=max(s['soc_max_kwh'] for s in summaries),
        mean_daily_soc_range_kwh=float(np.mean([s['soc_range_kwh'] for s in summaries])),
        total_actual_cost=sum(s['total_actual_cost'] for s in summaries),
        total_emergency_kwh=sum(s['emergency_kwh'] for s in summaries),
        total_adjustment_abs_kwh=sum(s['adjustment_abs_kwh'] for s in summaries),
        adjustment_nodes=sum(s['adjustment_nodes'] for s in summaries),
        settlement=args.settlement,
        template_headers_preserved=bool(template is not None),
        template_source=str(template.resolve()) if template is not None else None,
        lexicographic=args.lexicographic,lex_cost_eps=LEX_COST_EPS,
        model_role='risk_extension' if name.startswith('G') else ('baseline' if name=='B1' else 'main_D'),
        historical_candidate_check=name.startswith('G'),history_days=args.history_days,
        terminal_weight=args.terminal_weight,
        cross_day_forecast_used=False,
        workbook=str(target.resolve()) if target is not None else None,
        policy=asdict(Policy()))
    write_json(folder/'run_summary.json',report)
    return report


def export_preparation(data, folder, args):
    folder.mkdir(parents=True,exist_ok=True)
    mapping=[]
    for t in range(T):
        mapping.append(dict(
            slot=t,
            raw_sample_time=raw_sample_label(t),
            input_right_endpoint=raw_sample_label(t),
            interval_start=interval_start_label(t),
            interval_end=interval_end_label(t),
            physical_interval=physical_label(t),
            official_template_label=data.template_labels[t] if data.template_labels else '',
            note='右端点离散：原始采样时刻对应前一10分钟物理区间；正式xlsx保留官方标签，模型数值按slot顺序填入'))
    write_csv(folder/'time_mapping.csv',list(mapping[0]),mapping)

    notes = (
        '问题3最终实现口径\n'
        '1. 主模型是D系列；B1是无日内合同更新基准。G只作为预测误差风险增强扩展，不混入主模型。\n'
        '2. 输入功率统一乘DT=1/6换算为10分钟电量kWh；所有购电、SOC递推和费用结算在kWh口径下进行。\n'
        '3. 储能物理约束：1200~10800 kWh，最大充放电功率5000 kW，效率0.9，充放互斥。\n'
        '4. 仅2025-01-01 0:00初始储电量为6000 kWh；以后SOC跨日连续，不每日重置。\n'
        '5. refund主口径：合同费用=原计划费用+1.5倍调增费用-0.5倍调减退款；紧急购电另按5倍时价计。\n'
        '6. D方案在更新节点比较“保持当前合同”和“允许调整合同”的预测目标；仅允许调整更优时修改未来合同。\n'
        '7. 已执行时段不可追改；合同优化不读取未来真实负荷/PV。当前已结束区间观测只用于SOC/锚点和实时执行。\n'
        '8. 0/6/12/18发布的未来24小时PV预报仅截取到当日24:00；没有实现跨日滚动前瞻。\n'
        '9. 负荷未来值采用历史可解释预测；PV使用附件3对应发布时刻的官方预报。\n'
        '10. 日末6000仅作为滚动时域库存软价值参考，不是硬等式，也不是题目外的调整启动费。\n'
        '11. 两级求解仅用于同主目标最优解中的物理tie-break：在不恶化主目标的数值容差内优先吸收当前可用富余电量。\n'
        '12. 正式result3.xlsx从官方模板副本写出：保留四个sheet与原时间表头，只填数值；不改输入模板文件。\n'
        '13. 采用右端点离散：00:10→00:00-00:10，00:20→00:10-00:20，0:00+1→23:50-24:00；自生成论文表/图/CSV统一使用物理区间标签。time_mapping.csv同时保留原始采样时刻、区间起止和官方模板标签。\n'
        '14. comparison_main_D.csv只放B1与D主模型；comparison_risk_G.csv单独放G扩展，避免主模型和风险增强策略混淆。\n'
        '15. annual先共同运行1月预热，再从相同2月1日SOC对各方案独立评价；正式评价期为2月1日~12月31日334天。\n'
    )
    (folder/'model_notes.txt').write_text(
        notes+f'\n当前结算：{args.settlement}\n'
              f'历史窗口（仅G）：{args.history_days}天；末端软价值权重：{args.terminal_weight}\n'
              f'两级求解：{args.lexicographic}；主目标容差：{LEX_COST_EPS}元\n',
        encoding='utf-8')
    write_json(folder/'data_audit.json',dict(
        days=len(data.dates),shape=list(data.load.shape),issues=len(data.forecasts),
        forecast_values=24*len(data.forecasts),price_min=float(data.price.min()),
        price_max=float(data.price.max()),pv_surplus_slots=int((data.pv>data.load).sum()),
        max_surplus_kw=float((data.pv-data.load).max()/DT),
        template=str(data.files.get('result3_template')) if data.files.get('result3_template') else None))


def validate_full(data,args):
    # 右端点时间映射硬审计：只统一标签，不移动任何输入数组。
    require(raw_sample_label(0)=='00:10' and physical_label(0)=='00:00-00:10',
            '首个10分钟右端点映射错误')
    require(raw_sample_label(1)=='00:20' and physical_label(1)=='00:10-00:20',
            '第二个10分钟右端点映射错误')
    require(raw_sample_label(T-1)=='0:00+1' and physical_label(T-1)=='23:50-24:00',
            '日末右端点映射错误')
    # 可吸收富余量的小型物理例子：有容量、有功率余量，不应任意弃用。
    diag={}
    parts=optimize(np.array([-100.,200.]),np.array([1.,1.]),EMIN,Policy(),0.,
                   fixed_q=np.array([0.,200.]),time_limit=args.time_limit,
                   lexicographic=True,diagnostics=diag)
    require(parts[3][0]<TOL and abs(parts[1][0]-100)<TOL,'富余量吸收自检失败')
    require(diag['primary_objective_change']<=LEX_COST_EPS+1e-5,'费用优先级自检失败')
    # 电池满时允许弃用，不能为了w[0]=0突破SOC上限。
    parts=optimize(np.array([-100.]),np.array([1.]),EMAX,Policy(),0.,
                   fixed_q=np.array([0.]),time_limit=args.time_limit,lexicographic=True)
    require(abs(parts[3][0]-100)<TOL,'满电弃用边界检查失败')
    for rule,expectations in [('refund',[100.,90.,130.]),('no_refund',[100.,110.,130.])]:
        for final,expected in zip([100.,80.,120.],expectations):
            result=settlement_bill(np.array([100.]),np.array([final]),np.array([0.]),np.array([1.]),rule)
            require(abs(result['total_actual_cost']-expected)<1e-9,'结算分支错误')
    # original100/net80 => refund最优调减到80；原量100/net120 =>调增到120。
    for net,expected in [(80.,80.),(120.,120.)]:
        q,value=adjustment_problem(np.array([net]),np.array([1.]),EMIN,np.array([100.]),
                                   Policy(),0.,'refund',args.time_limit)
        require(abs(q[0]-expected)<TOL,'单区间调整优化错误')
    q,value=adjustment_problem(np.array([80.]),np.array([1.]),EMIN,np.array([100.]),
                               Policy(),0.,'no_refund',args.time_limit)
    require(abs(q[0]-100)<TOL,'不退款口径下调减支配性检查失败')
    # 0点计划不能因当天及未来真实曲线、6点及以后版本修改而改变。
    ds=date(2025,3,20);i=(ds-data.dates[0]).days
    q1,_,_=midnight_plan(data,ds,EINITIAL,'D6_12_18',args)
    altered=Data(data.dates,data.load.copy(),data.pv.copy(),data.price.copy(),
                 {k:v.copy() for k,v in data.forecasts.items()},data.template_labels,data.files)
    altered.load[i:]*=2;altered.pv[i:]=0
    for k in altered.forecasts:
        if k>=datetime.combine(ds,dtime(6)):altered.forecasts[k]*=2
    q2,_,_=midnight_plan(altered,ds,EINITIAL,'D6_12_18',args)
    require(np.allclose(q1,q2,atol=TOL),'0点合同读取未来信息')
    for hour in (6,12,18):
        dates1,r1=historical_net_residuals(data,ds,hour,args.history_days)
        dates2,r2=historical_net_residuals(altered,ds,hour,args.history_days)
        require(dates1==dates2 and all(d<ds for d in dates1),'历史检验使用了非法日期')
        require(len(r1)==len(r2) and all(np.array_equal(x,y) for x,y in zip(r1,r2)),
                '当天或未来数据改变了历史残差')
    issue=datetime.combine(ds,dtime(6))
    try:get_pv_forecast(issue,issue-timedelta(hours=1),data.forecasts[issue],0.)
    except ValueError:pass
    else:raise AssertionError('错误接受未来预报')
    print('自检通过：富余量优先吸收、满电边界、结算、合同优化、历史残差与0点信息隔离、发布时序。',flush=True)


def export_matched_errors(data, dates, folder):
    """仅事后诊断，同一天同目标整点比较，不将不同白天覆盖率的误差直接排名。"""
    rows=[]
    for ds in dates:
        index=(ds-data.dates[0]).days
        for old_hour,new_hour in [(0,6),(6,12),(12,18)]:
            old=data.forecasts[datetime.combine(ds,dtime(old_hour))]
            new=data.forecasts[datetime.combine(ds,dtime(new_hour))]
            target_hours=np.arange(new_hour+1,25)
            actual=data.pv[index,target_hours*6-1]/DT
            old_values=old[target_hours-old_hour-1]
            new_values=new[target_hours-new_hour-1]
            oe=old_values-actual;ne=new_values-actual
            rows.append(dict(date=str(ds),old_issue_hour=old_hour,new_issue_hour=new_hour,
                             first_target_hour=int(target_hours[0]),last_target_hour=24,
                             common_target_count=len(target_hours),
                             old_mae_kw=float(np.abs(oe).mean()),new_mae_kw=float(np.abs(ne).mean()),
                             old_rmse_kw=float(np.sqrt(np.mean(oe**2))),new_rmse_kw=float(np.sqrt(np.mean(ne**2))),
                             old_bias_kw=float(oe.mean()),new_bias_kw=float(ne.mean()),
                             ex_post_only=True))
    if rows:write_csv(folder/'matched_forecast_errors.csv',list(rows[0]),rows)


def _resolve_scenarios(spec):
    all_d=['B1']+['D'+'_'.join(map(str,h)) for n in (1,2,3) for h in itertools.combinations((6,12,18),n)]
    all_g=['G'+'_'.join(map(str,h)) for n in (1,2,3) for h in itertools.combinations((6,12,18),n)]
    presets={
        'core':['B1','D6','D6_12','D6_12_18'],
        'all-d':all_d,
        'risk':['B1','G6','G6_12','G6_12_18'],
        'all':all_d+all_g,
    }
    if spec in presets:
        return presets[spec]
    return [s.strip() for s in spec.split(',') if s.strip()]


def _comparison_row(name, summaries, start_soc, lam):
    row=dict(scenario=name,days=len(summaries),model_role=(
        'baseline' if name=='B1' else ('risk_extension' if name.startswith('G') else 'main_D')),
        initial_soc=start_soc,terminal_soc=summaries[-1]['soc_end'])
    for key in ['total_actual_cost','plan_cost','adjustment_net_cost','emergency_cost',
                'emergency_kwh','increase_kwh','decrease_kwh','adjustment_abs_kwh',
                'unused_kwh','adjustment_nodes']:
        row[key]=sum(s[key] for s in summaries)
    row['emergency_days']=sum(s['emergency_kwh']>1e-6 for s in summaries)
    row['soc_min_kwh']=min(s['soc_min_kwh'] for s in summaries)
    row['soc_max_kwh']=max(s['soc_max_kwh'] for s in summaries)
    row['mean_daily_soc_range_kwh']=float(np.mean([s['soc_range_kwh'] for s in summaries]))
    row['inventory_adjusted_cost']=row['total_actual_cost']-lam*(row['terminal_soc']-start_soc)
    return row


# =====================================================================
# 论文图：仅读取已完成的 annual/all-d 结果，不重新求解 MILP
# =====================================================================
Q3_COLORS = {
    'blue': '#104E8B', 'blue2': '#376B9E', 'blue3': '#5F89B1',
    'blue4': '#AFC3D8', 'mint': '#C5E9E3', 'grid': '#D7E1EB',
    'pink': '#F2DADA', 'rose': '#E5B5B5', 'salmon': '#D89090', 'red': '#B22222'
}
Q3_SCENARIO_ORDER = ['B1','D6','D12','D18','D6_12','D6_18','D12_18','D6_12_18']


def _read_csv_dict(path):
    path=Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    with path.open('r',encoding='utf-8-sig',newline='') as f:
        return list(csv.DictReader(f))


def _f(row,key,default=0.0):
    v=row.get(key,default)
    if v in (None,''):
        return float(default)
    return float(v)


def _configure_q3_font(plt):
    """Windows优先使用微软雅黑；找不到时按常见中文字体回退。"""
    try:
        from matplotlib import font_manager
        installed={f.name for f in font_manager.fontManager.ttflist}
        for name in ['Microsoft YaHei','SimHei','Noto Sans CJK SC','Source Han Sans SC',
                     'PingFang SC','Arial Unicode MS']:
            if name in installed:
                plt.rcParams['font.sans-serif']=[name]
                break
    except Exception:
        pass
    plt.rcParams['axes.unicode_minus']=False
    plt.rcParams['savefig.dpi']=320
    plt.rcParams['figure.dpi']=130
    plt.rcParams['font.size']=10.5


def _save_q3_figure(fig, outdir, stem, plt):
    outdir=Path(outdir); outdir.mkdir(parents=True,exist_ok=True)
    fig.savefig(outdir/f'{stem}.svg',format='svg',bbox_inches='tight',facecolor='white')
    plt.close(fig)


def _annual_result_root(args):
    if args.result_root is not None:
        return Path(args.result_root)
    return (Path(args.output_dir)/'annual'/args.settlement/
            ('lexicographic' if args.lexicographic else 'primary_only')/
            f'terminal{args.terminal_weight:g}')


def _main_comparison_rows(result_root):
    rows=_read_csv_dict(Path(result_root)/'comparison_main_D.csv')
    mp={r['scenario']:r for r in rows}
    missing=[s for s in Q3_SCENARIO_ORDER if s not in mp]
    if missing:
        raise ValueError(f'comparison_main_D.csv缺少all-d方案：{missing}；请先完整运行 --mode annual --scenarios all-d --resume')
    return [mp[s] for s in Q3_SCENARIO_ORDER]


def _auto_typical_date(result_root, scenario):
    """取该方案相对B1日节省额最接近正节省中位数的日期，避免只挑极端最好日。"""
    rows=_read_csv_dict(Path(result_root)/'comparison_by_day.csv')
    rr=[r for r in rows if r['scenario']==scenario]
    if not rr:
        raise ValueError(f'comparison_by_day.csv中没有{scenario}')
    positive=[r for r in rr if _f(r,'saving_vs_B1')>0]
    pool=positive or rr
    vals=np.array([_f(r,'saving_vs_B1') for r in pool],float)
    med=float(np.median(vals))
    return min(pool,key=lambda r:abs(_f(r,'saving_vs_B1')-med))['date']


def _scenario_detail(result_root, scenario):
    folder=Path(result_root)/scenario
    return (_read_csv_dict(folder/'daily_summary.csv'),
            _read_csv_dict(folder/'dispatch_detail.csv'),
            _read_csv_dict(folder/'update_log.csv') if (folder/'update_log.csv').exists() else [],
            _read_csv_dict(folder/'contract_versions.csv') if (folder/'contract_versions.csv').exists() else [])


def _plot_total_cost(comparison, outdir, plt):
    c=Q3_COLORS
    names=[r['scenario'] for r in comparison]
    vals=np.array([_f(r,'total_actual_cost') for r in comparison])/10000
    colors=[c['blue4'],c['blue3'],c['blue2'],c['mint'],c['blue2'],c['blue3'],c['mint'],c['blue']]
    fig,ax=plt.subplots(figsize=(11.6,5.4))
    bars=ax.bar(names,vals,color=colors,edgecolor='white',linewidth=.8)
    best=int(np.argmin(vals)); bars[best].set_edgecolor(c['red']); bars[best].set_linewidth(2.0)
    for i,(b,v) in enumerate(zip(bars,vals)):
        ax.text(b.get_x()+b.get_width()/2,v+max(vals)*.008,f'{v:.1f}',ha='center',va='bottom',fontsize=9,
                color=c['red'] if i==best else c['blue'])
    ax.set_ylabel('全年总费用（万元）'); ax.set_xlabel('预测更新方案')
    ax.set_title('不同预测更新方案全年实际费用对比')
    ax.grid(axis='y',color=c['grid'],alpha=.8); ax.set_axisbelow(True)
    ax.spines[['top','right']].set_visible(False)
    _save_q3_figure(fig,outdir,'Fig_01_方案全年费用对比',plt)


def _plot_cost_stack(comparison, outdir, plt):
    c=Q3_COLORS; names=[r['scenario'] for r in comparison]
    plan=np.array([_f(r,'plan_cost') for r in comparison])/10000
    adj=np.array([_f(r,'adjustment_net_cost') for r in comparison])/10000
    emergency=np.array([_f(r,'emergency_cost') for r in comparison])/10000
    x=np.arange(len(names))
    fig,ax=plt.subplots(figsize=(11.8,5.7))
    ax.bar(x,plan,color=c['blue2'],label='0点计划购电费')
    ax.bar(x,adj,bottom=plan,color=c['mint'],label='日内调整净费用')
    ax.bar(x,emergency,bottom=plan+adj,color=c['red'],label='紧急购电费')
    totals=plan+adj+emergency
    for i,v in enumerate(totals): ax.text(i,v+max(totals)*.008,f'{v:.1f}',ha='center',fontsize=8.6)
    ax.set_xticks(x,names); ax.set_ylabel('费用（万元）')
    ax.set_title('全年费用组成分解',pad=14)
    ax.legend(frameon=False,ncol=3,loc='upper center',bbox_to_anchor=(.5,-.10))
    ax.grid(axis='y',color=c['grid'],alpha=.8); ax.set_axisbelow(True); ax.spines[['top','right']].set_visible(False)
    fig.subplots_adjust(bottom=.20,top=.90)
    _save_q3_figure(fig,outdir,'Fig_02_费用组成堆叠图',plt)


def _plot_information_waterfall_3d(comparison, outdir, plt):
    """B1→D6→D6_12→D6_12_18 的三维增量收益瀑布。"""
    c=Q3_COLORS; mp={r['scenario']:r for r in comparison}
    costs={k:_f(mp[k],'total_actual_cost')/10000 for k in ['B1','D6','D6_12','D6_12_18']}
    inc=np.array([costs['B1']-costs['D6'], costs['D6']-costs['D6_12'], costs['D6_12']-costs['D6_12_18']])
    cum=np.r_[0,np.cumsum(inc)]
    from mpl_toolkits.mplot3d import Axes3D  # noqa: F401
    fig=plt.figure(figsize=(10.5,6.2)); ax=fig.add_subplot(111,projection='3d')
    xs=np.arange(3,dtype=float); ys=np.zeros(3); zs=cum[:-1]; dx=np.full(3,.62); dy=np.full(3,.72)
    colors=[c['blue3'],c['blue2'],c['red']]
    ax.bar3d(xs,ys,zs,dx,dy,inc,color=colors,shade=True,alpha=.93,edgecolor='white')
    ax.plot(xs+dx/2,np.full(3,.36),cum[1:],color=c['blue'],lw=1.8,marker='o')
    labels=['加入6:00','再加入12:00','再加入18:00']
    ax.set_xticks(xs+dx/2,labels); ax.set_yticks([]); ax.set_zlabel('累计节省费用（万元）')
    ax.set_title('预测信息价值三维瀑布图（B1 → D6 → D6_12 → D6_12_18）',pad=18)
    for i,v in enumerate(inc):
        ax.text(xs[i]+dx[i]/2,.38,zs[i]+v/2,f'+{v:.2f}',ha='center',va='center',fontsize=9)
    ax.view_init(elev=24,azim=-56); ax.grid(False)
    _save_q3_figure(fig,outdir,'Fig_03_预测信息价值三维瀑布图',plt)


def _plot_marginal_value(comparison, outdir, plt):
    c=Q3_COLORS; mp={r['scenario']:r for r in comparison}; b=_f(mp['B1'],'total_actual_cost')
    labels=['仅6:00','仅12:00','仅18:00','6:00增量','12:00增量\n(已有6:00)','18:00增量\n(已有6+12)']
    vals=[b-_f(mp['D6'],'total_actual_cost'), b-_f(mp['D12'],'total_actual_cost'), b-_f(mp['D18'],'total_actual_cost'),
          b-_f(mp['D6'],'total_actual_cost'), _f(mp['D6'],'total_actual_cost')-_f(mp['D6_12'],'total_actual_cost'),
          _f(mp['D6_12'],'total_actual_cost')-_f(mp['D6_12_18'],'total_actual_cost')]
    vals=np.array(vals)/10000
    colors=[c['blue3'],c['blue'],c['mint'],c['blue3'],c['blue2'],c['red']]
    fig,ax=plt.subplots(figsize=(11.5,6.0)); bars=ax.bar(np.arange(6),vals,color=colors)
    ax.axvline(2.5,color=c['grid'],lw=2)
    vmax=float(max(vals))
    ax.set_ylim(0,vmax*1.25)
    for b0,v in zip(bars,vals):
        ax.text(b0.get_x()+b0.get_width()/2,v+vmax*.018,f'{v:.2f}',ha='center',va='bottom',fontsize=9)
    ax.set_xticks(np.arange(6),labels); ax.set_ylabel('节省费用（万元）')
    ax.set_title('预测时刻的信息价值与边际收益',pad=18)
    # 两组说明放在坐标轴顶部空白区，避免遮挡柱顶数值。
    ax.text(.25,.955,'单时刻相对B1',transform=ax.transAxes,ha='center',va='top',color=c['blue'],fontsize=10.5)
    ax.text(.75,.955,'嵌套滚动方案的边际增益',transform=ax.transAxes,ha='center',va='top',color=c['blue'],fontsize=10.5)
    ax.grid(axis='y',color=c['grid'],alpha=.8); ax.set_axisbelow(True); ax.spines[['top','right']].set_visible(False)
    fig.subplots_adjust(top=.88,bottom=.17)
    _save_q3_figure(fig,outdir,'Fig_04_预测时刻边际收益',plt)


def _plot_emergency_compare(comparison, outdir, plt):
    c=Q3_COLORS; names=[r['scenario'] for r in comparison]
    vals=np.array([_f(r,'emergency_kwh') for r in comparison])/1000
    fig,ax=plt.subplots(figsize=(11.5,5.2)); bars=ax.bar(names,vals,color=c['salmon'])
    best=int(np.argmin(vals)); bars[best].set_color(c['blue'])
    for i,(b,v) in enumerate(zip(bars,vals)):
        ax.text(b.get_x()+b.get_width()/2,v+max(vals)*.012,f'{v:.1f}',ha='center',fontsize=8.8,
                color=c['blue'] if i==best else '#333333')
    ax.set_ylabel('紧急购电量（MWh）'); ax.set_xlabel('方案'); ax.set_title('日内预测更新对紧急购电风险的抑制效果')
    ax.grid(axis='y',color=c['grid'],alpha=.8); ax.set_axisbelow(True); ax.spines[['top','right']].set_visible(False)
    _save_q3_figure(fig,outdir,'Fig_05_紧急购电量比较',plt)


def _plot_timeline(outdir, plt):
    c=Q3_COLORS
    from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
    fig,ax=plt.subplots(figsize=(13.5,6.1)); ax.set_xlim(-.6,24.8); ax.set_ylim(-2.65,3.2); ax.axis('off')
    y=1.3; ax.plot([0,24],[y,y],color=c['blue'],lw=2.2)
    points=[(0,'0:00\n初始计划'),(6,'6:00\n更新'),(12,'12:00\n更新'),(18,'18:00\n更新'),(24,'24:00\n日终')]
    for h,label in points:
        ax.scatter([h],[y],s=110,color=c['blue'] if h in (0,24) else c['red'],zorder=3)
        ax.text(h,y+.45,label,ha='center',va='bottom',fontsize=10,color=c['blue'])
    boxes=[(0,'读取0点PV预测\n+历史负荷预测\n优化144时段计划'),
           (6,'冻结0–6点\n读取6点新PV预测\n比较保持/调整'),
           (12,'冻结已执行时段\n读取12点新PV预测\n比较保持/调整'),
           (18,'冻结已执行时段\n读取18点新PV预测\n比较保持/调整')]
    box_w=3.4; box_h=1.42; yy=-1.55
    for h,txt in boxes:
        # 0点框向右展开；其余框居中于更新节点。缩窄框宽，避免相邻框视觉重叠。
        x=-.10 if h==0 else h-box_w/2
        rect=FancyBboxPatch((x,yy),box_w,box_h,boxstyle='round,pad=.05,rounding_size=.08',
                            facecolor=c['mint'] if h==0 else c['pink'],edgecolor=c['blue2'],linewidth=1.2)
        ax.add_patch(rect); ax.text(x+box_w/2,yy+box_h/2,txt,ha='center',va='center',fontsize=8.8,color=c['blue'])
        ax.add_patch(FancyArrowPatch((h,y-.08),(h,yy+box_h+.03),arrowstyle='-|>',mutation_scale=12,color=c['blue3'],lw=1))
    ax.text(12,-2.28,'决策原则：只使用当前已发布预测；已执行时段锁定；若 C_adjust < C_stay 才修改未来合同',
            ha='center',fontsize=10.2,color=c['red'])
    ax.set_title('0/6/12/18多时刻滚动优化机制',fontsize=15,color=c['blue'],pad=12)
    _save_q3_figure(fig,outdir,'Fig_06_滚动优化机制时间轴',plt)


def _plot_typical_soc_and_balance(result_root, scenario, day, outdir, plt):
    c=Q3_COLORS; _,detail,_,_=_scenario_detail(result_root,scenario)
    rr=[r for r in detail if r['date']==day]
    if len(rr)!=T: raise ValueError(f'{scenario}/{day} dispatch_detail应有144行，实际{len(rr)}')
    rr.sort(key=lambda r:int(float(r['slot'])))
    soc=np.r_[_f(rr[0],'soc_start_kwh'),[_f(r,'soc_end_kwh') for r in rr]]
    x_soc=np.arange(T+1)*DT
    fig,ax=plt.subplots(figsize=(11.5,5.0)); ax.plot(x_soc,soc,color=c['blue'],lw=1.8,label='SOC')
    ax.fill_between(x_soc,soc,EMIN,color=c['blue4'],alpha=.25)
    ax.axhline(EMIN,color=c['red'],ls='--',lw=1,label='SOC下限 1200 kWh'); ax.axhline(EMAX,color=c['red'],ls='--',lw=1)
    for h in [6,12,18]: ax.axvline(h,color=c['grid'],ls=':',lw=1)
    ax.set(xlim=(0,24),xlabel='时刻（h）',ylabel='储能电量 SOC（kWh）',title=f'典型日SOC轨迹：{day}')
    ax.set_xticks([0,4,6,8,12,16,18,20,24]); ax.grid(color=c['grid'],alpha=.65); ax.legend(frameon=False)
    ax.spines[['top','right']].set_visible(False)
    _save_q3_figure(fig,outdir,'Fig_07_典型日SOC曲线',plt)

    t=(np.arange(T)+.5)*DT
    load=np.array([_f(r,'actual_load_kwh') for r in rr])/DT
    pv=np.array([_f(r,'actual_pv_kwh') for r in rr])/DT
    plan=np.array([_f(r,'original_plan_kwh') for r in rr])/DT
    contract=np.array([_f(r,'executed_contract_kwh') for r in rr])/DT
    emergency=np.array([_f(r,'emergency_kwh') for r in rr])/DT
    charge=np.array([_f(r,'charge_kwh') for r in rr])/DT
    discharge=np.array([_f(r,'discharge_kwh') for r in rr])/DT
    unused=np.array([_f(r,'unused_kwh') for r in rr])/DT
    net_supply=contract+pv+emergency+discharge-charge-unused
    fig,ax=plt.subplots(figsize=(12.4,5.8))
    ax.plot(t,load,color=c['blue'],lw=2,label='实际负荷')
    ax.plot(t,net_supply,color=c['red'],lw=1.2,ls='--',label='净供给（平衡校验）')
    ax.step(t,pv,where='mid',color=c['mint'],lw=1.2,label='实际光伏')
    ax.step(t,plan,where='mid',color=c['blue4'],lw=1,ls=':',label='0点原计划购电')
    ax.step(t,contract,where='mid',color=c['blue2'],lw=1.3,label='调整后合同购电')
    ax.step(t,emergency,where='mid',color=c['red'],lw=1.0,alpha=.75,label='紧急购电')
    ax.step(t,discharge-charge,where='mid',color=c['salmon'],lw=1.1,label='电池净输出(+放/-充)')
    for h in [6,12,18]: ax.axvline(h,color=c['grid'],ls=':',lw=1)
    ax.set(xlim=(0,24),xlabel='时刻（h）',ylabel='功率（kW）',title=f'典型日供需平衡与合同滚动调整：{day}')
    ax.set_xticks([0,4,6,8,12,16,18,20,24]); ax.grid(color=c['grid'],alpha=.55)
    ax.legend(frameon=False,ncol=4,loc='upper center',bbox_to_anchor=(.5,-.13))
    ax.spines[['top','right']].set_visible(False)
    fig.subplots_adjust(top=.90,bottom=.22)
    _save_q3_figure(fig,outdir,'Fig_08_典型日供需平衡',plt)


def _plot_monthly_adjustments(result_root, scenario, outdir, plt):
    c=Q3_COLORS; _,_,updates,_=_scenario_detail(result_root,scenario)
    if not updates: raise ValueError(f'{scenario}没有update_log.csv')
    months=sorted({r['date'][:7] for r in updates}); hours=[6,12,18]; vals={h:[] for h in hours}
    for m in months:
        for h in hours:
            rr=[r for r in updates if r['date'].startswith(m) and int(float(r['issue_hour']))==h]
            vals[h].append(sum(str(r['changed']).strip().lower() in {'true','1','yes'} for r in rr))
    x=np.arange(len(months)); width=.24
    fig,ax=plt.subplots(figsize=(12,5.4))
    for k,(h,color) in enumerate(zip(hours,[c['blue3'],c['blue'],c['red']])):
        ax.bar(x+(k-1)*width,vals[h],width,color=color,label=f'{h}:00 实际调整')
    ax.set_xticks(x,[m[5:]+'月' for m in months]); ax.set_ylabel('调整次数（天）'); ax.set_xlabel('月份')
    ax.set_title('各预测时刻月度调整次数',pad=14)
    ax.legend(frameon=False,ncol=3,loc='upper center',bbox_to_anchor=(.5,-.12))
    ax.grid(axis='y',color=c['grid'],alpha=.7); ax.set_axisbelow(True)
    ax.spines[['top','right']].set_visible(False)
    fig.subplots_adjust(bottom=.20,top=.90)
    _save_q3_figure(fig,outdir,'Fig_09_调整次数月分布',plt)


def _join_update_error(result_root, scenario):
    _,_,updates,_=_scenario_detail(result_root,scenario)
    errors=_read_csv_dict(Path(result_root)/'matched_forecast_errors.csv')
    emap={(r['date'],int(float(r['new_issue_hour']))):r for r in errors}
    joined=[]
    for u in updates:
        h=int(float(u['issue_hour'])); e=emap.get((u['date'],h))
        if e is None: continue
        joined.append({**u,'new_mae_kw':_f(e,'new_mae_kw'),'old_mae_kw':_f(e,'old_mae_kw'),
                       'forecast_mae_improvement_kw':_f(e,'old_mae_kw')-_f(e,'new_mae_kw')})
    return joined


def _plot_adjustment_scatter(result_root, scenario, outdir, plt):
    c=Q3_COLORS; rows=_join_update_error(result_root,scenario)
    if not rows: raise ValueError('无法拼接update_log与matched_forecast_errors')
    fig,ax=plt.subplots(figsize=(9.2,6.2)); colors={6:c['blue3'],12:c['blue'],18:c['red']}
    for h in [6,12,18]:
        rr=[r for r in rows if int(float(r['issue_hour']))==h]
        x=np.array([_f(r,'new_mae_kw') for r in rr]); y=np.array([_f(r,'revision_l1_kwh') for r in rr])
        changed=np.array([str(r['changed']).lower() in {'true','1','yes'} for r in rr])
        ax.scatter(x[changed],y[changed],s=23,alpha=.62,color=colors[h],label=f'{h}:00 调整')
        if np.any(~changed): ax.scatter(x[~changed],y[~changed],s=35,facecolors='none',edgecolors=colors[h],alpha=.85,label=f'{h}:00 保持')
    ax.set_xlabel('新发布光伏预测的事后 MAE（kW）'); ax.set_ylabel('合同调整电量 $\\|y-x\\|_1$（kWh）')
    ax.set_title('预测误差与合同调整幅度的二维关系')
    ax.grid(color=c['grid'],alpha=.65); ax.legend(frameon=False,ncol=2); ax.spines[['top','right']].set_visible(False)
    _save_q3_figure(fig,outdir,'Fig_10_调整行为二维散点图',plt)


def _plot_decision_region(result_root, scenario, outdir, plt):
    c=Q3_COLORS; rows=_join_update_error(result_root,scenario)
    x=np.array([_f(r,'forecast_mae_improvement_kw') for r in rows]); y=np.array([_f(r,'predicted_net_improvement') for r in rows])
    changed=np.array([str(r['changed']).lower() in {'true','1','yes'} for r in rows])
    fig,ax=plt.subplots(figsize=(9.5,6.2))
    xmin,xmax=float(np.min(x)),float(np.max(x)); ymin,ymax=float(np.min(y)),float(np.max(y))
    pad_y=max(1,(ymax-ymin)*.08); pad_x=max(1,(xmax-xmin)*.05)
    ax.axhspan(0,ymax+pad_y,color=c['mint'],alpha=.30,label='预测目标改善 > 0：调整候选区')
    ax.axhspan(ymin-pad_y,0,color=c['pink'],alpha=.55,label='预测目标改善 ≤ 0：保持区')
    ax.scatter(x[changed],y[changed],s=24,color=c['blue'],alpha=.62,label='实际执行调整')
    if np.any(~changed): ax.scatter(x[~changed],y[~changed],s=38,facecolors='none',edgecolors=c['red'],label='实际保持')
    ax.axhline(0,color=c['red'],lw=1.2); ax.axvline(0,color=c['grid'],lw=1,ls='--')
    ax.set_xlim(xmin-pad_x,xmax+pad_x); ax.set_ylim(ymin-pad_y,ymax+pad_y)
    ax.set_xlabel('新预测相对旧预测的MAE改善（kW；正值=更准）'); ax.set_ylabel('调整相对保持的预测净收益（元）')
    ax.set_title('调整决策二维区域图：预测精度变化 × 调整经济性')
    ax.grid(color=c['grid'],alpha=.45)
    ax.legend(frameon=False,loc='upper left',bbox_to_anchor=(0.02,-0.20),borderaxespad=0.,ncol=1)
    ax.spines[['top','right']].set_visible(False)
    fig.subplots_adjust(bottom=.30)
    _save_q3_figure(fig,outdir,'Fig_11_调整决策二维区域图',plt)


def _plot_joint_heatmaps(result_root, scenario, outdir, plt):
    c=Q3_COLORS; _,detail,_,_=_scenario_detail(result_root,scenario)
    dates=sorted({r['date'] for r in detail}); by={d:[] for d in dates}
    for r in detail: by[r['date']].append(r)
    soc=[]; pv=[]; price=[]
    for d in dates:
        rr=sorted(by[d],key=lambda r:int(float(r['slot'])))
        if len(rr)!=T: raise ValueError(f'{d}逐时明细不完整')
        soc.append([_f(r,'soc_end_kwh') for r in rr]); pv.append([_f(r,'actual_pv_kwh')/DT for r in rr]); price.append([_f(r,'price') for r in rr])
    soc=np.asarray(soc); pv=np.asarray(pv); price=np.asarray(price)
    from matplotlib.colors import LinearSegmentedColormap
    blue_map=LinearSegmentedColormap.from_list('q3blue',[c['grid'],c['mint'],c['blue4'],c['blue3'],c['blue']])
    red_map=LinearSegmentedColormap.from_list('q3red',[c['pink'],c['rose'],c['salmon'],c['red']])
    fig,axes=plt.subplots(3,1,figsize=(13,10.4),sharex=True,sharey=True,gridspec_kw={'hspace':.14})
    mats=[soc,pv,price]; cmaps=[blue_map,blue_map,red_map]; labels=['SOC（kWh）','实际光伏功率（kW）','交易电价（元/kWh）']
    for ax,mat,cmap,label in zip(axes,mats,cmaps,labels):
        im=ax.imshow(mat,aspect='auto',origin='upper',cmap=cmap); cb=fig.colorbar(im,ax=ax,pad=.012); cb.set_label(label); ax.set_ylabel('日期')
    month_pos=[]; month_lab=[]; last=None
    for i,d in enumerate(dates):
        m=d[:7]
        if m!=last: month_pos.append(i); month_lab.append(d[5:7]+'月'); last=m
    for ax in axes: ax.set_yticks(month_pos,month_lab)
    axes[-1].set_xticks([0,24,48,72,96,120,143],['0','4','8','12','16','20','24']); axes[-1].set_xlabel('日内时刻（h）')
    fig.suptitle('储能SOC—实际光伏—电价联合热力图',fontsize=14,color=c['blue'],y=.985)
    _save_q3_figure(fig,outdir,'Fig_12_SOC光伏电价联合热力图',plt)


def _export_q3_figure_data(result_root, scenario, typical_date, outdir):
    comparison=_main_comparison_rows(result_root); rows=[]
    for r in comparison:
        rows.append({k:r.get(k,'') for k in ['scenario','total_actual_cost','plan_cost','adjustment_net_cost','emergency_cost',
                                              'emergency_kwh','adjustment_abs_kwh','adjustment_nodes','saving_vs_B1','saving_pct_vs_B1']})
    write_csv(Path(outdir)/'图表核心数据.csv',list(rows[0]),rows)
    Path(outdir,'典型日说明.txt').write_text(
        f'典型日={typical_date}\n方案={scenario}\n选择规则=相对B1的正节省日中，取日节省额最接近中位数的日期；避免挑极端最好日。\n',encoding='utf-8')


def generate_q3_figures(data, result_root, outdir, scenario='D6_12_18', figure_date='auto'):
    """生成论文图；只读既有全年结果，不重新求解、不改result3.xlsx。"""
    result_root=Path(result_root); outdir=Path(outdir)
    if not result_root.exists(): raise FileNotFoundError(f'找不到全年结果目录：{result_root}')
    if not (result_root/'comparison_main_D.csv').exists(): raise FileNotFoundError(f'{result_root} 中缺少 comparison_main_D.csv；请先跑完all-d')
    if scenario not in Q3_SCENARIO_ORDER or scenario=='B1': raise ValueError('figure-scenario建议使用D方案，默认D6_12_18')
    if not (result_root/scenario/'dispatch_detail.csv').exists(): raise FileNotFoundError(f'缺少{scenario}/dispatch_detail.csv')
    try:
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise RuntimeError('画图需要matplotlib，请先运行：python -m pip install matplotlib') from exc
    _configure_q3_font(plt); outdir.mkdir(parents=True,exist_ok=True)
    comparison=_main_comparison_rows(result_root)
    typical=_auto_typical_date(result_root,scenario) if str(figure_date).lower()=='auto' else str(figure_date)

    _plot_total_cost(comparison,outdir,plt)
    _plot_cost_stack(comparison,outdir,plt)
    _plot_information_waterfall_3d(comparison,outdir,plt)
    _plot_marginal_value(comparison,outdir,plt)
    _plot_emergency_compare(comparison,outdir,plt)
    _plot_timeline(outdir,plt)
    _plot_typical_soc_and_balance(result_root,scenario,typical,outdir,plt)
    _plot_monthly_adjustments(result_root,scenario,outdir,plt)
    _plot_adjustment_scatter(result_root,scenario,outdir,plt)
    _plot_decision_region(result_root,scenario,outdir,plt)
    _plot_joint_heatmaps(result_root,scenario,outdir,plt)
    _export_q3_figure_data(result_root,scenario,typical,outdir)

    index=(
        '论文图清单（仅 SVG 矢量）\n\n'
        '01 方案全年费用对比：回答哪种更新组合最经济。\n'
        '02 费用组成堆叠图：计划购电费 / 调整净费用 / 紧急购电费。\n'
        '03 预测信息价值三维瀑布图：B1→6点→12点→18点的增量收益。\n'
        '04 预测时刻边际收益：单时刻价值与嵌套滚动边际收益。\n'
        '05 紧急购电量比较：展示风险控制效果。\n'
        '06 滚动优化机制时间轴：解释0/6/12/18的锁定—更新—判断机制。\n'
        '07 典型日SOC曲线：验证储能状态物理合理性。\n'
        '08 典型日供需平衡：展示负荷、PV、合同、紧急购电、电池净输出与平衡校验。\n'
        '09 调整次数月分布：解释各预测时刻的月度调整行为。\n'
        '10 调整行为二维散点：新预测事后MAE与合同调整幅度关系。\n'
        '11 调整决策二维区域：预测精度变化与经济性判断区域。\n'
        '12 SOC—光伏—电价联合热力图：全年日内运行状态的时空结构。\n\n'
        '注意：图10/11使用真实值计算的预测误差只用于事后解释，不参与在线决策。\n'
    )
    (outdir/'论文图说明.txt').write_text(index,encoding='utf-8')
    print('='*72,flush=True)
    print('论文图生成完成（不重新求解MILP，不修改result3.xlsx）',flush=True)
    print(f'全年结果目录：{result_root.resolve()}',flush=True)
    print(f'绘图方案：{scenario}；典型日：{typical}',flush=True)
    print(f'图件目录：{outdir.resolve()}',flush=True)
    print('共生成12张 SVG 图。',flush=True)
    print('='*72,flush=True)

def main():
    parser=argparse.ArgumentParser(description='问题3：预测信息价值驱动的多时刻滚动调度')
    parser.add_argument('--data-dir',type=Path,default=DATA_DIR,
                        help='默认C/data；可放原始xlsx或data/processed清洗CSV')
    parser.add_argument('--output-dir',type=Path,default=OUTPUT_DIR)
    parser.add_argument('--mode',choices=['demo','annual','audit','figures'],default=RUN_MODE)
    parser.add_argument('--start-date',default=DEMO_START)
    parser.add_argument('--days',type=int,default=DEMO_DAYS)
    parser.add_argument('--initial-soc',type=float,default=EINITIAL)
    parser.add_argument('--scenarios',default=SCENARIOS,
                        help='core / all-d / risk / all / B1,D6,D6_12等自定义')
    parser.add_argument('--settlement',choices=['refund','no_refund'],default=SETTLEMENT)
    parser.add_argument('--time-limit',type=float,default=TIME_LIMIT)
    parser.add_argument('--resume',action=argparse.BooleanOptionalAction,default=RESUME)
    parser.add_argument('--self-test',action=argparse.BooleanOptionalAction,default=RUN_SELF_TEST)
    parser.add_argument('--lexicographic',action=argparse.BooleanOptionalAction,default=USE_LEXICOGRAPHIC)
    parser.add_argument('--history-days',type=int,default=HISTORY_DAYS)
    parser.add_argument('--terminal-weight',type=float,default=TERMINAL_WEIGHT)
    parser.add_argument('--result-root',type=Path,default=None,
                        help='figures模式读取的全年结果根目录；默认自动定位q3_output/annual/.../terminal1')
    parser.add_argument('--figure-dir',type=Path,default=BASE_DIR/'figures_q3',
                        help='Q3论文图输出目录，默认C/figures_q3')
    parser.add_argument('--figure-scenario',default='D6_12_18',
                        help='典型日/行为图使用的D方案，默认全年最优D6_12_18')
    parser.add_argument('--figure-date',default='auto',
                        help='典型日YYYY-MM-DD；默认auto按正节省日中位数自动选取')
    args=parser.parse_args()

    require(np.isfinite(args.time_limit) and args.time_limit>0,'时限必须为正')
    require(np.isfinite(args.initial_soc) and EMIN<=args.initial_soc<=EMAX,'诊断SOC越界')
    require(args.history_days>=MIN_HISTORY_DAYS,'历史窗口不得小于3个已完成日期')
    require(np.isfinite(args.terminal_weight) and args.terminal_weight>=0,'末端软参考权重须非负有限')
    if args.settlement!='refund':
        print('[warning] 当前不是论文主口径refund；no_refund只建议做解释敏感性。',flush=True)

    names=_resolve_scenarios(args.scenarios)
    require(names and len(names)==len(set(names)),'方案为空或重复')
    for name in names:
        scenario_config(name)
    if any(n.startswith('G') for n in names):
        print('提示：G为风险增强扩展；comparison_main_D.csv不会把G混入主模型。',flush=True)

    data=read_data(args.data_dir)
    if args.mode=='figures':
        generate_q3_figures(data,_annual_result_root(args),args.figure_dir,args.figure_scenario,args.figure_date)
        return
    base_root=args.output_dir/args.mode/args.settlement/('lexicographic' if args.lexicographic else 'primary_only')
    base_root=base_root/f'terminal{args.terminal_weight:g}'

    if args.mode=='annual':
        root=base_root
        dates=None
    elif args.mode=='demo':
        start=date.fromisoformat(args.start_date)
        require(args.days>=1 and start>=date(2025,1,2),'demo需从1月2日后开始，days>=1')
        dates=[start+timedelta(days=i) for i in range(args.days)]
        require(dates[-1]<=date(2025,12,31),'诊断日期超范围')
        root=base_root/f'{dates[0]}_{dates[-1]}'
    else:
        root=base_root
        dates=None
    root.mkdir(parents=True,exist_ok=True)
    export_preparation(data,root,args)

    print(f'数据检查通过：{len(data.dates)}天，{len(data.forecasts)}次预报。',flush=True)
    print('时间口径：模型按00:00~24:00自然日物理区间；正式xlsx保留官方模板表头不改。',flush=True)
    if args.self_test:
        validate_full(data,args)
    if args.mode=='audit':
        print('audit完成：数据、结算、信息边界与物理约束自检通过。',flush=True)
        return

    if args.mode=='annual':
        _,warm=run_sequence(data,data.dates[:31],EINITIAL,'B1',root/'common_warmup',args)
        start_soc=warm[-1]['soc_end']
        dates=data.dates[31:]
        print(f'共同预热完成：2月1日初始SOC={start_soc:.6f} kWh，各方案从相同状态独立运行。',flush=True)
    else:
        start_soc=args.initial_soc

    comparison=[]
    scenario_daily={}
    export_matched_errors(data,dates,root)
    lam=float(data.price.min()/ETA)

    for name in names:
        folder=root/name
        results,summaries=run_sequence(data,dates,start_soc,name,folder,args)
        scenario_daily[name]=summaries
        export_run(data,results,summaries,folder,args)
        comparison.append(_comparison_row(name,summaries,start_soc,lam))

    if 'B1' in scenario_daily:
        base={s['date']:s for s in scenario_daily['B1']}
        paired=[]
        base_total=sum(s['total_actual_cost'] for s in scenario_daily['B1'])
        for name,summaries in scenario_daily.items():
            deltas=[]
            for s in summaries:
                ref=base[s['date']]
                delta=ref['total_actual_cost']-s['total_actual_cost']
                deltas.append(delta)
                paired.append(dict(
                    date=s['date'],scenario=name,actual_cost=s['total_actual_cost'],
                    baseline_actual_cost=ref['total_actual_cost'],saving_vs_B1=delta,
                    soc_start=s['soc_start'],baseline_soc_start=ref['soc_start'],
                    soc_end=s['soc_end'],baseline_soc_end=ref['soc_end']))
            row=next(r for r in comparison if r['scenario']==name)
            row['saving_vs_B1']=base_total-row['total_actual_cost']
            row['saving_pct_vs_B1']=100*row['saving_vs_B1']/base_total if base_total else None
            row['cheaper_days_vs_B1']=sum(v>1e-4 for v in deltas)
            row['costlier_days_vs_B1']=sum(v<-1e-4 for v in deltas)
        write_csv(root/'comparison_by_day.csv',list(paired[0]),paired)
    else:
        for row in comparison:
            row['saving_vs_B1']=None;row['saving_pct_vs_B1']=None
            row['cheaper_days_vs_B1']=None;row['costlier_days_vs_B1']=None

    write_csv(root/'comparison_all.csv',list(comparison[0]),comparison)
    main_rows=[r for r in comparison if r['scenario']=='B1' or r['scenario'].startswith('D')]
    risk_rows=[r for r in comparison if r['scenario'].startswith('G')]

    if main_rows:
        write_csv(root/'comparison_main_D.csv',list(main_rows[0]),main_rows)
        write_csv(root/'comparison.csv',list(main_rows[0]),main_rows)
        best=min(main_rows,key=lambda r:r['total_actual_cost'])
        best_inventory=min(main_rows,key=lambda r:r['inventory_adjusted_cost'])
        best_info={
            'criterion':'actual_total_cost',
            'best_main_scenario':best['scenario'],
            'best_actual_total_cost':best['total_actual_cost'],
            'best_inventory_adjusted_scenario':best_inventory['scenario'],
            'best_inventory_adjusted_cost':best_inventory['inventory_adjusted_cost'],
            'note':'B1也参与主策略比较；若B1最低，说明当前数据与结算下日内合同更新未产生净经济收益。G不参与主模型最优方案选择。'
        }
        write_json(root/'best_main_scenario.json',best_info)
        if args.mode=='annual':
            chosen=root/best['scenario']/'result3.xlsx'
            if chosen.exists():
                shutil.copy2(chosen,root/'result3.xlsx')
                best_info['final_result3']=str((root/'result3.xlsx').resolve())
                write_json(root/'best_main_scenario.json',best_info)

    if risk_rows:
        risk_out=([next(r for r in comparison if r['scenario']=='B1')]
                  if any(r['scenario']=='B1' for r in comparison) else [])+risk_rows
        write_csv(root/'comparison_risk_G.csv',list(risk_out[0]),risk_out)

    print('\n主模型比较（B1 + D；G不混入）：',flush=True)
    for r in main_rows:
        print(f'{r["scenario"]:12s} 费用={r["total_actual_cost"]:.2f}元 '
              f'紧急={r["emergency_kwh"]:.2f}kWh 调整={r["adjustment_abs_kwh"]:.2f}kWh '
              f'调整节点={r["adjustment_nodes"]} 期末SOC={r["terminal_soc"]:.2f}',flush=True)
    if risk_rows:
        print('\nG风险增强扩展（单独评价）：',flush=True)
        for r in risk_rows:
            print(f'{r["scenario"]:12s} 费用={r["total_actual_cost"]:.2f}元 '
                  f'紧急={r["emergency_kwh"]:.2f}kWh 期末SOC={r["terminal_soc"]:.2f}',flush=True)
    if main_rows:
        best_name=min(main_rows,key=lambda r:r['total_actual_cost'])['scenario']
        print(f'主模型事后最低实际费用方案：{best_name}',flush=True)
    print(f'输出目录：{root.resolve()}',flush=True)


if __name__=='__main__':
    try:main()
    except (ValueError,KeyError,FileNotFoundError,RuntimeError,AssertionError,PermissionError) as exc:
        print(f'运行停止：{exc}\n若Excel占用输出文件请先关闭；若配置改变请换输出目录。',file=sys.stderr)
        sys.exit(1)

