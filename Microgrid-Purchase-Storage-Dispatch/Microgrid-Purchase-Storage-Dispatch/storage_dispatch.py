from __future__ import annotations

import contextlib
import csv
import io
import math
import re
import shutil
from dataclasses import dataclass
from datetime import time as dt_time, datetime
from pathlib import Path

import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from scipy.optimize import Bounds, LinearConstraint, milp
from scipy.sparse import lil_matrix


# 模块：储能运行与经济调度
# ============================================================
# 0. 配置区
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"

# 输入文件
INPUT_XLSX = DATA_DIR / "附件1.xlsx"
TEMPLATE_XLSX = DATA_DIR / "附件5" / "result1.xlsx"

OUTPUT_XLSX = BASE_DIR / "result1_Q1_final.xlsx"
DETAIL_CSV = BASE_DIR / "q1_detail_final.csv"
SUMMARY_TXT = BASE_DIR / "Q1_结果摘要.txt"
PAPER_RESULT_XLSX = BASE_DIR / "Q1_论文结果表.xlsx"

SOC_FIG = BASE_DIR / "Q1_SOC曲线.svg"
PRICE_FIG = BASE_DIR / "Q1_电价曲线.svg"
BATTERY_FIG = BASE_DIR / "Q1_充放电曲线.svg"
BALANCE_FIG = BASE_DIR / "Q1_负荷光伏购电曲线.svg"
SENSITIVITY_CSV = BASE_DIR / "Q1_敏感性分析.csv"
SENSITIVITY_ETA_FIG = BASE_DIR / "Q1_敏感性分析_效率.svg"
SENSITIVITY_PMAX_FIG = BASE_DIR / "Q1_敏感性分析_功率上限.svg"
CAPACITY_SENSITIVITY_CSV = BASE_DIR / "Q1_储能容量敏感性.csv"
CAPACITY_SENSITIVITY_FIG = BASE_DIR / "Q1_储能容量敏感性.svg"

# 新增：用于论文核心可视化的两张图
PRICE_BATTERY_FIG = BASE_DIR / "Q1_电价充放电联合响应.svg"
BASELINE_COST_FIG = BASE_DIR / "Q1_M0_M1_M2购电费用对比.svg"

# 新增：4张高价值分析图
GRID_SHIFT_FIG = BASE_DIR / "Q1_储能削峰填谷效果.svg"
CUMULATIVE_COST_FIG = BASE_DIR / "Q1_累计购电成本与节省.svg"
STATE_TIMELINE_FIG = BASE_DIR / "Q1_储能运行状态与活跃约束.svg"
PV_UTIL_FIG = BASE_DIR / "Q1_光伏消纳改善.svg"

# 10 min = 1/6 h
DT_H = 10.0 / 60.0
N_EXPECTED = 144

# 题目储能参数
E0_KWH = 6000.0
E_MIN_KWH = 1200.0
E_MAX_KWH = 10800.0
P_BAT_MAX_KW = 5000.0

# 题面“充放电效率90%”：默认解释为单向效率均为0.9
ETA_C = 0.9
ETA_D = 0.9

# MILP设置
MIP_REL_GAP = 1e-10
TIME_LIMIT_S = 60.0

# 数值审计容差
AUDIT_TOL = 1e-6

# 词典序约束容差
LEXI_COST_TOL_YUAN = 1e-6
LEXI_CURTAIL_TOL_KWH = 1e-6


# ============================================================
# 1. 数据结构
# ============================================================

@dataclass
class InputData:
    raw_time: list
    price: np.ndarray
    load_kw: np.ndarray
    pv_kw: np.ndarray


@dataclass
class Solution:
    grid_kw: np.ndarray
    charge_kw: np.ndarray
    discharge_kw: np.ndarray
    curtail_kw: np.ndarray
    soc_kwh: np.ndarray
    mode_binary: np.ndarray

    total_cost_yuan: float
    stage1_cost_yuan: float
    stage2_curtail_kwh: float
    stage3_throughput_kwh: float
    solver_message: str


# ============================================================
# 2. 文件与时间显示
# ============================================================

def check_paths() -> None:
    print("[文件检查]")
    print(f"  附件1   : {INPUT_XLSX}")
    print(f"  result1 : {TEMPLATE_XLSX}")

    if not INPUT_XLSX.exists():
        raise FileNotFoundError(
            f"找不到附件1：{INPUT_XLSX.resolve()}\n"
            "请将 附件1.xlsx 与本程序放在同一文件夹。"
        )

    if not TEMPLATE_XLSX.exists():
        raise FileNotFoundError(
            f"找不到官方result1模板：{TEMPLATE_XLSX.resolve()}\n"
            "请将 result1.xlsx 与本程序放在同一文件夹。"
        )


def time_to_text(value) -> str:
    """仅用于显示，不参与模型索引。"""
    if isinstance(value, dt_time):
        return value.strftime("%H:%M")
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel时间小数：仅用于审计显示
        frac = float(value) % 1.0
        total_min = int(round(frac * 24 * 60)) % (24 * 60)
        return f"{total_min // 60}:{total_min % 60:02d}"
    return str(value).strip()


# ============================================================
# 3. 读取附件1：不做循环重排
# ============================================================

def load_input_xlsx(path: Path) -> InputData:
    """
    只做必要的数据审计：
    - 不平滑
    - 不标准化
    - 不删“异常值”
    - 不对时间顺序做循环移位
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]

    headers = [ws.cell(1, c).value for c in range(1, 5)]
    normalized = [str(x).strip() if x is not None else "" for x in headers]
    expected = ["时间", "电价", "小区负载", "光伏发电预测功率"]

    if normalized != expected:
        wb.close()
        raise ValueError(
            "附件1表头与预期不一致。\n"
            f"实际：{normalized}\n"
            f"预期：{expected}"
        )

    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 5)]

        if all(v is None for v in vals):
            continue

        if any(v is None for v in vals):
            wb.close()
            raise ValueError(f"附件1第{r}行存在缺失值：{vals}")

        rows.append(vals)

    wb.close()

    if len(rows) != N_EXPECTED:
        raise ValueError(
            f"附件1有效数据行数应为{N_EXPECTED}，实际为{len(rows)}。"
        )

    raw_time = [row[0] for row in rows]

    try:
        price = np.asarray([float(row[1]) for row in rows], dtype=float)
        load_kw = np.asarray([float(row[2]) for row in rows], dtype=float)
        pv_kw = np.asarray([float(row[3]) for row in rows], dtype=float)
    except (TypeError, ValueError) as exc:
        raise ValueError("电价/负荷/PV列存在无法转换为数值的单元格。") from exc

    for name, arr in [
        ("电价", price),
        ("小区负载", load_kw),
        ("光伏发电预测功率", pv_kw),
    ]:
        if not np.all(np.isfinite(arr)):
            raise ValueError(f"{name}存在NaN或无穷值。")

        if np.any(arr < 0):
            idx = np.where(arr < 0)[0][:5]
            raise ValueError(
                f"{name}存在负值，原始行索引示例：{idx.tolist()}"
            )

    if np.any(price <= 0):
        raise ValueError("问题1检测到0或负电价，请检查附件。")

    print("[数据审计] 144个记录读取成功，无缺失、NaN、负负荷或负PV。")
    print(
        f"[时间审计] 附件首时刻={time_to_text(raw_time[0])!r}，"
        f"末时刻={time_to_text(raw_time[-1])!r}"
    )
    print(
        "[时间策略] 不进行循环移位；附件第i条数据与官方result1第i条计划时段逐行对应。"
    )

    return InputData(
        raw_time=raw_time,
        price=price,
        load_kw=load_kw,
        pv_kw=pv_kw,
    )


# ============================================================
# 4. MILP骨架
# ============================================================

class ModelIndex:
    def __init__(self, n: int):
        self.n = n
        self.off_pg = 0
        self.off_pc = self.off_pg + n
        self.off_pd = self.off_pc + n
        self.off_pcur = self.off_pd + n
        self.off_e = self.off_pcur + n
        self.off_z = self.off_e + (n + 1)
        self.n_var = self.off_z + n


def build_model(
    data: InputData,
    *,
    allow_grid_charge: bool,
):
    """
    返回MILP的基础结构。

    allow_grid_charge=True:
        正式M2主模型，允许低价外网给储能充电。

    allow_grid_charge=False:
        M1对照场景，定义为“只允许满足负荷后的净光伏余量进入储能”：
            Pc <= max(PV - Load, 0)
        这是一个人为定义的对照情景，不是题目原始约束。
    """
    n = len(data.price)
    idx = ModelIndex(n)

    lb = np.zeros(idx.n_var, dtype=float)
    ub = np.full(idx.n_var, np.inf, dtype=float)

    # Pg >= 0：只允许购电，不允许反向售电
    # 题目未给外网购电上限，因此不擅自添加

    # 充放电功率限制
    ub[idx.off_pc:idx.off_pc + n] = P_BAT_MAX_KW
    ub[idx.off_pd:idx.off_pd + n] = P_BAT_MAX_KW

    # M1：仅允许净光伏余量进入储能
    if not allow_grid_charge:
        pv_surplus = np.maximum(data.pv_kw - data.load_kw, 0.0)
        ub[idx.off_pc:idx.off_pc + n] = np.minimum(
            ub[idx.off_pc:idx.off_pc + n],
            pv_surplus,
        )

    # 弃光
    ub[idx.off_pcur:idx.off_pcur + n] = data.pv_kw

    # SOC安全运行区间
    lb[idx.off_e:idx.off_e + n + 1] = E_MIN_KWH
    ub[idx.off_e:idx.off_e + n + 1] = E_MAX_KWH

    # 0-1状态变量
    ub[idx.off_z:idx.off_z + n] = 1.0
    integrality = np.zeros(idx.n_var, dtype=int)
    integrality[idx.off_z:idx.off_z + n] = 1

    # -------------------------
    # 等式约束
    # -------------------------
    # n个功率平衡 + n个SOC递推 + 2个首末SOC
    Aeq = lil_matrix((2 * n + 2, idx.n_var), dtype=float)
    beq = np.zeros(2 * n + 2, dtype=float)

    row = 0

    # (1) Pg - Pc + Pd - Pcur = Load - PV
    for t in range(n):
        Aeq[row, idx.off_pg + t] = 1.0
        Aeq[row, idx.off_pc + t] = -1.0
        Aeq[row, idx.off_pd + t] = 1.0
        Aeq[row, idx.off_pcur + t] = -1.0

        beq[row] = data.load_kw[t] - data.pv_kw[t]
        row += 1

    # (2) E[t+1]-E[t]-eta_c*Pc*dt+Pd*dt/eta_d = 0
    for t in range(n):
        Aeq[row, idx.off_e + t + 1] = 1.0
        Aeq[row, idx.off_e + t] = -1.0
        Aeq[row, idx.off_pc + t] = -ETA_C * DT_H
        Aeq[row, idx.off_pd + t] = DT_H / ETA_D
        row += 1

    # (3) 周期边界
    Aeq[row, idx.off_e] = 1.0
    beq[row] = E0_KWH
    row += 1

    Aeq[row, idx.off_e + n] = 1.0
    beq[row] = E0_KWH
    row += 1

    assert row == 2 * n + 2

    # -------------------------
    # 充放电互斥
    # -------------------------
    Aub = lil_matrix((2 * n, idx.n_var), dtype=float)
    bub = np.zeros(2 * n, dtype=float)

    row = 0
    for t in range(n):
        # Pc <= Pmax*z
        Aub[row, idx.off_pc + t] = 1.0
        Aub[row, idx.off_z + t] = -P_BAT_MAX_KW
        bub[row] = 0.0
        row += 1

        # Pd <= Pmax*(1-z)
        # 等价于 Pd + Pmax*z <= Pmax
        Aub[row, idx.off_pd + t] = 1.0
        Aub[row, idx.off_z + t] = P_BAT_MAX_KW
        bub[row] = P_BAT_MAX_KW
        row += 1

    base_constraints = [
        LinearConstraint(Aeq.tocsr(), beq, beq),
        LinearConstraint(
            Aub.tocsr(),
            np.full(2 * n, -np.inf, dtype=float),
            bub,
        ),
    ]

    bounds = Bounds(lb, ub)

    return idx, integrality, bounds, base_constraints


# ============================================================
# 5. 三级词典序优化
# ============================================================

def _run_milp(
    c: np.ndarray,
    integrality: np.ndarray,
    bounds: Bounds,
    constraints: list,
    stage_name: str,
):
    result = milp(
        c=c,
        integrality=integrality,
        bounds=bounds,
        constraints=constraints,
        options={
            "disp": False,
            "mip_rel_gap": MIP_REL_GAP,
            "time_limit": TIME_LIMIT_S,
        },
    )

    if not result.success or result.x is None:
        raise RuntimeError(
            f"{stage_name}求解失败。\n"
            f"status={result.status}\n"
            f"message={result.message}"
        )

    return result


def solve_dispatch(
    data: InputData,
    *,
    allow_grid_charge: bool = True,
    three_stage: bool = True,
) -> Solution:
    """
    正式主模型采用三级词典序：
      Stage 1: min 购电费用
      Stage 2: 在Stage1最优条件下 min 弃光量
      Stage 3: 在前两层最优条件下 min 总充放电吞吐量
    """
    n = len(data.price)

    idx, integrality, bounds, base_constraints = build_model(
        data,
        allow_grid_charge=allow_grid_charge,
    )

    # --------------------------------------------------------
    # Stage 1：购电费用最小
    # --------------------------------------------------------
    c_cost = np.zeros(idx.n_var, dtype=float)
    c_cost[idx.off_pg:idx.off_pg + n] = data.price * DT_H

    stage1 = _run_milp(
        c_cost,
        integrality,
        bounds,
        base_constraints,
        "Stage 1（购电费用最小）",
    )

    cost_star = float(stage1.fun)
    x_final = stage1.x
    curtail_star = float(
        np.sum(
            stage1.x[idx.off_pcur:idx.off_pcur + n]
        ) * DT_H
    )
    throughput_final = float(
        np.sum(
            stage1.x[idx.off_pc:idx.off_pc + n]
            + stage1.x[idx.off_pd:idx.off_pd + n]
        ) * DT_H
    )
    final_message = str(stage1.message)

    if three_stage:
        # ----------------------------------------------------
        # Stage 2：保持购电费用最优，最小弃光量
        # ----------------------------------------------------
        A_cost = lil_matrix((1, idx.n_var), dtype=float)
        A_cost[0, idx.off_pg:idx.off_pg + n] = data.price * DT_H

        cost_constraint = LinearConstraint(
            A_cost.tocsr(),
            np.array([-np.inf]),
            np.array([cost_star + LEXI_COST_TOL_YUAN]),
        )

        c_curtail = np.zeros(idx.n_var, dtype=float)
        c_curtail[idx.off_pcur:idx.off_pcur + n] = DT_H

        stage2_constraints = base_constraints + [cost_constraint]

        stage2 = _run_milp(
            c_curtail,
            integrality,
            bounds,
            stage2_constraints,
            "Stage 2（最低弃光）",
        )

        curtail_star = float(stage2.fun)
        x_stage2 = stage2.x

        # ----------------------------------------------------
        # Stage 3：保持购电费用、弃光量均最优，最小吞吐量
        # ----------------------------------------------------
        A_curtail = lil_matrix((1, idx.n_var), dtype=float)
        A_curtail[0, idx.off_pcur:idx.off_pcur + n] = DT_H

        curtail_constraint = LinearConstraint(
            A_curtail.tocsr(),
            np.array([-np.inf]),
            np.array([curtail_star + LEXI_CURTAIL_TOL_KWH]),
        )

        c_throughput = np.zeros(idx.n_var, dtype=float)
        c_throughput[idx.off_pc:idx.off_pc + n] = DT_H
        c_throughput[idx.off_pd:idx.off_pd + n] = DT_H

        stage3_constraints = (
            base_constraints
            + [cost_constraint, curtail_constraint]
        )

        stage3 = _run_milp(
            c_throughput,
            integrality,
            bounds,
            stage3_constraints,
            "Stage 3（最低储能吞吐量）",
        )

        x_final = stage3.x
        throughput_final = float(stage3.fun)
        final_message = str(stage3.message)

    Pg = x_final[idx.off_pg:idx.off_pg + n].copy()
    Pc = x_final[idx.off_pc:idx.off_pc + n].copy()
    Pd = x_final[idx.off_pd:idx.off_pd + n].copy()
    Pcur = x_final[idx.off_pcur:idx.off_pcur + n].copy()
    E = x_final[idx.off_e:idx.off_e + n + 1].copy()
    z = x_final[idx.off_z:idx.off_z + n].copy()

    total_cost = float(
        np.sum(data.price * Pg * DT_H)
    )
    curtail_kwh = float(
        np.sum(Pcur) * DT_H
    )
    throughput_kwh = float(
        np.sum(Pc + Pd) * DT_H
    )

    # 最终数值以最终解重新计算，避免只依赖solver返回fun
    sol = Solution(
        grid_kw=Pg,
        charge_kw=Pc,
        discharge_kw=Pd,
        curtail_kw=Pcur,
        soc_kwh=E,
        mode_binary=z,
        total_cost_yuan=total_cost,
        stage1_cost_yuan=cost_star,
        stage2_curtail_kwh=curtail_kwh,
        stage3_throughput_kwh=throughput_kwh,
        solver_message=final_message,
    )

    audit_solution(data, sol)

    # 词典序最优级别审计
    if three_stage:
        if sol.total_cost_yuan > cost_star + LEXI_COST_TOL_YUAN + 1e-5:
            raise AssertionError(
                "最终解购电费偏离Stage1最优值，词典序约束失效。"
            )

        if sol.stage2_curtail_kwh > curtail_star + LEXI_CURTAIL_TOL_KWH + 1e-5:
            raise AssertionError(
                "最终解弃光量偏离Stage2最优值，词典序约束失效。"
            )

    return sol


# ============================================================
# 6. 模型审计
# ============================================================

def audit_solution(
    data: InputData,
    sol: Solution,
    tol: float = AUDIT_TOL,
) -> dict[str, float]:
    Pg = sol.grid_kw
    Pc = sol.charge_kw
    Pd = sol.discharge_kw
    Pcur = sol.curtail_kw
    E = sol.soc_kwh

    balance_error = (
        Pg + data.pv_kw - Pcur + Pd
        - data.load_kw - Pc
    )

    soc_error = (
        E[1:] - E[:-1]
        - ETA_C * Pc * DT_H
        + Pd * DT_H / ETA_D
    )

    checks = {
        "功率平衡最大误差_kW":
            float(np.max(np.abs(balance_error))),
        "SOC递推最大误差_kWh":
            float(np.max(np.abs(soc_error))),
        "初始SOC误差_kWh":
            float(abs(E[0] - E0_KWH)),
        "终端SOC误差_kWh":
            float(abs(E[-1] - E0_KWH)),
        "最低SOC_kWh":
            float(np.min(E)),
        "最高SOC_kWh":
            float(np.max(E)),
        "最小购电功率_kW":
            float(np.min(Pg)),
        "最大充电功率_kW":
            float(np.max(Pc)),
        "最大放电功率_kW":
            float(np.max(Pd)),
        "最大弃光越界_kW":
            float(np.max(Pcur - data.pv_kw)),
        "最大同时充放电_kW":
            float(np.max(np.minimum(Pc, Pd))),
    }

    assert checks["功率平衡最大误差_kW"] <= tol, checks
    assert checks["SOC递推最大误差_kWh"] <= tol, checks
    assert checks["初始SOC误差_kWh"] <= tol, checks
    assert checks["终端SOC误差_kWh"] <= tol, checks

    assert np.min(E) >= E_MIN_KWH - tol, checks
    assert np.max(E) <= E_MAX_KWH + tol, checks

    assert np.min(Pg) >= -tol, checks

    assert np.min(Pc) >= -tol, checks
    assert np.max(Pc) <= P_BAT_MAX_KW + tol, checks

    assert np.min(Pd) >= -tol, checks
    assert np.max(Pd) <= P_BAT_MAX_KW + tol, checks

    assert np.min(Pcur) >= -tol, checks
    assert np.max(Pcur - data.pv_kw) <= tol, checks

    assert np.max(np.minimum(Pc, Pd)) <= tol, checks

    cost_check = float(
        np.sum(data.price * Pg * DT_H)
    )
    assert abs(cost_check - sol.total_cost_yuan) <= max(
        tol,
        1e-8 * max(1.0, abs(cost_check)),
    )

    print("[模型审计] 所有核心约束通过。")
    print(
        f"  功率平衡最大误差 = "
        f"{checks['功率平衡最大误差_kW']:.3e} kW"
    )
    print(
        f"  SOC递推最大误差 = "
        f"{checks['SOC递推最大误差_kWh']:.3e} kWh"
    )
    print(
        f"  SOC范围 = "
        f"[{checks['最低SOC_kWh']:.4f}, "
        f"{checks['最高SOC_kWh']:.4f}] kWh"
    )
    print(
        f"  SOC 0:00/24:00 = "
        f"{E[0]:.4f}/{E[-1]:.4f} kWh"
    )
    print(
        f"  最大充电功率 = "
        f"{checks['最大充电功率_kW']:.4f} / {P_BAT_MAX_KW:.4f} kW"
    )
    print(
        f"  最大放电功率 = "
        f"{checks['最大放电功率_kW']:.4f} / {P_BAT_MAX_KW:.4f} kW"
    )
    print("  [功率上限审计] 最大充/放电功率均未超过题设上限。")

    return checks


# ============================================================
# 7. 官方result1模板输出
# ============================================================

def read_template_plan_labels(
    template_path: Path,
) -> list[str]:
    wb = load_workbook(
        template_path,
        data_only=False,
        read_only=True,
    )

    required = {"计划购电量", "充放电量"}
    if not required.issubset(set(wb.sheetnames)):
        wb.close()
        raise ValueError(
            f"result1模板缺少工作表。实际={wb.sheetnames}，"
            f"需要={sorted(required)}"
        )

    ws = wb["计划购电量"]
    labels = [
        ws.cell(r, 1).value
        for r in range(2, 2 + N_EXPECTED)
    ]
    wb.close()

    if any(v is None for v in labels):
        raise ValueError("result1“计划购电量”时间段标签不足144行。")

    return [str(v).strip() for v in labels]


def fill_result1_template(
    template_path: Path,
    output_path: Path,
    sol: Solution,
) -> list[str]:
    """
    严格按官方模板行序写回：
        附件第i条数据 -> result1第i条计划购电量
    不做任何循环移位。
    """
    shutil.copy2(template_path, output_path)

    wb = load_workbook(output_path)
    ws_plan = wb["计划购电量"]
    ws_bat = wb["充放电量"]

    labels = [
        ws_plan.cell(r, 1).value
        for r in range(2, 2 + N_EXPECTED)
    ]

    if any(v is None for v in labels):
        wb.close()
        raise ValueError("result1计划购电量模板不足144行。")

    labels = [str(v).strip() for v in labels]

    # 计划购电量：kW -> kWh，严格逐行
    grid_kwh = sol.grid_kw * DT_H
    for i in range(N_EXPECTED):
        cell = ws_plan.cell(row=i + 2, column=2)
        cell.value = float(grid_kwh[i])
        cell.number_format = "0.0000"

    # 表2：按附件/模板给定顺序连续分成6个4h块
    for block in range(6):
        s = block * 24
        e = (block + 1) * 24

        charge_kwh = float(
            np.sum(sol.charge_kw[s:e]) * DT_H
        )
        discharge_kwh = float(
            np.sum(sol.discharge_kw[s:e]) * DT_H
        )

        c1 = ws_bat.cell(row=block + 2, column=2)
        c2 = ws_bat.cell(row=block + 2, column=3)

        c1.value = charge_kwh
        c2.value = discharge_kwh

        c1.number_format = "0.0000"
        c2.number_format = "0.0000"

    # 首末储电量
    ws_bat.cell(
        row=2,
        column=5,
        value=float(sol.soc_kwh[0]),
    )
    ws_bat.cell(
        row=3,
        column=5,
        value=float(sol.soc_kwh[-1]),
    )

    ws_bat.cell(row=2, column=5).number_format = "0.0000"
    ws_bat.cell(row=3, column=5).number_format = "0.0000"

    wb.save(output_path)
    wb.close()

    print(f"[官方结果] 已生成：{output_path.resolve()}")

    return labels


# ============================================================
# 8. 表1、表2
# ============================================================

def get_table1(
    sol: Solution,
    template_labels: list[str],
) -> list[tuple[str, float]]:
    """
    按官方result1模板标签寻找题目要求的6个时段。
    因主模型坚持逐行对应，标签所在行索引就是模型索引。
    """
    targets = [
        "10:00-10:10",
        "12:00-12:10",
        "14:00-14:10",
        "16:00-16:10",
        "18:00-18:10",
        "20:00-20:10",
    ]

    label_to_idx = {
        label: i
        for i, label in enumerate(template_labels)
    }

    grid_kwh = sol.grid_kw * DT_H
    result = []

    for label in targets:
        if label not in label_to_idx:
            raise ValueError(
                f"官方result1模板中未找到题目指定时段：{label}"
            )

        i = label_to_idx[label]
        result.append(
            (label, float(grid_kwh[i]))
        )

    return result


def get_table2(
    sol: Solution,
) -> list[tuple[str, float, float]]:
    names = [
        "0:00-4:00",
        "4:00-8:00",
        "8:00-12:00",
        "12:00-16:00",
        "16:00-20:00",
        "20:00-24:00",
    ]

    rows = []

    for block, name in enumerate(names):
        s = block * 24
        e = (block + 1) * 24

        charge_kwh = float(
            np.sum(sol.charge_kw[s:e]) * DT_H
        )
        discharge_kwh = float(
            np.sum(sol.discharge_kw[s:e]) * DT_H
        )

        rows.append(
            (name, charge_kwh, discharge_kwh)
        )

    return rows


def write_paper_result_xlsx(
    path: Path,
    data: InputData,
    sol: Solution,
    template_labels: list[str],
) -> None:
    """
    自动生成论文可直接使用的“表1 + 表2”Excel。

    说明：
    - 表1中的6个指定时段购电量直接调用 get_table1()，不重复手填；
    - 全天购电量、全天购电费直接由正式M2结果计算；
    - 表2中的6个4小时时段充放电量直接调用 get_table2()；
    - 0:00与24:00储电量直接取正式解的首末SOC；
    - 因此只要正式模型结果变化，本表会随 problem1.py 重新运行自动同步更新。
    """
    table1 = dict(get_table1(sol, template_labels))
    table2 = get_table2(sol)
    energy = energy_audit(data, sol)

    required_labels = [
        "10:00-10:10",
        "12:00-12:10",
        "14:00-14:10",
        "16:00-16:10",
        "18:00-18:10",
        "20:00-20:10",
    ]
    missing = [label for label in required_labels if label not in table1]
    if missing:
        raise ValueError(f"论文结果表缺少指定时段：{missing}")

    if len(table2) != 6:
        raise ValueError(f"论文结果表要求6个4小时时段，实际得到{len(table2)}个。")

    wb = Workbook()
    ws = wb.active
    ws.title = "Q1论文结果表"

    # -------------------------
    # 表1
    # -------------------------
    ws.merge_cells("A1:F1")
    ws["A1"] = "表 1   微网在指定时间段的购电量及全天的购电量和购电费"

    headers1 = ["时间段", "购电量", "时间段", "购电量", "时间段", "购电量"]
    for col, value in enumerate(headers1, start=1):
        ws.cell(row=3, column=col, value=value)

    row4 = [
        "10:00-10:10", table1["10:00-10:10"],
        "12:00-12:10", table1["12:00-12:10"],
        "14:00-14:10", table1["14:00-14:10"],
    ]
    row5 = [
        "16:00-16:10", table1["16:00-16:10"],
        "18:00-18:10", table1["18:00-18:10"],
        "20:00-20:10", table1["20:00-20:10"],
    ]
    for r, values in [(4, row4), (5, row5)]:
        for c, value in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=value)

    ws.merge_cells("A6:B6")
    ws["A6"] = "全天购电量"
    ws["C6"] = float(energy["grid_energy"])
    ws.merge_cells("D6:E6")
    ws["D6"] = "全天购电费"
    ws["F6"] = float(sol.total_cost_yuan)

    # -------------------------
    # 表2
    # -------------------------
    ws.merge_cells("A8:F8")
    ws["A8"] = "表 2   储能设备在指定时间段的充放电量及 0:00 和 24:00 的储电量"

    headers2 = ["时间段", "充电量", "放电量", "时间段", "充电量", "放电量"]
    for col, value in enumerate(headers2, start=1):
        ws.cell(row=10, column=col, value=value)

    # 两个4小时区间并排放置，与题目给定版式一致。
    for pair_idx in range(3):
        left = table2[2 * pair_idx]
        right = table2[2 * pair_idx + 1]
        r = 11 + pair_idx
        values = [
            left[0], float(left[1]), float(left[2]),
            right[0], float(right[1]), float(right[2]),
        ]
        for c, value in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=value)

    ws.merge_cells("A14:B14")
    ws["A14"] = "0:00 储电量"
    ws["C14"] = float(sol.soc_kwh[0])
    ws.merge_cells("D14:E14")
    ws["D14"] = "24:00 储电量"
    ws["F14"] = float(sol.soc_kwh[-1])

    ws.merge_cells("A16:F16")
    ws["A16"] = (
        "注：购电量、充电量、放电量及储电量单位均为 kWh；"
        "全天购电费单位为元。"
    )

    # -------------------------
    # 论文表格样式
    # -------------------------
    title_font = Font(name="微软雅黑", size=16, bold=True)
    header_font = Font(name="微软雅黑", size=11, bold=True)
    body_font = Font(name="微软雅黑", size=11)
    note_font = Font(name="微软雅黑", size=9, italic=True, color="555555")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(fill_type="solid", fgColor="F2F2F2")

    for cell in (ws["A1"], ws["A8"]):
        cell.font = title_font
        cell.alignment = center

    for row in (3, 10):
        for col in range(1, 7):
            cell = ws.cell(row=row, column=col)
            cell.font = header_font
            cell.alignment = center
            cell.fill = header_fill
            cell.border = border

    for row in range(4, 7):
        for col in range(1, 7):
            cell = ws.cell(row=row, column=col)
            cell.font = body_font
            cell.alignment = center
            cell.border = border

    for row in range(11, 15):
        for col in range(1, 7):
            cell = ws.cell(row=row, column=col)
            cell.font = body_font
            cell.alignment = center
            cell.border = border

    # 汇总行加粗。
    for row in (6, 14):
        for col in range(1, 7):
            ws.cell(row=row, column=col).font = Font(
                name="微软雅黑",
                size=11,
                bold=True,
            )

    # 数值统一保留4位小数。
    numeric_cells = [
        "B4", "D4", "F4",
        "B5", "D5", "F5",
        "C6", "F6",
        "B11", "C11", "E11", "F11",
        "B12", "C12", "E12", "F12",
        "B13", "C13", "E13", "F13",
        "C14", "F14",
    ]
    for addr in numeric_cells:
        ws[addr].number_format = "0.0000"

    ws["A16"].font = note_font
    ws["A16"].alignment = left

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[8].height = 30
    for row in list(range(3, 7)) + list(range(10, 15)):
        ws.row_dimensions[row].height = 24

    for col in ("A", "C", "E"):
        ws.column_dimensions[col].width = 18
    for col in ("B", "D", "F"):
        ws.column_dimensions[col].width = 16

    try:
        wb.save(path)
    except PermissionError as exc:
        raise PermissionError(
            f"无法写入 {path.resolve()}。"
            "请先关闭Excel/WPS中打开的 Q1_论文结果表.xlsx 后重新运行。"
        ) from exc

    print(f"[论文结果表] 已生成：{path.resolve()}")


# ============================================================
# 9. 完整明细CSV
# ============================================================

def write_detail_csv(
    path: Path,
    data: InputData,
    sol: Solution,
    template_labels: list[str],
) -> None:
    balance_error = (
        sol.grid_kw
        + data.pv_kw
        - sol.curtail_kw
        + sol.discharge_kw
        - data.load_kw
        - sol.charge_kw
    )

    with path.open(
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as f:
        writer = csv.writer(f)

        writer.writerow([
            "t",
            "附件原始时间",
            "result1时间段",
            "电价_元每kWh",
            "负荷_kW",
            "PV预测_kW",
            "净负荷_kW",
            "购电_kW",
            "购电_kWh",
            "充电_kW",
            "充电_kWh",
            "放电_kW",
            "放电_kWh",
            "弃光_kW",
            "弃光_kWh",
            "SOC期初_kWh",
            "SOC期末_kWh",
            "充电状态z",
            "平衡误差_kW",
        ])

        for t in range(N_EXPECTED):
            writer.writerow([
                t,
                time_to_text(data.raw_time[t]),
                template_labels[t],
                data.price[t],
                data.load_kw[t],
                data.pv_kw[t],
                data.load_kw[t] - data.pv_kw[t],
                sol.grid_kw[t],
                sol.grid_kw[t] * DT_H,
                sol.charge_kw[t],
                sol.charge_kw[t] * DT_H,
                sol.discharge_kw[t],
                sol.discharge_kw[t] * DT_H,
                sol.curtail_kw[t],
                sol.curtail_kw[t] * DT_H,
                sol.soc_kwh[t],
                sol.soc_kwh[t + 1],
                int(round(sol.mode_binary[t])),
                balance_error[t],
            ])

    print(f"[完整明细] 已生成：{path.resolve()}")


# ============================================================
# 10. M0/M1/M2价值分解
# ============================================================

def compute_baselines(
    data: InputData,
) -> dict[str, float]:
    """
    M0：无储能，PV优先满足负荷，余量弃光。
    M1：储能只吸收满足负荷后的净光伏余量，不允许外网低价充电。
    M2：正式主模型，在main中单独求解。

    M1是人为定义的解释性对照情景，不是题目原始约束。
    """
    # M0
    b0_grid_kw = np.maximum(
        data.load_kw - data.pv_kw,
        0.0,
    )
    b0_curtail_kw = np.maximum(
        data.pv_kw - data.load_kw,
        0.0,
    )

    b0_grid_kwh = float(
        np.sum(b0_grid_kw) * DT_H
    )
    b0_cost = float(
        np.sum(
            data.price * b0_grid_kw * DT_H
        )
    )
    b0_curtail_kwh = float(
        np.sum(b0_curtail_kw) * DT_H
    )

    # M1使用同样严谨的三级词典序
    b1 = solve_dispatch(
        data,
        allow_grid_charge=False,
        three_stage=True,
    )

    pv_surplus = np.maximum(
        data.pv_kw - data.load_kw,
        0.0,
    )
    b1_charge_surplus_violation = float(
        np.max(b1.charge_kw - pv_surplus)
    )

    return {
        "M0_grid_kwh": b0_grid_kwh,
        "M0_cost": b0_cost,
        "M0_curtail_kwh": b0_curtail_kwh,

        "M1_grid_kwh":
            float(np.sum(b1.grid_kw) * DT_H),
        "M1_cost":
            float(b1.total_cost_yuan),
        "M1_curtail_kwh":
            float(np.sum(b1.curtail_kw) * DT_H),
        "M1_charge_surplus_violation_kw":
            b1_charge_surplus_violation,
    }


def audit_baseline_definitions(
    data: InputData,
    sol: Solution,
    baselines: dict[str, float],
    tol: float = AUDIT_TOL,
) -> None:
    """
    检查M0/M1/M2的定义和包含关系是否与程序实现一致。

    M0：无储能，PV优先供负荷，余量弃光。
    M1：有储能，但充电功率不能超过当时的净PV富余。
    M2：正式模型，允许外网低价充电。

    因为M2可行域包含M1，M1又至少不应劣于M0，
    所以最优购电费应满足 M2 <= M1 <= M0。
    """
    b0_cost = float(baselines["M0_cost"])
    b1_cost = float(baselines["M1_cost"])
    b2_cost = float(sol.total_cost_yuan)

    # M0定义复算
    b0_grid_kw = np.maximum(data.load_kw - data.pv_kw, 0.0)
    b0_cost_check = float(np.sum(data.price * b0_grid_kw * DT_H))
    assert abs(b0_cost_check - b0_cost) <= max(
        tol,
        1e-8 * max(1.0, abs(b0_cost_check)),
    )

    # M1必须满足“充电不超过净PV富余”的定义
    assert baselines["M1_charge_surplus_violation_kw"] <= tol, baselines

    # 可行域包含关系对应的费用顺序
    assert b2_cost <= b1_cost + 1e-5, (b0_cost, b1_cost, b2_cost)
    assert b1_cost <= b0_cost + 1e-5, (b0_cost, b1_cost, b2_cost)

    print("[M0/M1/M2定义审计] 通过。")
    print("  M0：无储能，PV优先供负荷，余量弃光。")
    print("  M1：仅允许净PV富余进入储能，不允许外网低价充电。")
    print("  M2：正式经济调度模型，允许外网低价充电。")
    print(
        f"  费用关系：M2={b2_cost:.4f} <= "
        f"M1={b1_cost:.4f} <= M0={b0_cost:.4f} 元"
    )



def run_capacity_sensitivity(
    data: InputData,
    base_cost: float,
) -> None:
    """
    储能额定容量敏感性分析。

    基准储能参数为：
        额定容量 C = 12000 kWh
        SOC下限 = 10% C = 1200 kWh
        初始/终端SOC = 50% C = 6000 kWh
        SOC上限 = 90% C = 10800 kWh

    容量变化时保持上述SOC比例不变，仅改变储能容量相关参数；
    最大充放电功率 Pmax、充放电效率以及其余模型条件均保持不变。

    敏感性分析只比较第一目标“最低购电费”，因此使用
    three_stage=False，避免第二、三层词典序影响容量效应解释。
    """
    global E0_KWH, E_MIN_KWH, E_MAX_KWH

    output_csv = CAPACITY_SENSITIVITY_CSV

    # 基准额定容量由题设SOC上下限/初始值反推为12000 kWh。
    # 取基准值两侧若干代表点，便于观察边际收益是否递减。
    capacity_values = [
        6000.0,
        8000.0,
        10000.0,
        12000.0,
        14000.0,
        16000.0,
        18000.0,
    ]

    # 保存正式模型参数，函数结束后无条件恢复，避免污染后续结果。
    e0_base = float(E0_KWH)
    emin_base = float(E_MIN_KWH)
    emax_base = float(E_MAX_KWH)

    # 当前题设对应的额定容量：1200/10% = 6000/50% = 10800/90% = 12000。
    capacity_base = e0_base / 0.50

    # 校验现有参数是否确实符合10%/50%/90%的题设比例。
    ratio_tol = 1e-9
    if not (
        abs(emin_base - 0.10 * capacity_base) <= ratio_tol
        and abs(e0_base - 0.50 * capacity_base) <= ratio_tol
        and abs(emax_base - 0.90 * capacity_base) <= ratio_tol
    ):
        raise ValueError(
            "储能容量敏感性分析默认保持SOC下限10%、初始/终端50%、"
            "SOC上限90%的比例，但当前E_MIN/E0/E_MAX不满足该关系。"
        )

    rows: list[dict[str, float]] = []

    def solve_quietly() -> float:
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            s = solve_dispatch(
                data,
                allow_grid_charge=True,
                three_stage=False,
            )
        return float(s.total_cost_yuan)

    try:
        for capacity in capacity_values:
            # 容量变化时保持SOC安全比例及初末状态比例不变。
            E_MIN_KWH = 0.10 * capacity
            E0_KWH = 0.50 * capacity
            E_MAX_KWH = 0.90 * capacity

            if abs(capacity - capacity_base) <= 1e-12:
                cost = float(base_cost)
            else:
                cost = solve_quietly()

            rows.append({
                "额定容量_kWh": float(capacity),
                "SOC下限_kWh": float(E_MIN_KWH),
                "初始终端SOC_kWh": float(E0_KWH),
                "SOC上限_kWh": float(E_MAX_KWH),
                "可用SOC窗口_kWh": float(E_MAX_KWH - E_MIN_KWH),
                "最低购电费_元": float(cost),
                "相对基准变化_元": float(cost - base_cost),
                "相对基准变化率_%": float(
                    (cost - base_cost) / base_cost * 100.0
                ),
            })
    finally:
        # 无论中途是否异常，都恢复Q1正式模型参数。
        E0_KWH = e0_base
        E_MIN_KWH = emin_base
        E_MAX_KWH = emax_base

    # 理论一致性检查：在本模型设置下，容量增大扩大可行调度空间，
    # 最低购电费不应上升（允许极小数值误差）。
    costs = [float(r["最低购电费_元"]) for r in rows]
    assert all(
        costs[i + 1] <= costs[i] + 1e-5
        for i in range(len(costs) - 1)
    ), costs

    try:
        with output_csv.open(
            "w",
            newline="",
            encoding="utf-8-sig",
        ) as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "额定容量_kWh",
                    "SOC下限_kWh",
                    "初始终端SOC_kWh",
                    "SOC上限_kWh",
                    "可用SOC窗口_kWh",
                    "最低购电费_元",
                    "相对基准变化_元",
                    "相对基准变化率_%",
                ],
            )
            writer.writeheader()
            writer.writerows(rows)
    except PermissionError as exc:
        raise PermissionError(
            f"无法写入 {output_csv.resolve()}。"
            "请先关闭Excel/WPS中打开的 Q1_储能容量敏感性.csv 后重新运行。"
        ) from exc

    print("[储能容量敏感性分析] 完成。")
    print(
        "  容量口径：额定容量变化时保持SOC下限10%、"
        "初始/终端50%、SOC上限90%，其余参数不变。"
    )
    print(f"  基准额定容量：{capacity_base:.0f} kWh")
    print(f"  结果表：{output_csv.resolve()}")

    # 同步生成论文用容量敏感性SVG图。
    # 直接使用本轮敏感性求解得到的rows，不重复读取CSV，也不重复求解。
    make_capacity_sensitivity_plot(
        rows=rows,
        capacity_base=capacity_base,
    )

def run_sensitivity_analysis(
    data: InputData,
    base_cost: float,
) -> None:
    """
    一次只改变一个参数，其余参数保持基准值。

    1) 单向充/放电效率 eta：0.85, 0.875, 0.90, 0.925, 0.95
    2) 最大充/放电功率 Pmax：3000, 4000, 5000, 6000, 7000 kW

    敏感性分析只比较第一目标“最低购电费”，
    因此这里使用three_stage=False，避免第二、三层词典序影响解释。
    """
    global ETA_C, ETA_D, P_BAT_MAX_KW

    eta_c_base = float(ETA_C)
    eta_d_base = float(ETA_D)
    pmax_base = float(P_BAT_MAX_KW)

    eta_values = [0.85, 0.875, 0.90, 0.925, 0.95]
    pmax_values = [3000.0, 4000.0, 5000.0, 6000.0, 7000.0]

    rows: list[dict[str, float | str]] = []

    def solve_quietly() -> float:
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            s = solve_dispatch(
                data,
                allow_grid_charge=True,
                three_stage=False,
            )
        return float(s.total_cost_yuan)

    try:
        # 效率敏感性：eta_c = eta_d = eta
        P_BAT_MAX_KW = pmax_base
        for eta in eta_values:
            ETA_C = float(eta)
            ETA_D = float(eta)
            if (
                abs(eta - eta_c_base) <= 1e-12
                and abs(eta - eta_d_base) <= 1e-12
            ):
                cost = float(base_cost)
            else:
                cost = solve_quietly()

            rows.append({
                "参数": "单向充放电效率",
                "参数值": float(eta),
                "最低购电费_元": float(cost),
                "相对基准变化_元": float(cost - base_cost),
                "相对基准变化率_%": float(
                    (cost - base_cost) / base_cost * 100.0
                ),
            })

        # 最大充放电功率敏感性
        ETA_C = eta_c_base
        ETA_D = eta_d_base
        for pmax in pmax_values:
            P_BAT_MAX_KW = float(pmax)
            if abs(pmax - pmax_base) <= 1e-12:
                cost = float(base_cost)
            else:
                cost = solve_quietly()

            rows.append({
                "参数": "最大充放电功率_kW",
                "参数值": float(pmax),
                "最低购电费_元": float(cost),
                "相对基准变化_元": float(cost - base_cost),
                "相对基准变化率_%": float(
                    (cost - base_cost) / base_cost * 100.0
                ),
            })
    finally:
        # 无论中途是否异常，都恢复正式模型参数
        ETA_C = eta_c_base
        ETA_D = eta_d_base
        P_BAT_MAX_KW = pmax_base

    # 保存CSV；如果文件被Excel/WPS占用，给出明确提示而不是模糊报错
    try:
        with SENSITIVITY_CSV.open(
            "w",
            newline="",
            encoding="utf-8-sig",
        ) as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "参数",
                    "参数值",
                    "最低购电费_元",
                    "相对基准变化_元",
                    "相对基准变化率_%",
                ],
            )
            writer.writeheader()
            writer.writerows(rows)
    except PermissionError as exc:
        raise PermissionError(
            f"无法写入 {SENSITIVITY_CSV.resolve()}。"
            "请先关闭Excel/WPS中打开的 Q1_敏感性分析.csv 后重新运行。"
        ) from exc

    # 理论一致性检查：Pmax增加时可行域扩大，最低购电费不应上升
    p_rows = [
        r for r in rows
        if r["参数"] == "最大充放电功率_kW"
    ]
    p_costs = [float(r["最低购电费_元"]) for r in p_rows]
    assert all(
        p_costs[i + 1] <= p_costs[i] + 1e-5
        for i in range(len(p_costs) - 1)
    ), p_costs

    # 效率提高时，最优费用在本模型下也应不升
    eta_rows = [
        r for r in rows
        if r["参数"] == "单向充放电效率"
    ]
    eta_costs = [float(r["最低购电费_元"]) for r in eta_rows]
    assert all(
        eta_costs[i + 1] <= eta_costs[i] + 1e-5
        for i in range(len(eta_costs) - 1)
    ), eta_costs

    print("[敏感性分析] 完成。")
    print(f"  结果表：{SENSITIVITY_CSV.resolve()}")

    # 绘图单独交给统一论文风函数，避免求解逻辑与绘图逻辑混在一起。
    make_sensitivity_plots(
        rows=rows,
        eta_base=eta_c_base,
        pmax_base=pmax_base,
    )


# ============================================================
# 11. 能量审计
# ============================================================

def energy_audit(
    data: InputData,
    sol: Solution,
) -> dict[str, float]:
    load_energy = float(
        np.sum(data.load_kw) * DT_H
    )
    pv_energy = float(
        np.sum(data.pv_kw) * DT_H
    )
    grid_energy = float(
        np.sum(sol.grid_kw) * DT_H
    )
    charge_energy = float(
        np.sum(sol.charge_kw) * DT_H
    )
    discharge_energy = float(
        np.sum(sol.discharge_kw) * DT_H
    )
    curtail_energy = float(
        np.sum(sol.curtail_kw) * DT_H
    )

    # 系统侧严格能量闭合：
    # grid + PV - curtail + discharge = load + charge
    closure = (
        grid_energy
        + pv_energy
        - curtail_energy
        + discharge_energy
        - load_energy
        - charge_energy
    )

    pv_utilization = (
        (pv_energy - curtail_energy) / pv_energy
        if pv_energy > 0
        else float("nan")
    )

    return {
        "load_energy": load_energy,
        "pv_energy": pv_energy,
        "grid_energy": grid_energy,
        "charge_energy": charge_energy,
        "discharge_energy": discharge_energy,
        "curtail_energy": curtail_energy,
        "energy_closure": closure,
        "pv_utilization": pv_utilization,
    }


# ============================================================
# 12. 论文配图（科研论文风：白底 + 清晰坐标 + 参考线 + 注释）
# ============================================================

# 科研配色：严格采用用户提供的蓝-青-红配色板。
# 主色：#104E8B #376B9E #5F89B1 #AFC3D8 #C5E9E3 #D7E1EB
# 强调：#F2DADA #E5B5B5 #D89090 #B22222
C_BLUE = "#104E8B"       # 主深蓝：核心曲线 / 主结果
C_SKY = "#376B9E"        # 次蓝：辅助主线
C_GREEN = "#5F89B1"      # 灰蓝：第三组曲线 / 辅助信息
C_ORANGE = "#D89090"     # 暖红：次级强调 / 放电相关
C_RED = "#B22222"        # 深红：关键强调 / 高价 / 警戒
C_PURPLE = "#E5B5B5"     # 浅红：第四层级强调
C_YELLOW = "#C5E9E3"     # 薄荷青：淡色填充
C_GRAY = "#AFC3D8"       # 浅蓝灰：参考线
C_LIGHT = "#D7E1EB"      # 极浅蓝灰：静置 / 背景
C_PALE_RED = "#F2DADA"   # 极浅粉：淡背景
C_TEXT = "#333333"       # 正文深灰


def configure_plot_font() -> bool:
    """设置 Windows 常见中文字体；没有时退回 Matplotlib 默认字体。"""
    try:
        import matplotlib
        from matplotlib import font_manager

        available = {f.name for f in font_manager.fontManager.ttflist}
        candidates = [
            "Microsoft YaHei",
            "SimHei",
            "SimSun",
            "Noto Sans CJK SC",
            "Arial Unicode MS",
        ]
        for font_name in candidates:
            if font_name in available:
                matplotlib.rcParams["font.sans-serif"] = [font_name]
                matplotlib.rcParams["axes.unicode_minus"] = False
                matplotlib.rcParams["font.size"] = 11
                matplotlib.rcParams["axes.titlesize"] = 16
                matplotlib.rcParams["axes.labelsize"] = 12
                matplotlib.rcParams["legend.fontsize"] = 10.5
                # SVG 中保留文字对象，便于后续在 Illustrator / Inkscape 中编辑
                matplotlib.rcParams["svg.fonttype"] = "none"
                return True

        matplotlib.rcParams["axes.unicode_minus"] = False
        matplotlib.rcParams["svg.fonttype"] = "none"
    except Exception:
        pass
    return False


def _style_axis(ax, *, grid_axis: str = "both") -> None:
    """统一成用户参考图里的论文风：白底、灰网格、清晰黑灰边框。"""
    ax.set_facecolor("white")
    ax.grid(
        True,
        axis=grid_axis,
        color=C_LIGHT,
        linewidth=0.75,
        alpha=0.55,
        linestyle="-",
    )
    for side in ("left", "bottom", "top", "right"):
        ax.spines[side].set_color(C_GRAY)
        ax.spines[side].set_linewidth(0.85)
    ax.tick_params(
        axis="both",
        colors=C_TEXT,
        labelsize=10.5,
        direction="out",
        length=4,
        width=0.8,
    )
    ax.set_axisbelow(True)


def _legend(ax, **kwargs) -> None:
    """统一图例样式。"""
    defaults = dict(
        frameon=True,
        facecolor="white",
        edgecolor=C_GRAY,
        framealpha=0.92,
        fancybox=False,
    )
    defaults.update(kwargs)
    ax.legend(**defaults)


def _save_figure(fig, path: Path) -> None:
    """统一保存为 SVG 矢量图；放大不失真，适合论文排版。"""
    import matplotlib
    import matplotlib.pyplot as plt

    # 保留文字为 SVG 文本对象，便于后续编辑。
    matplotlib.rcParams["svg.fonttype"] = "none"

    path = Path(path).with_suffix(".svg")
    fig.savefig(
        path,
        format="svg",
        bbox_inches="tight",
        facecolor="white",
        transparent=False,
    )
    plt.close(fig)


def make_plots(data: InputData, sol: Solution) -> None:
    """
    生成4张基础图。

    风格重点：
    - 不做装饰性背景，不做3D；
    - 标题、坐标、图例、参考线清晰；
    - 用注释直接解释模型行为。
    """
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        print("[绘图提醒] 未安装matplotlib，跳过绘图。可执行：pip install matplotlib")
        return

    configure_plot_font()

    t = np.arange(N_EXPECTED, dtype=float) * DT_H
    t_soc = np.arange(N_EXPECTED + 1, dtype=float) * DT_H

    # --------------------------------------------------------
    # 图3：储能电量轨迹 + 上下界 + 首末状态
    # --------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11.0, 5.5))

    ax.step(
        t_soc,
        sol.soc_kwh,
        where="post",
        color=C_BLUE,
        linewidth=2.25,
        label="储能电量 $E_t$",
        zorder=3,
    )

    ax.axhline(
        E_MAX_KWH,
        color=C_GRAY,
        linewidth=1.35,
        linestyle="--",
        label=f"上限 {E_MAX_KWH:.0f} kWh",
        zorder=2,
    )
    ax.axhline(
        E_MIN_KWH,
        color=C_GRAY,
        linewidth=1.35,
        linestyle=":",
        label=f"下限 {E_MIN_KWH:.0f} kWh",
        zorder=2,
    )

    # 首末状态
    ax.scatter(
        [0, 24],
        [sol.soc_kwh[0], sol.soc_kwh[-1]],
        s=55,
        color=C_ORANGE,
        edgecolor="white",
        linewidth=0.8,
        zorder=5,
    )
    # 首末文字放到曲线旁边的空白区域，不用白底遮盖数据线
    ax.annotate(
        f"初始 {sol.soc_kwh[0]:.0f} kWh",
        xy=(0, sol.soc_kwh[0]),
        xytext=(1.25, 6550),
        textcoords="data",
        ha="left",
        va="bottom",
        fontsize=9.5,
        color=C_ORANGE,
        arrowprops=dict(arrowstyle="->", color=C_ORANGE, lw=1.0),
    )
    ax.annotate(
        f"24 h：{sol.soc_kwh[-1]:.0f} kWh",
        xy=(24, sol.soc_kwh[-1]),
        xytext=(21.45, 6550),
        textcoords="data",
        ha="left",
        va="bottom",
        fontsize=9.5,
        color=C_ORANGE,
        arrowprops=dict(arrowstyle="->", color=C_ORANGE, lw=1.0),
    )

    # 首次触及上下界的位置，用垂直参考线表示“活跃约束”
    low_hits = np.where(np.isclose(sol.soc_kwh, E_MIN_KWH, atol=1e-3))[0]
    high_hits = np.where(np.isclose(sol.soc_kwh, E_MAX_KWH, atol=1e-3))[0]

    if len(high_hits):
        ih = int(high_hits[0])
        th = float(t_soc[ih])
        ax.axvline(th, color=C_GREEN, linestyle="--", linewidth=1.1, alpha=0.85)
        ax.text(
            th + 0.12,
            E_MAX_KWH - 550,
            "首次触及上限",
            color=C_GREEN,
            fontsize=9.2,
            rotation=90,
            va="top",
        )

    if len(low_hits):
        il = int(low_hits[0])
        tl = float(t_soc[il])
        ax.axvline(tl, color=C_RED, linestyle="--", linewidth=1.1, alpha=0.85)
        ax.text(
            tl + 0.12,
            E_MIN_KWH + 450,
            "首次触及下限",
            color=C_RED,
            fontsize=9.2,
            rotation=90,
            va="bottom",
        )

    # 左右各留出0.6 h空白，避免0 h和24 h端点被坐标轴裁掉
    ax.set_xlim(-0.6, 24.6)
    ax.set_xticks(np.arange(0, 25, 4))
    ax.set_xlabel("时间 / h")
    ax.set_ylabel("储能电量 / kWh")
    ax.set_title("问题1：最优调度下储能电量轨迹及运行边界", pad=12)
    _style_axis(ax)
    _legend(ax, loc="best")
    fig.tight_layout()
    _save_figure(fig, SOC_FIG)

    # --------------------------------------------------------
    # 备用图：单独分时电价
    # --------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10.6, 4.8))
    ax.step(
        t,
        data.price,
        where="post",
        color=C_BLUE,
        linewidth=2.1,
        label="分时电价",
    )

    pmin_i = int(np.argmin(data.price))
    pmax_i = int(np.argmax(data.price))
    ax.scatter(t[pmin_i], data.price[pmin_i], color=C_GREEN, s=42, zorder=5)
    ax.scatter(t[pmax_i], data.price[pmax_i], color=C_RED, s=42, zorder=5)

    ax.annotate(
        f"最低 {data.price[pmin_i]:.3f}",
        (t[pmin_i], data.price[pmin_i]),
        xytext=(12, 18),
        textcoords="offset points",
        color=C_GREEN,
        fontsize=9.2,
        arrowprops=dict(arrowstyle="->", color=C_GREEN, lw=0.9),
    )
    ax.annotate(
        f"最高 {data.price[pmax_i]:.3f}",
        (t[pmax_i], data.price[pmax_i]),
        xytext=(-66, 18),
        textcoords="offset points",
        color=C_RED,
        fontsize=9.2,
        arrowprops=dict(arrowstyle="->", color=C_RED, lw=0.9),
    )

    ax.set_xlim(0, 24)
    ax.set_xticks(np.arange(0, 25, 4))
    ax.set_xlabel("时间 / h")
    ax.set_ylabel("电价 / (元/kWh)")
    ax.set_title("问题1：分时电价曲线", pad=10)
    _style_axis(ax)
    _legend(ax, loc="upper right")
    fig.tight_layout()
    _save_figure(fig, PRICE_FIG)

    # --------------------------------------------------------
    # 备用图：单独充放电功率
    # --------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10.6, 5.0))

    ax.step(
        t,
        sol.charge_kw,
        where="post",
        color=C_BLUE,
        linewidth=2.0,
        label="充电功率",
    )
    ax.step(
        t,
        -sol.discharge_kw,
        where="post",
        color=C_ORANGE,
        linewidth=2.0,
        label="放电功率（负值）",
    )
    ax.axhline(0, color="#444444", linewidth=0.9)

    # 标出最大充/放电功率
    ic = int(np.argmax(sol.charge_kw))
    idis = int(np.argmax(sol.discharge_kw))
    ax.scatter(t[ic], sol.charge_kw[ic], color=C_BLUE, s=45, zorder=5)
    ax.scatter(t[idis], -sol.discharge_kw[idis], color=C_ORANGE, s=45, zorder=5)
    # 文字移动到数据线附近的空白区，不使用白底文本框
    ax.annotate(
        f"最大充电 {sol.charge_kw[ic]:.0f} kW",
        xy=(t[ic], sol.charge_kw[ic]),
        xytext=(2.15, 4150),
        textcoords="data",
        ha="left",
        va="center",
        fontsize=9.2,
        color=C_BLUE,
        zorder=10,
        arrowprops=dict(
            arrowstyle="->",
            color=C_BLUE,
            lw=0.9,
            shrinkA=3,
            shrinkB=3,
        ),
    )
    ax.annotate(
        f"最大放电 {sol.discharge_kw[idis]:.0f} kW",
        xy=(t[idis], -sol.discharge_kw[idis]),
        xytext=(16.65, -3250),
        textcoords="data",
        ha="left",
        va="center",
        fontsize=9.2,
        color=C_ORANGE,
        zorder=10,
        arrowprops=dict(
            arrowstyle="->",
            color=C_ORANGE,
            lw=0.9,
            shrinkA=3,
            shrinkB=3,
        ),
    )

    # 给0 h与24 h两端留白，端点标记完整显示
    ax.set_xlim(-0.6, 24.6)
    ax.set_xticks(np.arange(0, 25, 4))
    ax.set_xlabel("时间 / h")
    ax.set_ylabel("功率 / kW")
    ax.set_title("问题1：储能最优充放电功率", pad=10)
    _style_axis(ax)
    _legend(ax, loc="lower left")
    fig.tight_layout()
    _save_figure(fig, BATTERY_FIG)

    # --------------------------------------------------------
    # 图1：负荷—光伏—外网购电联合调度
    # --------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11.4, 5.6))

    ax.step(
        t,
        data.load_kw,
        where="post",
        color=C_BLUE,
        linewidth=2.0,
        label="小区负荷",
        zorder=4,
    )
    ax.step(
        t,
        data.pv_kw,
        where="post",
        color=C_GREEN,
        linewidth=2.0,
        label="光伏预测功率",
        zorder=3,
    )
    ax.step(
        t,
        sol.grid_kw,
        where="post",
        color=C_ORANGE,
        linewidth=2.0,
        label="外网购电功率",
        zorder=5,
    )

    # 只标注一个典型 Pg=0 时段，避免图上到处都是点
    zero_idx = np.where(sol.grid_kw <= 1e-6)[0]
    if len(zero_idx):
        iz = int(zero_idx[len(zero_idx) // 2])
        ax.scatter(t[iz], 0, color=C_RED, s=45, zorder=6)
        ax.annotate(
            "$P_g=0$",
            xy=(t[iz], 0),
            xytext=(11.65, 900),
            textcoords="data",
            ha="left",
            va="center",
            fontsize=9.5,
            color=C_RED,
            zorder=10,
            arrowprops=dict(
                arrowstyle="->",
                color=C_RED,
                lw=0.9,
                shrinkA=4,
                shrinkB=3,
            ),
        )

    # 一个典型 Pg > PL 时段，解释凌晨外网给储能充电
    extra_idx = np.where(sol.grid_kw > data.load_kw + 100.0)[0]
    if len(extra_idx):
        ie = int(extra_idx[0])
        ax.annotate(
            "外网同时供负荷并给储能充电",
            xy=(t[ie], sol.grid_kw[ie]),
            xytext=(1.45, 7200),
            textcoords="data",
            ha="left",
            va="center",
            fontsize=9.2,
            color=C_ORANGE,
            zorder=10,
            arrowprops=dict(
                arrowstyle="->",
                color=C_ORANGE,
                lw=0.9,
                shrinkA=4,
                shrinkB=3,
            ),
        )

    ax.set_xlim(0, 24)
    ax.set_xticks(np.arange(0, 25, 4))
    ax.set_xlabel("时间 / h")
    ax.set_ylabel("功率 / kW")
    ax.set_title("问题1：负荷、光伏预测功率与最优外网购电功率", pad=12)
    _style_axis(ax)
    _legend(ax, loc="upper center", ncol=3)
    fig.tight_layout()
    _save_figure(fig, BALANCE_FIG)

    print("[论文配图] 已生成4张科研论文风基础图。")


def make_explanatory_plots(
    data: InputData,
    sol: Solution,
    baselines: dict[str, float],
) -> None:
    """
    两张正文核心解释图：
    1) 分时电价—储能充放电联合响应；
    2) M0/M1/M2购电费用与经济价值分解。
    """
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        print("[解释性配图] 未安装matplotlib，跳过。")
        return

    configure_plot_font()
    t = np.arange(N_EXPECTED, dtype=float) * DT_H

    # --------------------------------------------------------
    # 图2：分时电价与充放电联合响应（双Y轴，参考用户示例）
    # --------------------------------------------------------
    fig, ax_power = plt.subplots(figsize=(11.4, 5.8))
    ax_price = ax_power.twinx()

    charge_line = ax_power.step(
        t,
        sol.charge_kw,
        where="post",
        color=C_BLUE,
        linewidth=2.05,
        label="充电功率",
        zorder=4,
    )[0]
    discharge_line = ax_power.step(
        t,
        -sol.discharge_kw,
        where="post",
        color=C_ORANGE,
        linewidth=2.05,
        label="放电功率（负值）",
        zorder=4,
    )[0]
    price_line = ax_price.step(
        t,
        data.price,
        where="post",
        color=C_GREEN,
        linewidth=2.15,
        label="分时电价",
        zorder=5,
    )[0]

    ax_power.axhline(0, color="#444444", linewidth=0.9)

    # 最低价和最高价参考线
    pmin_i = int(np.argmin(data.price))
    pmax_i = int(np.argmax(data.price))
    t_low = float(t[pmin_i])
    t_high = float(t[pmax_i])

    ax_power.axvline(t_low, color=C_BLUE, linestyle="--", linewidth=1.1, alpha=0.8)
    ax_power.axvline(t_high, color=C_RED, linestyle="--", linewidth=1.1, alpha=0.8)

    ax_price.scatter(t_low, data.price[pmin_i], color=C_BLUE, s=42, zorder=7)
    ax_price.scatter(t_high, data.price[pmax_i], color=C_RED, s=42, zorder=7)

    ax_price.annotate(
        "最低电价",
        xy=(t_low, data.price[pmin_i]),
        # 文字向右下方移动，避开最低价附近的电价曲线和充放电曲线
        xytext=(18, -18),
        textcoords="offset points",
        ha="left",
        va="top",
        color=C_BLUE,
        fontsize=9.3,
        zorder=10,
        arrowprops=dict(arrowstyle="->", color=C_BLUE, lw=0.9),
    )
    ax_price.annotate(
        "最高电价",
        xy=(t_high, data.price[pmax_i]),
        xytext=(-88, 26),
        textcoords="offset points",
        ha="left",
        va="bottom",
        color=C_RED,
        fontsize=9.3,
        zorder=10,
        arrowprops=dict(arrowstyle="->", color=C_RED, lw=0.9),
    )

    ax_power.set_xlim(0, 24)
    ax_power.set_xticks(np.arange(0, 25, 4))
    ax_power.set_xlabel("时间 / h")
    ax_power.set_ylabel("储能充放电功率 / kW", color=C_TEXT)
    ax_price.set_ylabel("分时电价 / (元/kWh)", color=C_GREEN)
    ax_price.tick_params(axis="y", colors=C_GREEN)

    ax_power.set_title("问题1：分时电价与储能最优充放电响应", pad=12)
    _style_axis(ax_power)
    ax_price.grid(False)
    ax_price.spines["right"].set_color(C_GREEN)
    ax_price.spines["right"].set_linewidth(0.95)
    ax_price.spines["top"].set_color("#333333")

    # 图例仍放在坐标轴外，但改为单行横排并整体上移，
    # 让说明区紧贴主图下方，同时不遮挡横轴标题。
    fig.legend(
        [charge_line, discharge_line, price_line],
        ["充电功率", "放电功率（负值）", "分时电价"],
        loc="lower left",
        bbox_to_anchor=(0.075, 0.050),
        bbox_transform=fig.transFigure,
        ncol=3,
        frameon=True,
        facecolor="white",
        edgecolor="#C8C8C8",
        framealpha=0.92,
        fontsize=9.3,
        borderpad=0.45,
        labelspacing=0.35,
        columnspacing=1.35,
        handlelength=2.0,
    )

    # 套利判据与图例处在同一条底部说明带中，整体上移。
    fig.text(
        0.985,
        0.062,
        r"效率修正套利判据：$p_j>p_i/(\eta_c\eta_d)$",
        ha="right",
        va="bottom",
        fontsize=9.3,
        color="#555555",
    )

    # 减少底部留白，使图例/判据与主图距离更自然。
    fig.tight_layout(rect=(0.0, 0.115, 1.0, 1.0))
    _save_figure(fig, PRICE_BATTERY_FIG)

    # --------------------------------------------------------
    # 图4：M0/M1/M2 —— 改成“费用阶梯图”，不再画三个巨大柱子
    # --------------------------------------------------------
    b0 = float(baselines["M0_cost"])
    b1 = float(baselines["M1_cost"])
    b2 = float(sol.total_cost_yuan)
    save_pv = b0 - b1
    save_peak = b1 - b2
    total_save = b0 - b2
    saving_rate = total_save / b0 * 100.0

    x = np.arange(3, dtype=float)
    y = np.array([b0, b1, b2], dtype=float)

    fig, ax = plt.subplots(figsize=(9.6, 5.7))

    ax.plot(
        x,
        y,
        color=C_BLUE,
        linewidth=2.5,
        marker="o",
        markersize=8.5,
        markerfacecolor="white",
        markeredgewidth=2.0,
        markeredgecolor=C_BLUE,
        zorder=4,
    )

    # M2最终值用红色强调
    ax.scatter(
        [2],
        [b2],
        s=95,
        color=C_RED,
        edgecolor="white",
        linewidth=1.2,
        zorder=6,
    )

    # 费用标签
    for xi, yi in zip(x, y):
        ax.text(
            xi,
            yi + 450,
            f"{yi:,.2f} 元",
            ha="center",
            va="bottom",
            fontsize=10.8,
            color=C_TEXT,
        )

    # 两段下降用箭头 + 文本解释
    ax.annotate(
        "",
        xy=(1, b1),
        xytext=(0, b0),
        arrowprops=dict(arrowstyle="-|>", color=C_GREEN, lw=1.6),
    )
    ax.annotate(
        f"光伏余量时移价值\n−{save_pv:,.2f} 元",
        xy=(0.50, (b0 + b1) / 2),
        xytext=(0, 42),
        textcoords="offset points",
        ha="center",
        va="bottom",
        fontsize=9.5,
        color=C_GREEN,
        zorder=10,
    )

    ax.annotate(
        "",
        xy=(2, b2),
        xytext=(1, b1),
        arrowprops=dict(arrowstyle="-|>", color=C_ORANGE, lw=1.6),
    )
    ax.annotate(
        f"峰谷时移价值\n−{save_peak:,.2f} 元",
        xy=(1.50, (b1 + b2) / 2),
        xytext=(0, -50),
        textcoords="offset points",
        ha="center",
        va="top",
        fontsize=9.5,
        color=C_ORANGE,
        zorder=10,
    )

    # 以M0为基线，像用户示例中的“基线”一样显示
    ax.axhline(
        b0,
        color=C_GRAY,
        linestyle="--",
        linewidth=1.05,
        alpha=0.8,
        label=f"M0基线 = {b0:,.2f} 元",
    )

    ax.text(
        0.98,
        0.94,
        f"总节省 = {total_save:,.2f} 元\n总节费率 = {saving_rate:.2f}%",
        transform=ax.transAxes,
        ha="right",
        va="top",
        fontsize=10,
        color=C_TEXT,
    )

    ax.set_xticks(x)
    ax.set_xticklabels([
        "M0\n无储能",
        "M1\n仅净光伏余量储能",
        "M2\n完整经济调度",
    ])
    ax.set_xlim(-0.25, 2.25)
    # 不从0起，突出“费用下降”而不是柱高；折线图允许这样做
    pad = max(1200.0, (b0 - b2) * 0.14)
    ax.set_ylim(b2 - pad, b0 + pad * 1.35)
    ax.set_ylabel("购电费用 / 元")
    ax.set_title("问题1：不同储能利用机制下购电费用与价值分解", pad=12)
    _style_axis(ax)
    _legend(ax, loc="lower left")
    fig.tight_layout()
    _save_figure(fig, BASELINE_COST_FIG)

    print("[解释性配图] 已生成联合响应图和M0/M1/M2费用阶梯图。")


# ============================================================
# 13. 新增：问题1高价值分析图（非重复展示）
# ============================================================

def make_additional_value_plots(
    data: InputData,
    sol: Solution,
    baselines: dict[str, float],
) -> None:
    """
    额外生成4张真正有解释价值的图：

    1) 储能削峰填谷效果：M0无储能购电功率 vs M2最优购电功率；
    2) 累计购电成本与累计节省：展示节省在一天中何时形成；
    3) 储能运行状态与活跃约束时间轴：解释大量0值以及约束何时起作用；
    4) 光伏消纳改善：比较M0/M1/M2的PV利用率与弃光量。

    这些图只使用已求得的结果，不改变任何模型、约束或数值。
    """
    try:
        import matplotlib.pyplot as plt
        from matplotlib.patches import Patch
        from matplotlib.colors import ListedColormap
    except ImportError:
        print("[扩展配图] 未安装matplotlib，跳过新增高价值图。")
        return

    configure_plot_font()

    n = N_EXPECTED
    t = np.arange(n, dtype=float) * DT_H
    t_edge = np.arange(n + 1, dtype=float) * DT_H
    eps = 1e-6

    # M0：无储能时，PV优先供负荷，剩余缺口由外网承担
    b0_grid_kw = np.maximum(data.load_kw - data.pv_kw, 0.0)

    # ========================================================
    # 图7：储能对外网购电功率的削峰填谷作用
    # ========================================================
    fig, ax = plt.subplots(figsize=(11.2, 5.6))

    ax.step(
        t,
        b0_grid_kw,
        where="post",
        color=C_GRAY,
        linewidth=1.75,
        linestyle="--",
        label="M0 无储能购电功率",
        zorder=2,
    )
    ax.step(
        t,
        sol.grid_kw,
        where="post",
        color=C_BLUE,
        linewidth=2.25,
        label="M2 最优调度购电功率",
        zorder=3,
    )

    # M2 > M0：储能正在增加购电以低价充电
    ax.fill_between(
        t,
        b0_grid_kw,
        sol.grid_kw,
        where=(sol.grid_kw > b0_grid_kw + eps),
        step="post",
        color=C_ORANGE,
        alpha=0.22,
        label="低价增购用于充电",
        zorder=1,
    )

    # M2 < M0：储能放电减少外网购电，即削峰
    ax.fill_between(
        t,
        b0_grid_kw,
        sol.grid_kw,
        where=(sol.grid_kw < b0_grid_kw - eps),
        step="post",
        color=C_GREEN,
        alpha=0.22,
        label="储能放电削减购电",
        zorder=1,
    )

    reduction = b0_grid_kw - sol.grid_kw
    increase = sol.grid_kw - b0_grid_kw

    i_cut = int(np.argmax(reduction))
    i_fill = int(np.argmax(increase))

    if reduction[i_cut] > eps:
        ax.scatter(t[i_cut], sol.grid_kw[i_cut], s=45, color=C_GREEN, zorder=5)
        # 标注放在20~22 h附近的空白区域，不用白底覆盖数据
        ax.annotate(
            f"最大购电削减\n{reduction[i_cut]:.0f} kW",
            xy=(t[i_cut], sol.grid_kw[i_cut]),
            xytext=(21.10, 1450),
            textcoords="data",
            ha="left",
            va="center",
            fontsize=9.5,
            color=C_GREEN,
            arrowprops=dict(arrowstyle="->", color=C_GREEN, lw=1.0),
        )

    if increase[i_fill] > eps:
        ax.scatter(t[i_fill], sol.grid_kw[i_fill], s=45, color=C_ORANGE, zorder=5)
        # 0 h端点旁边上方有足够空白，文字移到右侧而不遮挡曲线
        ax.annotate(
            f"最大增购充电\n{increase[i_fill]:.0f} kW",
            xy=(t[i_fill], sol.grid_kw[i_fill]),
            xytext=(1.15, 7350),
            textcoords="data",
            ha="left",
            va="center",
            fontsize=9.5,
            color=C_ORANGE,
            arrowprops=dict(arrowstyle="->", color=C_ORANGE, lw=1.0),
        )

    # 两端留白，避免0 h处的最大增购点被裁掉
    ax.set_xlim(-0.6, 24.6)
    ax.set_xticks(np.arange(0, 25, 4))
    ax.set_xlabel("时间 / h")
    ax.set_ylabel("外网购电功率 / kW")
    ax.set_title("问题1：储能对外网购电功率的削峰填谷作用", pad=12)
    _style_axis(ax)
    _legend(ax, loc="upper center", ncol=2)
    fig.tight_layout()
    _save_figure(fig, GRID_SHIFT_FIG)

    # ========================================================
    # 图8：累计购电成本 + 累计节省
    # ========================================================
    b0_slot_cost = data.price * b0_grid_kw * DT_H
    b2_slot_cost = data.price * sol.grid_kw * DT_H

    cum_b0 = np.concatenate(([0.0], np.cumsum(b0_slot_cost)))
    cum_b2 = np.concatenate(([0.0], np.cumsum(b2_slot_cost)))
    cum_save = cum_b0 - cum_b2

    fig, (ax1, ax2) = plt.subplots(
        2,
        1,
        figsize=(11.0, 7.4),
        sharex=True,
        gridspec_kw={"height_ratios": [1.35, 1.0], "hspace": 0.10},
    )

    ax1.plot(
        t_edge,
        cum_b0,
        color=C_GRAY,
        linewidth=1.9,
        linestyle="--",
        label="M0 无储能累计购电费",
    )
    ax1.plot(
        t_edge,
        cum_b2,
        color=C_BLUE,
        linewidth=2.35,
        label="M2 最优调度累计购电费",
    )

    ax1.fill_between(
        t_edge,
        cum_b0,
        cum_b2,
        where=(cum_b0 >= cum_b2),
        color=C_GREEN,
        alpha=0.13,
        interpolate=True,
    )
    ax1.fill_between(
        t_edge,
        cum_b0,
        cum_b2,
        where=(cum_b0 < cum_b2),
        color=C_ORANGE,
        alpha=0.13,
        interpolate=True,
    )

    ax1.set_ylabel("累计购电费用 / 元")
    ax1.set_title("问题1：储能调度下累计购电成本与节省形成过程", pad=12)
    _style_axis(ax1)
    _legend(ax1, loc="upper left")

    ax2.plot(
        t_edge,
        cum_save,
        color=C_GREEN,
        linewidth=2.25,
        label="累计节省金额 M0−M2",
    )
    ax2.axhline(0.0, color=C_GRAY, linewidth=1.0, linestyle="--")
    ax2.fill_between(
        t_edge,
        0.0,
        cum_save,
        where=(cum_save >= 0.0),
        color=C_GREEN,
        alpha=0.18,
        interpolate=True,
    )
    ax2.fill_between(
        t_edge,
        0.0,
        cum_save,
        where=(cum_save < 0.0),
        color=C_ORANGE,
        alpha=0.18,
        interpolate=True,
    )

    total_save = float(cum_save[-1])
    save_rate = total_save / float(cum_b0[-1]) * 100.0 if cum_b0[-1] > 0 else float("nan")

    ax2.scatter([24.0], [total_save], s=58, color=C_RED, zorder=5)
    # 最终节省说明放在右下图的空白区域，不用白底遮住累计节省曲线
    ax2.annotate(
        f"全天最终节省\n{total_save:.2f} 元（{save_rate:.2f}%）",
        xy=(24.0, total_save),
        xytext=(19.05, total_save * 0.74),
        textcoords="data",
        ha="left",
        va="center",
        fontsize=9.8,
        color=C_RED,
        arrowprops=dict(arrowstyle="->", color=C_RED, lw=1.0),
    )

    # sharex会同步到上图；两端留白使24 h红点完整显示
    ax2.set_xlim(-0.6, 24.6)
    ax2.set_xticks(np.arange(0, 25, 4))
    ax2.set_xlabel("时间 / h")
    ax2.set_ylabel("累计节省 / 元")
    _style_axis(ax2)
    _legend(ax2, loc="upper left")

    fig.tight_layout()
    _save_figure(fig, CUMULATIVE_COST_FIG)

    # ========================================================
    # 图9：储能状态时间轴 + 关键活跃约束
    # ========================================================
    # 状态编码：0=放电，1=静置，2=充电
    state = np.ones(n, dtype=int)
    state[sol.discharge_kw > 1e-5] = 0
    state[sol.charge_kw > 1e-5] = 2

    charge_count = int(np.sum(state == 2))
    idle_count = int(np.sum(state == 1))
    discharge_count = int(np.sum(state == 0))

    # 采用三行布局：状态带 / 独立图例行 / 约束事件图。
    # 图例单独占一行，避免SVG中被下方子图遮挡。
    fig = plt.figure(figsize=(11.0, 6.6))
    gs = fig.add_gridspec(
        3,
        1,
        height_ratios=[0.75, 0.22, 1.6],
        hspace=0.08,
    )

    ax_state = fig.add_subplot(gs[0, 0])
    ax_legend = fig.add_subplot(gs[1, 0])
    ax_event = fig.add_subplot(gs[2, 0], sharex=ax_state)

    cmap = ListedColormap([C_RED, C_LIGHT, C_BLUE])
    ax_state.imshow(
        state.reshape(1, -1),
        aspect="auto",
        interpolation="nearest",
        cmap=cmap,
        vmin=0,
        vmax=2,
        extent=[0, 24, 0, 1],
    )
    ax_state.set_yticks([])
    ax_state.set_ylabel("运行状态")
    ax_state.set_title("问题1：储能运行状态与关键活跃约束时间轴", pad=12)
    ax_state.grid(False)
    ax_state.tick_params(axis="x", labelbottom=False)
    for side in ("left", "bottom", "top", "right"):
        ax_state.spines[side].set_color("#333333")
        ax_state.spines[side].set_linewidth(0.85)

    state_handles = [
        Patch(facecolor=C_BLUE, edgecolor="none", label=f"充电（{charge_count}个时段）"),
        Patch(facecolor=C_LIGHT, edgecolor="none", label=f"静置（{idle_count}个时段）"),
        Patch(facecolor=C_RED, edgecolor="none", label=f"放电（{discharge_count}个时段）"),
    ]

    # 独立图例轴：只显示图例，不显示坐标轴。
    ax_legend.axis("off")
    ax_legend.legend(
        handles=state_handles,
        loc="center",
        ncol=3,
        frameon=False,
        fontsize=9.5,
        handlelength=2.2,
        columnspacing=2.8,
    )

    # 事件用散点时间轴表示
    soc_start = sol.soc_kwh[:-1]
    event_specs = [
        ("SOC上界活跃", np.isclose(soc_start, E_MAX_KWH, atol=1e-3), C_GREEN, "s"),
        ("SOC下界活跃", np.isclose(soc_start, E_MIN_KWH, atol=1e-3), C_RED, "s"),
        ("充电功率达上限", np.isclose(sol.charge_kw, P_BAT_MAX_KW, atol=1e-3), C_BLUE, "o"),
        ("放电功率达上限", np.isclose(sol.discharge_kw, P_BAT_MAX_KW, atol=1e-3), C_PURPLE, "o"),
        ("外网购电为0", sol.grid_kw <= 1e-6, C_ORANGE, "|"),
    ]

    y_pos = np.arange(len(event_specs), dtype=float)
    for y, (label, mask, color, marker) in zip(y_pos, event_specs):
        times = t[mask]
        if len(times):
            size = 115 if marker == "|" else 30
            ax_event.scatter(
                times,
                np.full(len(times), y),
                s=size,
                color=color,
                marker=marker,
                linewidths=1.3,
                zorder=3,
            )

    ax_event.set_yticks(y_pos)
    ax_event.set_yticklabels([x[0] for x in event_specs])
    ax_event.set_ylim(-0.65, len(event_specs) - 0.35)
    ax_event.invert_yaxis()
    ax_event.set_xlim(0, 24)
    ax_event.set_xticks(np.arange(0, 25, 2))
    ax_event.set_xlabel("时间 / h")
    ax_event.set_ylabel("关键约束/事件")
    _style_axis(ax_event, grid_axis="x")

    # 固定边距，进一步避免SVG浏览器渲染时文字/图例被裁切或遮挡。
    fig.subplots_adjust(
        left=0.16,
        right=0.985,
        top=0.91,
        bottom=0.10,
        hspace=0.08,
    )
    _save_figure(fig, STATE_TIMELINE_FIG)

    # ========================================================
    # 图10：M0/M1/M2光伏消纳改善
    # ========================================================
    pv_energy = float(np.sum(data.pv_kw) * DT_H)
    if pv_energy <= 0:
        print("[扩展配图] PV总能量为0，跳过光伏消纳改善图。")
    else:
        b0_curtail = float(baselines["M0_curtail_kwh"])
        b1_curtail = float(baselines["M1_curtail_kwh"])
        b2_curtail = float(np.sum(sol.curtail_kw) * DT_H)

        curtail = np.array([b0_curtail, b1_curtail, b2_curtail], dtype=float)
        utilization = 100.0 * (1.0 - curtail / pv_energy)
        labels = ["M0\n无储能", "M1\n净光伏余量储能", "M2\n完整经济调度"]
        x3 = np.arange(3, dtype=float)

        fig, (ax_u, ax_c) = plt.subplots(1, 2, figsize=(11.2, 4.9))

        ax_u.plot(
            x3,
            utilization,
            color=C_GREEN,
            linewidth=2.2,
            marker="o",
            markersize=7.5,
        )
        for xi, yi in zip(x3, utilization):
            ax_u.text(xi, yi + 0.6, f"{yi:.2f}%", ha="center", va="bottom", fontsize=9.8)
        ax_u.set_xticks(x3)
        ax_u.set_xticklabels(labels)
        low = max(0.0, float(np.min(utilization)) - 4.0)
        ax_u.set_ylim(low, 102.5)
        ax_u.set_ylabel("光伏本地利用率 / %")
        ax_u.set_title("光伏本地利用率", pad=10)
        _style_axis(ax_u, grid_axis="y")

        ax_c.plot(
            x3,
            curtail,
            color=C_RED,
            linewidth=2.2,
            marker="o",
            markersize=7.5,
        )
        for xi, yi in zip(x3, curtail):
            offset = max(8.0, float(np.max(curtail)) * 0.025)
            ax_c.text(xi, yi + offset, f"{yi:.2f}", ha="center", va="bottom", fontsize=9.8)
        ax_c.set_xticks(x3)
        ax_c.set_xticklabels(labels)
        ax_c.set_ylabel("弃光电量 / kWh")
        ax_c.set_title("弃光电量", pad=10)
        _style_axis(ax_c, grid_axis="y")

        fig.suptitle("问题1：不同储能利用机制下的光伏消纳改善", fontsize=15, y=1.02)
        fig.tight_layout()
        _save_figure(fig, PV_UTIL_FIG)

    print("[扩展配图] 已生成4张新增高价值图。")
    print(f"  1) {GRID_SHIFT_FIG.resolve()}")
    print(f"  2) {CUMULATIVE_COST_FIG.resolve()}")
    print(f"  3) {STATE_TIMELINE_FIG.resolve()}")
    print(f"  4) {PV_UTIL_FIG.resolve()}")


def make_capacity_sensitivity_plot(
    rows: list[dict[str, float]],
    capacity_base: float,
) -> None:
    """
    将储能容量敏感性结果绘制为论文风SVG。

    横轴：储能额定容量 / kWh
    纵轴：最低购电费 / 元

    与效率、功率上限敏感性图保持同一配色和版式；
    不改变模型、不重新求解，只可视化run_capacity_sensitivity已得到的rows。
    """
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        print("[储能容量敏感性] 未安装matplotlib，仅保留CSV。")
        return

    configure_plot_font()

    capacity_x = np.asarray(
        [float(r["额定容量_kWh"]) for r in rows],
        dtype=float,
    )
    cost_y = np.asarray(
        [float(r["最低购电费_元"]) for r in rows],
        dtype=float,
    )

    if len(capacity_x) == 0:
        print("[储能容量敏感性] 无可绘制数据，跳过SVG。")
        return

    # 保证横轴按容量从小到大显示。
    order = np.argsort(capacity_x)
    capacity_x = capacity_x[order]
    cost_y = cost_y[order]

    fig, ax = plt.subplots(figsize=(9.0, 5.4))

    ax.plot(
        capacity_x,
        cost_y,
        color=C_BLUE,
        linewidth=2.35,
        marker="o",
        markersize=7.5,
        markerfacecolor="white",
        markeredgecolor=C_BLUE,
        markeredgewidth=1.7,
        label="最低购电费",
        zorder=3,
    )

    # 高亮题设基准容量12000 kWh（由当前SOC比例反推）。
    ib = int(np.argmin(np.abs(capacity_x - capacity_base)))
    base_x = float(capacity_x[ib])
    base_y = float(cost_y[ib])

    ax.axvline(
        base_x,
        color=C_RED,
        linestyle="--",
        linewidth=1.45,
        label=f"题设基准容量 {base_x:.0f} kWh",
        zorder=2,
    )
    ax.scatter(
        [base_x],
        [base_y],
        color=C_RED,
        s=68,
        zorder=5,
    )

    # 各容量点直接标出最低购电费，便于论文读数。
    yrange = float(np.max(cost_y) - np.min(cost_y))
    label_offset = max(45.0, yrange * 0.012)
    for xv, yv in zip(capacity_x, cost_y):
        ax.text(
            xv,
            yv + label_offset,
            f"{yv:,.0f}",
            ha="center",
            va="bottom",
            fontsize=8.8,
            color=C_TEXT,
        )

    # 基准说明移到曲线附近空白位置，不使用白底文本框，避免遮挡数据线。
    ax.annotate(
        f"基准费用\n{base_y:,.2f} 元",
        xy=(base_x, base_y),
        xytext=(24, 26),
        textcoords="offset points",
        ha="left",
        va="bottom",
        fontsize=9.5,
        color=C_RED,
        arrowprops=dict(
            arrowstyle="->",
            color=C_RED,
            lw=0.9,
            shrinkA=3,
            shrinkB=3,
        ),
    )

    # 两端留白，避免首末标记贴边或被裁切。
    xrange = float(np.max(capacity_x) - np.min(capacity_x))
    xpad = max(500.0, xrange * 0.06)
    ypad = max(650.0, yrange * 0.10)
    ax.set_xlim(float(np.min(capacity_x)) - xpad, float(np.max(capacity_x)) + xpad)
    ax.set_ylim(float(np.min(cost_y)) - ypad * 0.45, float(np.max(cost_y)) + ypad)

    ax.set_xticks(capacity_x)
    ax.set_xlabel("储能额定容量 / kWh")
    ax.set_ylabel("最低购电费 / 元")
    ax.set_title("问题1：储能容量敏感性分析", pad=12)
    _style_axis(ax)
    _legend(ax, loc="upper right")
    fig.tight_layout()
    _save_figure(fig, CAPACITY_SENSITIVITY_FIG)

    print(f"  容量敏感性图：{CAPACITY_SENSITIVITY_FIG.resolve()}")


def make_sensitivity_plots(
    rows: list[dict[str, float | str]],
    eta_base: float,
    pmax_base: float,
) -> None:
    """将敏感性分析画成用户参考图那种“基准虚线 + 注释 + 清晰折线”风格。"""
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        print("[敏感性分析] 未安装matplotlib，仅保留CSV。")
        return

    configure_plot_font()

    eta_rows = [r for r in rows if r["参数"] == "单向充放电效率"]
    p_rows = [r for r in rows if r["参数"] == "最大充放电功率_kW"]

    # --------------------------------------------------------
    # 图5：效率敏感性
    # --------------------------------------------------------
    eta_x = np.asarray([float(r["参数值"]) for r in eta_rows], dtype=float)
    eta_y = np.asarray([float(r["最低购电费_元"]) for r in eta_rows], dtype=float)

    fig, ax = plt.subplots(figsize=(9.0, 5.4))
    ax.plot(
        eta_x,
        eta_y,
        color=C_BLUE,
        linewidth=2.35,
        marker="o",
        markersize=7.5,
        label="最低购电费",
    )

    ib = int(np.argmin(np.abs(eta_x - eta_base)))
    ax.axvline(
        eta_base,
        color=C_RED,
        linestyle="--",
        linewidth=1.45,
        label=f"题设基准 η={eta_base:.2f}",
    )
    ax.scatter(eta_x[ib], eta_y[ib], color=C_RED, s=65, zorder=5)
    ax.annotate(
        f"基准费用\n{eta_y[ib]:,.2f} 元",
        (eta_x[ib], eta_y[ib]),
        xytext=(18, 22),
        textcoords="offset points",
        fontsize=9.5,
        color=C_TEXT,
        arrowprops=dict(arrowstyle="->", color=C_RED, lw=0.9),
        bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="#BFBFBF", alpha=0.95),
    )

    # 每个点标值，像用户示例中的风险曲线
    for xv, yv in zip(eta_x, eta_y):
        ax.text(xv, yv + 45, f"{yv:,.0f}", ha="center", va="bottom", fontsize=8.8)

    ax.set_xlabel("单向充放电效率 η")
    ax.set_ylabel("最低购电费 / 元")
    ax.set_title("问题1：储能效率敏感性分析", pad=12)
    _style_axis(ax)
    _legend(ax, loc="upper right")
    fig.tight_layout()
    _save_figure(fig, SENSITIVITY_ETA_FIG)

    # --------------------------------------------------------
    # 图6：最大充放电功率敏感性
    # --------------------------------------------------------
    p_x = np.asarray([float(r["参数值"]) for r in p_rows], dtype=float)
    p_y = np.asarray([float(r["最低购电费_元"]) for r in p_rows], dtype=float)

    fig, ax = plt.subplots(figsize=(9.0, 5.4))
    ax.plot(
        p_x,
        p_y,
        color=C_BLUE,
        linewidth=2.35,
        marker="o",
        markersize=7.5,
        label="最低购电费",
    )

    ib = int(np.argmin(np.abs(p_x - pmax_base)))
    ax.axvline(
        pmax_base,
        color=C_ORANGE,
        linestyle="--",
        linewidth=1.45,
        label=f"题设上限 {pmax_base:.0f} kW",
    )
    ax.scatter(p_x[ib], p_y[ib], color=C_RED, s=65, zorder=5)

    for xv, yv in zip(p_x, p_y):
        ax.text(xv, yv + 18, f"{yv:,.0f}", ha="center", va="bottom", fontsize=8.8)

    idx5000 = np.where(np.isclose(p_x, 5000.0))[0]
    idx6000 = np.where(np.isclose(p_x, 6000.0))[0]
    if len(idx5000) and len(idx6000):
        i5 = int(idx5000[0])
        i6 = int(idx6000[0])
        marginal = float(p_y[i5] - p_y[i6])
        ax.annotate(
            f"5000→6000 kW\n购电费仅再下降 {marginal:,.2f} 元",
            xy=((p_x[i5] + p_x[i6]) / 2, (p_y[i5] + p_y[i6]) / 2),
            xytext=(30, 38),
            textcoords="offset points",
            fontsize=9.3,
            color=C_GREEN,
            arrowprops=dict(arrowstyle="->", color=C_GREEN, lw=0.9),
            bbox=dict(boxstyle="round,pad=0.25", fc="white", ec=C_GREEN, alpha=0.95),
        )

    ax.set_xlabel("最大充放电功率 / kW")
    ax.set_ylabel("最低购电费 / 元")
    ax.set_title("问题1：最大充放电功率敏感性分析", pad=12)
    _style_axis(ax)
    _legend(ax, loc="upper right")
    fig.tight_layout()
    _save_figure(fig, SENSITIVITY_PMAX_FIG)

    print(f"  效率敏感性图：{SENSITIVITY_ETA_FIG.resolve()}")
    print(f"  功率上限敏感性图：{SENSITIVITY_PMAX_FIG.resolve()}")


# ============================================================
# 13. 结果摘要
# ============================================================

def build_summary_text(
    data: InputData,
    sol: Solution,
    template_labels: list[str],
    baselines: dict[str, float],
) -> str:
    table1 = get_table1(sol, template_labels)
    table2 = get_table2(sol)
    energy = energy_audit(data, sol)

    b0 = baselines["M0_cost"]
    b1 = baselines["M1_cost"]
    b2 = sol.total_cost_yuan

    total_saving = b0 - b2
    saving_rate = (
        total_saving / b0
        if b0 > 0
        else float("nan")
    )

    pv_value = b0 - b1
    arbitrage_value = b1 - b2

    lines = []

    lines.append("=" * 76)
    lines.append("C题问题1 最终审校版结果摘要")
    lines.append("=" * 76)

    lines.append("【模型口径】")
    lines.append(
        "附件原始144行与官方result1模板144个计划时段严格逐行对应，不做循环移位。"
    )
    lines.append(
        f"eta_c={ETA_C:.4f}, eta_d={ETA_D:.4f}, "
        f"往返效率={ETA_C * ETA_D:.4f}"
    )
    lines.append(
        "三级词典序：购电费最小 -> 弃光最小 -> 储能吞吐量最小"
    )
    lines.append("")

    lines.append("【三级词典序结果】")
    lines.append(
        f"Stage1最低购电费       = {sol.stage1_cost_yuan:.6f} 元"
    )
    lines.append(
        f"最终方案购电费         = {sol.total_cost_yuan:.6f} 元"
    )
    lines.append(
        f"最终弃光量             = {sol.stage2_curtail_kwh:.6f} kWh"
    )
    lines.append(
        f"最终储能总吞吐量       = {sol.stage3_throughput_kwh:.6f} kWh"
    )
    lines.append("")

    lines.append("【表1 微网指定时段计划购电量】")
    for label, value in table1:
        lines.append(
            f"{label:>13s} : {value:12.4f} kWh"
        )

    lines.append(
        f"全天购电量             = {energy['grid_energy']:.4f} kWh"
    )
    lines.append(
        f"全天购电费             = {sol.total_cost_yuan:.4f} 元"
    )
    lines.append("")

    lines.append("【表2 储能指定4小时时段充放电量】")
    for label, charge, discharge in table2:
        lines.append(
            f"{label:>13s} : "
            f"充电={charge:12.4f} kWh，"
            f"放电={discharge:12.4f} kWh"
        )

    lines.append(
        f"0:00储电量             = {sol.soc_kwh[0]:.4f} kWh"
    )
    lines.append(
        f"24:00储电量            = {sol.soc_kwh[-1]:.4f} kWh"
    )
    lines.append("")

    lines.append("【能量与可行性审计】")
    lines.append(
        f"全天负荷电量           = {energy['load_energy']:.4f} kWh"
    )
    lines.append(
        f"全天PV预测电量         = {energy['pv_energy']:.4f} kWh"
    )
    lines.append(
        f"储能总充电量           = {energy['charge_energy']:.4f} kWh"
    )
    lines.append(
        f"储能总放电量           = {energy['discharge_energy']:.4f} kWh"
    )
    lines.append(
        f"全天弃光量             = {energy['curtail_energy']:.4f} kWh"
    )
    lines.append(
        f"PV本地利用率           = {energy['pv_utilization'] * 100:.4f}%"
    )
    lines.append(
        f"系统能量闭合误差       = {energy['energy_closure']:.3e} kWh"
    )
    lines.append(
        f"SOC范围                = "
        f"[{sol.soc_kwh.min():.4f}, {sol.soc_kwh.max():.4f}] kWh"
    )
    lines.append("")

    lines.append("【M0/M1/M2 储能经济价值分解】")
    lines.append(
        f"M0 无储能                 = {b0:.4f} 元"
    )
    lines.append(
        f"M1 仅净光伏余量储能       = {b1:.4f} 元"
    )
    lines.append(
        f"M2 完整经济调度           = {b2:.4f} 元"
    )
    lines.append(
        f"光伏本地消纳价值 M0-M1   = {pv_value:.4f} 元"
    )
    lines.append(
        f"峰谷时移价值     M1-M2   = {arbitrage_value:.4f} 元"
    )
    lines.append(
        f"总节省           M0-M2   = {total_saving:.4f} 元"
    )
    lines.append(
        f"总节费率                  = {saving_rate * 100:.4f}%"
    )
    lines.append("")

    lines.append("【效率修正套利判据】")
    lines.append(
        "外网低价充电并在未来高价时段释放具有经济价值时，应满足："
    )
    lines.append(
        f"p_high > p_low/(eta_c*eta_d) "
        f"= p_low/{ETA_C * ETA_D:.4f}"
    )

    lines.append("=" * 76)

    return "\n".join(lines)


def save_and_print_summary(
    summary: str,
) -> None:
    print()
    print(summary)
    SUMMARY_TXT.write_text(
        summary,
        encoding="utf-8-sig",
    )
    print(f"\n[结果摘要] 已生成：{SUMMARY_TXT.resolve()}")


# ============================================================
# 14. 主程序
# ============================================================

def main() -> None:
    print("=" * 76)
    print("2026高教社杯 C题 问题1 —— 最终审校版")
    print("=" * 76)

    check_paths()

    # 1. 数据
    data = load_input_xlsx(INPUT_XLSX)

    # 2. 正式主模型M2
    sol = solve_dispatch(
        data,
        allow_grid_charge=True,
        three_stage=True,
    )

    # 3. 官方模板输出
    template_labels = fill_result1_template(
        TEMPLATE_XLSX,
        OUTPUT_XLSX,
        sol,
    )

    # 4. 完整明细
    write_detail_csv(
        DETAIL_CSV,
        data,
        sol,
        template_labels,
    )

    # 5. M0/M1对照
    baselines = compute_baselines(data)
    audit_baseline_definitions(data, sol, baselines)

    # 6. 结果摘要
    summary = build_summary_text(
        data,
        sol,
        template_labels,
        baselines,
    )
    save_and_print_summary(summary)

    # 6.1 生成论文可直接使用的表1 + 表2 Excel
    write_paper_result_xlsx(
        PAPER_RESULT_XLSX,
        data,
        sol,
        template_labels,
    )

    # 7. 论文配图
    make_plots(
        data,
        sol,
    )

    # 7.1 新增两张解释性核心图
    make_explanatory_plots(
        data,
        sol,
        baselines,
    )

    # 7.2 新增4张高价值分析图
    make_additional_value_plots(
        data,
        sol,
        baselines,
    )

    # 8. 参数敏感性分析
    run_capacity_sensitivity(data, sol.stage1_cost_yuan)

    run_sensitivity_analysis(
        data,
        sol.stage1_cost_yuan,
    )

    print("\n全部完成。主要输出：")
    print(f"  1) {OUTPUT_XLSX.name}")
    print(f"  2) {DETAIL_CSV.name}")
    print(f"  3) {SUMMARY_TXT.name}")
    print(f"  4) {PAPER_RESULT_XLSX.name}")
    print(f"  5) {SOC_FIG.name}")
    print(f"  6) {PRICE_FIG.name}")
    print(f"  7) {BATTERY_FIG.name}")
    print(f"  8) {BALANCE_FIG.name}")
    print(f"  9) {SENSITIVITY_CSV.name}")
    print(f" 10) {SENSITIVITY_ETA_FIG.name}")
    print(f" 11) {SENSITIVITY_PMAX_FIG.name}")
    print(f" 12) {CAPACITY_SENSITIVITY_CSV.name}")
    print(f" 13) {CAPACITY_SENSITIVITY_FIG.name}")
    print(f" 14) {PRICE_BATTERY_FIG.name}")
    print(f" 15) {BASELINE_COST_FIG.name}")
    print(f" 16) {GRID_SHIFT_FIG.name}")
    print(f" 17) {CUMULATIVE_COST_FIG.name}")
    print(f" 18) {STATE_TIMELINE_FIG.name}")
    print(f" 19) {PV_UTIL_FIG.name}")


if __name__ == "__main__":
    main()

