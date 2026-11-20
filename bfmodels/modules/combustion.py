"""燃烧、鼓风与气流类模型。

对应 Excel 工作表：理论燃烧温度 / 鼓风动能 / 实际风速 / 透气性指数 /
富氧率 / 热负荷。
"""

from __future__ import annotations

import math
from typing import Any, Dict, List

from ..base import (
    SRC_EXCEL,
    SRC_TYPICAL,
    Input,
    Model,
    Range,
    Result,
    numeric_partials,
)

PI = math.pi


# ===========================================================================
# 理论燃烧温度
# ===========================================================================


class TheoreticalCombustionTemperature(Model):
    key = "combustion_temp"
    title = "理论燃烧温度"
    subtitle = "风口前焦炭与喷吹物绝热燃烧所能达到的最高温度"
    group = "燃烧与鼓风"
    source_sheet = "理论燃烧温度"
    description = (
        "原表同一张表内并存 **三套公式**：块1(R19:R24) 用「湿分 %」+ 系数 6.003/3250，"
        "块2(R31:R36) 用「鼓风湿度 g/m³」+ 系数 6.033/3150。"
        "本工具把三套做成可切换的下拉项，并用同一套默认参数复算。"
    )
    formula_latex = (
        r"T_f = 1559 + 0.839\,BT - 6.033\,BW "
        r"+ \frac{4972\,O_2}{(BV + O_2/60)\times 60} "
        r"- \frac{3150\,PCI}{(BV + O_2/60)\times 60}"
    )
    formula_excel = (
        "=1559+0.839*BT-6.033*BW+4972*VO2/(VB*60+VO2)-3150*Pc/(VB*60+VO2)"
    )
    notes = [
        "原表 G24 / L24 把湿度项写成 `-6.003*BH/(273.15+BT)*(6.11*10^0.7)`，"
        "与 image2 主公式 `-6.003*BH*(273.15+BT)/(6.11*10^7)` 不等价："
        "(273.15+BT) 跑到了分母，且指数 0.7 与 7 不一致。"
        "同等参数下偏差 −1.23 ℃（G 列）与 −2.48 ℃（L 列）。",
        "原表 F24 把基准温度写死为 `(273.15+25)`，应为 F19（1220 ℃）；"
        "F 列与 D 列参数完全相同却因写法不同产生 0.0012 ℃ 差异。",
        "原表 H24:K24 无输入数据，输出 #DIV/0!，属预留模板位。",
        "两套公式的湿度定义不同（% 与 g/m³），系数也不同（6.003/3250 与 6.033/3150），"
        "跨块比较数值没有意义。",
    ]
    variant_label = "「公式版本」下拉可切换原表三套写法；默认参数对应块2 的 F 列（实测工况）。"

    inputs = [
        Input(
            key="bt", label="热风温度", symbol="BT", unit="℃", default=1198,
            step=1.0, fmt="%.6g",
            help="块2 F 列取值 1198 ℃",
        ),
        Input(
            key="bw", label="鼓风湿度 / 湿分", symbol="BW / BH", unit="g/m³ 或 %",
            default=13, step=0.1, fmt="%.6g",
            note="单位随所选公式版本：公式B 用 g/m³，公式A/原表G-L 用 %。",
            help="公式B(块2)取 g/m³；公式A(块1)取 %。",
        ),
        Input(
            key="vo2", label="富氧量", symbol="VO2", unit="m³/h", default=33815,
            step=1.0, fmt="%.8g",
        ),
        Input(
            key="vb", label="风量", symbol="VB", unit="m³/min", default=7829,
            step=1.0, fmt="%.8g",
        ),
        Input(
            key="pc", label="喷煤量", symbol="Pc", unit="kg/h", default=78900,
            step=100.0, fmt="%.8g",
        ),
        Input(
            key="variant", label="公式版本", kind="select", default="B",
            options=["B", "A", "G"],
            option_labels=[
                "公式B · 湿度按 g/m³（块2，已核对无误）",
                "公式A · 湿分按 %（块1，主公式正确写法）",
                "原表 G/L 列写法（6.11×10^0.7，含笔误）",
            ],
            help="切换以复现原表三种写法的差异。",
        ),
    ]

    reference = [
        {"key": "tf", "label": "理论燃烧温度", "symbol": "Tf", "unit": "℃",
         "excel": 2326.01467290564, "note": "块2 F 列；公式B 复算一致"},
    ]

    # ------------------------------------------------------------------
    def _tf(self, v: Dict[str, Any]) -> float:
        bt, bw, vo2, vb, pc = v["bt"], v["bw"], v["vo2"], v["vb"], v["pc"]
        den = (vb + vo2 / 60.0) * 60.0
        variant = v.get("variant", "B")
        if variant == "B":      # 公式B
            return (1559 + 0.839 * bt - 6.033 * bw
                    + 4972 * vo2 / den - 3150 * pc / den)
        if variant == "A":      # 公式A（正确写法，湿分按 %）
            return (1559 + 0.839 * bt
                    - 6.003 * bw * (273.15 + bt) / (6.11 * 1e7)
                    + 4972 * vo2 / den - 3250 * pc / den)
        # 原表 G/L 列写法（含笔误）
        return (1559 + 0.839 * bt
                - 6.003 * bw / (273.15 + bt) * (6.11 * 10 ** 0.7)
                + 4972 * vo2 / den - 3250 * pc / den)

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        tf = self._tf(v)
        variant = v.get("variant", "B")
        names = {
            "B": "公式B：Tf = 1559+0.839·BT − 6.033·BW + 4972·O₂/((BV+O₂/60)·60) − 3150·PCI/((BV+O₂/60)·60)",
            "A": "公式A：Tf = 1559+0.839·BT − 6.003·BH·(273.15+BT)/(6.11×10⁷) + 4972·O₂/((BV+O₂/60)·60) − 3250·PCI/((BV+O₂/60)·60)",
            "G": "原表 G/L 写法：Tf = 1559+0.839·BT − 6.003·BH/((273.15+BT)·(6.11×10^0.7)) + …（含笔误，仅供复现）",
        }
        results = [
            Result(
                key="tf", label="理论燃烧温度", value=tf, unit="℃", digits=2,
                symbol="Tf",
                rng=Range(2050, 2250, SRC_EXCEL,
                          "原表综合表给出 2150±100 ℃；不同炉型、煤比与喷煤比下需按现场标定。"),
                formula=names[variant],
                note=f"当前公式版本：{names[variant].split('：')[0]}",
            )
        ]
        if variant in ("A", "G"):
            results.append(Result(
                key="wet_correction", label="湿度修正项", value=(
                    -6.003 * v["bw"] * (273.15 + v["bt"]) / (6.11 * 1e7) if variant == 1
                    else -6.003 * v["bw"] / (273.15 + v["bt"]) * (6.11 * 10 ** 0.7)
                ), unit="℃", digits=6, symbol="ΔT_hum", primary=False,
                note="量级对比：公式A 的湿度项约 −0.0015 ℃（可忽略），"
                     "原表 G/L 写法约 −1.23 ℃（大 2 万倍）。",
            ))

        # 单因素灵敏度
        sens = numeric_partials(self._tf, v, ["bt", "bw", "vo2", "vb", "pc"])
        units = {"bt": "℃/℃", "bw": "℃/(g/m³ 或 %)", "vo2": "℃/(m³/h)",
                 "vb": "℃/(m³/min)", "pc": "℃/(kg/h)"}
        labels = {"bt": "热风温度", "bw": "鼓风湿度", "vo2": "富氧量",
                  "vb": "风量", "pc": "喷煤量"}
        for k in ("bt", "bw", "vo2", "vb", "pc"):
            results.append(Result(
                key=f"sens_{k}", label=f"∂Tf/∂{labels[k]}", value=sens[k],
                unit=units[k], digits=6, symbol="dTf/dx", primary=False,
                note="数值中心差分求得的偏导（以当前输入点为基准）。",
            ))
        return results


# ===========================================================================
# 鼓风动能
# ===========================================================================


class BlastMomentum(Model):
    key = "blast_momentum"
    title = "鼓风动能"
    subtitle = "鼓风所具有的动能及其在炉内的作用能力"
    group = "燃烧与鼓风"
    source_sheet = "鼓风动能"
    description = (
        "原表并列两套公式：左半区 D28 为经验式（结果量级 1.5~2.3 万，与综合表经验区间 "
        "「鼓风动能 15500±1000」对应）；右侧 N27 为按 0.5·ρ·v² 推导的标准动能式。"
        "**两者量纲不同**，下拉可切换。"
    )
    formula_latex = (
        r"E_k = 3.25\times10^{-8}\,(1.293-0.489\,BH)\,(273+BT)^2"
        r"\frac{BV^3}{n^3 S^2 (10.325+BP)^2}"
    )
    formula_excel = (
        "=3.25*10^-8*(1.293-0.489*BH)*(273+BT)^2*BV^3/n^3/S^2/(10.325+BP)^2"
    )
    notes = [
        "原表 J5 批注「公式有疑问」。",
        "右侧 N27 公式中 Pg 标注为 Mpa，但算例填 457（实为 kPa）；"
        "公式内 (0.101+Pg) 说明 Pg 的正确单位是 MPa，457 应写成 0.457。本工具右侧公式按 MPa 录入。",
        "两式结果不可混用：经验式量纲标注 kg·m/s；标准式 0.5·ρ·(Q₀/n)·v² 量纲实为功率（W）。",
    ]
    variant_label = "「公式版本」下拉切换经验式 / 标准动能式。"

    inputs = [
        Input(key="bh", label="湿分", symbol="BH", unit="%", default=0.0,
              step=0.1, fmt="%.6g", group="经验式（原表 D28）"),
        Input(key="bt1", label="风温", symbol="BT", unit="℃", default=1180,
              step=1.0, fmt="%.6g", group="经验式（原表 D28）"),
        Input(key="bv1", label="风量", symbol="BV", unit="m³/min", default=7608,
              step=1.0, fmt="%.8g", group="经验式（原表 D28）"),
        Input(key="n1", label="工作风口个数", symbol="n", unit="个", default=40,
              step=1.0, lo=1, fmt="%.6g", group="经验式（原表 D28）"),
        Input(key="s1", label="工作风口平均面积", symbol="S", unit="m²/个",
              default=0.0134775, step=1e-5, fmt="%.8g", group="经验式（原表 D28）"),
        Input(key="bp1", label="热风压力", symbol="BP", unit="kPa", default=457,
              step=1.0, fmt="%.6g", group="经验式（原表 D28）"),

        Input(key="rho", label="标准状态下风的密度", symbol="ρ", unit="kg/m³",
              default=1.2, step=0.01, fmt="%.6g", group="标准动能式（原表 N27）"),
        Input(key="q0", label="风量", symbol="Q0", unit="m³/min", default=7608,
              step=1.0, fmt="%.8g", group="标准动能式（原表 N27）"),
        Input(key="n2", label="风口数量", symbol="n", unit="个", default=40,
              step=1.0, lo=1, fmt="%.6g", group="标准动能式（原表 N27）"),
        Input(key="f", label="单个风口面积", symbol="f", unit="m²",
              default=0.0134775, step=1e-5, fmt="%.8g", group="标准动能式（原表 N27）"),
        Input(key="pg", label="风压", symbol="Pg", unit="kPa", default=457,
              step=1.0, fmt="%.6g", group="标准动能式（原表 N27）",
              help="原表标注 Mpa 却填了 457（实为 kPa）；公式内 (0.101+Pg) 与 457 相加存在量纲混用，"
                   "下拉可切换压力口径。"),
        Input(key="t", label="热风温度", symbol="t", unit="℃", default=1180,
              step=1.0, fmt="%.6g", group="标准动能式（原表 N27）"),

        Input(key="variant", label="公式版本", kind="select", default="exp",
              options=["exp", "std"],
              option_labels=["经验式（原表 D28，kg·m/s 量级）",
                             "标准动能式（原表 N27，0.5·ρ·(Q₀/n)·v²）"],
              group="输出"),
        Input(key="pg_mode", label="标准式压力口径", kind="select", default="raw",
              options=["raw", "fixed"],
              option_labels=["原表写法：(0.101 + Pg)，Pg 以 kPa 代入",
                             "量纲自洽：(101.325 + Pg)，Pg 以 kPa 代入"],
              group="输出",
              help="原表把标准大气压写成 0.101（MPa）却与 kPa 的 Pg 相加。"),
    ]

    reference = [
        {"key": "ek_exp", "label": "鼓风动能（经验式）", "symbol": "Ek", "unit": "kg·m/s",
         "excel": 15388.2042441461, "note": "原表 D28 算例"},
        {"key": "ek_std", "label": "鼓风动能（标准式）", "symbol": "E", "unit": "W 量级",
         "excel": 31433.1891898091, "note": "原表 N27 算例"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        ek_exp = (3.25e-8 * (1.293 - 0.489 * v["bh"]) * (273 + v["bt1"]) ** 2
                  * v["bv1"] ** 3 / v["n1"] ** 3 / v["s1"] ** 2
                  / (10.325 + v["bp1"]) ** 2)

        p_total = (0.101 + v["pg"]) if v.get("pg_mode", "raw") == "raw" \
            else (101.325 + v["pg"])
        v_blast = (v["q0"] * (273 + v["t"]) * 0.101 / v["n2"] / v["f"] / 273
                   / p_total)
        ek_std = 0.5 * v["rho"] * v["q0"] / v["n2"] * v_blast ** 2

        # 量纲自洽口径下的对照值
        v_blast_fix = (v["q0"] * (273 + v["t"]) * 0.101 / v["n2"] / v["f"] / 273
                       / (101.325 + v["pg"]))
        ek_std_fix = 0.5 * v["rho"] * v["q0"] / v["n2"] * v_blast_fix ** 2

        results = [
            Result(
                key="ek_exp", label="鼓风动能（经验式）", value=ek_exp,
                unit="kg·m/s", digits=1, symbol="Ek",
                rng=Range(14500, 16500, SRC_EXCEL,
                          "原表综合表「鼓风动能：15500±1000」。原表 F 列算例 22888 已超出该区间。"),
                formula="Ek = 3.25e-8·(1.293−0.489·BH)·(273+BT)²·BV³/n³/S²/(10.325+BP)²",
                note="与综合表经验区间口径一致。",
            ),
            Result(
                key="ek_std", label="鼓风动能（标准式）", value=ek_std,
                unit="W 量级", digits=1, symbol="E",
                formula="E = 0.5·ρ·(Q₀/n)·[Q₀·(273+t)·0.101/n/f/273/P总]²",
                note="量纲为 0.5ρv²×质量流率（W），与经验式不可直接比较；原表亦标注为 kg·m/s。"
                     + ("　当前 P总 = 0.101 + Pg（原表写法）。" if v.get("pg_mode", "raw") == "raw"
                        else "　当前 P总 = 101.325 + Pg（量纲自洽）。"),
            ),
            Result(
                key="v_blast", label="单风口实际风速", value=v_blast, unit="m/s",
                digits=2, symbol="v", primary=False,
                formula="v = Q₀·(273+t)·0.101/n/f/273/P总",
                note="由标准动能式反推的单风口风速，可与「实际风速」Tab 交叉验证。",
            ),
            Result(
                key="ek_std_fixed", label="鼓风动能（标准式·量纲自洽）", value=ek_std_fix,
                unit="W 量级", digits=1, symbol="E", primary=False,
                formula="E = 0.5·ρ·(Q₀/n)·[Q₀·(273+t)·0.101/n/f/273/(101.325+Pg)]²",
                note="与原表写法的差异全部来自标准大气压项 (0.101 vs 101.325 kPa)。",
            ),
        ]
        if v.get("variant", "exp") == "std":
            results[0].primary = False
            results[1].primary = True
        return results


# ===========================================================================
# 实际风速
# ===========================================================================


class ActualBlastVelocity(Model):
    key = "actual_velocity"
    title = "实际风速"
    subtitle = "风口平面的实际（热态、实压）风速"
    group = "燃烧与鼓风"
    source_sheet = "实际风速"
    description = (
        "把标况风量与富氧折算到实际温度压力下，再除以风口总面积，"
        "并作湿度密度修正。"
    )
    formula_latex = (
        r"v = \frac{BV/60 + O_2/3600}{S_{风口}}"
        r"\cdot\frac{273.15+BT}{273.15}\cdot\frac{101.325}{101.325+BP}"
        r"\cdot\frac{1}{1 - BH\,(273.15+BT)/(6.11\times10^{7}\,/\,804)}"
    )
    formula_excel = (
        "=(BV/60+O2/3600)/S风口*(273.15+BT)/273.15*101.325/(101.325+BP)"
        "*1/(1-BH*(273.15+BT)/6.11/10000000/804)"
    )
    notes = [
        "公式中 BH 的单位标注为 %，但湿度修正项按 g/m³ 量纲（除以 6.11×10⁷）书写，"
        "两种定义混用；不过该项在 15% 输入下仅约 4.5×10⁻⁷，对结果无实质影响。",
        "结果量级约 250~300 m/s，与现代高炉实际风口风速（150~280 m/s）相符；"
        "注意不要与标况风速（BV/S）混淆。",
    ]

    inputs = [
        Input(key="bv", label="风量", symbol="BV", unit="Nm³/min", default=3397,
              step=1.0, fmt="%.8g"),
        Input(key="o2", label="氧量", symbol="O2", unit="Nm³/h", default=20000,
              step=100.0, fmt="%.8g"),
        Input(key="sf", label="风口面积", symbol="S风口", unit="m²", default=0.218,
              step=0.001, fmt="%.6g"),
        Input(key="bh", label="湿分", symbol="BH", unit="%", default=15,
              step=0.1, fmt="%.6g"),
        Input(key="bt", label="风温", symbol="BT", unit="℃", default=1213,
              step=1.0, fmt="%.6g"),
        Input(key="bp", label="风压", symbol="BP", unit="kPa", default=432,
              step=1.0, fmt="%.6g"),
    ]

    reference = [
        {"key": "v", "label": "实际风速", "symbol": "v", "unit": "m/s",
         "excel": 294.799135394861, "note": "原表 D28 算例，复算一致"},
        {"key": "v_std", "label": "标况风速", "symbol": "v0", "unit": "m/s",
         "excel": None, "note": "BV/S = 3397/60/0.218 ≈ 259.7 m/s"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        flow = v["bv"] / 60.0 + v["o2"] / 3600.0
        v_act = (flow / v["sf"]
                 * (273.15 + v["bt"]) / 273.15
                 * 101.325 / (101.325 + v["bp"])
                 * 1.0 / (1 - v["bh"] * (273.15 + v["bt"]) / 6.11 / 1e7 / 804))
        v_std = flow / v["sf"]
        return [
            Result(
                key="v", label="实际风速", value=v_act, unit="m/s", digits=2, symbol="v",
                rng=Range(150, 280, SRC_TYPICAL,
                          "现代大高炉实际风口风速常见区间；喷煤比高、鼓风强度大的炉子偏上限。"),
                formula="v = (BV/60+O₂/3600)/S风口 ·(273.15+BT)/273.15 ·101.325/(101.325+BP) ·[1/(1−BH·(273.15+BT)/6.11/10⁷/804)]",
            ),
            Result(
                key="v_std", label="标况风速", value=v_std, unit="m/s", digits=2,
                symbol="v0", primary=False,
                formula="v₀ = (BV/60 + O₂/3600) / S风口",
                note="不含温度压力换算，仅作对照。",
            ),
        ]


# ===========================================================================
# 透气性指数
# ===========================================================================


class PermeabilityIndex(Model):
    key = "permeability"
    title = "透气性指数"
    subtitle = "以风量与压差之比表征炉料/煤气流通过的顺畅程度"
    group = "燃烧与鼓风"
    source_sheet = "透气性指数"
    formula_latex = r"T_z = \frac{BV}{BP - TP}"
    formula_excel = "=BV/(BP-TP)"
    notes = [
        "原表未给出经验区间，本工具给出的是通用参考值（透气性指数高通常对应炉况顺行、"
        "风量富余），请按本厂历史数据重新标定。",
    ]

    inputs = [
        Input(key="bp", label="风压", symbol="BP", unit="kPa", default=441,
              step=1.0, fmt="%.6g"),
        Input(key="tp", label="顶压", symbol="TP", unit="kPa", default=246,
              step=1.0, fmt="%.6g"),
        Input(key="bv", label="风量", symbol="BV", unit="Nm³/min", default=3831,
              step=1.0, fmt="%.8g"),
    ]

    reference = [
        {"key": "tz", "label": "透气性指数", "symbol": "Tz", "unit": "m³/(kPa·min)",
         "excel": 19.6461538461538, "note": "原表 D23 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        dp = v["bp"] - v["tp"]
        tz = v["bv"] / dp if dp != 0 else float("nan")
        return [
            Result(
                key="tz", label="透气性指数", value=tz, unit="m³/(kPa·min)",
                digits=3, symbol="Tz",
                rng=Range(18, 30, SRC_TYPICAL,
                          "一般认为 >20 透气性良好；取值随风量、顶压操作制度变化，需按厂标定。"),
                formula="Tz = BV / (BP − TP)",
            ),
            Result(
                key="dp", label="压差", value=dp, unit="kPa", digits=2,
                symbol="ΔP", primary=False,
                formula="ΔP = BP − TP",
            ),
        ]


# ===========================================================================
# 富氧率
# ===========================================================================


class OxygenEnrichment(Model):
    key = "oxygen_rate"
    title = "富氧率"
    subtitle = "鼓风含氧量相对自然空气的提高幅度"
    group = "燃烧与鼓风"
    source_sheet = "富氧率"
    formula_latex = (
        r"VO_2 = 100\left[\frac{0.21\left(BV - O_2/60\right) + O_2/60}{BV} - 0.21\right]"
    )
    formula_excel = "=100*((0.21*(BV-O2/60)+O2/60)/BV-0.21)"
    notes = [
        "原表 D25 批注：「这三个数据，第一个能对上，后两个结果对不上」。"
        "E/F 两例的富氧率偏差较大，建议核对现场报表的富氧量口径（标况/工况）。",
    ]

    inputs = [
        Input(key="bv", label="风量", symbol="BV", unit="Nm³/min", default=3800,
              step=1.0, fmt="%.8g"),
        Input(key="o2", label="氧量", symbol="O2", unit="Nm³/h", default=23000,
              step=100.0, fmt="%.8g"),
    ]

    reference = [
        {"key": "vo2", "label": "富氧率", "symbol": "VO2", "unit": "%",
         "excel": 7.96929824561403, "note": "原表 D23 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        bv, o2 = v["bv"], v["o2"]
        o2_min = o2 / 60.0
        vo2 = 100.0 * ((0.21 * (bv - o2_min) + o2_min) / bv - 0.21)
        o2_total = 0.21 * (bv - o2_min) + o2_min
        return [
            Result(
                key="vo2", label="富氧率", value=vo2, unit="%", digits=3, symbol="VO2",
                rng=Range(2, 8, SRC_TYPICAL,
                          "常规高炉富氧率 2%~8%，富氧强度大的炉子可更高。"),
                formula="VO2 = 100·[(0.21·(BV−O₂/60) + O₂/60)/BV − 0.21]",
            ),
            Result(
                key="o2_total", label="鼓风总氧量", value=o2_total, unit="Nm³/min",
                digits=2, symbol="O2总", primary=False,
                formula="O₂总 = 0.21·(BV − O₂/60) + O₂/60",
            ),
        ]


# ===========================================================================
# 热负荷
# ===========================================================================


class HeatLoad(Model):
    key = "heat_load"
    title = "热负荷"
    subtitle = "冷却水带走的热量（炉体热负荷）"
    group = "燃烧与鼓风"
    source_sheet = "热负荷"
    formula_latex = r"Q = (T_{出} - T_{进}) \times 4.18 \times Q_0 \quad [\mathrm{MJ/h}]"
    formula_excel = "=(T出-T进)*4.18*Q0"
    notes = [
        "原表单位标注为「10MJ/h」，C24 批注「单位是否有问题，应为 MJ/h」。"
        "本工具按 MJ/h 输出：Q₀(m³/h)×4.18(kJ/kg·℃)×1000(kg/m³) ÷ 1000 = MJ/h，量纲自洽。",
        "原表 D22 冷却水流量 879.9 m³/h 数值偏大（通常为 m³/h 量级的小时流量，请核对单位与取数口径）。",
    ]

    inputs = [
        Input(key="t_out", label="冷却水出水温度", symbol="T出", unit="℃",
              default=41.06, step=0.01, fmt="%.6g"),
        Input(key="t_in", label="冷却水进水温度", symbol="T进", unit="℃",
              default=38.6, step=0.01, fmt="%.6g"),
        Input(key="q0", label="冷却水流量", symbol="Q0", unit="m³/h", default=879.9,
              step=0.1, fmt="%.6g"),
    ]

    reference = [
        {"key": "q", "label": "热负荷", "symbol": "Q", "unit": "MJ/h",
         "excel": 9047.83572, "note": "原表 D23 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        dt = v["t_out"] - v["t_in"]
        q = dt * 4.18 * v["q0"]
        return [
            Result(
                key="q", label="热负荷", value=q, unit="MJ/h", digits=1, symbol="Q",
                rng=Range(3000, 15000, SRC_TYPICAL,
                          "炉体热负荷参考区间，随炉缸直径、冷却水量与水温差变化；请按本厂标定。"),
                formula="Q = (T出 − T进) × 4.18 × Q0   [MJ/h]",
                note="= 2.5 MW 级（按 9047 MJ/h 折算）。",
            ),
            Result(
                key="dt", label="冷却水温差", value=dt, unit="℃", digits=3,
                symbol="δt", primary=False, formula="δt = T出 − T进",
            ),
        ]


MODELS = [
    TheoreticalCombustionTemperature(),
    BlastMomentum(),
    ActualBlastVelocity(),
    PermeabilityIndex(),
    OxygenEnrichment(),
    HeatLoad(),
]
