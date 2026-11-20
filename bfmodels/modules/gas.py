"""煤气流动类模型。

对应 Excel 工作表：炉腹煤气 / 炉腹煤气指数 / 煤气流速 / 炉顶煤气量。
"""

from __future__ import annotations

import math
from typing import Any, Dict, List

from ..base import SRC_EXCEL, SRC_REPORT, SRC_TYPICAL, Input, Model, Range, Result

# 与 Excel 原表保持一致（3.1415926 而非 math.pi），确保复算结果逐位对齐
PI = 3.1415926


# ===========================================================================
# 炉腹煤气量
# ===========================================================================


class BellyGasVolume(Model):
    key = "belly_gas"
    title = "炉腹煤气量"
    subtitle = "到达炉腹的煤气体积流量"
    group = "煤气流动"
    source_sheet = "炉腹煤气"
    description = (
        "炉腹煤气量由三部分构成：焦炭燃烧产气、喷吹物挥发分/含氢产气、以及鼓风中的水分汽化部分。"
    )
    formula_latex = (
        r"V_{bg} = 1.21\,BV + VO_2 "
        r"+ 44.8\,BH\,(273.15+BT)\,(BV+VO_2)\big/(6.11\times18000\times10^{7}) "
        r"+ 22.4\,PCI\,H/120"
    )
    formula_excel = (
        "=1.21*BV+VO2+44.8*BH*(273.15+BT)*(BV+VO2)/6.11/18000/10000000"
        "+22.4*PCI*H/120"
    )
    notes = [
        "原表 I23 批注：「这里数值为斌鑫数据，但是得不到斌鑫的结果，若煤粉含氢量均为 0，"
        "则结果与斌鑫数据一致。」即含氢产气项是差异来源（算例中该项 +4.98 m³/min）。",
        "原表 D19 的 VO2 单位标注 m³/h，但公式中直接与 m³/min 的项相加，未做 /60 折算。"
        "本工具提供开关复现两种口径。",
        "水分汽化项 44.8·BH·(273.15+BT)·(BV+VO₂)/6.11/18000/10⁷ 在算例下仅 0.0025 m³/min，"
        "量级可忽略；若湿度项按 g/m³ 定义需重新推导分母。",
    ]
    variant_label = "「富氧量口径」开关用于复现原表（不折算）与量纲自洽（/60）两种写法。"

    inputs = [
        Input(key="bv", label="风量", symbol="BV", unit="Nm³/min", default=3800,
              step=1.0, fmt="%.8g"),
        Input(key="vo2", label="富氧量", symbol="VO2", unit="m³/h", default=375,
              step=1.0, fmt="%.8g",
              help="原表按 m³/h 填写，但公式中未 /60 折算为 m³/min。"),
        Input(key="bh", label="湿分", symbol="BH", unit="%", default=10,
              step=0.1, fmt="%.6g"),
        Input(key="bt", label="风温", symbol="BT", unit="℃", default=1219,
              step=1.0, fmt="%.6g"),
        Input(key="pci", label="煤量", symbol="PCI", unit="t/h", default=44.5,
              step=0.1, fmt="%.6g"),
        Input(key="hh", label="煤粉含氢量", symbol="H", unit="%", default=0.6,
              step=0.05, fmt="%.6g"),
        Input(key="vo2_div60", label="富氧量口径", kind="select", default="raw",
              options=["raw", "div60"],
              option_labels=["按原表：不折算（VO2 直接相加）",
                             "量纲自洽：VO2/60 折算为 m³/min"],
              help="切换以观察该口径对结果的影响。"),
        Input(key="with_h", label="计入喷煤含氢产气项", kind="bool", default=True,
              help="取消勾选可复现「含氢量均为 0 时与斌鑫数据一致」的结论。"),
    ]

    reference = [
        {"key": "vbg", "label": "炉腹煤气量", "symbol": "Vbg", "unit": "m³/min",
         "excel": 4977.98653765899, "note": "原表 D24 算例（默认参数），复算一致"},
        {"key": "vbg_no_h", "label": "炉腹煤气量（不含氢产气）", "symbol": "Vbg", "unit": "m³/min",
         "excel": None, "note": "原表批注指出此时与斌鑫数据一致（≈4973 m³/min）"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        bv = v["bv"]
        vo2 = v["vo2"] / 60.0 if v.get("vo2_div60") == "div60" else v["vo2"]
        pci = v["pci"]
        h = v["hh"] if v.get("with_h", True) else 0.0

        term_coke = 1.21 * bv
        term_o2 = vo2
        term_moist = (44.8 * v["bh"] * (273.15 + v["bt"]) * (bv + vo2)
                      / 6.11 / 18000 / 1e7)
        term_h2 = 22.4 * pci * h / 120

        vbg = term_coke + term_o2 + term_moist + term_h2
        vbg_no_h = vbg - term_h2

        return [
            Result(
                key="vbg", label="炉腹煤气量", value=vbg, unit="m³/min", digits=1,
                symbol="Vbg",
                rng=Range(3800, 6200, SRC_TYPICAL,
                          "炉腹煤气量通常为鼓风量的 1.0~1.2 倍（含富氧与喷煤产气）；"
                          "请按本厂炉况标定。"),
                formula="Vbg = 1.21·BV + VO₂ + 44.8·BH·(273.15+BT)·(BV+VO₂)/6.11/18000/10⁷ + 22.4·PCI·H/120",
            ),
            Result(
                key="vbg_no_h", label="炉腹煤气量（不计含氢产气）", value=vbg_no_h,
                unit="m³/min", digits=1, symbol="Vbg", primary=False,
                formula="Vbg' = Vbg − 22.4·PCI·H/120",
                note="原表批注：含氢量取 0 时结果与斌鑫数据一致。",
            ),
            Result(
                key="ratio", label="炉腹煤气量 / 鼓风量", value=vbg / bv if bv else float("nan"),
                unit="", digits=4, symbol="Vbg/BV", primary=False,
                formula="Vbg / BV",
            ),
            Result(
                key="term_h2", label="含氢产气项", value=term_h2, unit="m³/min",
                digits=4, symbol="22.4·PCI·H/120", primary=False,
                formula="22.4·PCI·H/120",
            ),
            Result(
                key="term_moist", label="水分汽化项", value=term_moist, unit="m³/min",
                digits=6, symbol="ΔV水", primary=False,
                formula="44.8·BH·(273.15+BT)·(BV+VO₂)/6.11/18000/10⁷",
                note="量级 0.0025 m³/min，对结果无实质影响。",
            ),
        ]


# ===========================================================================
# 炉腹煤气指数
# ===========================================================================


class BellyGasIndex(Model):
    key = "belly_gas_index"
    title = "炉腹煤气指数"
    subtitle = "单位炉缸面积的炉腹煤气流量"
    group = "煤气流动"
    source_sheet = "炉腹煤气指数"
    formula_latex = r"\alpha_{bg} = \frac{4\,V_{bg}}{\pi\,d^2}"
    formula_excel = "=4*Vbg/3.1415926/d/d"
    notes = [
        "原表算例（Vbg=5600、d=8.5）得 98.69，明显高于经验区间 58~65，"
        "说明该算例的炉腹煤气量或炉缸直径与本厂实际不匹配，需核对取数。",
    ]

    inputs = [
        Input(key="vbg", label="炉腹煤气量", symbol="Vbg", unit="m³/min", default=5600,
              step=10.0, fmt="%.8g"),
        Input(key="d", label="炉缸直径", symbol="d", unit="m", default=8.5,
              step=0.1, fmt="%.6g"),
    ]

    reference = [
        {"key": "abg", "label": "炉腹煤气指数", "symbol": "αbg", "unit": "m³/(min·m²)",
         "excel": 98.6870805833081, "note": "原表 D21 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        area = PI * v["d"] ** 2 / 4
        abg = 4 * v["vbg"] / PI / v["d"] / v["d"]
        return [
            Result(
                key="abg", label="炉腹煤气指数", value=abg, unit="m³/(min·m²)",
                digits=3, symbol="αbg",
                rng=Range(58, 65, SRC_EXCEL,
                          "原表综合表给出「指数范围：62±2，58-65」。"),
                formula="αbg = 4·Vbg/(π·d²)",
            ),
            Result(
                key="area", label="炉缸截面积", value=area, unit="m²", digits=3,
                symbol="S缸", primary=False, formula="S = π·d²/4",
            ),
        ]


# ===========================================================================
# 煤气流速
# ===========================================================================


class GasVelocity(Model):
    key = "gas_velocity"
    title = "煤气流速"
    subtitle = "炉缸内煤气的平均流速"
    group = "煤气流动"
    source_sheet = "煤气流速"
    formula_latex = (
        r"v = \frac{V_{炉缸}\,P\,(273+T_f)\,101.325}{86400\,\varepsilon\,S_{炉缸}\,(101.325+BP)\,273}"
    )
    formula_excel = (
        "=V炉缸*P*(273+Tf)*101.325/86400/eps/S炉缸/(101.325+BP)/273"
    )
    notes = [
        "原表 I28 批注：「公式没问题，且与文献中公式一致，但是结果对应不上，报表中显示为 2.6」。"
        "本工具复算 5.25 m/s，与报表 2.6 相差约 2 倍，需核对 ε（炉料空隙系数）、"
        "炉缸截面积与产量的口径是否一致。",
    ]

    inputs = [
        Input(key="v_h", label="炉缸容积", symbol="V炉缸", unit="m³", default=1250,
              step=10.0, fmt="%.8g"),
        Input(key="p", label="高炉产量", symbol="P", unit="t/d", default=5661,
              step=1.0, fmt="%.8g"),
        Input(key="eps", label="炉料空隙系数", symbol="ε", unit="", default=0.55,
              step=0.01, fmt="%.6g"),
        Input(key="s_h", label="炉缸横截面积", symbol="S炉缸", unit="m²", default=51.5,
              step=0.1, fmt="%.6g"),
        Input(key="tf", label="理论燃烧温度", symbol="Tf", unit="℃", default=2382,
              step=1.0, fmt="%.6g"),
        Input(key="bp", label="热风压力", symbol="BP", unit="kPa", default=441,
              step=1.0, fmt="%.6g"),
    ]

    reference = [
        {"key": "v", "label": "煤气流速", "symbol": "v", "unit": "m/s",
         "excel": 5.25385850979254, "note": "原表 D28 算例，复算一致"},
        {"key": "v_report", "label": "现场报表值", "symbol": "v报表", "unit": "m/s",
         "excel": 2.6, "note": "原表批注给出的报表值，公式与该值不匹配"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        vel = (v["v_h"] * v["p"] * (273 + v["tf"]) * 101.325 / 86400
               / v["eps"] / v["s_h"] / (101.325 + v["bp"]) / 273)
        return [
            Result(
                key="v", label="煤气流速", value=vel, unit="m/s", digits=3, symbol="v",
                rng=Range(2.0, 4.0, SRC_REPORT,
                          "按现场报表 2.6 m/s 给出参考区间；原表公式复算为 5.25 m/s，超出上限。"),
                formula="v = V炉缸·P·(273+Tf)·101.325/(86400·ε·S炉缸·(101.325+BP)·273)",
                note="与现场报表 2.6 m/s 相差约 2 倍，待查。",
            ),
            Result(
                key="v_report", label="现场报表值（对照）", value=2.6, unit="m/s",
                digits=3, symbol="v报表", primary=False,
                note="原表批注记录，仅作对照，不参与计算。",
            ),
        ]


# ===========================================================================
# 炉顶煤气量
# ===========================================================================


class TopGasVolume(Model):
    key = "top_gas"
    title = "炉顶煤气量"
    subtitle = "氮气平衡法计算的高炉煤气总量"
    group = "煤气流动"
    source_sheet = "炉顶煤气量"
    formula_latex = r"Q_{炉顶煤气} = BV \cdot \frac{N_{2,air}}{N_{2,gas}}"
    formula_excel = "=BV*N2air/N2gas"
    notes = [
        "氮气平衡法假设：鼓入空气中的氮全部进入炉顶煤气（不计入喷吹带入的氮与耗损）。"
        "喷煤比高时该假设会引入偏差。",
    ]

    inputs = [
        Input(key="bv", label="风量", symbol="BV", unit="m³/h", default=3755,
              step=1.0, fmt="%.8g"),
        Input(key="n2air", label="空气中氮气含量占比", symbol="N2air", unit="%",
              default=78, step=0.1, fmt="%.6g"),
        Input(key="n2gas", label="高炉煤气中氮气比例", symbol="N2gas", unit="%",
              default=44.69, step=0.01, fmt="%.6g"),
    ]

    reference = [
        {"key": "q", "label": "炉顶煤气量", "symbol": "Q炉顶煤气量", "unit": "m³/h",
         "excel": 6553.81517117924, "note": "原表 D21 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        q = v["bv"] * v["n2air"] / v["n2gas"]
        return [
            Result(
                key="q", label="炉顶煤气量", symbol="Q炉顶煤气量", value=q, unit="m³/h",
                digits=1,
                rng=Range(1.55 * v["bv"], 1.95 * v["bv"], SRC_TYPICAL,
                          "按炉顶煤气量/鼓风量 1.55~1.95 的经验比例给出（随喷煤比变化）。"),
                formula="Q = BV·N2air/N2gas",
            ),
            Result(
                key="ratio", label="炉顶煤气量 / 鼓风量", value=q / v["bv"] if v["bv"] else float("nan"),
                unit="", digits=4, symbol="Q/BV", primary=False,
                formula="Q / BV",
            ),
        ]


MODELS = [
    BellyGasVolume(),
    BellyGasIndex(),
    GasVelocity(),
    TopGasVolume(),
]
