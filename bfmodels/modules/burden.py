"""冶炼过程类模型。

对应 Excel 工作表：冶炼强度 / 冶炼周期（批）/ 冶炼周期（h）/ 煤热滞后。
"""

from __future__ import annotations

from typing import Any, Dict, List

from ..base import SRC_EXCEL, SRC_TYPICAL, Input, Model, Range, Result


# ===========================================================================
# 冶炼强度
# ===========================================================================


class SmeltingIntensity(Model):
    key = "smelting_intensity"
    title = "冶炼强度"
    subtitle = "单位高炉有效容积的单位时间焦炭消耗量"
    group = "冶炼过程"
    source_sheet = "冶炼强度"
    formula_latex = (
        r"I = \frac{C_h + 0.9\,C_R + 0.8\,PCI}{V\,(T_{日历} - T_{休风})}"
    )
    formula_excel = "=(Ch+0.9*CR+0.8*PCI)*(T日历-T休风)/V"
    notes = [
        "原表 J4 批注：「首先公式有误，应该是除以时间不是乘。」本工具默认采用修正式"
        "（除以有效冶炼天数），下拉可切回原表写法对比。",
        "原表 J5 批注：「按照公式将焦丁和煤折算，无法得到镔鑫的结果，若只计算干焦，"
        "可以得到镔鑫的结果」——即 0.9/0.8 的折算系数是差异来源。",
        "原表算例中 T日历−T休风 = 1，两种写法数值恰好相同（均为 1.3524），"
        "无法据此判断正误，故两种写法都保留。",
    ]
    variant_label = "「公式版本」下拉切换 修正式（÷时间）/ 原表式（×时间）。"

    inputs = [
        Input(key="ch", label="干焦量", symbol="Ch", unit="t", default=1690.527,
              step=1.0, fmt="%.8g"),
        Input(key="cr", label="焦丁量", symbol="CR", unit="t", default=0.0,
              step=1.0, fmt="%.8g"),
        Input(key="pci", label="煤量", symbol="PCI", unit="t", default=0.0,
              step=1.0, fmt="%.8g"),
        Input(key="v", label="高炉有效容积", symbol="V", unit="m³", default=1250,
              step=10.0, fmt="%.8g"),
        Input(key="t_cal", label="日历天数", symbol="T日历", unit="天", default=2,
              step=0.5, fmt="%.6g"),
        Input(key="t_rest", label="休风天数", symbol="T休风", unit="天", default=1,
              step=0.5, fmt="%.6g"),
        Input(key="variant", label="公式版本", kind="select", default="fixed",
              options=["fixed", "raw"],
              option_labels=["修正式：÷ (V × 有效天数)",
                             "原表式：× 有效天数 ÷ V"],
              group="输出"),
    ]

    reference = [
        {"key": "i", "label": "冶炼强度", "symbol": "I", "unit": "t/(m³·d)",
         "excel": 1.3524216, "note": "原表 D29 算例，复算一致（两式在此算例下同值）"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        days = v["t_cal"] - v["t_rest"]
        coke = v["ch"] + 0.9 * v["cr"] + 0.8 * v["pci"]
        if v.get("variant", "fixed") == "fixed":
            intensity = coke / (v["v"] * days) if days else float("nan")
            formula = "I = (Ch + 0.9·CR + 0.8·PCI) / (V · (T日历 − T休风))"
        else:
            intensity = coke * days / v["v"]
            formula = "I = (Ch + 0.9·CR + 0.8·PCI) × (T日历 − T休风) / V（原表写法，量纲存疑）"
        return [
            Result(
                key="i", label="冶炼强度", value=intensity, unit="t/(m³·d)", digits=4,
                symbol="I",
                rng=Range(0.90, 1.60, SRC_TYPICAL,
                          "现代大高炉 1.1~1.6 t/(m³·d)；小高炉偏低。请按本厂标定。"),
                formula=formula,
            ),
            Result(
                key="coke", label="折算焦炭量", value=coke, unit="t", digits=1,
                symbol="折算焦", primary=False,
                formula="Ch + 0.9·CR + 0.8·PCI",
            ),
            Result(
                key="days", label="有效冶炼天数", value=days, unit="天", digits=2,
                symbol="T有效", primary=False, formula="T日历 − T休风",
            ),
        ]


# ===========================================================================
# 冶炼周期（批）
# ===========================================================================


class SmeltingCycleBatch(Model):
    key = "smelting_cycle_batch"
    title = "冶炼周期（批）"
    subtitle = "高炉有效容积对应的料批数与料速"
    group = "冶炼过程"
    source_sheet = "冶炼周期（批）"
    formula_latex = (
        r"t_{批} = \frac{V\,(1-k)}{\displaystyle\sum_i \frac{m_i}{a_i}}"
    )
    formula_excel = "=V*0.835/(m1*a1+m2*a2+m3*a3+m4*a4+m5*a5)/(1-k)"
    notes = [
        "原表 J13 批注：「公式是否应该是质量除以体积而不是乘」——原表把 Σ(mᵢ·aᵢ)"
        "（质量×堆比重）放在分母，量纲为 t²/m³，正确写法应为 Σ(mᵢ/aᵢ)（体积）。"
        "下拉可切换两种写法。",
        "两种写法在算例下相差约 9 倍（9.48 批 vs 85.28 批）。"
        "原表式的 9.48 与「冶炼周期（h）」表中的「小时料速 6.8 批/h」同量级，"
        "说明原表实际把它当作「料速」使用，而标题写作「冶炼周期（批）」，命名与用途不一致。",
        "系数 0.835 是高炉有效容积中料柱所占的比例（容积利用系数）。",
    ]
    variant_label = "「公式版本」下拉切换 修正式（质量÷体积）/ 原表式（质量×体积）。"

    inputs = [
        Input(key="v", label="高炉有效容积", symbol="V", unit="m³", default=2157,
              step=10.0, fmt="%.8g", group="总体参数"),
        Input(key="k", label="炉料在炉内的压缩率", symbol="k", unit="", default=0.12,
              step=0.01, fmt="%.6g", group="总体参数"),
        Input(key="m1", label="烧结矿重量", symbol="m1", unit="t", default=40,
              step=1.0, fmt="%.6g", group="炉料组成"),
        Input(key="a1", label="烧结矿堆比重", symbol="a1", unit="t/m³", default=3,
              step=0.1, fmt="%.6g", group="炉料组成"),
        Input(key="m2", label="球团矿重量", symbol="m2", unit="t", default=10,
              step=1.0, fmt="%.6g", group="炉料组成"),
        Input(key="a2", label="球团矿堆比重", symbol="a2", unit="t/m³", default=3,
              step=0.1, fmt="%.6g", group="炉料组成"),
        Input(key="m3", label="块矿重量", symbol="m3", unit="t", default=10,
              step=1.0, fmt="%.6g", group="炉料组成"),
        Input(key="a3", label="块矿堆比重", symbol="a3", unit="t/m³", default=3,
              step=0.1, fmt="%.6g", group="炉料组成"),
        Input(key="m4", label="焦炭重量", symbol="m4", unit="t", default=12,
              step=1.0, fmt="%.6g", group="炉料组成"),
        Input(key="a4", label="焦炭堆比重", symbol="a4", unit="t/m³", default=3,
              step=0.1, fmt="%.6g", group="炉料组成"),
        Input(key="m5", label="辅料重量", symbol="m5", unit="t", default=0,
              step=1.0, fmt="%.6g", group="炉料组成"),
        Input(key="a5", label="辅料堆比重", symbol="a5", unit="t/m³", default=1,
              step=0.1, fmt="%.6g", group="炉料组成"),
        Input(key="variant", label="公式版本", kind="select", default="raw",
              options=["raw", "fixed"],
              option_labels=["原表式：Σ(mᵢ·aᵢ) 质量×比重口径（复刻原表）",
                             "修正式：Σ(mᵢ/aᵢ) 体积口径（批注所指）"],
              group="输出"),
    ]

    reference = [
        {"key": "t_batch", "label": "冶炼周期（批）", "symbol": "t(批)", "unit": "批",
         "excel": 9.4754577020202, "note": "原表 D43 算例（原表式），复算一致"},
        {"key": "t_batch_fixed", "label": "冶炼周期（批）·修正式", "symbol": "t(批)",
         "unit": "批", "excel": None, "note": "修正式算例值 85.28 批，量纲为批"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        pairs = [(v["m1"], v["a1"]), (v["m2"], v["a2"]), (v["m3"], v["a3"]),
                 (v["m4"], v["a4"]), (v["m5"], v["a5"])]
        vol_sum = sum(m / a for m, a in pairs if a)
        mass_sum = sum(m * a for m, a in pairs)

        raw = v["v"] * 0.835 / mass_sum / (1 - v["k"]) if mass_sum else float("nan")
        fixed = v["v"] * 0.835 / vol_sum / (1 - v["k"]) if vol_sum else float("nan")

        if v.get("variant", "raw") == "fixed":
            value, formula = fixed, "t(批) = V·0.835 / Σ(mᵢ/aᵢ) / (1−k)"
            note = "修正式：分母为料批体积（质量÷堆比重），与批注一致。"
        else:
            value, formula = raw, "t(批) = V·0.835 / Σ(mᵢ·aᵢ) / (1−k)"
            note = "原表式：分母为质量×堆比重（量纲 t²/m³），存疑；结果实际是「料速」量级。"

        return [
            Result(
                key="t_batch", label="冶炼周期（批）", value=value, unit="批", digits=4,
                symbol="t(批)",
                rng=Range(5.0, 12.0, SRC_TYPICAL,
                          "该量实际对应「料速（批/h）」的常见范围；原表标题与下游用法不一致。"),
                formula=formula,
                note=note,
            ),
            Result(
                key="t_batch_fixed", label="冶炼周期（批）·修正式", value=fixed, unit="批",
                digits=4, symbol="t(批)", primary=False,
                formula="t(批) = V·0.835/Σ(mᵢ/aᵢ)/(1−k)",
            ),
            Result(
                key="t_batch_raw", label="冶炼周期（批）·原表式", value=raw, unit="批",
                digits=4, symbol="t(批)", primary=False,
                formula="t(批) = V·0.835/Σ(mᵢ·aᵢ)/(1−k)",
            ),
            Result(
                key="vol_sum", label="料批体积", value=vol_sum, unit="m³", digits=4,
                symbol="Σm/a", primary=False, formula="Σ(mᵢ/aᵢ)",
            ),
            Result(
                key="mass_sum", label="Σ(质量×堆比重)", value=mass_sum, unit="t²/m³",
                digits=2, symbol="Σm·a", primary=False, formula="Σ(mᵢ·aᵢ)",
            ),
        ]


# ===========================================================================
# 冶炼周期（h）
# ===========================================================================


class SmeltingCycleHours(Model):
    key = "smelting_cycle_hours"
    title = "冶炼周期（h）"
    subtitle = "一次出铁间隔对应的冶炼时间"
    group = "冶炼过程"
    source_sheet = "冶炼周期（h)"
    formula_latex = r"t_{(h)} = \frac{t_{(批)}}{\text{小时料速}}"
    formula_excel = "=t(批)/小时料速"
    notes = [
        "本工具严格按「各 Tab 独立」实现：冶炼周期（批）与小时料速均为独立输入。",
        "原表算例 42.3 批 ÷ 6.8 批/h = 6.22 h，与现代大高炉冶炼周期（4~8 h）相符。",
    ]

    inputs = [
        Input(key="t_batch", label="冶炼周期（批）", symbol="t(批)", unit="批",
              default=42.3, step=0.1, fmt="%.6g"),
        Input(key="rate", label="小时料速", symbol="批/h", unit="批/h", default=6.8,
              step=0.1, fmt="%.6g"),
    ]

    reference = [
        {"key": "t_h", "label": "冶炼周期（h）", "symbol": "t(h)", "unit": "h",
         "excel": 6.22058823529412, "note": "原表 C20 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        t_h = v["t_batch"] / v["rate"] if v["rate"] else float("nan")
        return [
            Result(
                key="t_h", label="冶炼周期（h）", value=t_h, unit="h", digits=3,
                symbol="t(h)",
                rng=Range(4.0, 8.0, SRC_TYPICAL,
                          "现代大高炉冶炼周期一般 4~8 h；大炉偏上限、小炉偏下限。"),
                formula="t(h) = t(批) / 小时料速",
            ),
        ]


# ===========================================================================
# 煤热滞后
# ===========================================================================


class CoalThermalLag(Model):
    key = "coal_thermal_lag"
    title = "煤热滞后"
    subtitle = "喷吹物从风口到完成燃烧所需的时间"
    group = "冶炼过程"
    source_sheet = "煤热滞后"
    formula_latex = r"t = \frac{V}{V_{料批}\cdot v}"
    formula_excel = "=V/V料批/v"
    notes = [
        "煤热滞后时间决定喷煤在炉内的实际热效率：滞后越长，煤粉未完全燃烧、"
        "利用率越低。",
    ]

    inputs = [
        Input(key="v", label="高炉反应区至风口总容积", symbol="V", unit="m³",
              default=2000, step=10.0, fmt="%.8g"),
        Input(key="v_batch", label="料批体积", symbol="V料批", unit="m³/批",
              default=12, step=0.1, fmt="%.6g"),
        Input(key="rate", label="料速", symbol="v", unit="批/h", default=15,
              step=0.5, fmt="%.6g"),
    ]

    reference = [
        {"key": "t", "label": "煤热滞后时间", "symbol": "t", "unit": "h",
         "excel": 11.1111111111111, "note": "原表 D22 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        t = v["v"] / v["v_batch"] / v["rate"] if v["v_batch"] and v["rate"] else float("nan")
        return [
            Result(
                key="t", label="煤热滞后时间", value=t, unit="h", digits=3, symbol="t",
                rng=Range(4.0, 8.0, SRC_TYPICAL,
                          "煤热滞后过长会降低喷煤利用率；算例 11.1 h 明显偏长。"),
                formula="t = V/(V料批 · v)",
            ),
        ]


MODELS = [
    SmeltingIntensity(),
    SmeltingCycleBatch(),
    SmeltingCycleHours(),
    CoalThermalLag(),
]
