"""炉缸、炉衬与出铁类模型。

对应 Excel 工作表：安全容铁 / 出铁次数 / 渣皮厚度 / 耐材厚度（导热计算法）。
"""

from __future__ import annotations

from typing import Any, Dict, List

from ..base import SRC_EXCEL, SRC_REPORT, SRC_TYPICAL, Input, Model, Range, Result

PI = 3.1415926


# ===========================================================================
# 安全容铁
# ===========================================================================


class SafeHotMetal(Model):
    key = "safe_hot_metal"
    title = "安全容铁"
    subtitle = "炉缸内允许的最大安全铁水量"
    group = "炉缸与炉衬"
    source_sheet = "安全容铁"
    formula_latex = r"T_{容} = 0.6\,\pi\,r^2\,d\,\rho"
    formula_excel = "=0.6*3.1415926*r*r*d*ρ      （文献式，原表 I23:L23）"
    notes = [
        "原表 I3 批注「公式怀疑有问题，文献结果如下」。原表主公式 D23 多乘了炉缸容积 V："
        "T = 0.6·V炉缸·π·r²·d·ρ，算例得 3.04×10⁸ kg，量纲明显不自洽；"
        "右侧文献式 L23 = 0.6·π·r²·d·ρ 得 704744 kg ≈ 705 t，"
        "与「出铁次数」表中使用的 705 t 完全一致。",
        "本工具默认采用文献式（正确），下拉可切回原表式复现 3.04×10⁸ kg。",
        "系数 0.6 表示允许铁水液面最低可降至铁口中心线以下 60% 炉缸半径处。",
    ]
    variant_label = "「公式版本」下拉切换 文献式 / 原表式（多乘炉缸容积）。"

    inputs = [
        Input(key="v_h", label="炉缸容积", symbol="V炉缸", unit="m³", default=1250,
              step=10.0, fmt="%.8g",
              help="仅原表式使用；文献式不需要。"),
        Input(key="r", label="炉缸半径", symbol="r", unit="m", default=4.05,
              step=0.01, fmt="%.6g"),
        Input(key="d", label="铁口中心线到风口中心线的距离", symbol="d", unit="m",
              default=2.9, step=0.05, fmt="%.6g",
              help="原表左算例取 1 m，右侧文献算例取 2.9 m（默认采用文献算例）。"),
        Input(key="rho", label="铁的密度", symbol="ρ", unit="kg/m³", default=7860,
              step=10.0, fmt="%.8g"),
        Input(key="variant", label="公式版本", kind="select", default="lit",
              options=["lit", "raw"],
              option_labels=["文献式：0.6·π·r²·d·ρ（原表 L23）",
                             "原表式：0.6·V炉缸·π·r²·d·ρ（原表 D23，量纲存疑）"],
              group="输出"),
    ]

    reference = [
        {"key": "t_cap", "label": "安全容铁量", "symbol": "T容", "unit": "kg",
         "excel": 704744.517560682, "note": "原表 L23 文献式算例（r=4.05, d=2.9），复算一致"},
        {"key": "t_cap_raw", "label": "安全容铁量（原表式）", "symbol": "T容", "unit": "kg",
         "excel": None,
         "note": "原表 D23 用 d=1 m 得 3.04×10⁸ kg；本工具默认 d=2.9 m（文献式取值），"
                 "故复算值不同。把「铁口中心线到风口中心线的距离」改为 1 m 可复现原表值。"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        base = 0.6 * PI * v["r"] * v["r"] * v["d"] * v["rho"]
        raw = base * v["v_h"]
        value = raw if v.get("variant", "lit") == "raw" else base
        formula = ("T容 = 0.6·V炉缸·π·r²·d·ρ（原表式，量纲存疑）"
                   if v.get("variant", "lit") == "raw"
                   else "T容 = 0.6·π·r²·d·ρ（文献式）")
        return [
            Result(
                key="t_cap", label="安全容铁量", value=value, unit="kg", digits=0,
                symbol="T容",
                rng=Range(500e3, 900e3, SRC_TYPICAL,
                          "按 8~12 次/昼夜出铁反推的常见区间；应与「出铁次数」Tab 联立标定。"),
                formula=formula,
                note=f"= {value/1000:,.1f} t",
            ),
            Result(
                key="t_cap_raw", label="安全容铁量（原表式）", value=raw, unit="kg",
                digits=0, symbol="T容", primary=False,
                formula="T容 = 0.6·V炉缸·π·r²·d·ρ",
                note="原表 D23 的 d=1 m 时为 3.04×10⁸ kg；把 d 改为 1 m 可复现该值。",
            ),
            Result(
                key="t_cap_t", label="安全容铁量（t）", value=value / 1000.0, unit="t",
                digits=1, symbol="T容", primary=False,
                formula="T容 / 1000",
            ),
        ]


# ===========================================================================
# 出铁次数
# ===========================================================================


class TappingCount(Model):
    key = "tapping_count"
    title = "出铁次数"
    subtitle = "昼夜出铁次数"
    group = "炉缸与炉衬"
    source_sheet = "出铁次数"
    formula_latex = r"n = \frac{1.2\,P}{T_{容}}"
    formula_excel = "=1.2*P/T容"
    notes = [
        "原表 H21 批注：「使用之前算到的安全容铁量，得不到生产报表中 12 次出铁次数」。"
        "按默认参数算得 9.64 次；要达到 12 次，安全容铁量需为 566 t（即 1.2×5661/12）。",
        "系数 1.2 表示允许铁水液面比安全位置再下降 20%。",
    ]

    inputs = [
        Input(key="p", label="高炉昼夜出铁量", symbol="P", unit="t", default=5661,
              step=1.0, fmt="%.8g"),
        Input(key="t_cap", label="炉缸安全容铁量", symbol="T容", unit="t", default=705,
              step=1.0, fmt="%.8g",
              help="可由「安全容铁」Tab 得出（该 Tab 结果为 704.7 t）。"),
    ]

    reference = [
        {"key": "n", "label": "昼夜出铁次数", "symbol": "n", "unit": "次",
         "excel": 9.63574468085106, "note": "原表 D20 算例，复算一致"},
        {"key": "t_cap_needed", "label": "达成 12 次所需安全容铁量", "symbol": "T容*",
         "unit": "t", "excel": None, "note": "= 1.2×P/12"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        n = 1.2 * v["p"] / v["t_cap"] if v["t_cap"] else float("nan")
        needed = 1.2 * v["p"] / 12.0
        return [
            Result(
                key="n", label="昼夜出铁次数", value=n, unit="次", digits=3, symbol="n",
                rng=Range(8.0, 12.0, SRC_REPORT,
                          "按现场报表 12 次/昼夜反推的可接受区间；现代大高炉常见 8~12 次。"),
                formula="n = 1.2·P/T容",
            ),
            Result(
                key="t_cap_needed", label="达成 12 次所需安全容铁量", value=needed,
                unit="t", digits=1, symbol="T容*", primary=False,
                formula="T容* = 1.2·P/12",
                note="与报表 12 次对不上时，用该值回推并核对「安全容铁」Tab 的输入口径。",
            ),
        ]


# ===========================================================================
# 渣皮厚度
# ===========================================================================


class SlagSkinThickness(Model):
    key = "slag_skin"
    title = "渣皮厚度"
    subtitle = "炉缸内渣层的挂渣厚度"
    group = "炉缸与炉衬"
    source_sheet = "渣皮厚度"
    formula_latex = r"d = 9.64\,t^{0.3298}"
    formula_excel = "=9.64*t^0.3298"
    notes = [
        "t 为挂渣时间（min），d 为渣皮厚度（mm）。",
    ]

    inputs = [
        Input(key="t", label="挂渣时间", symbol="t", unit="min", default=1800,
              step=30.0, fmt="%.8g"),
    ]

    reference = [
        {"key": "d", "label": "渣皮厚度", "symbol": "d", "unit": "mm",
         "excel": 114.199947981178, "note": "原表 D14 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        d = 9.64 * v["t"] ** 0.3298
        return [
            Result(
                key="d", label="渣皮厚度", value=d, unit="mm", digits=2, symbol="d",
                rng=Range(60, 150, SRC_TYPICAL,
                          "常见渣皮厚度 60~150 mm；过厚影响散热与铁水温度，过薄易脱落。"),
                formula="d = 9.64·t^0.3298",
            ),
        ]


# ===========================================================================
# 耐材厚度（导热计算法）
# ===========================================================================


class RefractoryThickness(Model):
    key = "refractory_thickness"
    title = "耐材厚度（导热计算法）"
    short_title = "耐材厚度"
    subtitle = "炉衬剩余厚度"
    group = "炉缸与炉衬"
    source_sheet = "耐材厚度（导热计算法）"
    formula_latex = (
        r"L = \frac{\lambda\,(T_{热}-T_{冷})\,F\times3600}"
        r"{M\times1000\times a\,\delta t}"
    )
    formula_excel = "=λ*(T热-T冷)*F*a*δt/M"
    notes = [
        "原表 H3 批注「公式怀疑有误」，且算例无输入（输出 #DIV/0!）。",
        "原表公式把 a×δt 放在分子、M 放在分母，量纲不自洽（λ·ΔT·F 是 W，"
        "而 a·δt/M 是 J/(kg·℃)·℃/(m³/h)，两者不匹配）。",
        "本工具的修正式按稳态导热与冷却水吸热平衡推导："
        "λ·F·ΔT/L = (M·1000/3600)·a·δt，整理即得 L = λ·ΔT·F·3600/(M·1000·a·δt)。",
        "注意原表算例的冷却水流量 879.9 m³/h 与冷却面积 F 若取 100 m² 不匹配，"
        "会算出极小的剩余厚度；该模型目前缺少可靠的现场标定数据。",
    ]
    variant_label = "「公式版本」下拉切换 修正式（量纲自洽）/ 原表式（量纲存疑）。"

    inputs = [
        Input(key="lam", label="耐材导热系数", symbol="λ", unit="W/(m·℃)",
              default=1.5, step=0.1, fmt="%.6g"),
        Input(key="t_hot", label="耐材热端温度", symbol="T热", unit="℃",
              default=800, step=10.0, fmt="%.8g"),
        Input(key="t_cold", label="耐材冷端温度", symbol="T冷", unit="℃",
              default=100, step=10.0, fmt="%.8g"),
        Input(key="f_area", label="冷却设备面积", symbol="F", unit="m²",
              default=100, step=1.0, fmt="%.8g"),
        Input(key="m_flow", label="单位时间冷却水流量", symbol="M", unit="m³/h",
              default=100, step=1.0, fmt="%.8g"),
        Input(key="a_cp", label="水的比热容", symbol="a", unit="J/(kg·℃)",
              default=4180, step=1.0, fmt="%.8g"),
        Input(key="dt", label="冷却设备进出水温差", symbol="δt", unit="℃",
              default=5.0, step=0.5, fmt="%.6g"),
        Input(key="variant", label="公式版本", kind="select", default="fixed",
              options=["fixed", "raw"],
              option_labels=["修正式：λ·ΔT·F·3600/(M·1000·a·δt)",
                             "原表式：λ·ΔT·F·a·δt/M"],
              group="输出"),
    ]

    reference = [
        {"key": "l_raw", "label": "耐材侵蚀剩余长度（原表式）", "symbol": "L耐材剩余",
         "unit": "m", "excel": "#DIV/0!", "note": "原表 D33 无输入数据，输出 #DIV/0!"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        dt_lam = v["t_hot"] - v["t_cold"]
        denom = v["m_flow"] * 1000 * v["a_cp"] * v["dt"]
        fixed = (v["lam"] * dt_lam * v["f_area"] * 3600 / denom) if denom else float("nan")
        raw = (v["lam"] * dt_lam * v["f_area"] * v["a_cp"] * v["dt"] / v["m_flow"]
               if v["m_flow"] else float("nan"))
        value = raw if v.get("variant", "fixed") == "raw" else fixed
        formula = ("L = λ·(T热−T冷)·F·a·δt/M（原表式，量纲存疑）"
                   if v.get("variant", "fixed") == "raw"
                   else "L = λ·(T热−T冷)·F·3600/(M·1000·a·δt)")
        return [
            Result(
                key="l_thick", label="耐材侵蚀剩余长度", value=value, unit="m",
                digits=5, symbol="L耐材剩余",
                rng=Range(0.10, 0.60, SRC_TYPICAL,
                          "炉衬剩余厚度报警区间，取决于炉缸直径与砌筑结构，需按本厂标定。"),
                formula=formula,
            ),
            Result(
                key="l_raw", label="耐材侵蚀剩余长度（原表式）", value=raw, unit="m",
                digits=6, symbol="L", primary=False,
                formula="L = λ·ΔT·F·a·δt/M",
            ),
            Result(
                key="heat_flux", label="炉衬导热热流", value=v["lam"] * dt_lam * v["f_area"],
                unit="W", digits=1, symbol="λ·F·ΔT", primary=False,
                formula="Q导 = λ·F·(T热−T冷)",
            ),
            Result(
                key="water_heat", label="冷却水吸热功率",
                value=v["m_flow"] * 1000 / 3600 * v["a_cp"] * v["dt"] if v["dt"] else float("nan"),
                unit="W", digits=1, symbol="ṁ·a·δt", primary=False,
                formula="Q吸 = (M·1000/3600)·a·δt",
            ),
        ]


MODELS = [
    SafeHotMetal(),
    TappingCount(),
    SlagSkinThickness(),
    RefractoryThickness(),
]
