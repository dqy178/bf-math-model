"""风口回旋区与布料落点类模型。

对应 Excel 工作表：风口回旋区长度 / 回旋区面积 / 回旋区面积占比 / 落点计算。
"""

from __future__ import annotations

import math
from typing import Any, Dict, List

from ..base import SRC_EXCEL, SRC_TYPICAL, Input, Model, Range, Result

PI = 3.1415926


# ===========================================================================
# 风口回旋区长度
# ===========================================================================


class RacewayLength(Model):
    key = "raceway_length"
    title = "风口回旋区长度"
    subtitle = "回旋区在炉缸内的轴向长度"
    group = "回旋区与布料"
    source_sheet = "风口回旋区长度"
    formula_latex = r"D_r = 0.88 + k_E\,E_k - 0.00031\,\frac{PCI}{n}"
    formula_excel = "=0.88+0.000095*Ek-0.00031*PCI/n"
    notes = [
        "系数 0.000095 出自本表；综合计算表用的是 0.000092，两者不一致，"
        "下拉可切换以复现综合表的结果。",
        "原表算例 D 列（Ek=2200）得 1.088 m、E 列（Ek=75）得 0.887 m，"
        "均低于经验区间 1.8~2.2 m，提示算例中的鼓风动能取值口径与本式不一致"
        "（综合表用的鼓风动能为 kg·m²/s² 量级，本表标注 kg·m/s）。",
    ]
    variant_label = "「经验式系数」下拉切换 0.000095（本表）/ 0.000092（综合表）。"

    inputs = [
        Input(key="ek", label="鼓风动能", symbol="Ek", unit="kg·m/s", default=2200,
              step=10.0, fmt="%.8g",
              help="注意：综合表中同名的鼓风动能量纲为 kg·m²/s²，两者不可直接互换。"),
        Input(key="pci", label="煤量", symbol="PCI", unit="kg/h", default=43,
              step=1.0, fmt="%.8g"),
        Input(key="n", label="风口个数", symbol="n", unit="个", default=22,
              step=1.0, lo=1, fmt="%.6g"),
        Input(key="coef", label="经验式系数", kind="select", default="0.000095",
              options=["0.000095", "0.000092"],
              option_labels=["0.000095（本表 D19）", "0.000092（综合计算表 F35）"]),
    ]

    reference = [
        {"key": "dr", "label": "风口回旋区长度", "symbol": "Dr", "unit": "m",
         "excel": 1.08839409090909, "note": "原表 D19 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        coef = 0.000095 if v.get("coef", "0.000095") == "0.000095" else 0.000092
        dr = 0.88 + coef * v["ek"] - 0.00031 * v["pci"] / v["n"]
        return [
            Result(
                key="dr", label="风口回旋区长度", value=dr, unit="m", digits=4,
                symbol="Dr",
                rng=Range(1.8, 2.2, SRC_EXCEL,
                          "原表综合表给出「回旋区长度：2±0.2」m。"),
                formula="Dr = 0.88 + 0.000095·Ek − 0.00031·PCI/n",
                note=f"当前系数 {coef}",
            ),
        ]


# ===========================================================================
# 回旋区面积
# ===========================================================================


class RacewayArea(Model):
    key = "raceway_area"
    title = "风口回旋区面积"
    subtitle = "炉缸截面上回旋区所占的面积"
    group = "回旋区与布料"
    source_sheet = "回旋区面积"
    formula_latex = r"S_r = \pi r^2 - \pi (r - D_r)^2"
    formula_excel = "=3.1415926*r*r-3.1415926*(r-Dr)*(r-Dr)"
    notes = [
        "本 Tab 按「各 Tab 独立」原则，回旋区长度 Dr 作为独立输入录入，不与其他 Tab 联动。",
    ]

    inputs = [
        Input(key="dr", label="风口回旋区长度", symbol="Dr", unit="m", default=1.088,
              step=0.001, fmt="%.6g"),
        Input(key="r", label="炉缸半径", symbol="r", unit="m", default=4.25,
              step=0.01, fmt="%.6g"),
    ]

    reference = [
        {"key": "sr", "label": "风口回旋区面积", "symbol": "Sr", "unit": "m²",
         "excel": 25.3346069741056, "note": "原表 D19 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        r, dr = v["r"], v["dr"]
        sr = PI * r * r - PI * (r - dr) * (r - dr)
        ratio = sr / (PI * r * r) if r else float("nan")
        return [
            Result(
                key="sr", label="风口回旋区面积", value=sr, unit="m²", digits=3,
                symbol="Sr",
                rng=Range(20.0, 32.0, SRC_TYPICAL,
                          "对应炉缸半径 4.25 m、面积占比 45%~55% 的取值区间；请按本厂标定。"),
                formula="Sr = π·r² − π·(r − Dr)²",
            ),
            Result(
                key="ratio", label="回旋区面积占比", value=ratio, unit="", digits=4,
                symbol="Rr", primary=False, formula="Rr = Sr/(π·r²)",
                note="与「回旋区面积占比」Tab 的模型一致，此处作为中间量显示。",
            ),
        ]


# ===========================================================================
# 回旋区面积占比
# ===========================================================================


class RacewayAreaRatio(Model):
    key = "raceway_area_ratio"
    title = "回旋区面积占比"
    subtitle = "回旋区面积占炉缸截面积的比例"
    group = "回旋区与布料"
    source_sheet = "回旋区面积占比"
    formula_latex = r"R_r = \frac{S_r}{\pi r^2}"
    formula_excel = "=Sr/3.1415926/r/r"
    notes = [
        "原表算例得 0.4465，与综合表经验值 50% 接近。",
    ]

    inputs = [
        Input(key="sr", label="风口回旋区面积", symbol="Sr", unit="m²", default=25.335,
              step=0.001, fmt="%.6g"),
        Input(key="r", label="炉缸半径", symbol="r", unit="m", default=4.25,
              step=0.01, fmt="%.6g"),
    ]

    reference = [
        {"key": "rr", "label": "回旋区面积占比", "symbol": "Rr", "unit": "",
         "excel": 0.446470926174663, "note": "原表 C19 算例，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        rr = v["sr"] / PI / v["r"] / v["r"] if v["r"] else float("nan")
        return [
            Result(
                key="rr", label="回旋区面积占比", value=rr, unit="", digits=4,
                symbol="Rr",
                rng=Range(0.45, 0.55, SRC_EXCEL,
                          "原表综合表给出「回旋区面积：50%」。"),
                formula="Rr = Sr/(π·r²)",
            ),
        ]


# ===========================================================================
# 落点计算
# ===========================================================================


class BurdenDistribution(Model):
    key = "burden_distribution"
    title = "落点计算"
    subtitle = "炉料经溜槽布到料面的落点位置"
    group = "回旋区与布料"
    source_sheet = "落点计算"
    formula_latex = (
        r"C_1 = \sqrt{2g\,(l_0-e\tan\beta)(\sin\beta-\mu\cos\beta)"
        r"+4\pi^2\omega^2 (l_0-e\tan\beta)^2\cos\beta(\cos\beta+\mu\sin\beta)}"
    )
    formula_excel = (
        "=SQRT(2*9.81*(l0-e*TAN(β))*(SIN(β)-μ*COS(β))"
        "+4*3.1415926^2*ω^2*(l0-e*TAN(β))^2*COS(β)*(COS(β)+μ*SIN(β)))"
    )
    notes = [
        "原表 H39 批注：「公式已经核对过，但没有实际数据。」",
        "**角度单位**：原表把「角度 β = 6」直接送入 TAN/SIN/COS，"
        "Excel 按弧度计算，因此原表数值对应的是「6 弧度」（约 344°）。"
        "下拉可切换为工程上正确的「度→弧度」换算，两者在该算例下相差约 22%。",
        "算例结果 Lx≈174 m、n≈202 m，与炉缸半径（米级）明显不同量级，"
        "说明算例中「一块炉料的质量 m=1 kg、重量 Q=10 N」等取值与真实工况不符，"
        "该模型目前只能作为公式框架使用。",
    ]
    variant_label = "「角度单位」下拉切换 原表弧度制 / 工程度数制。"

    inputs = [
        Input(key="l0", label="溜槽长度", symbol="l0", unit="m", default=1.0,
              step=0.1, fmt="%.6g"),
        Input(key="e", label="溜槽倾动距", symbol="e", unit="m", default=1.0,
              step=0.1, fmt="%.6g"),
        Input(key="beta", label="角度", symbol="β", unit="°", default=6.0,
              step=0.5, fmt="%.6g"),
        Input(key="angle_mode", label="角度单位", kind="select", default="rad",
              options=["rad", "deg"],
              option_labels=["原表写法：弧度制（β 直接代入 SIN/TAN/COS）",
                             "工程写法：度 → 弧度换算"], group="输出"),
        Input(key="mu", label="摩擦系数", symbol="μ", unit="", default=0.6,
              step=0.05, fmt="%.6g"),
        Input(key="h1", label="溜槽末端到料线零尺的距离", symbol="h1", unit="m",
              default=0.2, step=0.05, fmt="%.6g"),
        Input(key="h0", label="料线深度", symbol="h0", unit="m", default=1.8,
              step=0.1, fmt="%.6g"),
        Input(key="p_gas", label="上升煤气阻力", symbol="P", unit="Pa", default=1.0,
              step=0.1, fmt="%.6g"),
        Input(key="m", label="一块炉料的质量", symbol="m", unit="kg", default=1.0,
              step=0.1, fmt="%.6g"),
        Input(key="q", label="一块炉料的重量", symbol="Q", unit="N", default=10.0,
              step=0.1, fmt="%.6g"),
        Input(key="omega", label="溜槽转速", symbol="ω", unit="圈/s", default=7.5,
              step=0.1, fmt="%.6g"),
    ]

    reference = [
        {"key": "c1", "label": "炉料在溜槽末端的速度", "symbol": "C1", "unit": "m/s",
         "excel": 52.8654278278716, "note": "原表 D37 算例（弧度制），复算一致"},
        {"key": "lx", "label": "离开溜槽末端后 x 方向投影", "symbol": "Lx", "unit": "m",
         "excel": 174.242159582869, "note": "原表 D38 算例（弧度制），复算一致"},
        {"key": "n", "label": "落点到中心距离", "symbol": "n", "unit": "m",
         "excel": 201.597452489082, "note": "原表 D39 算例（弧度制），复算一致"},
    ]

    def _calc(self, v: Dict[str, Any], use_deg: bool) -> Dict[str, float]:
        l0, e = v["l0"], v["e"]
        beta = math.radians(v["beta"]) if use_deg else v["beta"]
        mu, h1, h0 = v["mu"], v["h1"], v["h0"]
        pg, m, q, om = v["p_gas"], v["m"], v["q"], v["omega"]

        h = l0 - e * math.tan(beta)
        c1_sq = (2 * 9.81 * h * (math.sin(beta) - mu * math.cos(beta))
                 + 4 * PI * PI * om * om * h * h
                 * math.cos(beta) * (math.cos(beta) + mu * math.sin(beta)))
        c1 = math.sqrt(c1_sq) if c1_sq >= 0 else float("nan")

        radicand = (math.tan(beta) ** 2
                    + 2 * (q - pg) / m / c1 / c1 / math.cos(beta) / math.cos(beta)
                    * (l0 * (1 - math.sin(beta)) - e * math.cos(beta) + h0 + h1))
        lx = (m * c1 * c1 * math.cos(beta) * math.cos(beta) / (q - pg)
              * (math.sqrt(radicand) - math.tan(beta))) if radicand >= 0 else float("nan")

        term_a = l0 * math.cos(beta) - e * math.sin(beta)
        n_sq = (term_a * term_a + 2 * term_a * lx + 1
                + 4 * PI * PI * om * om * h * h * lx * lx / c1 / c1)
        n_dist = math.sqrt(n_sq) if n_sq >= 0 else float("nan")
        return {"c1": c1, "lx": lx, "n": n_dist}

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        use_deg = v.get("angle_mode", "rad") == "deg"
        r = self._calc(v, use_deg)
        r_deg = self._calc(v, True)
        c1, lx, n_dist = r["c1"], r["lx"], r["n"]

        return [
            Result(
                key="c1", label="炉料在溜槽末端的速度", value=c1, unit="m/s", digits=3,
                symbol="C1",
                formula="C1 = √{2g(l₀−e·tanβ)(sinβ−μcosβ) + 4π²ω²(l₀−e·tanβ)²cosβ(cosβ+μsinβ)}",
                note="当前按「度→弧度」换算计算。" if use_deg else "当前按原表弧度制计算。",
            ),
            Result(
                key="lx", label="离开溜槽末端后 x 方向投影", value=lx, unit="m",
                digits=3, symbol="Lx",
                formula="Lx = m·C1²·cos²β/(Q−P)·[√(tan²β + 2(Q−P)/(m·C1²·cos²β)·(l₀(1−sinβ)−e·cosβ+h₀+h₁)) − tanβ]",
            ),
            Result(
                key="n", label="落点到中心距离", value=n_dist, unit="m", digits=3,
                symbol="n",
                formula="n = √{(l₀cosβ−e·sinβ)² + 2(l₀cosβ−e·sinβ)·Lx + 1 + 4π²ω²(l₀−e·tanβ)²·Lx²/C1²}",
                note="算例取值与炉缸半径不同量级，模型缺少实测标定。",
            ),
            Result(
                key="c1_deg", label="C1（按度换算对照）", value=r_deg["c1"], unit="m/s",
                digits=3, symbol="C1°", primary=False,
                formula="同上，但 β 先做 度→弧度 换算",
                note="工程上正确的角度处理；与原表弧度制结果相差约 22%。",
            ),
        ]


MODELS = [
    RacewayLength(),
    RacewayArea(),
    RacewayAreaRatio(),
    BurdenDistribution(),
]
