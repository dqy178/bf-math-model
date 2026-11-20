"""综合计算表：把炉腹煤气量、煤气量指数、鼓风动能、回旋区尺寸等串成一条链。

对应 Excel 工作表：高炉炉腹煤气量指数 鼓风动能 风口回旋区面积占比计算。
"""

from __future__ import annotations

from typing import Any, Dict, List

from ..base import SRC_EXCEL, Input, Model, Range, Result


class IntegratedBlastFurnace(Model):
    key = "integrated"
    title = "综合计算（煤气量指数/鼓风动能/回旋区）"
    short_title = "综合计算"
    subtitle = "一次算完炉腹煤气量、煤气量指数、鼓风动能、回旋区尺寸与压量关系 K"
    group = "综合计算"
    source_sheet = "高炉炉腹煤气量指数 鼓风动能 风口回旋区面积占比计算"
    description = (
        "原表把 10 余个公式串成一条链：鼓风 → 风口风速 → 热风质量 → 鼓风动能 → "
        "回旋区长度 → 回旋区高度/面积比 → 炉腹煤气量 → 煤气量指数 → 压量关系 K。"
        "本 Tab 为「一站式」入口，其余 Tab 仍可单独使用。"
    )
    formula_latex = (
        r"K = \frac{P_B^2 - P_{顶}^2}{V_{bg}^{1.7}},\qquad"
        r"R_r = 1 - \frac{(D - 2D_r)^2}{D^2}"
    )
    formula_excel = (
        "压量关系K = (PB^2-P顶^2)/Vbg^1.7\n"
        "风口风速 = (VB+VO2/60)/60/Sf*TB/T大气/PB*100*803.6/(803.6-WB)\n"
        "热风质量 = (1.2507*0.79*VB + 1.4289*(0.21*VB+VO2/60) + (VB+VO2/60)*WB/1000/(1-WB/803.6))*(1/60/9.81/n)\n"
        "鼓风动能 = 0.5*风口风速^2*热风质量\n"
        "回旋区长度 = 0.88+0.000092*鼓风动能-0.00031*Pc/n\n"
        "回旋区高度 = 22.856*(风口风速^2/9.81/d焦)^(-0.404)*回旋区长度^1.286/d焦^0.286\n"
        "回旋区面积比 = 1-(D-2*回旋区长度)^2/D^2\n"
        "炉腹煤气量 = 1.21*VB + 2*VO2/60 + 44.8*WB*(VB+VO2/60)/18000 + 22.4*Pc*H/120\n"
        "煤气量指数 = 炉腹煤气量/(3.14/4*D^2)\n"
        "理论燃烧温度 = 1559+0.839*热风温度-6.033*WB+4972*VO2/(VB*60+VO2)-3150*Pc/(VB*60+VO2)\n"
        "铁口深度 = 0.45*D/2"
    )
    notes = [
        "原表「3 经验参数范围」给出的六项基准：指数 62±2（58~65）、回旋区长度 2±0.2 m、"
        "回旋区面积 50%、理论燃烧温度 2150±100 ℃、风口面积 0.475~0.495 m²、鼓风动能 15500±1000。",
        "注意口径差异：原表这里的鼓风动能 = 0.5·v²·m，量纲为 kg·m²/s²，"
        "与「鼓风动能」Tab 的经验式（kg·m/s）不是同一量，因此默认算例 31262 "
        "并不能落在 15500±1000 区间内。",
        "原表 A29 标注「高炉产量 P产量 单位 t/d/m3」，实为 t/d，单位标注有误；"
        "该值在本模型中未参与公式。",
        "回旋区长度用 0.000092 系数（与「风口回旋区长度」Tab 的 0.000095 不同）。",
    ]

    inputs = [
        Input(key="vb", label="风量", symbol="VB", unit="m³/min", default=7829,
              step=1.0, fmt="%.8g", group="鼓风与富氧"),
        Input(key="vo2", label="富氧", symbol="VO2", unit="m³/h", default=33815,
              step=1.0, fmt="%.8g", group="鼓风与富氧"),
        Input(key="wb", label="鼓风湿度", symbol="WB", unit="g/m³", default=1,
              step=0.1, fmt="%.6g", group="鼓风与富氧"),
        Input(key="bt", label="热风温度", symbol="BT", unit="℃", default=1198,
              step=1.0, fmt="%.8g", group="鼓风与富氧"),
        Input(key="pb", label="热风压力", symbol="PB", unit="kPa", default=385,
              step=1.0, fmt="%.8g", group="鼓风与富氧"),
        Input(key="n", label="风口数目", symbol="n", unit="个", default=38,
              step=1.0, lo=1, fmt="%.6g", group="鼓风与富氧"),
        Input(key="sf", label="风口面积", symbol="Sf", unit="m²", default=0.485,
              step=0.001, fmt="%.6g", group="鼓风与富氧",
              help="原表经验范围 0.475~0.495 m²。"),
        Input(key="t_amb", label="大气温度", symbol="T大气", unit="℃", default=35,
              step=1.0, fmt="%.6g", group="鼓风与富氧"),

        Input(key="d", label="炉缸直径", symbol="D", unit="m", default=14.2,
              step=0.1, fmt="%.6g", group="炉体与几何"),
        Input(key="v_furnace", label="高炉容积", symbol="V", unit="m³", default=4328,
              step=10.0, fmt="%.8g", group="炉体与几何"),
        Input(key="pc", label="喷煤量", symbol="Pc", unit="kg/h", default=78900,
              step=100.0, fmt="%.8g", group="炉体与几何"),
        Input(key="hh", label="氢含量", symbol="H", unit="%", default=0.04,
              step=0.01, fmt="%.6g", group="炉体与几何"),
        Input(key="p_top", label="炉顶压力", symbol="P顶", unit="kPa", default=240,
              step=1.0, fmt="%.8g", group="炉体与几何"),
        Input(key="d_coke", label="焦炭平均直径", symbol="d焦", unit="m", default=0.062,
              step=0.001, fmt="%.6g", group="炉体与几何"),
    ]

    reference = [
        {"key": "vbg", "label": "炉腹煤气量", "symbol": "Vbg", "unit": "m³/min",
         "excel": 11210.2648740741, "note": "原表 C36，复算一致"},
        {"key": "abg", "label": "煤气量指数", "symbol": "", "unit": "m³/(min·m²)",
         "excel": 70.8222187873076, "note": "原表 C37，复算一致；经验区间 58~65"},
        {"key": "tf", "label": "理论燃烧温度", "symbol": "Tf", "unit": "℃",
         "excel": 2398.41067290564, "note": "原表 C38，复算一致；经验区间 2050~2250"},
        {"key": "k_rel", "label": "压量关系 K", "symbol": "K", "unit": "",
         "excel": 0.0118277199644723, "note": "原表 C39，复算一致"},
        {"key": "dr", "label": "回旋区长度", "symbol": "Dr", "unit": "m",
         "excel": 3.11245970816113, "note": "原表 F35，复算一致；经验区间 1.8~2.2"},
        {"key": "rr", "label": "回旋区面积比", "symbol": "Rr", "unit": "",
         "excel": 0.68457691769414, "note": "原表 F38，复算一致；经验值 50%"},
        {"key": "h_rec", "label": "回旋区高度", "symbol": "h回", "unit": "m",
         "excel": 1.53999473963315, "note": "原表 F37，复算一致"},
        {"key": "v_duct", "label": "风口风速", "symbol": "v", "unit": "m/s",
         "excel": 358.215746437513, "note": "原表 F29，复算一致"},
        {"key": "ek", "label": "鼓风动能", "symbol": "Ek", "unit": "kg·m²/s²",
         "excel": 31262.1478575866, "note": "原表 F31，复算一致"},
        {"key": "taphole", "label": "铁口深度", "symbol": "h铁口", "unit": "m",
         "excel": 3.195, "note": "原表 F39，复算一致"},
    ]

    def compute(self, v: Dict[str, Any]) -> List[Result]:
        vb, vo2, wb = v["vb"], v["vo2"], v["wb"]
        bt, pb, n, sf = v["bt"], v["pb"], v["n"], v["sf"]
        d, pc, hh, p_top = v["d"], v["pc"], v["hh"], v["p_top"]

        tb = bt + 273.0
        t_amb = v["t_amb"] + 273.0

        # 风口风速
        v_duct = ((vb + vo2 / 60.0) / 60.0 / sf
                  * tb / t_amb / pb * 100.0 * 803.6 / (803.6 - wb))
        # 热风质量（kg/个/…，原表口径）
        m_heat = ((1.2507 * 0.79 * vb
                   + 1.4289 * (0.21 * vb + vo2 / 60.0)
                   + (vb + vo2 / 60.0) * wb / 1000.0 / (1 - wb / 803.6))
                  * (1.0 / 60.0 / 9.81 / n))
        ek = 0.5 * v_duct * v_duct * m_heat

        dr = 0.88 + 0.000092 * ek - 0.00031 * pc / n
        h_rec = (22.856 * (v_duct * v_duct / 9.81 / v["d_coke"]) ** (-0.404)
                 * dr ** 1.286 / v["d_coke"] ** 0.286)
        rr = 1 - (d - 2 * dr) ** 2 / d / d
        taphole = 0.45 * d / 2

        vbg = (1.21 * vb + 2 * vo2 / 60.0
               + 44.8 * wb * (vb + vo2 / 60.0) / 18000
               + 22.4 * pc * hh / 120)
        abg = vbg / (3.14 / 4 * d * d)

        tf = (1559 + 0.839 * bt - 6.033 * wb
              + 4972 * vo2 / (vb * 60 + vo2)
              - 3150 * pc / (vb * 60 + vo2))
        k_rel = (pb ** 2 - p_top ** 2) / (vbg ** 1.7)

        return [
            Result(
                key="vbg", label="炉腹煤气量", value=vbg, unit="m³/min", digits=1,
                symbol="Vbg",
                formula="Vbg = 1.21·VB + 2·VO₂/60 + 44.8·WB·(VB+VO₂/60)/18000 + 22.4·Pc·H/120",
            ),
            Result(
                key="abg", label="煤气量指数", value=abg, unit="m³/(min·m²)", digits=3,
                symbol="αbg",
                rng=Range(58, 65, SRC_EXCEL, "原表综合表「指数范围：62±2，58-65」。"),
                formula="αbg = Vbg/(3.14/4·D²)",
            ),
            Result(
                key="tf", label="理论燃烧温度", value=tf, unit="℃", digits=2, symbol="Tf",
                rng=Range(2050, 2250, SRC_EXCEL, "原表综合表「理论燃烧温度：2150±100」。"),
                formula="Tf = 1559+0.839·BT−6.033·WB+4972·VO₂/(VB·60+VO₂)−3150·Pc/(VB·60+VO₂)",
            ),
            Result(
                key="dr", label="回旋区长度", value=dr, unit="m", digits=4, symbol="Dr",
                rng=Range(1.8, 2.2, SRC_EXCEL, "原表综合表「回旋区长度：2±0.2」。"),
                formula="Dr = 0.88 + 0.000092·Ek − 0.00031·Pc/n",
            ),
            Result(
                key="rr", label="回旋区面积比", value=rr, unit="", digits=4, symbol="Rr",
                rng=Range(0.45, 0.55, SRC_EXCEL, "原表综合表「回旋区面积：50%」。"),
                formula="Rr = 1 − (D − 2Dr)²/D²",
            ),
            Result(
                key="ek", label="鼓风动能", value=ek, unit="kg·m²/s²", digits=1,
                symbol="Ek",
                rng=Range(14500, 16500, SRC_EXCEL,
                          "原表综合表「鼓风动能：15500±1000」；但本项量纲为 kg·m²/s²，"
                          "与「鼓风动能」Tab 的 kg·m/s 不是同一口径，默认算例 31262 必然超出该区间。"),
                formula="Ek = 0.5·v²·m",
            ),
            Result(
                key="k_rel", label="压量关系 K", value=k_rel, unit="", digits=6,
                symbol="K",
                formula="K = (PB² − P顶²)/Vbg^1.7",
            ),
            Result(
                key="h_rec", label="回旋区高度", value=h_rec, unit="m", digits=4,
                symbol="h回", primary=False,
                formula="h回 = 22.856·(v²/9.81/d焦)^(−0.404)·Dr^1.286/d焦^0.286",
            ),
            Result(
                key="v_duct", label="风口风速", value=v_duct, unit="m/s", digits=2,
                symbol="v", primary=False,
                formula="v = (VB+VO₂/60)/60/Sf·TB/T大气/PB·100·803.6/(803.6−WB)",
            ),
            Result(
                key="m_heat", label="热风质量", value=m_heat, unit="kg", digits=5,
                symbol="m", primary=False,
                formula="m = (1.2507·0.79·VB + 1.4289·(0.21·VB+VO₂/60) + (VB+VO₂/60)·WB/1000/(1−WB/803.6))·(1/60/9.81/n)",
            ),
            Result(
                key="taphole", label="铁口深度", value=taphole, unit="m", digits=3,
                symbol="h铁口", primary=False,
                formula="h铁口 = 0.45·D/2",
            ),
            Result(
                key="sf", label="风口面积（输入）", value=sf, unit="m²", digits=4,
                symbol="Sf", primary=False,
                rng=Range(0.475, 0.495, SRC_EXCEL, "原表综合表「风口面积：0.475-0.495」。"),
            ),
        ]


MODELS = [IntegratedBlastFurnace()]
