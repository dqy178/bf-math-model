"""模型基础框架：输入项、经验范围、计算结果与模型基类。

设计要点
--------
每个模型（Model）自包含：输入定义、公式、计算逻辑、经验范围。
各模型之间**不共享任何可变状态**，由 Streamlit 层按 model.key 隔离存储，
从而满足"各 Tab 数据完全隔离、不联动"的要求。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence

# ---------------------------------------------------------------------------
# 输入项
# ---------------------------------------------------------------------------


@dataclass
class Input:
    """一个输入框。

    kind:
        number  数值输入
        select  下拉选择（options 为字符串列表）
        bool    开关
    """

    key: str
    label: str
    symbol: str = ""
    unit: str = ""
    default: Any = 0.0
    kind: str = "number"
    step: Optional[float] = None
    lo: Optional[float] = None
    hi: Optional[float] = None
    fmt: str = "%.6g"
    help: str = ""
    note: str = ""
    options: Sequence[str] = ()
    option_labels: Sequence[str] = ()
    group: str = ""

    @property
    def caption(self) -> str:
        bits = [b for b in (self.symbol, self.unit) if b]
        return " / ".join(bits)


# ---------------------------------------------------------------------------
# 经验范围
# ---------------------------------------------------------------------------

# 范围来源说明
SRC_EXCEL = "Excel 综合表「3 经验参数范围」"
SRC_TYPICAL = "工程参考值（待现场标定）"
SRC_REPORT = "现场报表实测值（待核对）"
SRC_NONE = "未标定"


@dataclass
class Range:
    """结果的正常区间。low/high 为 None 表示未标定。"""

    low: Optional[float] = None
    high: Optional[float] = None
    source: str = SRC_TYPICAL
    note: str = ""
    editable: bool = True

    def judge(self, value: Optional[float]) -> str:
        if value is None or self.low is None or self.high is None:
            return "none"
        if value < self.low:
            return "low"
        if value > self.high:
            return "high"
        return "ok"


# ---------------------------------------------------------------------------
# 计算结果
# ---------------------------------------------------------------------------


@dataclass
class Result:
    key: str
    label: str
    value: Optional[float]
    unit: str = ""
    digits: int = 4
    rng: Optional[Range] = None
    formula: str = ""          # 该结果自身的计算式（纯文本）
    latex: str = ""            # 该结果的计算式（LaTeX，可选）
    note: str = ""
    primary: bool = True       # False 表示中间量/辅助量
    symbol: str = ""


# ---------------------------------------------------------------------------
# 模型基类
# ---------------------------------------------------------------------------


class Model:
    """所有计算模型的基类。"""

    key: str = ""
    title: str = ""
    short_title: str = ""      # Tab 标签用的短名（留空则用 title）
    subtitle: str = ""
    group: str = "其他"
    source_sheet: str = ""          # 对应 Excel 工作表名
    description: str = ""
    formula_latex: str = ""         # 主公式（LaTeX）
    formula_excel: str = ""         # 原表公式（Excel 语法）
    inputs: List[Input] = []
    notes: List[str] = []           # 已知问题 / 原表批注
    variant_label: str = ""         # 变体（公式版本）选择的说明

    # 与 Excel 原表算例的对照数据。每项：
    #   {"key","label","symbol","unit","excel","note"}
    # "excel" 为原表算例值（数值或字符串，如 "#DIV/0!"）
    reference: List[Dict[str, Any]] = []

    # ------------------------------------------------------------------
    def compute(self, v: Dict[str, Any]) -> List[Result]:
        raise NotImplementedError

    # 可选：修正版公式的结果。返回 None 表示没有修正版。
    def compute_fixed(self, v: Dict[str, Any]) -> Optional[List[Result]]:
        return None

    @property
    def has_variant(self) -> bool:
        return self.compute_fixed is not Model.compute_fixed

    # ------------------------------------------------------------------
    def default_values(self) -> Dict[str, Any]:
        return {i.key: i.default for i in self.inputs}

    # ------------------------------------------------------------------
    def reference_rows(self) -> List[Dict[str, Any]]:
        """把原表算例值与本工具用默认参数的复算结果对照。"""
        try:
            computed = {r.key: r.value for r in self.compute(self.default_values())}
        except Exception:  # noqa: BLE001
            computed = {}

        rows: List[Dict[str, Any]] = []
        for ref in self.reference:
            key = ref.get("key", "")
            excel = ref.get("excel")
            mine = computed.get(key)
            dev = ""
            if isinstance(excel, (int, float)) and isinstance(mine, (int, float)):
                d = mine - excel
                rel = d / excel * 100 if excel else float("nan")
                dev = f"{d:+.6g}  ({rel:+.3f}%)" if abs(rel) > 1e-9 else "0"
            elif isinstance(excel, str):
                dev = "原表为错误值" if excel.startswith("#") else ""
            rows.append(
                {
                    "项目": ref.get("label", key),
                    "符号": ref.get("symbol", ""),
                    "单位": ref.get("unit", ""),
                    "原表算例值": _excel_cell(excel),
                    "本工具复算": _f(mine),
                    "偏差": dev,
                    "备注": ref.get("note", ""),
                }
            )
        return rows


def _excel_cell(excel: Any) -> str:
    """原表算例值统一转成字符串，避免 DataFrame 列类型混合导致 Arrow 序列化告警。"""
    if excel is None:
        return "—"
    if isinstance(excel, bool):
        return str(excel)
    if isinstance(excel, (int, float)):
        return f"{excel:,}" if float(excel).is_integer() else f"{excel:,.6f}"
    return str(excel)


def _f(v: Optional[float]) -> str:
    if v is None:
        return "—"
    if isinstance(v, float) and v != v:
        return "NaN"
    if v != 0 and abs(v) < 1e-4:
        return f"{v:.6e}"
    return f"{v:,.6f}"


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

PI = 3.1415926


def safe_div(a: float, b: float, default: float = float("nan")) -> float:
    try:
        if b == 0:
            return default
        return a / b
    except ZeroDivisionError:
        return default


def fmt_range(rng: Optional[Range]) -> str:
    if rng is None or rng.low is None or rng.high is None:
        return "未标定区间"
    return f"[{rng.low:,.4g}, {rng.high:,.4g}]"


def numeric_partials(fn, v: Dict[str, Any], keys: Sequence[str],
                     rel: float = 1e-4) -> Dict[str, float]:
    """以中心差分求 fn 对各输入的偏导，用于灵敏度展示。"""
    out: Dict[str, float] = {}
    for k in keys:
        x = v.get(k)
        if not isinstance(x, (int, float)):
            continue
        h = abs(float(x)) * rel if float(x) != 0 else 1e-3
        vp = dict(v)
        vm = dict(v)
        vp[k] = x + h
        vm[k] = x - h
        try:
            out[k] = (fn(vp) - fn(vm)) / (2 * h)
        except Exception:  # noqa: BLE001
            out[k] = float("nan")
    return out
