"""Streamlit 渲染层：统一的模型页（输入区 + 公式区 + 结果区 + 经验范围比对）。

数据隔离规则
------------
每个模型的输入值与计算结果分别存放在 session_state 的
``__inp__<model.key>__<input.key>`` 与 ``__res__<model.key>`` 键下，
不同模型之间键名不冲突，因此各 Tab 完全隔离、互不联动。
"""

from __future__ import annotations

import html
from typing import Any, Dict, List, Optional

import streamlit as st

from .base import Input, Model, Range, Result

INPUT_PREFIX = "__inp__"
RESULT_PREFIX = "__res__"
VARIANT_PREFIX = "__var__"


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------


def _fmt_num(v: Optional[float], digits: int = 4) -> str:
    if v is None:
        return "—"
    if isinstance(v, float) and (v != v):  # NaN
        return "NaN"
    if v is not None and abs(v) >= 1e6:
        return f"{v:,.{digits}f}"
    if v is not None and abs(v) < 1e-4 and v != 0:
        return f"{v:.3e}"
    return f"{v:,.{digits}f}"


def _judge_style(status: str) -> tuple:
    """返回 (徽标, 颜色, 文本)。"""
    if status == "ok":
        return "正常", "#16a34a", "处于经验区间内"
    if status == "low":
        return "偏低", "#dc2626", "低于经验区间下限"
    if status == "high":
        return "偏高", "#d97706", "高于经验区间上限"
    return "未标定", "#64748b", "尚无经验区间可比对"


def _range_bar_html(value: Optional[float], rng: Optional[Range]) -> str:
    """绘制一条区间比对条：灰色底 + 绿色正常区间 + 数值指针。"""
    if rng is None or rng.low is None or rng.high is None or value is None:
        return ""
    try:
        low, high, val = float(rng.low), float(rng.high), float(value)
    except (TypeError, ValueError):
        return ""
    if not (high > low):
        return ""

    span = high - low
    # 视图区间：正常区间左右各留 40% 余量，方便看到越界幅度
    pad = span * 0.4
    view_lo, view_hi = low - pad, high + pad
    vspan = view_hi - view_lo

    def pct(v: float) -> float:
        return max(0.0, min(100.0, (v - view_lo) / vspan * 100.0))

    seg_left = pct(low)
    seg_width = pct(high) - seg_left
    marker = pct(val)

    if val < low:
        color = "#dc2626"
    elif val > high:
        color = "#d97706"
    else:
        color = "#16a34a"

    return f"""
<div style="margin:6px 0 2px 0;">
  <div style="position:relative;height:22px;background:#eef2f7;border-radius:6px;overflow:hidden;">
    <div style="position:absolute;left:{seg_left:.2f}%;width:{seg_width:.2f}%;top:0;bottom:0;
                background:#dcfce7;border-left:2px solid #86efac;border-right:2px solid #86efac;"></div>
    <div style="position:absolute;left:{marker:.2f}%;top:0;bottom:0;width:3px;
                background:{color};transform:translateX(-1.5px);"></div>
  </div>
  <div style="display:flex;justify-content:space-between;font-size:11px;color:#64748b;margin-top:2px;">
    <span>{_fmt_num(low, 4)}</span>
    <span style="color:{color};font-weight:600;">当前 {_fmt_num(val, 4)}</span>
    <span>{_fmt_num(high, 4)}</span>
  </div>
</div>
"""


# ---------------------------------------------------------------------------
# 输入渲染
# ---------------------------------------------------------------------------


def _render_input(inp: Input, prefix: str) -> Any:
    sk = f"{INPUT_PREFIX}{prefix}__{inp.key}"
    if sk not in st.session_state:
        st.session_state[sk] = (float(inp.default) if inp.kind == "number"
                                else inp.default)

    label = f"{inp.label}"
    if inp.caption:
        label += f"（{inp.caption}）"

    if inp.kind == "select":
        labels = list(inp.option_labels) if inp.option_labels else list(inp.options)
        label_map = dict(zip(inp.options, labels))
        if st.session_state[sk] not in inp.options:
            st.session_state[sk] = inp.default
        return st.selectbox(
            label,
            options=list(inp.options),
            format_func=lambda o: label_map[o],
            key=sk,
            help=inp.help,
        )

    if inp.kind == "bool":
        return st.checkbox(label, key=sk, help=inp.help)

    step = inp.step
    if step is None:
        step = 1.0
    # Streamlit 要求 value/min_value/max_value/step 类型一致，统一为 float
    return st.number_input(
        label,
        min_value=float(inp.lo) if inp.lo is not None else None,
        max_value=float(inp.hi) if inp.hi is not None else None,
        step=float(step),
        key=sk,
        format=inp.fmt or "%.10g",
        help=inp.help,
    )


def _chunked(items: List[Any], size: int) -> List[List[Any]]:
    """按每行 size 个分块，避免 zip(cols, items) 截断多余的结果卡片。"""
    return [items[i:i + size] for i in range(0, len(items), size)]


# ---------------------------------------------------------------------------
# 结果渲染
# ---------------------------------------------------------------------------


def _render_result_card(res: Result, model: Model) -> None:
    rng = res.rng
    if rng is not None and rng.editable:
        # 允许用户现场标定区间，存入独立 key，不影响其他模型
        lo_key = f"{RESULT_PREFIX}{model.key}__rng__{res.key}__lo"
        hi_key = f"{RESULT_PREFIX}{model.key}__rng__{res.key}__hi"
        if lo_key not in st.session_state:
            st.session_state[lo_key] = rng.low if rng.low is not None else 0.0
        if hi_key not in st.session_state:
            st.session_state[hi_key] = rng.high if rng.high is not None else 0.0

    st.markdown(
        f"<div style='font-size:13px;color:#475569;margin-bottom:2px;'>"
        f"{html.escape(res.label)}"
        + (f" <span style='color:#94a3b8;'>{html.escape(res.unit)}</span>" if res.unit else "")
        + "</div>",
        unsafe_allow_html=True,
    )
    st.metric(label=res.label, value=_fmt_num(res.value, res.digits), delta=None,
              label_visibility="collapsed")
    if res.note:
        st.caption(res.note)

    if rng is not None:
        # 读取本 Tab 现场标定后的区间（优先 session_state，其次原始范围）
        cur_lo = st.session_state.get(lo_key, rng.low)
        cur_hi = st.session_state.get(hi_key, rng.high)
        cur_rng = Range(low=cur_lo, high=cur_hi, source=rng.source, note=rng.note,
                        editable=rng.editable)

        status = cur_rng.judge(res.value)
        badge, color, text = _judge_style(status)
        bar = _range_bar_html(res.value, cur_rng)
        src = f" · {cur_rng.source}" if cur_rng.source else ""
        st.markdown(
            f"<div style='font-size:12px;color:{color};font-weight:600;'>● {badge}"
            f"<span style='color:#64748b;font-weight:400;'> · {html.escape(text)}</span></div>"
            + bar
            + f"<div style='font-size:11px;color:#94a3b8;'>区间 {html.escape(_fmt_num(cur_rng.low,4))}"
            f" ~ {html.escape(_fmt_num(cur_rng.high,4))}{html.escape(src)}</div>",
            unsafe_allow_html=True,
        )
        if cur_rng.note:
            st.caption(f"📌 {cur_rng.note}")

        if cur_rng.editable and cur_rng.low is not None and cur_rng.high is not None:
            with st.expander("现场标定区间", expanded=False):
                # 用独立 form 承载：输入不触发整页重载，点「应用区间」才提交生效
                with st.form(f"rngform_{model.key}__{res.key}"):
                    c1, c2 = st.columns(2)
                    with c1:
                        lo_in = st.number_input("下限", key=lo_key, format="%.6g")
                    with c2:
                        hi_in = st.number_input("上限", key=hi_key, format="%.6g")
                    apply = st.form_submit_button("应用区间", type="primary")
                if apply:
                    # form 提交时 Streamlit 已把 widget 的新值写入
                    # session_state[lo_key] / [hi_key]，此处不可再手动赋值
                    st.rerun()
                st.caption("修改上下限后点击「应用区间」立即刷新比对结果，仅本 Tab 生效。")

    if res.formula:
        with st.expander("该结果的计算式", expanded=False):
            st.code(res.formula, language="text")


# ---------------------------------------------------------------------------
# 模型页
# ---------------------------------------------------------------------------


def render_model(model: Model) -> None:
    st.subheader(model.title)
    if model.subtitle:
        st.caption(model.subtitle)
    if model.description:
        st.info(model.description, icon="📘")

    # ---------------- 输入区 ----------------
    with st.form(f"form_{model.key}"):
        groups: Dict[str, List[Input]] = {}
        for inp in model.inputs:
            groups.setdefault(inp.group or "输入参数", []).append(inp)

        values: Dict[str, Any] = {}
        for gname, items in groups.items():
            if gname != "输入参数":
                st.markdown(f"**{gname}**")
            cols = st.columns(2)
            for idx, inp in enumerate(items):
                with cols[idx % 2]:
                    values[inp.key] = _render_input(inp, model.key)
        submitted = st.form_submit_button("计算", type="primary", width="stretch")

    if submitted:
        try:
            results = model.compute(values)
            st.session_state[f"{RESULT_PREFIX}{model.key}"] = results
            st.session_state[f"{RESULT_PREFIX}{model.key}__values"] = values
            st.session_state[f"{RESULT_PREFIX}{model.key}__err"] = None
        except Exception as exc:  # noqa: BLE001
            st.session_state[f"{RESULT_PREFIX}{model.key}__err"] = f"{type(exc).__name__}: {exc}"

    # ---------------- 公式区 ----------------
    if model.formula_latex or model.formula_excel:
        with st.expander("计算公式", expanded=False):
            if model.formula_latex:
                st.markdown("**公式（排版）**")
                st.latex(model.formula_latex)
            if model.formula_excel:
                st.markdown("**原表 Excel 公式**")
                st.code(model.formula_excel, language="excel")

    if model.variant_label:
        st.caption(model.variant_label)

    # ---------------- 变体选择 ----------------
    results = st.session_state.get(f"{RESULT_PREFIX}{model.key}")
    err = st.session_state.get(f"{RESULT_PREFIX}{model.key}__err")
    if err:
        st.error(f"计算失败：{err}")

    if results is None:
        st.caption("填写左侧参数后点击「计算」。")
        return

    # ---------------- 结果区 ----------------
    st.markdown("---")
    st.markdown("**计算结果**")

    primary = [r for r in results if r.primary]
    secondary = [r for r in results if not r.primary]

    if primary:
        for chunk in _chunked(primary, 3):
            cols = st.columns(len(chunk))
            for col, res in zip(cols, chunk):
                with col:
                    with st.container(border=True):
                        _render_result_card(res, model)

    if secondary:
        with st.expander("中间量 / 辅助计算结果", expanded=False):
            for chunk in _chunked(secondary, 3):
                cols = st.columns(len(chunk))
                for col, res in zip(cols, chunk):
                    with col:
                        with st.container(border=True):
                            _render_result_card(res, model)

    # ---------------- 与 Excel 原表数值对照 ----------------
    with st.expander("与 Excel 原表数值对照", expanded=False):
        st.caption("默认参数取自《高炉检测软件数学模型.xlsx》对应工作表的算例。")
        st.dataframe(
            _excel_reference_table(model),
            width="stretch",
            hide_index=True,
        )

    if model.notes:
        with st.expander("⚠️ 已知问题与原表批注", expanded=True):
            for n in model.notes:
                st.warning(n, icon="⚠️")


def _excel_reference_table(model: Model) -> List[Dict[str, Any]]:
    return model.reference_rows()


# ---------------------------------------------------------------------------
# 应用外壳
# ---------------------------------------------------------------------------



def render_app(models: List[Model], source_file: str = "") -> None:
    from collections import OrderedDict

    st.set_page_config(page_title="高炉检测软件数学模型计算器", page_icon="🔥", layout="wide")

    st.title("🔥 高炉检测软件数学模型计算器")
    st.caption(
        "每个 Tab 为一个独立计算器：参数独立、结果独立，不与其他模块联动。"
        + (f"　数据来源：{source_file}" if source_file else "")
    )

    grouped: "OrderedDict[str, List[Model]]" = OrderedDict()
    for m in models:
        grouped.setdefault(m.group, []).append(m)

    with st.sidebar:
        st.markdown("### 模型导航")
        st.caption(f"共 {len(models)} 个计算模块 / {len(grouped)} 个分类")
        st.caption("（顶部 Tab 较多，可横向滚动；也可在本清单中按名称查找）")
        for gname, items in grouped.items():
            st.markdown(f"**{gname}**")
            for m in items:
                st.markdown(f"- {m.title}")
        st.divider()
        st.markdown("**使用说明**")
        st.markdown(
            "- 顶部为平铺的 Tab，每个 Tab 一个模型\n"
            "- 修改参数后需点击该 Tab 的「计算」\n"
            "- 结果卡片下可展开「现场标定区间」调整经验上下限\n"
            "- 「与 Excel 原表数值对照」可核对复算偏差"
        )

    tabs = st.tabs([m.short_title or m.title for m in models])
    for tab, model in zip(tabs, models):
        with tab:
            st.caption(f"分类：{model.group}　|　原表工作表：{model.source_sheet}")
            render_model(model)
