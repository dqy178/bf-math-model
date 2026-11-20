"""高炉检测软件数学模型计算器 —— Streamlit 入口。

启动：
    streamlit run app.py
"""

from __future__ import annotations

import os

import streamlit as st
import streamlit.components.v1 as components

from bfmodels.registry import MODELS
from bfmodels.ui import render_app

SOURCE_FILE = "高炉检测软件数学模型.xlsx"

# 页面样式：
# 1) 屏蔽 Streamlit 1.51 自带的「AI 技能」提示气泡
#    —— 它固定在页面右下角，会遮挡结果卡片右下方的按钮（现场标定区间的「应用区间」）。
# 2) 让顶部 Tab 栏的横向滚动条常显
#    —— Streamlit 默认用 scrollbar-width:none 隐藏滚动条，23 个 Tab 溢出时用户
#       看不到可横向滚动，容易误以为后面的 Tab 不可用。
APP_CSS = """
<style>
/* 1) 屏蔽提示气泡 */
[data-testid="stSkillsNudge"],
[data-testid="stSkillsNudgeAnchor"],
[data-testid="stAppViewContainer"] > div[data-testid="stToolbar"],
div[data-testid="stStatusWidget"] {
    display: none !important;
}

/* 2) 顶部 Tab 栏横向滚动条常显 */
div[data-testid="stTabs"] > div > div > div,
div[data-testid="stTabs"] > div > div {
    scrollbar-width: auto !important;
    scrollbar-color: #cbd5e1 transparent !important;
    padding-bottom: 2px;
}
div[data-testid="stTabs"] > div > div > div::-webkit-scrollbar,
div[data-testid="stTabs"] > div > div::-webkit-scrollbar {
    height: 9px;
}
div[data-testid="stTabs"] > div > div > div::-webkit-scrollbar-track,
div[data-testid="stTabs"] > div > div::-webkit-scrollbar-track {
    background: #f1f5f9;
    border-radius: 5px;
}
div[data-testid="stTabs"] > div > div > div::-webkit-scrollbar-thumb,
div[data-testid="stTabs"] > div > div::-webkit-scrollbar-thumb {
    background: #cbd5e1;
    border-radius: 5px;
}
div[data-testid="stTabs"] > div > div > div::-webkit-scrollbar-thumb:hover,
div[data-testid="stTabs"] > div > div::-webkit-scrollbar-thumb:hover {
    background: #94a3b8;
    background-repeat: no-repeat;
    background-size: 40px 100%, 40px 100%, 14px 100%, 14px 100%;
    background-attachment: local, local, scroll, scroll;
}
</style>
"""

# 让顶部 Tab 栏支持鼠标滚轮横向滚动。
# Streamlit 把 tab 栏做成 overflow-x:scroll + scrollbar-width:none 的横向滚动容器，
# 但浏览器不会把垂直滚轮自动转成横向滚动，且滚动条被隐藏，
# 导致 23 个 Tab 溢出后「看不到也滚不动」。
# 这里通过同源 iframe 注入 wheel 事件转换（iframe 与主页面同源，可访问 parent.document）。
TAB_SCROLL_JS = """
<script>
(function () {
  var P = window.parent;
  function findTabScroller() {
    var doc = P.document;
    var tabs = Array.prototype.slice.call(
      doc.querySelectorAll('[role="tab"]'));
    if (!tabs.length) return null;
    var e = tabs[0];
    while (e && e !== doc.body) {
      var cs = P.getComputedStyle(e);
      if (cs.overflowX === 'scroll' || cs.overflowX === 'auto') return e;
      e = e.parentElement;
    }
    return null;
  }
  function attach() {
    var c = findTabScroller();
    if (!c || c.dataset.wheelBound === '1') return;
    c.dataset.wheelBound = '1';
    c.addEventListener('wheel', function (ev) {
      if (!ev.deltaY) return;
      var before = c.scrollLeft;
      var max = c.scrollWidth - c.clientWidth;
      var next = before + ev.deltaY;
      if (next < 0) next = 0;
      if (next > max) next = max;
      if (next !== before) {
        c.scrollLeft = next;
        ev.preventDefault();
      }
    }, { passive: false });
  }
  function tick() { try { attach(); } catch (e) {} }
  tick();
  setInterval(tick, 400);
})();
</script>
"""


def main() -> None:
    here = os.path.dirname(os.path.abspath(__file__))
    src = os.path.join(here, SOURCE_FILE)
    st.markdown(APP_CSS, unsafe_allow_html=True)
    components.html(TAB_SCROLL_JS, height=0, width=0)
    render_app(MODELS, source_file=SOURCE_FILE if os.path.exists(src) else SOURCE_FILE)


if __name__ == "__main__":
    main()
