"""Streamlit 应用冒烟测试：无头运行 app.py，检查异常、Tab 数量与「计算」按钮行为。

运行：
    python tests/test_app_smoke.py
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from streamlit.testing.v1 import AppTest  # noqa: E402

APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")


def main() -> int:
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()

    if at.exception:
        print("❌ 应用抛出异常：")
        for e in at.exception:
            print("   ", e.value)
        return 1

    print(f"✓ 应用首屏渲染无异常")
    print(f"  标题：{[t.value for t in at.title]}")
    print(f"  Tab 数量：{len(at.tabs)}")

    titles = [t.label for t in at.tabs]
    print("  Tab 列表：")
    for i, t in enumerate(titles, 1):
        print(f"    {i:2d}. {t}")

    buttons = at.button
    print(f"  按钮数量：{len(buttons)}（每个 Tab 一个「计算」按钮）")

    # 点击第一个「计算」按钮，验证结果能渲染出来
    calc_buttons = [b for b in buttons if b.label == "计算"]
    if not calc_buttons:
        print("❌ 未找到「计算」按钮")
        return 1
    print(f"  「计算」按钮数量：{len(calc_buttons)}")

    calc_buttons[0].click()
    at.run()

    if at.exception:
        print("❌ 点击「计算」后抛出异常：")
        for e in at.exception:
            print("   ", e.value)
        return 1

    metrics = [m for m in at.metric if m.value not in ("", None)]
    print(f"  点击第 1 个 Tab 的「计算」后，渲染出 {len(metrics)} 个指标值")
    for m in metrics[:6]:
        print(f"    · {m.label or '(主指标)'} = {m.value}")

    # 逐个点击所有「计算」按钮，确认每个模型都能算通
    print("\n  逐个 Tab 点击「计算」：")
    at2 = AppTest.from_file(APP, default_timeout=120)
    at2.run()
    calc_buttons = [b for b in at2.button if b.label == "计算"]
    ok = 0
    for idx, b in enumerate(calc_buttons):
        b.click()
        at2.run()
        if at2.exception:
            print(f"    ✗ Tab {idx + 1} 异常：{at2.exception[0].value}")
            continue
        ok += 1
    print(f"  可正常计算的 Tab：{ok}/{len(calc_buttons)}")

    if at2.exception:
        return 1
    print("\n✅ 冒烟测试通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
