"""全模型真实验证（浏览器端）：模拟用户操作路径校验「现场标定区间」。

对每个模型：点计算 → 展开区间标定 expander → 改上限 → 点「应用区间」→ 校验区间显示已更新。
"""

from __future__ import annotations

import re
import sys

from playwright.sync_api import sync_playwright

URL = "http://localhost:8501"


def read_state(page):
    t = page.locator("body").inner_text()
    badges = re.findall(r"●\s*(正常|偏低|偏高)", t)
    iv = re.search(r"区间\s*([\d,.]+)\s*~\s*([\d,.]+)", t)
    return badges, (iv.groups() if iv else None)


def main():
    sys.path.insert(0, r"F:\高炉检测软件数学模型")
    from bfmodels.registry import MODELS

    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=True)
        page = browser.new_page(viewport={"width": 1680, "height": 1200})

        for idx, model in enumerate(MODELS):
            page.goto(URL, wait_until="networkidle", timeout=90000)
            page.wait_for_timeout(3000)

            tab_name = model.short_title or model.title
            page.get_by_role("tab", name=tab_name).first.click(timeout=10000)
            page.wait_for_timeout(1200)
            page.get_by_role("button", name="计算").first.click(timeout=10000)
            page.wait_for_timeout(2500)
            page.mouse.wheel(0, 600)
            page.wait_for_timeout(600)

            # 展开当前可见的「现场标定区间」expander
            togs = page.locator("summary:visible")
            opened = False
            for i in range(togs.count()):
                if "现场标定区间" in togs.nth(i).inner_text():
                    togs.nth(i).click()
                    opened = True
                    break
            if not opened:
                results.append((model.title, "无区间标定项", ""))
                print(f"[{idx + 1:2d}] {model.title:<24} 无区间标定项")
                continue
            page.wait_for_timeout(1200)

            inp = page.locator('input[aria-label="上限"]:visible')
            if inp.count() == 0:
                results.append((model.title, "找不到上限输入框", ""))
                print(f"[{idx + 1:2d}] {model.title:<24} ✗ 找不到上限输入框")
                continue

            before_badges, before_iv = read_state(page)
            old = inp.first.input_value()
            try:
                new = float(old) * 1.5 + 1.0
            except ValueError:
                new = 99.0

            inp.first.click(timeout=8000)
            page.keyboard.press("Control+a")
            page.keyboard.type(f"{new:g}")
            page.wait_for_timeout(500)

            btn = page.locator('button:has-text("应用区间"):visible')
            if btn.count() == 0:
                results.append((model.title, "找不到应用区间按钮", ""))
                print(f"[{idx + 1:2d}] {model.title:<24} ✗ 找不到应用区间按钮")
                continue
            btn.first.click(timeout=8000)
            page.wait_for_timeout(2500)

            after_badges, after_iv = read_state(page)
            # 用相对容差：显示格式(%.6g / 千分位)会对大数值做有效位截断
            tol = max(1.0, abs(new) * 1e-5)
            ok = after_iv is not None and abs(
                float(after_iv[1].replace(",", "")) - new) <= tol
            detail = f"{before_iv} -> {after_iv}" if before_iv and after_iv else str(after_iv)
            results.append((model.title, "OK" if ok else "未更新", detail))
            print(f"[{idx + 1:2d}] {model.title:<24} {'✓' if ok else '✗'} {detail}")

        browser.close()

    ok_n = sum(1 for _, s, _ in results if s == "OK")
    skipped = sum(1 for _, s, _ in results if s == "无区间标定项")
    print("=" * 84)
    print(f"区间标定生效 {ok_n} 个，无区间项 {skipped} 个，失败 {len(results) - ok_n - skipped} 个")
    fails = [r for r in results if r[1] not in ("OK", "无区间标定项")]
    for f in fails:
        print("  •", f)
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
