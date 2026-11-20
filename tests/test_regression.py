"""回归测试：每个模型用 Excel 原表算例的默认参数复算，核对与表内缓存值是否一致。

运行：
    python tests/test_regression.py
退出码 0 表示全部通过。
"""

from __future__ import annotations

import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bfmodels.registry import MODELS  # noqa: E402


def close(a: float, b: float, tol_rel: float = 1e-9, tol_abs: float = 1e-9) -> bool:
    if isinstance(a, str) or isinstance(b, str):
        return str(a) == str(b)
    if math.isnan(a) or math.isnan(b):
        return math.isnan(a) and math.isnan(b)
    if abs(b) < 1e-12:
        return abs(a - b) <= tol_abs
    return abs(a - b) / abs(b) <= tol_rel


def main() -> int:
    total = 0
    passed = 0
    skipped = 0
    failures = []

    for model in MODELS:
        try:
            values = model.default_values()
            results = {r.key: r for r in model.compute(values)}
        except Exception as exc:  # noqa: BLE001
            failures.append(f"[{model.key}] compute 抛异常：{type(exc).__name__}: {exc}")
            continue

        print("=" * 92)
        print(f"■ {model.title}   ({model.key})   ← {model.source_sheet}")
        print("=" * 92)
        for ref in model.reference:
            total += 1
            key = ref.get("key")
            excel = ref.get("excel")
            label = ref.get("label", key)
            if key not in results:
                failures.append(f"[{model.key}] 结果缺少 key={key}")
                print(f"  ✗ {label:<28} 缺少结果 key={key}")
                continue
            mine = results[key].value
            if isinstance(excel, str):
                skipped += 1
                print(f"  – {label:<28} 原表={excel}  复算={mine}")
                continue
            if excel is None:
                skipped += 1
                print(f"  – {label:<28} 原表=—（参考项）  复算={mine}")
                continue
            ok = close(mine, excel)
            if ok:
                passed += 1
                print(f"  ✓ {label:<28} 原表={excel!r:<24} 复算={mine!r}")
            else:
                diff = mine - excel
                rel = diff / excel * 100 if excel else float("nan")
                failures.append(
                    f"[{model.key}] {label}: 复算={mine!r} 原表={excel!r} 偏差={diff:+.6g} ({rel:+.4f}%)"
                )
                print(f"  ✗ {label:<28} 原表={excel!r:<24} 复算={mine!r}  偏差={diff:+.6g}")
        # 附带展示所有结果键，便于检查
        print(f"    结果键: {', '.join(results.keys())}")
        print()

    print("=" * 92)
    print(f"合计 {total} 项：核对通过 {passed} 项，失败 {total - passed - skipped} 项，"
          f"参考项（无原表对照值）{skipped} 项")
    if failures:
        print("\n失败明细：")
        for f in failures:
            print("  •", f)
        return 1
    print("全部通过 ✅")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
