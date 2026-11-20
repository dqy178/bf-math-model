"""高炉检测软件数学模型计算器 —— 模型层。

每个模块（Model）自带输入、公式、经验范围与计算逻辑，模块之间互不引用、
互不共享状态，从而满足「各 Tab 数据完全隔离、不联动」的要求。
"""

from .base import Input, Model, Range, Result

__all__ = ["Input", "Model", "Range", "Result"]
