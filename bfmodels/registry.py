"""模型注册表：集中登记所有计算模块及其展示顺序。"""

from __future__ import annotations

from typing import List

from .base import Model
from .modules.burden import MODELS as BURDEN_MODELS
from .modules.combustion import MODELS as COMBUSTION_MODELS
from .modules.gas import MODELS as GAS_MODELS
from .modules.integrated import MODELS as INTEGRATED_MODELS
from .modules.raceway import MODELS as RACEWAY_MODELS
from .modules.vessel import MODELS as VESSEL_MODELS

# 展示顺序：与用户习惯一致，先列高频单模型 Tab，综合表放最后。
_MODELS: List[Model] = [
    # 回旋区与布料
    *RACEWAY_MODELS,
    # 煤气流动
    *GAS_MODELS,
    # 燃烧与鼓风
    *COMBUSTION_MODELS,
    # 冶炼过程
    *BURDEN_MODELS,
    # 炉缸与炉衬
    *VESSEL_MODELS,
    # 综合计算
    *INTEGRATED_MODELS,
]

MODELS = _MODELS

# 便于测试：key -> model
BY_KEY = {m.key: m for m in MODELS}

__all__ = ["MODELS", "BY_KEY"]
