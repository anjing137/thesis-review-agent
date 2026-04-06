#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文评价工具模块包 (v0.2.2+)

统一导出接口
"""

from .converter import Converter, convert_paper
from .stats import stats_from_xml
from .xml_analyzer import analyze_xml

__all__ = [
    # 转换
    'Converter',
    'convert_paper',
    # 统计
    'stats_from_xml',
    # XML 分析（docx 路径）
    'analyze_xml',
]

__version__ = "0.4.0"
