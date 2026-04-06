#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统计模块 - 从 xml_analyzer 结果提取统计数据（活跃函数 stats_from_xml）
废弃的 Stats/stats_from_markdown 类已移除（依赖已删除的 marker.py）
"""

import re
from typing import Dict


def stats_from_xml(xml_result: dict) -> Dict:
    """
    便捷函数：从 xml_analyzer 的结构化结果中提取统计数据。

    xml_result 由 xml_analyzer.analyze_xml() 返回，包含：
      - char_count / chinese_char_count / english_word_count
      - abstract_length
      - reference_count / recent_references_count
      - books_count / theses_count / journals_count
      - total_tables / native_table_count / screenshot_table_count
      - media_image_count / standalone_image_count

    Returns格式与 stats_from_markdown() 保持一致，供 Reviewer 共用。

    Args:
        xml_result: xml_analyzer.analyze_xml() 的返回值

    Returns:
        完整统计字典
    """
    # 题目字数（从 title 字段估算）
    title_text = xml_result.get("title") or ""
    # 去除章节编号等前缀
    title_clean = re.sub(r"^\d+[\.、]\s*", "", title_text).strip()
    title_words = len(title_clean)  # 中文按字符计

    abstract_text = xml_result.get("abstract") or ""
    body_text     = xml_result.get("body_text") or ""

    # 中文字符数（正文）
    chinese_chars_body = xml_result.get("chinese_char_count", 0)
    # 英文词数（正文）
    english_words_body = xml_result.get("english_word_count", 0)
    # 数字个数（正文）
    digit_count_body   = xml_result.get("digit_count", 0)

    # 字数合计（中文 char + 英文 word + 数字）
    body_word_count = chinese_chars_body + english_words_body + digit_count_body

    # 摘要字数（已有字符数，直接用）
    abstract_word_count = xml_result.get("abstract_length", 0)

    ref_count       = xml_result.get("reference_count", 0)
    recent_ref      = xml_result.get("recent_references_count", 0)
    books           = xml_result.get("books_count", 0)
    theses          = xml_result.get("theses_count", 0)
    journals        = xml_result.get("journals_count", 0)
    foreign         = xml_result.get("foreign_count", 0)
    total_tables = xml_result.get("total_tables", 0)
    native_tbl   = xml_result.get("native_table_count", 0)
    screenshot_tbl = xml_result.get("screenshot_table_count", 0)
    media_imgs   = xml_result.get("media_image_count", 0)
    standalone_imgs = xml_result.get("standalone_image_count", 0)

    # 第一人称统计（从正文中检测）
    first_person_i   = len(re.findall(r'\b我\b', body_text))
    first_person_we   = len(re.findall(r'\b我们\b', body_text))

    return {
        'word_count': {
            'title':   title_words,
            'abstract': abstract_word_count,
            'body':    body_word_count,
        },
        'references': {
            'total':   ref_count,
            'foreign': foreign,
            'journals': journals,
            'books':   books,
            'theses':  theses,
            'recent':  recent_ref,
        },
        'tables': {
            'total':       total_tables,
            'native':      native_tbl,
            'screenshots': screenshot_tbl,
        },
        'images': {
            'media_images':      media_imgs,
            'standalone_images': standalone_imgs,
        },
        'writing_specs': {
            'first_person_count': first_person_i + first_person_we,
        },
        'marker_status': {
            # xml_analyzer 直接提取，不需要 markdown 标记
            'abstract':   bool(abstract_text),
            'body':       bool(body_text),
            'references': bool(ref_count > 0),
            'fully_marked': bool(abstract_text and body_text),
            'from_xml': True,
        },
        # 保留原始 xml 数据供其他模块使用
        '_xml': xml_result,
    }
