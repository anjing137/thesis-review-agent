"""
xml_analyzer.py
从 docx 文件的原始 XML 中提取全量结构化信息，替代 Pandoc 转 Markdown 的方案。

核心优势：
- 表格：<w:tbl> 标签直接计数，可区分原生表格 vs 截图表格
- 图片：word/media/ 目录文件精确计数
- 语义结构：标题层级/章节边界均有 XML 标签可定位
- 文字提取：直接读 XML 文字节点，无格式丢失
"""

import zipfile
import re
import os
from pathlib import Path
from typing import Optional, List, Tuple
from collections import Counter
import xml.etree.ElementTree as ET
from datetime import datetime


# ── XML 命名空间 ──────────────────────────────────────────────────────────────
NS = {
    "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o":  "urn:schemas-microsoft-com:office:office",
    "v":  "urn:schemas-microsoft-com:vml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic":"http://schemas.openxmlformats.org/drawingml/2006/picture",
}

# 注册命名空间（避免 ns0: 前缀）
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def _tag(local: str) -> str:
    """生成带命名空间的标签名，如 'w:p'"""
    prefix, name = local.split(":")
    return f"{{{NS[prefix]}}}{name}"


# ── 全文本索引（marker 边界检测的根基）───────────────────────────────────────

def _build_text_index(xml_root: ET.Element) -> Tuple[str, List[int]]:
    """
    从 XML 中提取所有段落文本，拼成完整纯文本，同时记录每个段落的起始字符位置。

    Returns:
        full_text:           所有段落拼接而成的完整纯文本（含 \n 分隔）
        para_start_positions: 第 i 个段落在 full_text 中的起始字符下标
                              即 full_text[para_start_positions[i]] == 段落 i 的首字
    """
    all_paragraphs = xml_root.findall(f".//{_tag('w:p')}")
    parts: List[str] = []
    starts: List[int] = []

    for p in all_paragraphs:
        text = "".join(t.text or "" for t in p.findall(f".//{_tag('w:t')}"))
        starts.append(sum(len(x) for x in parts))
        parts.append(text + "\n")

    full_text = "".join(parts)
    return full_text, starts


def _char_pos_to_para_idx(char_pos: int, para_starts: List[int]) -> int:
    """
    把 full_text 中的字符位置映射回段落下标。
    使用 bisect 找到最后一个 start <= char_pos 的段落。
    """
    from bisect import bisect_right
    idx = bisect_right(para_starts, char_pos) - 1
    return max(0, idx)


# ── Marker 风格边界检测（移植自 marker.py）──────────────────────────────────

def _marker_abstract_boundary(full_text: str) -> Tuple[int, int]:
    """
    找摘要边界。移植自 marker.py Marker.find_abstract_boundary()。
    Returns (abstract_start, abstract_end) 字符位置。
    """
    # 移除非摘要标记（防止干扰）
    content = re.sub(r'<!--\s*摘要开始\s*-->|<!--\s*摘要结束\s*-->', '', full_text)

    abstract_start = None

    # 格式1: # 摘要 或 # 摘　要
    for match in re.finditer(r'^#\s+摘\s?要', content, re.MULTILINE):
        abstract_start = match.start()
        break

    # 格式2: **摘要** 或 **摘　要**
    if abstract_start is None:
        for match in re.finditer(r'\*\*摘\s?要\*\*', content):
            abstract_start = match.start()
            break

    # 格式3: 摘　要（全角空格开头，无标记）
    if abstract_start is None:
        for match in re.finditer(r'^摘\s?要\s*$', content, re.MULTILINE):
            abstract_start = match.start()
            break

    if abstract_start is None:
        return (-1, -1)

    # 摘要结束：找中文关键词行（全文搜索，不限范围，再过滤必须在 abstract_start 之后）
    # 注意：允许前导空格/缩进（WORD 常用），关键词后支持全角/半角冒号
    kw_patterns = [
        r'\*\*关键词[：:站][^\n]+',
        r'^\s*关键词[：:站][^\n]+',
    ]

    kw_match = None
    for pattern in kw_patterns:
        # 全文搜索，再验证在 abstract_start 之后（防止把文档开头附近的无关文本当关键词）
        for m in re.finditer(pattern, content, re.MULTILINE):
            if m.start() > abstract_start:
                kw_match = m
                break
        if kw_match:
            break

    if kw_match:
        # abstract_end = 关键词行末（该行之后连续换行处）
        # 注意：关键词行本身以 \n 结尾，p28 空段落又贡献一个 \n，
        # 实际是 \n\n（两个连续），用 \n+ 匹配（允许1个或多个）
        line_end = kw_match.end()
        blank_match = re.search(r'\n+', content[line_end:])
        if blank_match:
            abstract_end = line_end + blank_match.start() + blank_match.group().__len__()
        else:
            abstract_end = line_end
    else:
        # 找不到关键词（罕见），保守以 abstract_start + 较长距离为上界
        abstract_end = abstract_start + 2000

    return (abstract_start, abstract_end)


def _marker_body_boundary(full_text: str) -> Tuple[int, int]:
    """
    找正文边界。移植自 marker.py Marker.find_body_boundary()。
    Returns (body_start, body_end) 字符位置。
    """
    # 移除旧标记
    content = re.sub(r'<!--\s*正文开始\s*-->|<!--\s*正文结束\s*-->', '', full_text)

    body_start = None
    body_end = None

    # === 找正文结束位置（参考文献之前）===
    ref_patterns = [
        r'^#{1,3}\s+参考文献',
        r'\*\*参考文献\*\*',
        r'\[]{#_Toc\d+\s*\.[a-z]+\}参考文献',
        r'参考文献\s*\[',
    ]

    ref_start_in_body = None
    for pattern in ref_patterns:
        for match in re.finditer(pattern, content, re.MULTILINE):
            pos = match.start()
            if pos > 5:
                prev_context = content[max(0, pos - 20):pos]
                if re.search(r'\[参考文献\s*\[\d+\]', prev_context):
                    continue
                if '\n' in prev_context and len(prev_context.split('\n')[-1]) < 30:
                    continue
            ref_start_in_body = pos
            break
        if ref_start_in_body is not None:
            break

    body_end = ref_start_in_body if ref_start_in_body is not None else len(content)

    # === 找正文开始位置（目录结束后）===

    # 方法1：目录项格式 [章节名\l]
    toc_items = []
    for match in re.finditer(r'\[([^\]]+)\]\(\\[lst]\)', content):
        toc_items.append(match)

    # 标准链接格式 [章节名](#anchor))
    for match in re.finditer(r'\[([^\]]+)\]\(#([^)]+)\)\)', content):
        toc_items.append(match)

    if toc_items:
        last_toc = toc_items[-1]
        last_toc_end = last_toc.end()
        body_start = content.find('\n\n', last_toc_end)
        if body_start != -1:
            body_start += 2
        else:
            body_start = last_toc_end

    # 方法2：备用 - 找第一个章节标题
    if body_start is None or body_start < 0:
        chapter_patterns = [
            r'^#{1,3}\s+[一二三四五六七八九十百]+[、，]',
            r'^#{1,3}\s+第[一二三四五六七八九十百]+[章节节]',
            r'^#{1,3}\s+\([一二三四五六七]\)',
            r'\n\n[一二三四五六七八九十]+、',
        ]
        for pattern in chapter_patterns:
            match = re.search(pattern, content)
            if match:
                body_start = match.start()
                break

    return (body_start if body_start else -1, body_end if body_end else -1)


def _marker_references_boundary(full_text: str) -> Tuple[int, int]:
    """
    找参考文献边界。移植自 marker.py Marker.find_references_boundary()。
    Returns (ref_start, ref_end) 字符位置。
    """
    # 移除旧标记
    content = re.sub(r'<!--\s*参考文献开始\s*-->|<!--\s*参考文献结束\s*-->', '', full_text)

    ref_start = None

    # 格式1: # 参考文献 或 ## 参考文献
    for match in re.finditer(r'^#{1,3}\s+参考文献', content, re.MULTILINE):
        ref_start = match.start()
        break

    # 格式2: **参考文献**
    if ref_start is None:
        for match in re.finditer(r'\*\*参考文献\*\*', content):
            ref_start = match.start()
            break

    # 格式3: Pandoc 锚点格式
    if ref_start is None:
        for match in re.finditer(r'\[]{#_Toc\d+\s*\.[a-z]+\}[参考文献\(]+', content):
            pos = match.start()
            if pos > 0 and content[pos - 1] == '[':
                continue
            ref_start = pos
            break

    # 格式4: 参考文献 [ 格式（排除目录中的）
    if ref_start is None:
        for match in re.finditer(r'参考文献\s*\[', content):
            pos = match.start()
            if pos > 5:
                prev_context = content[max(0, pos - 20):pos]
                if re.search(r'\[参考文献\s*\[\d+\]', prev_context):
                    continue
            ref_start = pos
            break

    if ref_start is None:
        return (-1, -1)

    ref_end = len(content)
    return (ref_start, ref_end)


def _marker_boundaries(full_text: str) -> dict:
    """
    用 marker 风格正则一次性检测摘要/正文/参考文献三个边界。
    Returns dict:
        abstract_start, abstract_end,
        body_start, body_end,
        ref_start, ref_end
    全部为字符位置，失败返回全 -1。
    """
    abs_s, abs_e = _marker_abstract_boundary(full_text)
    bod_s, bod_e = _marker_body_boundary(full_text)
    ref_s,  ref_e  = _marker_references_boundary(full_text)
    return {
        "abstract_start": abs_s,
        "abstract_end": abs_e,
        "body_start": bod_s,
        "body_end": bod_e,
        "ref_start": ref_s,
        "ref_end": ref_e,
    }


# ── 封面信息提取 ──────────────────────────────────────────────────────────────

def _extract_cover_info(xml_root: ET.Element) -> dict:
    """
    从 document.xml 中提取封面信息。
    封面通常是一个表格（<w:tbl>），包含姓名、学号等专业信息。
    也可能是一系列段落，靠标签文字（如 '姓名：'）来定位。
    """
    result = {
        "title": None,
        "student_name": None,
        "student_id": None,
        "school": None,
        "major": None,
        "advisor": None,
    }

    # 策略1：从封面表格中提取（大多数论文用表格做封面）
    tables = xml_root.findall(f".//{_tag('w:tbl')}")
    for tbl in tables:
        cells_text = []
        for tc in tbl.findall(f".//{_tag('w:tc')}"):
            cell_paragraphs = tc.findall(f".//{_tag('w:p')}")
            cell_text = "".join(
                "".join(t.text or "" for t in p.findall(f".//{_tag('w:t')}"))
                for p in cell_paragraphs
            ).strip()
            cells_text.append(cell_text)

        cell_str = "\n".join(cells_text)
        # 去除所有空格（Word 分散对齐会在标签/内容中间插入空格）
        cell_str_ns = cell_str.replace(" ", "")

        # 尝试从表格单元格中匹配标签
        label_patterns = {
            "student_name": (r"姓名[:：]\s*(.+?)(?:\n|$)", None),
            "student_id":   (r"学号[:：]\s*(\d+)", None),
            "major":        (r"专业[:：]\s*(.+?)(?:\n|$)", None),
            "school":       (r"学院[:：]\s*(.+?)(?:\n|$)", None),
            "advisor":      (r"指导教师[:：]\s*(.+?)(?:\n|$)", None),
        }
        for field, (pattern, _) in label_patterns.items():
            m = re.search(pattern, cell_str_ns, re.IGNORECASE)
            if m and result[field] is None:
                result[field] = m.group(1).strip()

    # 策略2：从纯段落中提取（封面没有表格的情况）
    # 只取前20个段落（封面区域），避免正文中的 esttab LaTeX 代码干扰
    all_paragraphs = xml_root.findall(f".//{_tag('w:p')}")
    cover_paragraphs = all_paragraphs[:20]
    full_text = "\n".join(
        "".join(t.text or "" for t in p.findall(f".//{_tag('w:t')}"))
        for p in cover_paragraphs
    )
    # 同样去除所有空格后匹配
    full_text_ns = full_text.replace(" ", "")

    # 排除 esttab 输出的 LaTeX 代码片段（正文中的表格标签会干扰正则）
    LATEX_SKIP_PATTERNS = ["title(", "cells(", "scalar(", "nomtitlenumerator",
                            "meansd", "bVIF", "esttab", "using", "replace"]

    def _looks_like_garbage(text: str) -> bool:
        return any(p in text for p in LATEX_SKIP_PATTERNS)

    # 题目：包含"题目"字样的行（非 LaTeX）
    title_match = re.search(r"(?:题目|Title)[:：]?\s*\n?\s*(.+?)(?:\n|$)", full_text_ns, re.IGNORECASE)
    if title_match:
        candidate = title_match.group(1).strip()
        if not _looks_like_garbage(candidate):
            result["title"] = candidate
    # 备选：取封面区域第2~10行中的较长文字
    # （第1行通常是学校/学期名称，需跳过）
    # 过滤掉含"届"、"学年论文"等模板占位文字的行
    if result["title"] is None:
        lines = [l.strip() for l in full_text_ns.split("\n") if l.strip()]
        skip_patterns = ["届", "学年论文", "毕业设计", "毕业论文"]
        for line in lines[1:10]:
            if 10 < len(line) <= 50 and not any(kw in line for kw in skip_patterns):
                if not _looks_like_garbage(line):
                    result["title"] = line
                    break

    # 补充匹配（未被表格策略捕获的字段）
    label_patterns2 = {
        "student_name": (r"姓名[:：]\s*(.+?)(?:\n|$)", None),
        "student_id":   (r"学号[:：]\s*(\d+)", None),
        "major":        (r"专业[:：]\s*(.+?)(?:\n|$)", None),
        "school":       (r"学院[:：]\s*(.+?)(?:\n|$)", None),
        "advisor":      (r"指导教师[:：]\s*(.+?)(?:\n|$)", None),
    }
    for field, (pattern, _) in label_patterns2.items():
        if result[field] is None:
            m = re.search(pattern, full_text_ns, re.IGNORECASE)
            if m:
                result[field] = m.group(1).strip()

    return result


# ── 摘要与关键词提取 ─────────────────────────────────────────────────────────

def _extract_abstract(xml_root: ET.Element, marker_bounds: dict = None) -> dict:
    """
    提取：中文摘要、英文摘要、中文关键词、英文关键词
    标准论文格式：
    摘要 → 关键词： → Abstract → Keywords:

    marker_bounds: 可选，marker 边界检测结果（含 abstract_start/end 字符位置）。
                  若传入，则用字符位置精确截取；否则退化为状态机逻辑。
    """
    all_paragraphs = xml_root.findall(f".//{_tag('w:p')}")
    paragraphs = []
    for p in all_paragraphs:
        p_text = "".join(t.text or "" for t in p.findall(f".//{_tag('w:t')}")).strip()
        if p_text:
            paragraphs.append({"text": p_text, "element": p})

    # 分开变量：中文 / 英文
    abstract_text = ""          # 中文摘要
    english_abstract_text = ""  # 英文摘要
    keywords = []               # 中文关键词
    english_keywords = []       # 英文关键词

    # ── Marker 边界精确路径 ──────────────────────────────────────────────────
    if marker_bounds and marker_bounds.get("abstract_start", -1) >= 0:
        abs_start = marker_bounds["abstract_start"]
        full_text, para_starts = _build_text_index(xml_root)

        # 使用全量段落的索引体系（与 _char_pos_to_para_idx 对齐）
        start_idx = _char_pos_to_para_idx(abs_start, para_starts)
        # 保守取较大范围，状态机自行终止
        end_idx = min(start_idx + 80, len(all_paragraphs) - 1)

        # 状态机：DEFAULT → CN_KEYWORDS → EN_ABSTRACT → EN_KEYWORDS → DONE
        state = "DEFAULT"
        # 英文摘要终止关键词
        EN_ABSTRACT_STOP = ["目录", "目次", "引言", "绪论", "Table of Contents",
                            "一、", "二、", "三、", "四、", "五、", "六、", "七、"]

        for i in range(start_idx, min(end_idx + 1, len(all_paragraphs))):
            text = "".join(t.text or "" for t in all_paragraphs[i].findall(f".//{_tag('w:t')}")).strip()
            if not text:
                continue

            # 跳过摘要标题本身
            if re.match(r"^摘\s?要\s*$", text, re.IGNORECASE):
                continue

            # 提取中文关键词
            cn_kw_match = re.match(r"^关键词[:：]\s*(.+)", text)
            if cn_kw_match:
                kw_text = cn_kw_match.group(1).strip()
                keywords = [k.strip() for k in re.split(r"[,，;；\s]+", kw_text) if k.strip()]
                state = "CN_KEYWORDS"
                continue

            # 提取英文关键词（进入 DONE 状态）
            en_kw_match = re.match(r"^(Key words?|Keywords|Key Words)[:：]\s*(.+)", text, re.IGNORECASE)
            if en_kw_match:
                kw_text = en_kw_match.group(2).strip()
                # 英文关键词用分号或逗号分隔，词组内部含空格，不能按\s拆
                english_keywords = [k.strip() for k in re.split(r"[,;]+", kw_text) if k.strip()]
                state = "DONE"
                break

            if state == "DEFAULT":
                # 中文摘要正文
                if abstract_text:
                    abstract_text += "\n"
                abstract_text += text

            elif state == "CN_KEYWORDS":
                # 英文标题行（很短，< 50 字符，全 ASCII）→ 跳过
                if len(text) < 50 and re.match(r"^[A-Za-z\s,\-\'\"]+$", text):
                    continue
                # 作者姓名（拉丁字母格式）→ 跳过
                if re.match(r"^[A-Z][a-z]+ [A-Z][a-z]+", text):
                    continue
                # 机构名（括号包裹）→ 跳过
                if re.match(r"^[（\(][^）\)]+[）\)]$", text):
                    continue
                # Abstract 本身（章节标题）→ 跳过
                if re.match(r"^Abstract$", text, re.IGNORECASE):
                    continue
                # 目录/正文标题 → 终止英文摘要
                ns = re.sub(r'\s+', '', text)
                stopped = False
                for stop_kw in EN_ABSTRACT_STOP:
                    if ns.startswith(stop_kw):
                        state = "DONE"
                        stopped = True
                        break
                if not stopped and re.search(r"^[一二三四五六七八九十0-9]+[、.．。]?\S*\d{1,3}$", ns):
                    state = "DONE"
                    stopped = True
                if stopped:
                    break
                # 正式进入英文摘要
                state = "EN_ABSTRACT"
                english_abstract_text += text + "\n"

            elif state == "EN_ABSTRACT":
                # 英文摘要终止检测
                ns = re.sub(r'\s+', '', text)
                stopped = False
                for stop_kw in EN_ABSTRACT_STOP:
                    if ns.startswith(stop_kw):
                        stopped = True
                        break
                if not stopped and re.search(r"^[一二三四五六七八九十0-9]+[、.．。]?\S*\d{1,3}$", ns):
                    stopped = True
                if stopped:
                    break
                # 英文摘要正文
                english_abstract_text += text + "\n"

        # 处理英文摘要：去掉末尾空行
        english_abstract_text = english_abstract_text.strip()

        # 摘要字数：与正文统一逻辑（中文字 + 英文词 + 数字组），不用 len()
        _cn = len(re.findall(r'[\u4e00-\u9fff]', abstract_text))
        _en = len(re.findall(r'[a-zA-Z]{3,}', abstract_text))
        _nm = len(re.findall(r'\d+\.?\d*', abstract_text))

        return {
            "abstract": abstract_text.strip(),
            "english_abstract": english_abstract_text.strip(),
            "keywords": keywords,
            "english_keywords": english_keywords,
            "abstract_length": _cn + _en + _nm,
            "keyword_count": len(keywords),
            "english_keyword_count": len(english_keywords),
        }

    # ── Fallback：状态机路径（原逻辑） ──────────────────────────────────────
    in_abstract = False
    in_english_abstract = False

    # 英文摘要终止关键词（遇到这些说明已离开英文摘要区）
    EN_ABSTRACT_STOP = ["目录", "目次", "引言", "绪论", "Table of Contents",
                        "一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、"]

    for para in paragraphs:
        text = para["text"]

        if re.match(r"^摘要$", text, re.IGNORECASE):
            in_abstract = True
            in_english_abstract = False
            continue

        if re.match(r"^(Abstract|ABSTRACT|英文摘要)$", text):
            in_english_abstract = True
            in_abstract = False
            continue

        cn_kw_match = re.match(r"^关键词[:：]\s*(.+)", text)
        if cn_kw_match:
            kw_text = cn_kw_match.group(1).strip()
            keywords = [k.strip() for k in re.split(r"[,，;；\s]+", kw_text) if k.strip()]
            in_abstract = False
            continue

        en_kw_match = re.match(r"^(Key words?|Keywords|Key Words)[:：]\s*(.+)", text, re.IGNORECASE)
        if en_kw_match:
            kw_text = en_kw_match.group(2).strip()
            # 英文关键词用分号或逗号分隔，词组内部含空格，不能按\s拆
            english_keywords = [k.strip() for k in re.split(r"[,;]+", kw_text) if k.strip()]
            in_english_abstract = False
            break

        # 英文摘要终止检测（防止目录/正文被误识别为英文摘要）
        if in_english_abstract:
            ns = re.sub(r'\s+', '', text)
            stopped = False
            for stop_kw in EN_ABSTRACT_STOP:
                if ns.startswith(stop_kw):
                    in_english_abstract = False
                    stopped = True
                    break
            # 也检测目录条目格式（如 "摘要2"、"一、引言4"）
            if not stopped and re.search(r"^[一二三四五六七八九十0-9a-zA-Z]+[、.．。]?\S*\d{1,3}$", ns):
                in_english_abstract = False
                stopped = True
            if stopped:
                continue

        if in_abstract:
            abstract_text += text + "\n"
        if in_english_abstract:
            english_abstract_text += text + "\n"

    # 摘要字数：与正文统一逻辑（中文字 + 英文词 + 数字组），不用 len()
    _cn = len(re.findall(r'[\u4e00-\u9fff]', abstract_text))
    _en = len(re.findall(r'[a-zA-Z]{3,}', abstract_text))
    _nm = len(re.findall(r'\d+\.?\d*', abstract_text))

    return {
        "abstract": abstract_text.strip(),
        "english_abstract": english_abstract_text.strip(),
        "keywords": keywords,
        "english_keywords": english_keywords,
        "abstract_length": _cn + _en + _nm,
        "keyword_count": len(keywords),
        "english_keyword_count": len(english_keywords),
    }

# ── 正文与章节结构提取 ───────────────────────────────────────────────────────

def _heading_level(p_element: ET.Element) -> Optional[int]:
    """检测段落是否为标题，并返回标题级别（1~4）

    标题层级编码规范（经管学院统一要求）：
        一级：一、二、三、……
        二级：（一）（二）（三）……
        三级：1. 2. 3. ……
        四级：(1) (2) (3) ……
    """
    pStyle = p_element.find(_tag("w:pPr"))
    if pStyle is None:
        return None
    style_val = pStyle.find(_tag("w:pStyle"))
    if style_val is None:
        return None
    style_name = style_val.get(_tag("w:val")) or ""
    # 标准样式名：Heading/标题/Title
    m = re.search(r"Heading|标题|一级|二级|三级|四级|Title", style_name, re.IGNORECASE)
    if m:
        # 尝试从样式名提取级别数字
        num_m = re.search(r"\d+", style_name)
        if num_m:
            lvl = int(num_m.group())
            if 1 <= lvl <= 6:
                return lvl
        return 1

    # 自定义样式编号回退：通过段落文本内容辅助判断
    # 很多论文模板用自定义样式名（纯数字编号如24、25）代替标准 Heading
    # 通过文本模式判断是否为章节标题
    text = "".join(t.text or "" for t in p_element.findall(f".//{_tag('w:t')}")).strip()
    if not text or len(text) > 80:
        return None

    # 一级标题：一、xxx / 第一章 / 引言 / 绪论 / 结论
    # 排除目录条目（末尾带页码，如 "一、 引言4"）
    if re.match(r"^[一二三四五六七八九十]+[、．.]\s*\S+", text) and not re.search(r"\d{1,3}$", text):
        return 1
    if re.match(r"^(引言|绪论|结论|结语|文献综述|理论分析|研究设计|实证分析|结论与建议)\s*$", text):
        return 1
    # 二级标题：（一）xxx / (一)xxx
    if re.match(r"^[（\(][一二三四五六七八九十]+[）\)]\s*\S+", text) and not re.search(r"\d{1,3}$", text):
        return 2
    # 三级标题：1. xxx / 1.1 xxx
    if re.match(r"^\d+\.\d+\s+\S+", text):
        return 3
    if re.match(r"^\d+[、．.]\s*\S+", text):
        return 3
    # 四级标题：(1) xxx / (2) xxx —— 半角小括号+阿拉伯数字
    if re.match(r"^\(\d+\)\s*\S+", text):
        return 4

    return None


def _extract_body(xml_root: ET.Element, marker_bounds: dict = None) -> dict:
    """
    提取正文段落（排除封面、摘要、目录、参考文献、致谢）。
    纯状态机版本，忽略 marker_bounds 参数。
    """
    all_paragraphs = xml_root.findall(f".//{_tag('w:p')}")

    sections = []
    current_section = {"level": 0, "title": "", "body": ""}
    body_paragraphs = []

    in_cn_zone = False
    in_en_zone = False
    in_toc = False
    body_started = False

    TOC_KEYWORDS = ["目录", "目次", "Table of Contents", "Contents"]
    STOP_KEYWORDS = ["参考文献", "References", "致谢", "谢辞", "Acknowledgment", "注释", "附录"]

    # 正文首章标题的回退检测模式（当目录退出条件不满足时使用）
    # 注意：需排除目录条目（末尾带页码，如 "一、 引言4"）
    BODY_START_PATTERNS = [
        r"^[一二三四五六七八九十]+[、．.]\s*\S+(?<!\d)$",    # 一、引言 (排除 "一、引言4")
        r"^第[一二三四五六七八九十]+[章节]\s*\S+",            # 第一章、第二节
        r"^引言\s*$",                                         # 独立的"引言"
        r"^绪论\s*$",                                         # 独立的"绪论"
    ]

    # 自定义标题样式编号映射（常见论文模板样式名 → 标题级别）
    # 当标准 Heading 检测失败时，用样式名数字映射作为回退
    CUSTOM_HEADING_STYLES = {
        "24": 1, "25": 2, "26": 3,  # 本科论文常见自定义样式
        "1": 1, "2": 2, "3": 3, "4": 4,  # 部分模板
        "10": 1, "20": 2, "30": 3, "40": 4,  # 部分模板
    }

    # 基于文本内容的章节标题检测（回退1）
    # 标题层级编码规范：一级"一、"、二级"（一）"、三级"1."、四级"(1)"
    TEXT_HEADING_PATTERNS = [
        (1, r"^[一二三四五六七八九十]+[、．.]\s*\S+(?<!\d)$"),   # 一、引言 (排除目录条目)
        (2, r"^（[一二三四五六七八九十]+）\s*\S+(?<!\d)$"),      # （一）研究背景
        (2, r"^\([一二三四五六七八九十]+\)\s*\S+(?<!\d)$"),      # (一)研究背景
        (3, r"^\d+[、．.]\s*\S+(?<!\d)$"),                       # 1. 理论基础
        (3, r"^\d+\.\d+\s+\S+"),                                  # 1.1 理论基础
        (4, r"^\(\d+\)\s*\S+"),                                   # (1) 样本选择
    ]

    def _detect_heading_from_text(text):
        """基于文本内容检测章节级别，返回 level 或 None"""
        for level, pattern in TEXT_HEADING_PATTERNS:
            if re.match(pattern, text):
                return level
        return None

    def _get_heading_level_robust(p_element, text):
        """鲁棒的标题级别检测：标准检测 + 文本模式 + 自定义样式回退"""
        lvl = _heading_level(p_element)
        if lvl is not None:
            return lvl
        # 回退1：文本模式检测
        lvl = _detect_heading_from_text(text)
        if lvl is not None:
            return lvl
        # 回退2：自定义样式编号 + 文本长度辅助判断
        if text and len(text) < 50:
            pPr = p_element.find(_tag("w:pPr"))
            if pPr is not None:
                pStyle = pPr.find(_tag("w:pStyle"))
                if pStyle is not None:
                    style_name = pStyle.get(_tag("w:val")) or ""
                    if style_name in CUSTOM_HEADING_STYLES:
                        return CUSTOM_HEADING_STYLES[style_name]
        return None

    def _detect_heading_from_text(text):
        """基于文本内容检测章节级别，返回 level 或 None"""
        for level, pattern in TEXT_HEADING_PATTERNS:
            if re.match(pattern, text):
                return level
        return None

    for p in all_paragraphs:
        text = "".join(t.text or "" for t in p.findall(f".//{_tag('w:t')}")).strip()
        if not text:
            continue

        lvl = _heading_level(p)

        if re.match(r"^摘要$", text, re.IGNORECASE):
            in_cn_zone = True
            in_en_zone = False
            continue
        if re.match(r"^关键词[:：]", text, re.IGNORECASE):
            in_cn_zone = False
            continue
        if re.match(r"^Abstract$", text, re.IGNORECASE):
            in_en_zone = True
            in_cn_zone = False
            continue
        if re.match(r"^(Key words?|Keywords)[:：]", text, re.IGNORECASE):
            in_en_zone = False
            body_started = True
            in_toc = True
            continue

        if in_cn_zone or in_en_zone:
            continue

        if not body_started:
            ns = re.sub(r'\s+', '', text)
            for kw in TOC_KEYWORDS:
                if re.match(rf"^{kw}[\s：:]*$", ns, re.IGNORECASE):
                    in_toc = True
                    break
            if in_toc:
                # 目录条目：点号引导+页码
                if re.search(r"\.{2,}\s*\d+$", ns):
                    continue
                # 目录条目：标题+页码（无点号），如 "一、引言4"
                if re.search(r"^[一二三四五六七八九十0-9a-zA-ZⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+[、.．。]?\S*\d+$", ns):
                    continue
                # heading 样式的非目录行 → 退出目录
                if lvl is not None and not re.search(r"\d+(\.\d+)*\s*.{3,}\d+", ns):
                    in_toc = False
                    body_started = True
                    # 不 continue，当前行是正文首行，进入下面的正文处理
                else:
                    # 回退检测：匹配正文首章标题模式
                    for bp in BODY_START_PATTERNS:
                        if re.match(bp, text):
                            in_toc = False
                            body_started = True
                            break
                    if in_toc:
                        continue

        if in_toc:
            ns = re.sub(r'\s+', '', text)
            if re.search(r"^[一二三四五六七八九十0-9a-zA-ZⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+[、.．。]?\S*\d+$", ns):
                continue
            # 回退检测：匹配正文首章标题
            for bp in BODY_START_PATTERNS:
                if re.match(bp, text):
                    in_toc = False
                    break
            if in_toc:
                continue

        for kw in STOP_KEYWORDS:
            if re.match(rf"^{kw}[\s：:]*$", text, re.IGNORECASE):
                in_toc = False
                body_started = False
                break
        if not body_started:
            continue

        # 检测章节标题：优先用 _heading_level，回退到文本模式+自定义样式检测
        detected_lvl = _get_heading_level_robust(p, text)

        if detected_lvl is not None and len(text) < 100:
            if current_section["title"] or current_section["body"]:
                sections.append(current_section)
            current_section = {"level": detected_lvl, "title": text, "body": ""}
        else:
            current_section["body"] += text + "\n"
            body_paragraphs.append(text)

    if current_section["title"] or current_section["body"]:
        sections.append(current_section)

    full_body = "\n".join(body_paragraphs)
    return {
        "full_text": full_body,
        "char_count": len(full_body),
        "chinese_char_count": len(re.findall(r"[\u4e00-\u9fff]", full_body)),
        "english_word_count": len(re.findall(r"[a-zA-Z]{3,}", full_body)),
        "digit_count": len(re.findall(r"\d+\.?\d*", full_body)),
        "sections": sections,
        "section_count": len(sections),
    }


# ── 表格分析 ─────────────────────────────────────────────────────────────────

def _count_tables_deep(xml_root: ET.Element) -> dict:
    """
    深入分析所有表格：
    - 原生表格：有 <w:tbl> 但单元格内不含图片
    - 截图表格：<w:tbl> 单元格内嵌入了 <w:drawing>（Stata 截图）
    """
    all_tables = xml_root.findall(f".//{_tag('w:tbl')}")

    native_table_count = 0      # 纯原生 Word 表格
    screenshot_table_count = 0  # 含截图的表格

    for tbl in all_tables:
        # 检查是否有图片嵌入（<w:drawing>）在任意单元格中
        drawings = tbl.findall(f".//{_tag('w:drawing')}")
        inline_shapes = tbl.findall(f".//{_tag('wp:inline')}") + tbl.findall(f".//{_tag('wp:anchor')}")
        has_screenshot = len(drawings) > 0 or len(inline_shapes) > 0
        if has_screenshot:
            screenshot_table_count += 1
        else:
            native_table_count += 1

    # 额外：检查游离图片（不在任何表格内的 <w:drawing>）
    all_drawings = xml_root.findall(f".//{_tag('w:drawing')}")
    table_drawings = set()
    for tbl in all_tables:
        for d in tbl.findall(f".//{_tag('w:drawing')}"):
            table_drawings.add(d)
    standalone_images = len(all_drawings) - len(table_drawings)

    return {
        "total_tables": len(all_tables),
        "native_table_count": native_table_count,
        "screenshot_table_count": screenshot_table_count,
        "standalone_image_count": standalone_images,
    }


# ── 图片统计 ─────────────────────────────────────────────────────────────────

def _count_images(docx_path: str) -> dict:
    """
    统计 word/media/ 目录下的图片文件数量。
    注意：Stata 截图和正常图片都以图片形式存在，无法单从文件类型区分。
    """
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            media_files = [f for f in z.namelist() if f.startswith("word/media/")]
            media_count = len(media_files)
    except Exception:
        media_count = 0

    return {"media_image_count": media_count}


# ── 参考文献提取 ─────────────────────────────────────────────────────────────

def _extract_references(xml_root: ET.Element, marker_bounds: dict = None) -> dict:
    """
    提取参考文献列表。
    marker_bounds: 可选，marker 边界检测结果（含 ref_start/end 字符位置）。
                   若传入，则用字符位置精确截取；否则退化为正则匹配逻辑。
    """
    all_paragraphs = xml_root.findall(f".//{_tag('w:p')}")
    paragraphs = []
    for p in all_paragraphs:
        text = "".join(t.text or "" for t in p.findall(f".//{_tag('w:t')}")).strip()
        if text:
            paragraphs.append(text)

    # ── Marker 边界精确路径 ──────────────────────────────────────────────────
    if marker_bounds and marker_bounds.get("ref_start", -1) >= 0:
        ref_start = marker_bounds["ref_start"]
        ref_end   = marker_bounds["ref_end"]
        full_text, para_starts = _build_text_index(xml_root)

        start_idx = _char_pos_to_para_idx(ref_start, para_starts)
        end_idx   = _char_pos_to_para_idx(ref_end,   para_starts)

        # para_starts 基于 all_paragraphs（含空段落），需遍历全量段落
        ref_entries = []
        for i in range(start_idx, min(end_idx + 1, len(all_paragraphs))):
            text = "".join(t.text or "" for t in all_paragraphs[i].findall(f".//{_tag('w:t')}")).strip()
            if not text:
                continue
            # 跳过参考文献标题本身
            if re.match(r"^(参考文献|References)[：:。]*(?<![0-9])$", text, re.IGNORECASE):
                continue
            # 遇到致谢/附录/后记，停止收集
            if re.match(r"^(致谢|谢辞|附录|后记|Acknowledgment)", text):
                break
            ref_entries.append(text)
    else:
        # ── Fallback：正则匹配路径（原逻辑） ──────────────────────────────
        ref_entries = []
        in_ref = False

        for i, text in enumerate(paragraphs):
            if re.match(r"^(参考文献|References)[：:。]*(?<![0-9])$", text, re.IGNORECASE):
                in_ref = True
                continue

            if in_ref:
                if re.match(r"^(致谢|谢辞|附录|后记|Acknowledgment)", text):
                    break
                ref_entries.append(text)

    ref_text = "\n".join(ref_entries)

    # 统计
    ref_count = len(ref_entries)
    has_superscript_cite = bool(re.search(r"\[\d+\]", ref_text))

    foreign_count = 0
    journals_count = 0
    books_count = 0
    theses_count = 0
    for entry in ref_entries:
        entry_clean = entry.strip()
        if not entry_clean:
            continue
        # 去掉编号前缀（如 [1]、[6] 等），以便正确识别外文文献
        entry_no_num = re.sub(r"^\[\d+\]\s*", "", entry_clean)
        if re.match(r"^[a-zA-Z]", entry_no_num):
            foreign_count += 1
        if re.search(r"\[J\]|学报|期刊", entry_clean):
            journals_count += 1
        if re.search(r"\[M\]|出版社", entry_clean):
            books_count += 1
        if re.search(r"\[D\]|博士论文|硕士论文|博士學位論文|碩士學位論文", entry_clean):
            theses_count += 1

    current_year = datetime.now().year
    recent_threshold = current_year - 5
    recent_count = 0
    for entry in ref_entries:
        years = re.findall(r"\b(20[12]\d)\b", entry)
        if years:
            last_year = int(years[-1])
            if last_year >= recent_threshold:
                recent_count += 1

    return {
        "reference_text": ref_text,
        "reference_count": ref_count,
        "has_superscript_cite": has_superscript_cite,
        "books_count": books_count,
        "theses_count": theses_count,
        "journals_count": journals_count,
        "foreign_count": foreign_count,
        "recent_references_count": recent_count,
    }


# ── 致谢提取 ─────────────────────────────────────────────────────────────────

def _extract_acknowledgment(xml_root: ET.Element) -> dict:
    """提取致谢内容"""
    all_paragraphs = xml_root.findall(f".//{_tag('w:p')}")
    paragraphs = []
    for p in all_paragraphs:
        text = "".join(t.text or "" for t in p.findall(f".//{_tag('w:t')}")).strip()
        if text:
            paragraphs.append(text)

    ack_entries = []
    in_ack = False

    ACK_KEYWORDS = ["致谢", "谢辞", "Acknowledgment", "致 谢"]

    for text in paragraphs:
        for kw in ACK_KEYWORDS:
            if re.match(rf"^{kw}[\s：:]*$", text, re.IGNORECASE):
                in_ack = True
                continue
        if in_ack:
            ack_entries.append(text)

    return {
        "acknowledgment_text": "\n".join(ack_entries),
        "has_acknowledgment": in_ack,
    }


# ── 主函数 ────────────────────────────────────────────────────────────────────

def analyze_xml(docx_path: str) -> dict:
    """
    入口函数：对给定 docx 文件执行完整的 XML 结构分析。
    返回一个包含所有提取结果的字典。

    边界检测：优先使用 marker 风格正则（marker_boundaries）精确截取；
    若检测失败则退化为各模块内置的状态机/正则逻辑。
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"文件不存在: {docx_path}")

    # 读取 docx（zip）中的 document.xml
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            xml_content = z.read("word/document.xml")
    except Exception as e:
        raise ValueError(f"无法读取 docx 内容: {e}")

    # 解析 XML
    root = ET.fromstring(xml_content)

    # ── 统一边界检测（marker 风格，一次性检测三个边界）──────────────────────
    # 建立全文本索引，供 boundary 函数和下游提取函数共用
    full_text, _ = _build_text_index(root)
    marker_bounds = _marker_boundaries(full_text)

    # ── 各模块提取（均传入 marker_bounds 以启用精确路径）─────────────────────
    cover_info    = _extract_cover_info(root)
    abstract_info = _extract_abstract(root, marker_bounds)
    body_info     = _extract_body(root, marker_bounds)
    table_info    = _count_tables_deep(root)
    image_info    = _count_images(docx_path)
    ref_info      = _extract_references(root, marker_bounds)
    ack_info      = _extract_acknowledgment(root)

    # 合并结果
    result = {
        # 封面信息
        "title":        cover_info["title"],
        "student_name": cover_info["student_name"],
        "student_id":   cover_info["student_id"],
        "school":       cover_info["school"],
        "major":        cover_info["major"],
        "advisor":      cover_info["advisor"],

        # 摘要
        "abstract":           abstract_info["abstract"],
        "abstract_length":    abstract_info["abstract_length"],
        "keywords":           abstract_info["keywords"],
        "keyword_count":      abstract_info["keyword_count"],
        "english_abstract":   abstract_info["english_abstract"],
        "english_keywords":   abstract_info["english_keywords"],
        "english_keyword_count": abstract_info.get("english_keyword_count", 0),

        # 正文
        "body_text":           body_info["full_text"],
        "char_count":          body_info["char_count"],
        "chinese_char_count":  body_info["chinese_char_count"],
        "english_word_count":  body_info["english_word_count"],
        "digit_count":         body_info["digit_count"],
        "sections":           body_info["sections"],
        "section_count":       body_info["section_count"],

        # 表格与图片
        "total_tables":           table_info["total_tables"],
        "native_table_count":     table_info["native_table_count"],
        "screenshot_table_count": table_info["screenshot_table_count"],
        "standalone_image_count": table_info["standalone_image_count"],
        "media_image_count":      image_info["media_image_count"],

        # 参考文献
        "reference_text":         ref_info["reference_text"],
        "reference_count":        ref_info["reference_count"],
        "has_superscript_cite":   ref_info["has_superscript_cite"],
        "books_count":            ref_info["books_count"],
        "theses_count":           ref_info["theses_count"],
        "journals_count":         ref_info["journals_count"],
        "foreign_count":         ref_info["foreign_count"],
        "recent_references_count":ref_info["recent_references_count"],

        # 致谢
        "acknowledgment_text": ack_info["acknowledgment_text"],
        "has_acknowledgment":  ack_info["has_acknowledgment"],

        # 原始 docx 路径（便于追溯）
        "source_docx": docx_path,
    }

    return result


# ── CLI 调试入口 ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import json, sys, io

    # 强制 UTF-8 输出（Windows PowerShell 默认 GBK 会截断非ASCII字符）
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    if len(sys.argv) < 2:
        print("用法: python xml_analyzer.py <path_to_docx>")
        sys.exit(1)

    docx_path = sys.argv[1]
    result = analyze_xml(docx_path)

    # 输出 JSON（CLI 调试用）
    print(json.dumps(result, ensure_ascii=False, indent=2))
