#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标记识别与内容提取模块:

识别AI在Markdown中添加的标记（HTML注释格式），提取各部分内容
提供自动标记功能，自动检测论文结构并添加标记

[DEPRECATED] 此模块已废弃，仅用于 .md 旧路径的向后兼容。
新流程（.docx 直接分析）请使用 xml_analyzer 模块。
"""

import re
from typing import Dict, Optional, Tuple


class Marker:
    """标记识别器"""

    # 标记模式定义
    MARKERS = {
        'abstract': {
            'start': r'<!--\s*摘要开始\s*-->',
            'end': r'<!--\s*摘要结束\s*-->',
        },
        'body': {
            'start': r'<!--\s*正文开始\s*-->',
            'end': r'<!--\s*正文结束\s*-->',
        },
        'references': {
            'start': r'<!--\s*参考文献开始\s*-->',
            'end': r'<!--\s*参考文献结束\s*-->',
        },
    }

    def __init__(self, md_path: str):
        """
        初始化标记器

        Args:
            md_path: Markdown文件路径
        """
        self.md_path = md_path
        self.content = self._load_content()
        self.marks = self._find_all_marks()

    def _load_content(self) -> str:
        """加载Markdown文件内容"""
        try:
            with open(self.md_path, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            return ""

    def _find_all_marks(self) -> Dict[str, Dict]:
        """查找所有标记及其位置"""
        marks = {}

        for name, pattern in self.MARKERS.items():
            start_match = re.search(pattern['start'], self.content)
            end_match = re.search(pattern['end'], self.content)

            marks[name] = {
                'start': start_match,
                'end': end_match,
                'start_pos': start_match.start() if start_match else None,
                'end_pos': end_match.end() if end_match else None,
            }

        return marks

    def has_mark(self, mark_name: str) -> bool:
        """
        检查是否存在指定标记

        Args:
            mark_name: 标记名（abstract/body/references）

        Returns:
            是否存在
        """
        if mark_name not in self.marks:
            return False
        mark = self.marks[mark_name]
        return mark['start'] is not None and mark['end'] is not None

    def is_fully_marked(self) -> bool:
        """检查是否所有必需标记都存在"""
        return all(self.has_mark(name) for name in ['abstract', 'body', 'references'])

    def extract(self, mark_name: str) -> Optional[str]:
        """
        提取指定标记区间的内容

        Args:
            mark_name: 标记名（abstract/body/references）

        Returns:
            标记区间内的内容，如果标记不存在则返回None
        """
        if mark_name not in self.marks:
            return None

        mark = self.marks[mark_name]
        if mark['start_pos'] is None or mark['end_pos'] is None:
            return None

        # 内容在结束标记的开始位置之后
        content = self.content[mark['start_pos']:mark['end_pos']]

        # 移除开始和结束标记
        start_marker = self.MARKERS[mark_name]['start']
        end_marker = self.MARKERS[mark_name]['end']

        content = re.sub(start_marker, '', content)
        content = re.sub(end_marker, '', content)

        return content.strip()

    def extract_all(self) -> Dict[str, Optional[str]]:
        """
        提取所有标记区间的内容

        Returns:
            包含各部分内容的字典
        """
        return {
            'abstract': self.extract('abstract'),
            'body': self.extract('body'),
            'references': self.extract('references'),
            'full_content': self.content,
        }

    def get_positions(self) -> Dict[str, Optional[Dict]]:
        """
        获取各标记的位置信息

        Returns:
            包含位置信息的字典
        """
        positions = {}
        for name, mark in self.marks.items():
            if mark['start'] and mark['end']:
                positions[name] = {
                    'start': mark['start_pos'],
                    'end': mark['end_pos'],
                    'start_line': self.content[:mark['start_pos']].count('\n') + 1,
                    'end_line': self.content[:mark['end_pos']].count('\n') + 1,
                }
            else:
                positions[name] = None
        return positions

    @staticmethod
    def generate_markers(abstract: str = None, body: str = None, references: str = None) -> str:
        """
        生成带有标记的内容（用于测试或模板）

        Args:
            abstract: 摘要内容
            body: 正文内容
            references: 参考文献内容

        Returns:
            带标记的Markdown文本
        """
        parts = []

        if abstract:
            abstract_content = "<!-- 摘要开始 -->\n" + str(abstract) + "\n<!-- 摘要结束 -->"
            parts.append(abstract_content)

        if body:
            parts.append(f"<!-- 正文开始 -->\n{body}\n<!-- 正文结束 -->")

        if references:
            parts.append(f"<!-- 参考文献开始 -->\n{references}\n<!-- 参考文献结束 -->")

        return '\n\n'.join(parts)

    def find_abstract_boundary(self) -> Tuple[int, int]:
        """
        自动查找摘要的边界位置

        支持多种格式：
        - # 摘要 (标准Markdown)
        - **摘要** (加粗格式)
        - 摘　要 (全角空格)
        - 关键词格式：**关键词：**xxx 或 关键词：xxx

        Returns:
            (start_pos, end_pos) 元组
            - start_pos: 摘要开始位置
            - end_pos: 摘要结束位置（中文关键词段落的结束，不包含英文摘要）
        """
        # 先移除旧的标记，避免干扰
        content = self._remove_markers()

        # 摘要开始：支持多种格式
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

        # 摘要结束：找到关键词行结束
        # 关键词格式：
        # - **关键词：**xxx
        # - **关键词：** xxx（带空格）
        # - Keywords: xxx (英文)
        # 结束位置应该在英文标题 **The Impact of... 之前

        abstract_end = None

        # 找中文关键词所在行
        # 支持多种格式
        kw_patterns = [
            r'\*\*关键词[：:]\*\*[^\n]+',  # **关键词：**xxx
            r'\*\*关键词[：:]\*\*\s*[^\n]+',  # **关键词：** xxx
            r'^关键词[：:][^\n]+',  # 关键词：xxx
        ]

        kw_match = None
        for pattern in kw_patterns:
            kw_match = re.search(pattern, content[abstract_start:], re.MULTILINE)
            if kw_match:
                kw_match = re.search(pattern, content, re.MULTILINE)
                break

        if kw_match:
            kw_end = kw_match.end()
            # 找到关键词所在行的结束（最后一个字符后的位置）
            line_end = kw_end  # 这是关键词最后一个字符之后的位置

            # 摘要结束位置：关键词行结束后，跳过1-2个换行，找到英文标题/作者/Abstract之前
            # 论文格式通常是：
            # **关键词：**xxx
            # \n
            # **The Impact of...** (英文标题)
            # 或
            # **关键词：**xxx
            # \n
            # Zhang Shengnan (作者)

            # 跳过换行和空白
            next_pos = line_end
            while next_pos < len(content) and content[next_pos] in '\n \t':
                next_pos += 1

            # 如果跳过后正好是英文标题或作者名或 Abstract，说明找到了
            # 否则继续往后找空行后的内容

            # 尝试匹配英文标题
            eng_title_match = re.match(r'\*\*[A-Z]', content[next_pos:])
            author_match = re.match(r'[A-Z][a-z]+ [A-Z][a-z]+', content[next_pos:])
            abstract_match = re.match(r'\*\*Abstract', content[next_pos:], re.IGNORECASE)

            if eng_title_match or abstract_match:
                # 英文标题或 Abstract 在 next_pos，摘要结束于 next_pos（空行之后，标题之前）
                abstract_end = next_pos
            elif author_match:
                # 作者名在 next_pos，摘要结束于 next_pos
                abstract_end = next_pos
            else:
                # 没找到明确的边界，找空行
                # 关键词行后可能是单换行（+空行）或双换行
                blank_match = re.search(r'\n\n+', content[line_end:])
                if blank_match:
                    # abstract_end 是第二个换行后（下一个内容开始之前）
                    # 但我们需要的是第一个换行后（空行开始）
                    abstract_end = line_end + blank_match.start() + 1
                else:
                    # 只有单换行
                    abstract_end = next_pos

        if abstract_end is None:
            abstract_end = abstract_start + 1000

        return (abstract_start, abstract_end)

    def find_body_boundary(self) -> Tuple[int, int]:
        r"""
        自动查找正文的边界位置

        支持多种目录格式：
        - [一、绪论 1](\l) - 无页码链接
        - [参考文献 [16](#_Toc11061)] - 带页码和锚点链接
        - [摘要 I](#page) - 带罗马数字页码

        Returns:
            (start_pos, end_pos) 元组
            - start_pos: 正文开始位置（目录结束后的下一个章节）
            - end_pos: 正文结束位置（参考文献开始前）
        """
        # 先移除旧的标记，避免干扰
        content = self._remove_markers()

        body_start = None
        body_end = None

        # === 找正文结束位置（参考文献之前）===
        # 参考文献标题格式（多种）
        ref_patterns = [
            r'^#{1,3}\s+参考文献',                    # # 参考文献 或 ## 参考文献
            r'\*\*参考文献\*\*',                       # **参考文献**
            r'\[]{#_Toc\d+\s*\.[a-z]+\}参考文献',   # Pandoc锚点格式
            r'参考文献\s*\[',                          # 参考文献 [ 格式
        ]

        # 先找参考文献开始位置
        ref_start_in_body = None
        for pattern in ref_patterns:
            for match in re.finditer(pattern, content, re.MULTILINE):
                pos = match.start()
                # 确保这不是目录中的参考文献
                # 检查前面是否有 [ 和数字（目录格式 [参考文献 14]）
                if pos > 5:
                    prev_context = content[max(0, pos-20):pos]
                    # 如果前面是 [参考文献 14] 格式，跳过
                    if re.search(r'\[参考文献\s*\[\d+\]', prev_context):
                        continue
                    # 如果是目录行的某一项（短行），跳过
                    if '\n' in prev_context and len(prev_context.split('\n')[-1]) < 30:
                        continue
                ref_start_in_body = pos
                break
            if ref_start_in_body is not None:
                break

        if ref_start_in_body is not None:
            body_end = ref_start_in_body
        else:
            body_end = len(content)

        # === 找正文开始位置（目录结束后）===

        # 方法1：找目录结束的标志
        # 目录通常以 [参考文献 XXX] 或 [Abstract] 等结尾
        # 之后有空行，然后是正文

        # 找所有目录项
        toc_items = []
        for match in re.finditer(r'\[([^\]]+)\]\(\\[lst]\)', content):  # \l 格式
            toc_items.append(match)

        # 也找标准链接格式
        for match in re.finditer(r'\[([^\]]+)\]\(#([^)]+)\)\)', content):
            toc_items.append(match)

        if toc_items:
            # 最后一个目录项之后
            last_toc = toc_items[-1]
            last_toc_end = last_toc.end()

            # 找空行（目录结束标志）
            body_start = content.find('\n\n', last_toc_end)
            if body_start != -1:
                body_start += 2  # 跳过空行
            else:
                body_start = last_toc_end

        # 方法2：备用 - 找第一个章节标题
        if body_start is None or body_start < 0:
            # 章节标题格式：一、二、三 或 第一、第二
            chapter_patterns = [
                r'#{1,3}\s+[一二三四五六七八九十百]+[、，]',  # 一、研究背景
                r'#{1,3}\s+第[一二三四五六七八九十百]+[章节节]',  # 第一章
                r'#{1,3}\s+\([一二三四五六七]\)',  # (一)
                r'\n\n[一二三四五六七八九十]+、',  # 换行后的一、
            ]
            for pattern in chapter_patterns:
                match = re.search(pattern, content)
                if match:
                    body_start = match.start()
                    break

        return (body_start if body_start else -1, body_end if body_end else -1)

    def _remove_markers(self) -> str:
        """移除所有标记，返回干净的内容（保留换行结构）"""
        content = self.content
        for name, patterns in self.MARKERS.items():
            # 移除开始标记和它后面的单个换行（如果有）
            content = re.sub(patterns['start'] + r'\n?', '', content)
            # 移除结束标记和它后面的单个换行（如果有）
            content = re.sub(patterns['end'] + r'\n?', '', content)
        return content

    def find_references_boundary(self) -> Tuple[int, int]:
        """
        自动查找参考文献的边界位置

        支持多种格式：
        - # 参考文献 或 ## 参考文献
        - **参考文献**
        - []{#_Toc11061 .anchor}参考文献 (Pandoc锚点格式)
        - 参考文献 [页码] 格式

        Returns:
            (start_pos, end_pos) 元组
        """
        # 先移除旧的标记，避免干扰
        content = self._remove_markers()

        ref_start = None

        # 查找参考文献开始位置
        # 格式1: # 参考文献 或 ## 参考文献
        for match in re.finditer(r'^#{1,3}\s+参考文献', content, re.MULTILINE):
            ref_start = match.start()
            break

        # 格式2: **参考文献**
        if ref_start is None:
            for match in re.finditer(r'\*\*参考文献\*\*', content):
                ref_start = match.start()
                break

        # 格式3: Pandoc锚点格式 []{#_Tocxxx .anchor}参考文献
        if ref_start is None:
            for match in re.finditer(r'\[]{#_Toc\d+\s*\.[a-z]+\}[参考文献\(]+', content):
                pos = match.start()
                # 确保这不是目录中的
                if pos > 0 and content[pos - 1] == '[':
                    continue
                ref_start = pos
                break

        # 格式4: 参考文献 前面有 [ 后面有数字或空
        if ref_start is None:
            for match in re.finditer(r'参考文献\s*\[', content):
                pos = match.start()
                # 确保这不是目录中的
                if pos > 5:
                    prev_context = content[max(0, pos-20):pos]
                    if re.search(r'\[参考文献\s*\[\d+\]', prev_context):
                        continue
                ref_start = pos
                break

        if ref_start is None:
            return (-1, -1)

        # 参考文献结束：文件末尾
        ref_end = len(content)

        return (ref_start, ref_end)

    def preview_mark(self) -> dict:
        """
        预览自动标记结果（不写入文件）

        Returns:
            dict: 包含检测结果的字典
                - abs_start, abs_end: 摘要边界
                - body_start, body_end: 正文边界
                - ref_start, ref_end: 参考文献边界
                - abstract_preview: 摘要内容预览
                - body_preview: 正文内容预览
                - references_preview: 参考文献预览
                - success: 是否成功
                - error: 错误信息（如有）
        """
        result = {
            'success': False,
            'error': None,
            'abs_start': -1, 'abs_end': -1,
            'body_start': -1, 'body_end': -1,
            'ref_start': -1, 'ref_end': -1,
            'abstract_preview': '',
            'body_preview': '',
            'references_preview': '',
        }

        # 获取清理后的干净内容
        clean_content = self._remove_markers()

        # 查找各部分边界
        abs_start, abs_end = self.find_abstract_boundary()
        body_start, body_end = self.find_body_boundary()
        ref_start, ref_end = self.find_references_boundary()

        result['abs_start'] = abs_start
        result['abs_end'] = abs_end
        result['body_start'] = body_start
        result['body_end'] = body_end
        result['ref_start'] = ref_start
        result['ref_end'] = ref_end

        if abs_start < 0 or body_start < 0 or ref_start < 0:
            result['error'] = "无法找到论文结构"
            return result

        # 生成预览内容
        if abs_start >= 0 and abs_end > abs_start:
            result['abstract_preview'] = clean_content[abs_start:abs_end].strip()
        if body_start >= 0 and body_end > body_start:
            result['body_preview'] = clean_content[body_start:body_start+300].strip()
        if ref_start >= 0:
            result['references_preview'] = clean_content[ref_start:ref_start+200].strip()

        result['success'] = True
        return result

    def confirm_and_mark(self) -> bool:
        """
        交互式确认后添加标记

        显示预览信息，询问用户确认，确认后写入文件

        Returns:
            是否成功添加标记
        """
        import sys

        # 先预览
        result = self.preview_mark()

        if not result['success']:
            print(f"❌ 错误: {result['error']}")
            return False

        print("\n" + "="*60)
        print("📋 自动标记预览")
        print("="*60)

        print(f"\n【摘要】({result['abs_start']} - {result['abs_end']})")
        preview = result['abstract_preview']
        if len(preview) > 300:
            preview = preview[:300] + "..."
        print(preview)

        print(f"\n【正文开始】({result['body_start']} - {result['body_end']})")
        preview = result['body_preview']
        if len(preview) > 300:
            preview = preview[:300] + "..."
        print(preview)

        print(f"\n【参考文献】({result['ref_start']} - {result['ref_end']})")
        preview = result['references_preview']
        if len(preview) > 200:
            preview = preview[:200] + "..."
        print(preview)

        print("\n" + "="*60)

        # 询问确认
        print("\n请确认标记是否正确：")
        print("  y - 确认并写入文件")
        print("  n - 取消，不写入")
        print("  q - 退出")

        while True:
            try:
                choice = input("\n您的选择 (y/n/q): ").strip().lower()
                if choice == 'y':
                    return self._write_markers()
                elif choice == 'n':
                    print("已取消，未写入文件。")
                    return False
                elif choice == 'q':
                    print("已退出。")
                    sys.exit(0)
                else:
                    print("无效选项，请输入 y、n 或 q")
            except EOFError:
                # 非交互环境（piped input）
                print("检测到非交互环境，自动确认并写入。")
                return self._write_markers()

    def _write_markers(self) -> bool:
        """
        内部方法：将标记写入文件

        Returns:
            是否成功
        """
        # 查找各部分边界
        abs_start, abs_end = self.find_abstract_boundary()
        body_start, body_end = self.find_body_boundary()
        ref_start, ref_end = self.find_references_boundary()

        if abs_start < 0 or body_start < 0 or ref_start < 0:
            print("写入失败：无法找到论文结构")
            return False

        # 获取清理后的干净内容
        clean_content = self._remove_markers()

        # 插入标记（按位置从后往前插入，避免位置偏移）
        insertions = [
            (ref_end, '\n<!-- 参考文献结束 -->'),
            (ref_start, '<!-- 参考文献开始 -->\n'),
            (body_end, '\n<!-- 正文结束 -->'),
            (body_start, '<!-- 正文开始 -->\n'),
            (abs_end, '\n<!-- 摘要结束 -->'),
            (abs_start, '<!-- 摘要开始 -->\n'),
        ]

        # 按位置从后往前排序
        insertions.sort(key=lambda x: -x[0])

        new_content = clean_content
        for pos, marker in insertions:
            if 0 <= pos <= len(new_content):
                new_content = new_content[:pos] + marker + new_content[pos:]

        # 写回文件
        try:
            with open(self.md_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            self.content = new_content
            self.marks = self._find_all_marks()
            print("✅ 标记已写入文件！")
            return True
        except Exception as e:
            print(f"写入文件失败: {e}")
            return False

    def auto_mark(self) -> bool:
        """
        自动检测论文结构并添加标记（直接写入，不确认）

        Returns:
            是否成功添加标记
        """
        # 先移除旧的标记
        for name, patterns in self.MARKERS.items():
            self.content = re.sub(patterns['start'] + r'\s*\n?', '', self.content)
            self.content = re.sub(patterns['end'] + r'\s*\n?', '', self.content)

        # 重新加载
        self.marks = self._find_all_marks()

        # 查找各部分边界
        abs_start, abs_end = self.find_abstract_boundary()
        body_start, body_end = self.find_body_boundary()
        ref_start, ref_end = self.find_references_boundary()

        print(f"自动标记检测结果:")
        print(f"  摘要: {abs_start} - {abs_end}")
        print(f"  正文: {body_start} - {body_end}")
        print(f"  参考文献: {ref_start} - {ref_end}")

        if abs_start < 0 or body_start < 0 or ref_start < 0:
            print("自动标记失败：无法找到论文结构")
            return False

        return self._write_markers()


def extract_marked_content(md_path: str) -> dict:
    """
    便捷函数：从Markdown文件中提取标记内容

    Args:
        md_path: Markdown文件路径

    Returns:
        包含各部分内容的字典
    """
    marker = Marker(md_path)
    return marker.extract_all()


def preview_mark_paper(md_path: str) -> dict:
    """
    便捷函数：预览自动标记结果（不写入文件）

    Args:
        md_path: Markdown文件路径

    Returns:
        包含检测结果的字典
    """
    marker = Marker(md_path)
    return marker.preview_mark()


def confirm_and_mark_paper(md_path: str) -> bool:
    """
    便捷函数：用户确认后添加标记

    Args:
        md_path: Markdown文件路径

    Returns:
        是否成功
    """
    marker = Marker(md_path)
    return marker.confirm_and_mark()


def auto_mark_paper(md_path: str) -> bool:
    """
    便捷函数：自动为论文添加标记（直接写入，不确认）

    Args:
        md_path: Markdown文件路径

    Returns:
        是否成功
    """
    marker = Marker(md_path)
    return marker.auto_mark()