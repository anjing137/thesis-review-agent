#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学生信息提取模块:

从论文封面/Markdown中提取学生信息

[DEPRECATED] 此模块已废弃，仅用于 .md 旧路径的向后兼容。
新流程（.docx 直接分析）请使用 xml_analyzer 模块。
"""

import re
import os
from typing import Dict, Optional


class Extractor:
    """学生信息提取器"""

    def __init__(self, md_path: str):
        """
        初始化提取器

        Args:
            md_path: Markdown文件路径
        """
        self.md_path = md_path
        self._load_content()

    def _load_content(self):
        """加载文件内容"""
        try:
            with open(self.md_path, 'r', encoding='utf-8') as f:
                self.content = f.read()
        except FileNotFoundError:
            print(f"⚠️ 文件不存在: {self.md_path}")
            self.content = ""
        except PermissionError:
            print(f"⚠️ 无权限读取文件: {self.md_path}")
            self.content = ""
        except Exception as e:
            print(f"⚠️ 读取文件失败 {self.md_path}: {e}")
            self.content = ""

    def extract(self) -> Dict:
        """
        提取学生信息

        Returns:
            学生信息字典
        """
        return {
            'name': self.extract_name(),
            'student_id': self.extract_student_id(),
            'class_name': self.extract_class_name(),
            'advisor': self.extract_advisor(),
            'paper_title': self.extract_paper_title(),
            'paper_type': self.extract_paper_type(),
        }

    def _preprocess_line(self, line: str) -> str:
        """
        预处理单行内容

        Args:
            line: 原始行内容

        Returns:
            清理后的内容
        """
        # 移除星号
        line = re.sub(r'\*+', '', line)
        # 清理 [.underline}] 格式，如 [张胜楠 ]{.underline}
        line = re.sub(r'\[([^\]]+)\]\s*\{[^}]*\}', r'\1', line)
        # 清理残留的方括号
        line = line.replace('[', '').replace(']', '')
        # 移除引用标记
        line = re.sub(r'^>\s*', '', line)
        return line.strip()

    def _find_value_after_label(self, lines: list, label: str, max_lines: int = 3) -> Optional[str]:
        """
        在标签后查找值

        Args:
            lines: 行列表
            label: 标签名（如"姓名"）
            max_lines: 最大查找范围

        Returns:
            找到的值，未找到返回None
        """
        for i, line in enumerate(lines):
            if label in line:
                # 在接下来的几行中查找值
                for j in range(i, min(i + max_lines, len(lines))):
                    value_line = self._preprocess_line(lines[j])
                    # 移除标签本身
                    value_line = re.sub(f'{label}[：:]*', '', value_line)
                    value_line = value_line.strip()
                    # 移除方括号
                    value_line = value_line.replace('[', '').replace(']', '')
                    if value_line and len(value_line) > 0:
                        return value_line
        return None

    def extract_name(self) -> Optional[str]:
        """
        提取学生姓名

        Returns:
            姓名，未找到返回None
        """
        lines = self.content.split('\n')[:60]  # 封面通常在前60行

        # 尝试多种模式（注意"姓名"中间可能有空格）
        patterns = [
            r'姓\s*名[：:]\s*\[?([^\]\s]+)',
            r'^\[([^\]]+)\]\s*$',  # [姓名] 格式
        ]

        for i, line in enumerate(lines):
            clean = self._preprocess_line(line)

            for pattern in patterns:
                match = re.search(pattern, clean)
                if match:
                    name = match.group(1).strip()
                    # 过滤常见非姓名内容
                    skip_words = ['学院', '专业', '班级', '学号', '指导教师', '摘要', 'Abstract']
                    if name and not any(sw in name for sw in skip_words):
                        return name

        return None

    def extract_student_id(self) -> Optional[str]:
        """
        提取学号

        Returns:
            学号，未找到返回None
        """
        lines = self.content.split('\n')[:60]

        patterns = [
            r'学\s*号[：:]*\s*\[?(\d+)',
            r'学\s*号[：:]*(\d+)',
        ]

        for line in lines:
            clean = self._preprocess_line(line)

            for pattern in patterns:
                match = re.search(pattern, clean)
                if match:
                    return match.group(1).strip()

        return None

    def extract_class_name(self) -> Optional[str]:
        """
        提取班级/专业信息

        Returns:
            班级/专业，未找到返回None
        """
        lines = self.content.split('\n')[:60]

        # 先尝试"专业"
        patterns_major = [
            r'专\s*业[：:]\s*\[?([^\]\s]+)',
        ]

        for line in lines:
            clean = self._preprocess_line(line)
            for pattern in patterns_major:
                match = re.search(pattern, clean)
                if match:
                    major = match.group(1).strip()
                    if major and len(major) > 1:
                        return major

        # 再尝试"班级"
        patterns_class = [
            r'班\s*级[：:]\s*\[?([^\]\s]+)',
        ]

        for line in lines:
            clean = self._preprocess_line(line)
            for pattern in patterns_class:
                match = re.search(pattern, clean)
                if match:
                    class_name = match.group(1).strip()
                    if class_name and len(class_name) > 1:
                        return class_name

        return None

    def extract_advisor(self) -> Optional[str]:
        """
        提取指导教师

        Returns:
            指导教师，未找到返回None
        """
        lines = self.content.split('\n')[:60]

        patterns = [
            r'指\s*导\s*教\s*师[：:]\s*\[?([^\]\s]+)',
        ]

        for line in lines:
            clean = self._preprocess_line(line)
            for pattern in patterns:
                match = re.search(pattern, clean)
                if match:
                    advisor = match.group(1).strip()
                    if advisor and len(advisor) > 1:
                        return advisor

        return None

    def extract_paper_title(self) -> Optional[str]:
        """
        提取论文题目

        Returns:
            论文题目，未找到返回None
        """
        lines = self.content.split('\n')[:80]

        # 模式1: 加粗标题 **题目**
        for i, line in enumerate(lines[:30]):
            if '<!--' in line:  # 跳过注释
                continue
            match = re.match(r'^\*\*([^\*]{5,50})\*\*$', line.strip())
            if match:
                title = match.group(1).strip()
                # 排除非标题内容
                skip_words = ['河南科技学院', '学年', '学期', '学号', '姓名', '专业', '学院',
                             '指导教师', '完成时间', '摘要', 'Abstract', '目录']
                if title and not any(sw in title for sw in skip_words):
                    return title

        # 模式2: 在"学年论文"/"毕业论文"后查找
        for i, line in enumerate(lines):
            if '学年论文' in line or '毕业论文' in line:
                for j in range(i + 1, min(i + 5, len(lines))):
                    candidate = self._preprocess_line(lines[j])
                    if len(candidate) > 5:
                        skip_words = ['学院', '专业', '班级', '学号', '姓名', '指导教师']
                        if not any(sw in candidate for sw in skip_words):
                            return candidate

        return None

    def extract_paper_type(self) -> str:
        """
        判断论文类型

        Returns:
            论文类型（实证性/学理性）
        """
        # 合并前500行内容用于判断
        content_sample = self.content[:5000]

        # 实证性论文关键词
        empirical_keywords = [
            '回归', '回归分析', '实证', '数据', '样本', '变量',
            '显著性', 'P值', 'p值', 'R²', 'r²', '计量模型',
            '稳健性', '内生性', '工具变量', '面板数据', '时间序列',
            'GDP', 'CPI', '股价', '收益率', '上市公司', '并购',
            '实证研究', '实证分析', '实证检验'
        ]

        # 学理性论文关键词
        theoretical_keywords = [
            '现状分析', '问题分析', '原因分析', '对策建议',
            '对策研究', '建议', '措施', '策略',
            '现状', '问题', '原因', '对策'
        ]

        empirical_count = sum(1 for kw in empirical_keywords if kw in content_sample)
        theoretical_count = sum(1 for kw in theoretical_keywords if kw in content_sample)

        if empirical_count > theoretical_count:
            return '实证性'
        else:
            return '学理性'


def extract_from_markdown(md_path: str) -> Dict:
    """
    便捷函数：从Markdown文件提取学生信息

    Args:
        md_path: Markdown文件路径

    Returns:
        学生信息字典
    """
    if not os.path.exists(md_path):
        return {
            'error': f'文件不存在：{md_path}'
        }

    extractor = Extractor(md_path)
    return extractor.extract()