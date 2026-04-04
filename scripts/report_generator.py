#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文评价报告生成器

功能：
- 生成结构化评价报告（Markdown格式）
- 生成Word格式报告
"""

import os
import json
from datetime import datetime
from typing import Dict, List, Optional


class ReportGenerator:
    """评价报告生成器"""

    def __init__(self, evaluation_result: Dict):
        """
        初始化报告生成器

        Args:
            evaluation_result: AI评价结果（字典格式）
        """
        self.data = evaluation_result
        self.metadata = evaluation_result.get('metadata', {})
        self.student_info = evaluation_result.get('student_info', {})
        self.basic_stats = evaluation_result.get('basic_stats', {})
        self.dimensions = evaluation_result.get('dimensions', {})
        self.suggestions = evaluation_result.get('suggestions', {})
        self.overview = evaluation_result.get('overview', '')

    def generate_markdown(self) -> str:
        """生成Markdown格式报告"""
        md = []
        md.append("# 论文评价报告\n")

        # 基本信息
        md.append("## 基本信息")
        md.append(f"- 学生姓名：{self.student_info.get('name', '待填写')}")
        md.append(f"- 学号：{self.student_info.get('student_id', '待填写')}")
        md.append(f"- 班级/专业：{self.student_info.get('class_name', '待填写')}")
        md.append(f"- 论文题目：{self.metadata.get('paper_title', self.student_info.get('paper_title', '待填写'))}")
        md.append(f"- 论文类型：{self.metadata.get('paper_type_cn', '待识别')}")
        md.append(f"- 评价日期：{datetime.now().strftime('%Y年%m月%d日')}")
        md.append("\n---\n")

        # 总体评价
        md.append("## 一、总体评价\n")

        # 综合评分
        overall_score = self.data.get('overall_score', 0)
        grade = self._get_grade(overall_score)
        md.append(f"### 1.1 综合评分：{overall_score}分（{grade}）\n")

        # 各维度得分
        md.append("### 1.2 各维度得分\n")
        md.append("| 评价维度 | 权重 | 得分 | 关键问题 |")
        md.append("|---------|------|------|---------|")

        dimension_weights = {
            '选题与研究问题': '15%',
            '参考文献与学术规范': '15%',
            '内容创新性': '15%',
            '框架与逻辑结构': '20%',
            '方法与论证严谨性': '25%',
            '语言与表达': '10%',
        }

        dimension_keys = [
            'topic',      # 选题与研究问题
            'references', # 参考文献与学术规范
            'innovation', # 内容创新性
            'structure',  # 框架与逻辑结构
            'methodology', # 方法与论证严谨性
            'language',   # 语言与表达
        ]

        dimension_names_cn = {
            'topic': '选题与研究问题',
            'references': '参考文献与学术规范',
            'innovation': '内容创新性',
            'structure': '框架与逻辑结构',
            'methodology': '方法与论证严谨性',
            'language': '语言与表达',
        }

        for key in dimension_keys:
            dim = self.dimensions.get(key, {})
            score = dim.get('score', 0)
            weight = dimension_weights.get(dimension_names_cn.get(key, ''), '-')
            summary = dim.get('summary', '-')
            md.append(f"| {dimension_names_cn.get(key, key)} | {weight} | {score}分 | {summary} |")

        md.append("")

        # 计算过程
        md.append(f"**计算过程**：")
        md.append(f"综合评分 = 15%×{self.dimensions.get('topic', {}).get('score', 0)} + "
                 f"15%×{self.dimensions.get('references', {}).get('score', 0)} + "
                 f"15%×{self.dimensions.get('innovation', {}).get('score', 0)} + "
                 f"20%×{self.dimensions.get('structure', {}).get('score', 0)} + "
                 f"25%×{self.dimensions.get('methodology', {}).get('score', 0)} + "
                 f"10%×{self.dimensions.get('language', {}).get('score', 0)}")
        md.append("")

        # 评价概述
        md.append("### 1.3 评价概述\n")
        md.append(self.overview if self.overview else "（待AI评价后填写）")
        md.append("\n---\n")

        # 各维度详细评价
        md.append("## 二、各维度详细评价\n")

        for key in dimension_keys:
            dim = self.dimensions.get(key, {})
            if not dim:
                continue

            dim_name = dimension_names_cn.get(key, key)
            md.append(f"### 2.{dimension_keys.index(key)+1} {dim_name}（{dim.get('score', 0)}分）\n")

            # 优点
            advantages = dim.get('advantages', [])
            if advantages:
                md.append("**优点**：")
                for adv in advantages:
                    md.append(f"- {adv}")
                md.append("")

            # 问题
            problems = dim.get('problems', [])
            if problems:
                md.append("**问题**：")
                for i, prob in enumerate(problems, 1):
                    md.append(f"{i}. {prob}")
                md.append("")

            # 修改建议
            dim_suggestions = dim.get('suggestions', [])
            if dim_suggestions:
                md.append("**修改建议**：")
                md.append("| 优先级 | 建议 |")
                md.append("|-------|------|")
                for sug in dim_suggestions:
                    priority = sug.get('priority', '中')
                    content = sug.get('content', '')
                    md.append(f"| {priority} | {content} |")
                md.append("")

            md.append("---\n")

        # 修改建议汇总
        md.append("## 三、修改建议汇总\n")

        high_priority = self.suggestions.get('high', [])
        medium_priority = self.suggestions.get('medium', [])
        low_priority = self.suggestions.get('low', [])

        if high_priority:
            md.append("### 高优先级修改建议（必须修改，否则影响论文合格）\n")
            md.append("| 序号 | 问题 | 修改方法 |")
            md.append("|-----|------|---------|")
            for i, sug in enumerate(high_priority, 1):
                md.append(f"| {i} | {sug.get('problem', '')} | {sug.get('method', '')} |")
            md.append("")

        if medium_priority:
            md.append("### 中优先级修改建议（建议修改，可显著提升质量）\n")
            md.append("| 序号 | 问题 | 修改建议 |")
            md.append("|-----|------|---------|")
            for i, sug in enumerate(medium_priority, 1):
                md.append(f"| {i} | {sug.get('problem', '')} | {sug.get('method', '')} |")
            md.append("")

        if low_priority:
            md.append("### 低优先级修改建议（可选修改，进一步完善）\n")
            for i, sug in enumerate(low_priority, 1):
                md.append(f"{i}. {sug.get('problem', '')}")
            md.append("")

        # 总结
        md.append("## 四、总结\n")
        summary = self.data.get('summary', '（待AI评价后填写）')
        md.append(summary)
        md.append("\n---\n")

        # 评价人信息
        md.append(f"**评价人**：AI评价系统")
        md.append(f"**评价日期**：{datetime.now().strftime('%Y年%m月%d日')}")

        return '\n'.join(md)

    def generate_markdown_simple(self, ai_evaluation: str) -> str:
        """
        生成简化版Markdown报告（用于接收AI评价结果后填充）

        Args:
            ai_evaluation: AI评价报告内容（Markdown格式）
        """
        # 生成报告头部
        header = []
        header.append("# 论文评价报告\n")
        header.append("## 基本信息")
        header.append(f"- 学生姓名：{self.student_info.get('name', '待填写')}")
        header.append(f"- 学号：{self.student_info.get('student_id', '待填写')}")
        header.append(f"- 班级/专业：{self.student_info.get('class_name', '待填写')}")
        header.append(f"- 论文题目：{self.metadata.get('paper_title', self.student_info.get('paper_title', '待填写'))}")
        header.append(f"- 论文类型：{self.metadata.get('paper_type_cn', '待识别')}")
        header.append(f"- 评价日期：{datetime.now().strftime('%Y年%m月%d日')}")

        # 返回完整报告
        return '\n'.join(header) + "\n\n---\n\n" + ai_evaluation

    def _get_grade(self, score: int) -> str:
        """根据分数确定等级"""
        if score >= 90:
            return "优秀（优+）"
        elif score >= 85:
            return "良好上（优）"
        elif score >= 80:
            return "良好中（良+）"
        elif score >= 75:
            return "良好下（良）"
        elif score >= 70:
            return "中等上（中+）"
        elif score >= 65:
            return "中等（中）"
        elif score >= 60:
            return "中等下（中-）"
        else:
            return "不合格"

    def save_markdown(self, content: str, output_path: str):
        """保存Markdown报告"""
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✅ Markdown报告已保存：{output_path}")


def create_report_data_template() -> Dict:
    """创建空的报告数据模板"""
    return {
        'metadata': {
            'file_path': '',
            'file_name': '',
            'paper_type': '',
            'paper_type_cn': '',
            'paper_title': '',
            'analysis_time': '',
        },
        'student_info': {
            'name': '',
            'student_id': '',
            'class_name': '',
            'paper_title': '',
            'advisor': '',
        },
        'basic_stats': {
            'word_count': 0,
            'word_count_ok': False,
            'ref_count': 0,
            'ref_count_ok': False,
            'foreign_ref_count': 0,
            'foreign_ref_ok': False,
            'journal_ratio': 0.0,
            'recent_5yr_ratio': 0.0,
        },
        'structure': {
            'sections': [],
            'has_required': {},
            'issues': [],
        },
        'writing_specs': {
            'first_person_count': 0,
            'has_software_screenshot': False,
            'table_count': 0,
            'issues': [],
        },
        'abstract': '',
        'overall_score': 0,
        'grade': '',
        'overview': '',
        'dimensions': {
            'topic': {
                'score': 0,
                'summary': '',
                'advantages': [],
                'problems': [],
                'suggestions': [],
            },
            'references': {
                'score': 0,
                'summary': '',
                'advantages': [],
                'problems': [],
                'suggestions': [],
            },
            'innovation': {
                'score': 0,
                'summary': '',
                'advantages': [],
                'problems': [],
                'suggestions': [],
            },
            'structure': {
                'score': 0,
                'summary': '',
                'advantages': [],
                'problems': [],
                'suggestions': [],
            },
            'methodology': {
                'score': 0,
                'summary': '',
                'advantages': [],
                'problems': [],
                'suggestions': [],
            },
            'language': {
                'score': 0,
                'summary': '',
                'advantages': [],
                'problems': [],
                'suggestions': [],
            },
        },
        'suggestions': {
            'high': [],
            'medium': [],
            'low': [],
        },
        'summary': '',
    }


def main():
    """测试入口"""
    # 创建示例数据
    template = create_report_data_template()

    # 生成报告
    generator = ReportGenerator(template)
    report = generator.generate_markdown()

    print(report)


if __name__ == '__main__':
    main()
