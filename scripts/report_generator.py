#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文评价报告生成器

功能：
- 生成AI填充用报告骨架
"""

from datetime import datetime
from typing import Dict


# 评价维度配置
DIMENSION_WEIGHTS = {
    'topic': {'name': '选题与研究问题', 'weight': '15%'},
    'references': {'name': '参考文献与学术规范', 'weight': '15%'},
    'innovation': {'name': '内容创新性', 'weight': '15%'},
    'structure': {'name': '框架与逻辑结构', 'weight': '20%'},
    'methodology': {'name': '方法与论证严谨性', 'weight': '25%'},
    'language': {'name': '语言与表达', 'weight': '10%'},
}


class ReportGenerator:
    """评价报告生成器"""

    @staticmethod
    def generate_report_skeleton(analysis_result: Dict, markdown_path: str) -> str:
        """
        生成AI填充用报告骨架

        Args:
            analysis_result: PaperAnalyzer返回的分析结果字典
            markdown_path: 论文markdown文件路径

        Returns:
            Markdown格式报告骨架
        """
        metadata = analysis_result.get('metadata', {})
        student_info = analysis_result.get('student_info', {})
        basic_stats = analysis_result.get('basic_stats', {})
        structure = analysis_result.get('structure', {})
        writing_specs = analysis_result.get('writing_specs', {})
        model_spec = analysis_result.get('model_spec', {})
        abstract = analysis_result.get('abstract', '')

        paper_type = metadata.get('paper_type_cn', '待识别')
        is_empirical = '实证' in paper_type

        md = []
        today = datetime.now().strftime('%Y年%m月%d日')

        # === 封面 ===
        md.append("# 论文评价报告\n")
        md.append(f"> **评价日期**：{today}")
        md.append("> **状态**：待AI评价 - 请填充以下各部分内容\n")
        md.append("---\n")

        # === 基本信息 ===
        md.append("## 基本信息\n")
        md.append(f"- **学生姓名**：{student_info.get('name', '待识别')}")
        md.append(f"- **学号**：{student_info.get('student_id', '待识别')}")
        md.append(f"- **专业**：{student_info.get('class_name', '待识别')}")
        md.append(f"- **论文题目**：{metadata.get('paper_title', student_info.get('paper_title', '待识别'))}")
        md.append(f"- **论文类型**：{paper_type}")
        md.append(f"- **字数**：{basic_stats.get('word_count', 0)}字 {'✅' if basic_stats.get('word_count_ok') else '❌'}")

        if basic_stats.get('title_word_count'):
            md.append(f"- **标题字数**：{basic_stats.get('title_word_count')}字 {'✅' if basic_stats.get('title_word_count_ok') else '❌'}")
        md.append("")

        # === 字数统计 ===
        md.append("## 字数统计\n")
        has_markers = basic_stats.get('has_markers', False)
        word_count = basic_stats.get('word_count', 0)
        if has_markers:
            abstract_count = basic_stats.get('abstract_count', 0)
            md.append(f"**正文字数**：{word_count}字（AI标记）| 要求：≥8000字 | {'✅' if word_count >= 8000 else '❌'}")
            md.append(f"**摘要字数**：{abstract_count}字（AI标记）| 要求：300-500字 | {'✅' if 300 <= abstract_count <= 500 else '❌'}")
        elif word_count >= 8000:
            md.append(f"**正文字数**：{word_count}字 | 要求：≥8000字 | ✅")
            md.append(f"**摘要字数**：需AI标记后统计 | 要求：300-500字")
        else:
            md.append(f"**正文字数（正则）**：{word_count}字 | 要求：≥8000字 | ❌")
            md.append(f"**摘要字数**：需AI标记后统计 | 要求：300-500字")
            md.append("")
            md.append("**【AI标记】**：请读取论文markdown，在摘要和正文处添加标记：")
            md.append("```")
            md.append("【摘要】")
            md.append("...摘要内容...")
            md.append("【摘要结束】")
            md.append("")
            md.append("【正文内容】")
            md.append("...正文内容...")
            md.append("【正文内容结束】")
            md.append("```")
        md.append("")
        md.append("---\n")

        # === 参考文献 ===
        md.append("## 参考文献（Python检测）\n")
        md.append(f"- **参考文献总数**：{basic_stats.get('ref_count', 0)}篇 {'✅' if basic_stats.get('ref_count_ok') else '❌'}")
        md.append(f"- **外文文献**：{basic_stats.get('foreign_ref_count', 0)}篇 {'✅' if basic_stats.get('foreign_ref_ok') else '❌'}")
        md.append(f"- **期刊占比**：{basic_stats.get('journal_ratio', 0)*100:.0f}%")
        md.append(f"- **近5年占比**：{basic_stats.get('recent_5yr_ratio', 0)*100:.0f}%" if basic_stats.get('recent_5yr_ratio', 0) > 0 else "- **近5年占比**：未检测")
        md.append("")
        md.append("---\n")

        # === 写作规范 ===
        md.append("## 写作规范（Python检测）\n")
        md.append(f"- **第一人称使用**：{writing_specs.get('first_person_count', 0)}处 {'⚠️需检查' if writing_specs.get('first_person_count', 0) > 0 else '✅'}")
        md.append(f"- **表格数量**：{writing_specs.get('table_count', 0)}")
        md.append(f"- **软件截图**：{'有' if writing_specs.get('has_software_screenshot') else '无'}")
        if writing_specs.get('issues'):
            for issue in writing_specs.get('issues', [])[:3]:
                md.append(f"  - ⚠️ {issue}")
        md.append("")
        md.append("---\n")

        # === 模型设定（仅实证性论文）===
        if is_empirical and model_spec:
            md.append("## 模型设定分析（Python检测 + AI深度评价）\n")
            variables = model_spec.get('variables', [])
            formulas = model_spec.get('formulas', [])
            causal_chains = model_spec.get('causal_chains', [])
            over_control = model_spec.get('over_control_issues', [])

            md.append(f"- **提取变量数**：{len(variables)} {'(Python提取)' if len(variables) > 0 else '(AI需深度提取)'}")
            md.append(f"- **提取公式数**：{len(formulas)}")
            md.append(f"- **提取因果链**：{len(causal_chains)}")
            md.append(f"- **过度控制问题**：{len(over_control)}处")

            if variables:
                md.append("\n**变量定义表**：")
                md.append("| 变量名 | 角色 | 定义 |")
                md.append("|--------|------|------|")
                for v in variables[:10]:
                    role = v.get('role', 'unknown')
                    role_cn = {'dependent': '被解释', 'independent': '核心解释', 'control': '控制', 'mediator': '中介'}.get(role, role)
                    md.append(f"| {v.get('name_en', v.get('name_cn', '-'))} | {role_cn} | {v.get('measurement', '-') or v.get('definition_text', '-')} |")

            if over_control:
                md.append("\n**⚠️ 过度控制问题**：")
                for issue in over_control:
                    md.append(f"- **{issue.get('mediator_var', '-')}**：{issue.get('explanation', '').split(chr(10))[0]}")

            endogeneity = model_spec.get('endogeneity_check')
            if endogeneity and endogeneity.get('has_endogeneity'):
                md.append(f"\n- **内生性处理**：{endogeneity.get('method_used', '有讨论')}")
                md.append(f"  - 处理充分性：{'✅' if endogeneity.get('is_adequate') else '❌ 不足'}")
            md.append("")
            md.append("---\n")

        # === 论文原文链接 ===
        md.append(f"**论文全文**：{markdown_path}\n")
        md.append("---\n")

        # === AI评价区域（待填充）===
        md.append("# AI评价（请填充以下内容）\n")
        md.append("> 以下内容请AI基于论文全文和上述检测数据完成评价后填写\n")

        # 总体评价
        md.append("## 一、总体评价\n")
        md.append("### 1.1 综合评分：__分（__）\n")
        md.append("### 1.2 各维度得分\n")
        md.append("| 评价维度 | 权重 | 得分 | 关键问题 |")
        md.append("|---------|------|------|---------|")
        for key, info in DIMENSION_WEIGHTS.items():
            md.append(f"| {info['name']} | {info['weight']} | __分 |  |")
        md.append("")
        md.append("### 1.3 评价概述\n")
        md.append("[100-200字概括论文整体质量，突出主要优点和核心问题]\n")
        md.append("")
        md.append("---\n")

        # 各维度详细评价
        md.append("## 二、各维度详细评价\n")
        dimension_descriptions = {
            'topic': '题目质量（≤20字）、研究必要性、研究价值',
            'references': '数量≥15篇、外文≥3篇、期刊占比',
            'innovation': '研究视角、边际贡献、文献评述',
            'structure': '章节完整、逻辑清晰',
            'methodology': '理论分析、稳健性检验（实证论文强制）、内生性',
            'language': '第三人称、术语规范',
        }

        for i, (key, info) in enumerate(DIMENSION_WEIGHTS.items(), 1):
            md.append(f"### 2.{i} {info['name']}（__分）\n")
            md.append(f"**评价要点**：{dimension_descriptions.get(key, '')}\n")
            md.append("**优点**：\n")
            md.append("- [引用原文说明优点]\n")
            md.append("**问题**：\n")
            md.append("1. [引用原文说明问题]\n")
            md.append("**修改建议**：\n")
            md.append("| 优先级 | 建议 |")
            md.append("|-------|------|")
            md.append("| 高 | [必须修改的问题] |")
            md.append("| 中 | [建议修改的问题] |")
            md.append("| 低 | [可选完善的问题] |")
            md.append("")
            md.append("---\n")

        # 修改建议汇总
        md.append("## 三、修改建议汇总\n")
        md.append("### 高优先级（必须修改，影响论文合格）\n")
        md.append("1. [具体可操作的修改建议]\n")
        md.append("")
        md.append("### 中优先级（建议修改，提升论文质量）\n")
        md.append("1. [具体可操作的修改建议]\n")
        md.append("")
        md.append("### 低优先级（可选修改，完善细节）\n")
        md.append("1. [可选完善的问题]\n")
        md.append("")
        md.append("---\n")

        # 总结
        md.append("## 四、总结\n")
        md.append("[总结性评价，对论文给出最终结论和修改意见]\n")
        md.append("")
        md.append("---\n")
        md.append(f"*本报告骨架由Python规则层生成 | {today}*\n")

        return '\n'.join(md)
