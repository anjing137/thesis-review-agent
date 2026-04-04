#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文评价 Agent - 主入口

功能：
- 分析论文（Python规则检测层）
- 生成结构化数据供AI评价
- 生成评价报告
"""

import os
import sys
import json
import argparse

# 添加scripts目录到路径
script_dir = os.path.dirname(os.path.abspath(__file__))
scripts_dir = os.path.join(script_dir, 'scripts')
if scripts_dir not in sys.path:
    sys.path.insert(0, scripts_dir)

from paper_analyzer import PaperAnalyzer
from report_generator import ReportGenerator


def analyze_paper(paper_path: str, output_dir: str = None) -> dict:
    """
    分析单篇论文

    Args:
        paper_path: 论文文件路径
        output_dir: 输出目录（可选）

    Returns:
        分析结果字典
    """
    analyzer = PaperAnalyzer(paper_path)
    result = analyzer.analyze()

    # 如果指定了输出目录，保存JSON结果
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, 'analysis_result.json')

        # 转换为可序列化格式
        result_dict = {
            'metadata': result.metadata,
            'student_info': result.student_info,
            'basic_stats': result.basic_stats,
            'structure': result.structure,
            'writing_specs': result.writing_specs,
            'abstract': result.abstract,
            'reference_section': result.reference_section,
            'model_spec': result.model_spec,
        }

        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result_dict, f, ensure_ascii=False, indent=2)

        print(f"\n✅ 分析结果已保存至：{output_file}")

    return {
        'metadata': result.metadata,
        'student_info': result.student_info,
        'basic_stats': result.basic_stats,
        'structure': result.structure,
        'writing_specs': result.writing_specs,
        'abstract': result.abstract,
        'reference_section': result.reference_section,
        'model_spec': result.model_spec,
    }


def main():
    """CLI入口"""
    parser = argparse.ArgumentParser(
        description='论文评价 Agent - 分析论文并生成结构化数据'
    )
    parser.add_argument('paper_path', help='论文文件路径（.docx格式）')
    parser.add_argument('--output', '-o', help='输出目录')
    parser.add_argument('--format', '-f', choices=['json', 'report', 'both'],
                        default='json', help='输出格式')
    parser.add_argument('--json-only', action='store_true',
                        help='仅输出JSON数据，不调用AI评价')

    args = parser.parse_args()

    if not os.path.exists(args.paper_path):
        print(f"❌ 文件不存在：{args.paper_path}")
        sys.exit(1)

    if not args.paper_path.endswith('.docx'):
        print("❌ 仅支持 .docx 格式的论文")
        sys.exit(1)

    # 分析论文
    result = analyze_paper(args.paper_path, args.output)

    # 如果只是查看数据，输出简要信息
    if args.json_only:
        print("\n📊 分析结果摘要：")
        print(f"   论文类型：{result['metadata']['paper_type_cn']}")
        print(f"   论文题目：{result['student_info'].get('paper_title', '未识别')}")
        print(f"   学生姓名：{result['student_info'].get('name', '未识别')}")
        print(f"   正文字数：{result['basic_stats']['word_count']}字")
        print(f"   参考文献：{result['basic_stats']['ref_count']}篇")
        print(f"   外文文献：{result['basic_stats']['foreign_ref_count']}篇")
        print(f"   第一人称：{result['writing_specs']['first_person_count']}处")


if __name__ == '__main__':
    main()
