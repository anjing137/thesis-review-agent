#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量处理模块:

批量处理论文文件夹，生成独立评价报告和汇总表
"""

import os
import glob
import json
import csv
from typing import List, Dict
from datetime import datetime


class BatchProcessor:
    """批量处理器"""

    def __init__(self, folder_path: str, output_dir: str = None):
        """
        初始化批量处理器

        Args:
            folder_path: 论文文件夹路径
            output_dir: 输出目录（默认为 folder_path/review_output）
        """
        self.folder_path = folder_path
        self.output_dir = output_dir or os.path.join(folder_path, 'review_output')

        # 确保输出目录存在
        os.makedirs(self.output_dir, exist_ok=True)

    def find_docx_files(self) -> List[str]:
        """
        查找文件夹下所有docx/doc文件

        Returns:
            docx/doc文件路径列表
        """
        files = []
        # 查找 .docx 和 .doc 文件
        for ext in ['*.docx', '*.doc']:
            pattern = os.path.join(self.folder_path, ext)
            files.extend(glob.glob(pattern))

        # 递归搜索子目录
        for root, dirs, files_in_dir in os.walk(self.folder_path):
            if root == self.folder_path:
                continue
            for f in files_in_dir:
                if f.endswith('.docx') or f.endswith('.doc'):
                    full_path = os.path.join(root, f)
                    if full_path not in files:
                        files.append(full_path)

        return sorted(files)

    def generate_summary_csv(self, results: List[Dict], output_path: str = None) -> str:
        """
        生成汇总CSV表

        Args:
            results: 处理结果列表
            output_path: 输出文件路径

        Returns:
            CSV文件路径
        """
        if output_path is None:
            output_path = os.path.join(self.output_dir, '评价汇总.csv')

        if not results:
            # 创建空文件
            with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['姓名', '学号', '论文题目', '指导教师', '综合得分', '评级', '评价报告路径'])
            return output_path

        # 准备数据
        rows = [['姓名', '学号', '论文题目', '指导教师', '综合得分', '评级', '评价报告路径']]

        for result in results:
            student_info = result.get('student_info', {})
            evaluation = result.get('evaluation', {})

            row = [
                student_info.get('name', '未知'),
                student_info.get('student_id', '未知'),
                student_info.get('paper_title', '未知'),
                student_info.get('advisor', '未知'),
                evaluation.get('total_score', '待评价'),
                evaluation.get('grade', '待评价'),
                result.get('report_path', ''),
            ]
            rows.append(row)

        # 写入CSV
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(rows)

        return output_path

    def process_batch(self, converter_func=None, extractor_func=None, stats_func=None) -> Dict:
        """
        批量处理所有论文

        Args:
            converter_func: 文档转换函数
            extractor_func: 信息提取函数
            stats_func: 统计函数

        Returns:
            批量处理结果
        """
        docx_files = self.find_docx_files()

        if not docx_files:
            return {
                'success': False,
                'error': f'未找到docx文件：{self.folder_path}',
                'processed': 0,
                'results': []
            }

        results = []
        summary = {
            'total': len(docx_files),
            'processed': 0,
            'failed': 0,
            'output_dir': self.output_dir,
        }

        for i, paper_path in enumerate(docx_files):
            paper_result = {
                'paper_path': paper_path,
                'paper_name': os.path.basename(paper_path),
                'success': False,
                'report_path': None,
                'student_info': {},
                'stats': {},
                'error': None,
            }

            try:
                # 1. 转换文档
                if converter_func:
                    convert_result = converter_func(paper_path, self.output_dir)
                    if not convert_result.get('success'):
                        paper_result['error'] = f"转换失败：{convert_result.get('error')}"
                        results.append(paper_result)
                        summary['failed'] += 1
                        continue

                    md_path = convert_result.get('md_path')

                    # 2. 提取学生信息
                    if extractor_func:
                        paper_result['student_info'] = extractor_func(md_path)

                    # 3. 统计
                    if stats_func:
                        paper_result['stats'] = stats_func(md_path)

                    paper_result['success'] = True
                    paper_result['md_path'] = md_path
                    summary['processed'] += 1

            except FileNotFoundError as e:
                paper_result['error'] = f"文件未找到: {e}"
                print(f"⚠️ 文件未找到: {paper_path}")
                summary['failed'] += 1
            except PermissionError as e:
                paper_result['error'] = f"权限不足: {e}"
                print(f"⚠️ 权限不足: {paper_path}")
                summary['failed'] += 1
            except (OSError, IOError) as e:
                paper_result['error'] = f"IO错误: {e}"
                print(f"⚠️ IO错误 {paper_path}: {e}")
                summary['failed'] += 1
            except Exception as e:
                paper_result['error'] = f"未知错误: {e}"
                print(f"⚠️ 处理失败 {paper_path}: {e}")
                summary['failed'] += 1

            results.append(paper_result)

        # 生成汇总表
        summary_csv = self.generate_summary_csv(results)
        summary['summary_csv'] = summary_csv

        return {
            'success': True,
            'summary': summary,
            'results': results,
        }


def batch_process(folder_path: str, output_dir: str = None) -> Dict:
    """
    便捷函数：批量处理论文文件夹

    Args:
        folder_path: 论文文件夹路径
        output_dir: 输出目录

    Returns:
        批量处理结果
    """
    processor = BatchProcessor(folder_path, output_dir)
    return processor.process_batch()