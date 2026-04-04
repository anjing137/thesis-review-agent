#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文分析器 - 规则检测层

功能：
- 文档格式转换
- 学生信息提取
- 论文类型检测
- 字数统计
- 参考文献分析
- 结构分析
- 写作规范检查
"""

import re
import os
import subprocess
import json
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict
from datetime import datetime

# 导入模型设定分析器
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from model_spec_analyzer import ModelSpecAnalyzer, model_spec_to_dict


@dataclass
class StudentInfo:
    """学生信息"""
    name: str = ""
    student_id: str = ""
    class_name: str = ""
    paper_title: str = ""
    advisor: str = ""


@dataclass
class BasicStats:
    """基础统计数据"""
    word_count: int = 0
    word_count_ok: bool = False
    ref_count: int = 0
    ref_count_ok: bool = False
    foreign_ref_count: int = 0
    foreign_ref_ok: bool = False
    journal_ratio: float = 0.0
    recent_5yr_ratio: float = 0.0


@dataclass
class StructureInfo:
    """结构信息"""
    sections: List[str] = None
    has_required: Dict[str, bool] = None
    issues: List[str] = None

    def __post_init__(self):
        if self.sections is None:
            self.sections = []
        if self.has_required is None:
            self.has_required = {}
        if self.issues is None:
            self.issues = []


@dataclass
class WritingSpecs:
    """写作规范"""
    first_person_count: int = 0
    first_person_locations: List[Dict] = None
    has_software_screenshot: bool = False
    table_count: int = 0
    uses_three_line_table: bool = False
    issues: List[str] = None

    def __post_init__(self):
        if self.first_person_locations is None:
            self.first_person_locations = []
        if self.issues is None:
            self.issues = []


@dataclass
class AnalysisResult:
    """完整分析结果"""
    metadata: Dict = None
    student_info: Dict = None
    basic_stats: Dict = None
    structure: Dict = None
    writing_specs: Dict = None
    abstract: str = ""
    reference_section: str = ""
    model_spec: Dict = None  # 模型设定分析结果

    def __post_init__(self):
        if self.metadata is None:
            self.metadata = {}
        if self.student_info is None:
            self.student_info = {}
        if self.basic_stats is None:
            self.basic_stats = {}
        if self.structure is None:
            self.structure = {}
        if self.writing_specs is None:
            self.writing_specs = {}
        if self.model_spec is None:
            self.model_spec = {}


class PaperAnalyzer:
    """论文分析器"""

    def __init__(self, paper_path: str):
        self.paper_path = paper_path
        self.content = None

    def convert_to_markdown(self) -> Optional[str]:
        """使用 pandoc 将 Word 文档转换为 Markdown"""
        if not os.path.exists(self.paper_path):
            print(f"❌ 文件不存在：{self.paper_path}")
            return None

        md_path = self.paper_path.replace('.docx', '_temp.md')
        try:
            subprocess.run(
                ['pandoc', self.paper_path, '-o', md_path],
                check=True,
                capture_output=True
            )
            with open(md_path, 'r', encoding='utf-8') as f:
                self.content = f.read()
            os.remove(md_path)
            return self.content
        except subprocess.CalledProcessError as e:
            print(f"❌ Pandoc转换失败：{e}")
            return None
        except FileNotFoundError:
            print("❌ 未找到 pandoc，请先安装：brew install pandoc")
            return None

    def extract_student_info(self) -> StudentInfo:
        """从论文内容中提取学生信息"""
        info = StudentInfo()
        if not self.content:
            return info

        lines = self.content.split('\n')

        for line in lines[:50]:
            # 移除星号和空格
            clean_line = re.sub(r'\*+', '', line)
            clean_line = re.sub(r'\s+', '', clean_line)

            # 提取姓名
            if not info.name:
                match = re.search(r'姓名[：:](.+)', clean_line)
                if match and match.group(1):
                    info.name = match.group(1).strip()

            # 提取学号
            if not info.student_id:
                match = re.search(r'学号[：:](\d+)', clean_line)
                if match:
                    info.student_id = match.group(1).strip()

            # 提取专业/班级
            if not info.class_name:
                match = re.search(r'专业[：:](.+)', clean_line)
                if match and match.group(1):
                    info.class_name = match.group(1).strip()
                else:
                    match = re.search(r'班级[：:](.+)', clean_line)
                    if match and match.group(1):
                        info.class_name = match.group(1).strip()

            # 提取指导教师
            if not info.advisor:
                match = re.search(r'指导教师[：:]?(.+)', clean_line)
                if match and match.group(1):
                    info.advisor = match.group(1).strip()

        # 提取论文题目（封面页）
        for i, line in enumerate(lines[:30]):
            if '学年论文' in line or '毕业论文' in line:
                for j in range(i+1, min(i+5, len(lines))):
                    candidate = re.sub(r'\*+', '', lines[j]).strip()
                    # 过滤掉无效内容
                    if candidate and len(candidate) > 10:
                        # 排除常见的非题目内容
                        skip_words = ['学院', '专业', '班级', '学号', '姓名', '指导教师']
                        if not any(sw in candidate for sw in skip_words):
                            info.paper_title = candidate
                            break
                break

        return info

    def detect_paper_type(self) -> str:
        """自动检测论文类型"""
        if not self.content:
            return 'theoretical'

        empirical_keywords = [
            '回归分析', '显著性', '模型', '变量', 'OLS', '固定效应',
            '随机效应', '稳健性检验', '异质性分析', '内生性',
            '面板数据', '双重差分', '断点回归', '工具变量',
            '描述性统计', '相关性分析', 'F检验', '豪斯曼检验',
            '聚类稳健标准误', '实证分析', '计量模型', '回归结果'
        ]

        theoretical_keywords = [
            '现状分析', '问题提出', '原因分析', '对策建议',
            '政策建议', '现状描述', '问题研究', '发展对策',
            '现状及问题', '影响因素', '对策研究'
        ]

        empirical_count = sum(1 for kw in empirical_keywords if kw in self.content)
        theoretical_count = sum(1 for kw in theoretical_keywords if kw in self.content)

        return 'empirical' if empirical_count > theoretical_count else 'theoretical'

    def count_words(self) -> int:
        """统计正文字数（不含摘要、参考文献、致谢）"""
        if not self.content:
            return 0

        content = self.content

        # 1. 移除参考文献部分
        ref_patterns = [
            r'\n##?\s*\*{0,2}参考文献\*{0,2}\s*\n',
            r'\n\[\]\{[^}]+\}\*?\*{0,2}参考文献\*{0,2}\*?\s*\n',
        ]
        for pattern in ref_patterns:
            if re.search(pattern, content):
                content = re.split(pattern, content)[0]
                break

        # 2. 移除致谢部分
        ack_patterns = [
            r'\n##?\s*\*{0,2}致谢\*{0,2}\s*\n',
            r'\n\[\]\{[^}]+\}\*?\*{0,2}致谢\*{0,2}\*?\s*\n',
        ]
        for pattern in ack_patterns:
            if re.search(pattern, content):
                content = re.split(pattern, content)[0]
                break

        # 3. 移除摘要部分
        abstract_match = re.search(r'摘要[：:]?\s*[\s\S]{100,800}', content)
        if abstract_match:
            content = content[abstract_match.end():]

        # 4. 统计中文字符数
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', content))
        return chinese_chars

    def extract_reference_section(self) -> str:
        """提取参考文献章节内容"""
        if not self.content:
            return ""

        patterns = [
            # 模式1: ## 参考文献 或 **参考文献** 或 # 参考文献
            r'##?\s*\*{0,2}参考文献\*{0,2}\s*\n([\s\S]+?)(?:\n##|\Z)',
            # 模式2: []{anchor}**参考文献** 格式（带Markdown anchor）
            r'\[\]\{[^}]+\}.*?\*{0,2}参考文献\*{0,2}\s*\n([\s\S]+?)(?:\n##|\Z)',
            # 模式3: 参考文献作为加粗文本，四级标题后跟列表
            r'\*{0,3}参考文献\*{0,3}\s*\n\s*(?:<!--.*?-->\s*)?([\s\S]+?)(?:\n##|\n#\s|\Z)',
            # 模式4: 简单的"参考文献"后面跟数字开头的行
            r'(?<=\n)\s*参考文献\s*\n([\s\S]+?)(?=\n#\s|\Z)',
        ]

        for pattern in patterns:
            match = re.search(pattern, self.content)
            if match:
                section = match.group(1)
                # 检查是否包含参考文献条目（数字开头）
                if re.search(r'^\s*\d+\.', section, re.MULTILINE):
                    return section

        return ""

    def count_references(self) -> Dict:
        """统计参考文献详细信息"""
        ref_section = self.extract_reference_section()
        if not ref_section:
            return {
                'total': 0, 'foreign': 0, 'journals': 0,
                'books': 0, 'recent_5yr': 0, 'journal_ratio': 0
            }

        lines = ref_section.split('\n')
        total = 0
        foreign = 0
        journals = 0
        books = 0
        recent_5yr = 0
        current_year = 2026

        # 外文文献特征
        foreign_patterns = [
            r'[A-Z][a-z]+,\s*[A-Z]\.',  # Smith, J.
            r'[A-Z][a-z]+\s+[A-Z][a-z]+,',  # John Smith,
            r'[A-Z][a-z]+\s+[A-Z]\s*,',  # Tang C ,
            r'et\s+al',
            r'Journal\s+of', r'Review\s+of', r'International\s+',
            r'Economics?', r'Management', r'Finance',
        ]

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 检查是否是参考文献条目
            if not re.match(r'^\s*\[?\d+\]?[\.\s]', line):
                continue

            # 排除空条目（只有数字和标点）
            cleaned = re.sub(r'[\s\d\.\[\]]', '', line)
            if len(cleaned) < 10:
                continue

            total += 1

            # 检测外文文献
            for pattern in foreign_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    foreign += 1
                    break

            # 检测期刊类[J]
            if re.search(r'\[J\]', line) or re.search(r'\d+\(\d+\)', line):
                journals += 1

            # 检测书籍[M]
            if re.search(r'\[M\]', line):
                books += 1

            # 检测近五年文献
            year_match = re.search(r'(20\d{2})', line)
            if year_match:
                year = int(year_match.group(1))
                if year >= current_year - 5:
                    recent_5yr += 1

        return {
            'total': total,
            'foreign': foreign,
            'journals': journals,
            'books': books,
            'recent_5yr': recent_5yr,
            'journal_ratio': journals / total if total > 0 else 0,
        }

    def analyze_structure(self, paper_type: str) -> StructureInfo:
        """分析论文结构"""
        info = StructureInfo()
        if not self.content:
            return info

        all_titles = []

        # 模式1: # 标题格式
        h1_pattern = r'^#\s+(.+)$'
        h2_pattern = r'^##?\s+(.+)$'
        h1_titles = re.findall(h1_pattern, self.content, re.MULTILINE)
        h2_titles = re.findall(h2_pattern, self.content, re.MULTILINE)
        all_titles.extend(h1_titles)
        all_titles.extend(h2_titles)

        # 模式2: 编号列表格式（如 "1.  []{#anchor}引言" 或 "1. 引言"）
        list_pattern = r'^\s*\d+\.\s+(?:\[.*?\])?\s*\*?\**([^\*]+?)\*?\s*(?:\n|$)'
        list_titles = re.findall(list_pattern, self.content, re.MULTILINE)
        # 过滤掉太短的或明显不是标题的内容
        list_titles = [t.strip() for t in list_titles if len(t.strip()) > 2 and len(t.strip()) < 30]
        all_titles.extend(list_titles)

        # 去重
        all_titles = list(dict.fromkeys(all_titles))
        info.sections = all_titles

        # 实证性论文必要章节
        if paper_type == 'empirical':
            required_sections = {
                '引言': ['引言', '绪论'],
                '文献综述': ['文献综述'],
                '理论分析': ['理论分析'],
                '研究假设': ['研究假设', '假设'],
                '研究设计': ['研究设计'],
                '实证分析': ['实证分析', '实证研究'],
                '结论': ['结论', '结语', '结论与建议']
            }
        else:
            required_sections = {
                '引言': ['引言', '绪论'],
                '现状分析': ['现状'],
                '问题分析': ['问题'],
                '原因分析': ['原因'],
                '对策建议': ['对策', '建议', '政策建议'],
                '结语': ['结语']
            }

        # 检查必要章节（合并所有文本进行搜索）
        all_text = ' '.join(all_titles)
        for key, keywords in required_sections.items():
            info.has_required[key] = any(kw in all_text for kw in keywords)

        # 检查引言三要素
        intro_text = self._extract_intro_text()
        intro_keywords = ['研究背景', '问题提出', '研究目的', '研究意义']
        info.has_required['intro_complete'] = sum(1 for kw in intro_keywords if kw in intro_text) >= 3

        # 检查结构问题
        if any('问题提出' in t for t in h1_titles):
            info.issues.append("'问题提出'不应作为一级标题，应作为'引言'的二级标题")

        if len([t for t in h1_titles if '文献综述' in t]) > 1:
            info.issues.append("文献综述出现多次，可能存在冗余")

        return info

    def _extract_intro_text(self) -> str:
        """提取引言部分文本"""
        patterns = [
            r'#\s*引言\s*\n([\s\S]+?)(?=\n##|\n#\s*文献)',
            r'#\s*绪论\s*\n([\s\S]+?)(?=\n##|\n#\s*文献)',
        ]
        for pattern in patterns:
            match = re.search(pattern, self.content, re.MULTILINE)
            if match:
                return match.group(1)[:2000]  # 只取前2000字
        return ""

    def check_writing_specs(self) -> WritingSpecs:
        """检查基础写作规范"""
        specs = WritingSpecs()
        if not self.content:
            return specs

        # 1. 检测第一人称（注意：本文是正确用法，不应被检测）
        first_person_patterns = [
            r'\b[我我们][认为觉得研究发现]\b',
            r'\b我们[发现在]\b',
            r'\b我[计计算]\b',
            r'\b[我我们][的][\w]+?\b',  # 我的、我们的等所有格形式
        ]

        lines = self.content.split('\n')
        for i, line in enumerate(lines):
            for pattern in first_person_patterns:
                matches = re.finditer(pattern, line)
                for match in matches:
                    specs.first_person_count += 1
                    specs.first_person_locations.append({
                        'line': i + 1,
                        'text': line.strip()[:100]
                    })

        # 2. 检测软件截图
        software_names = ['Stata', 'SPSS', 'EViews', 'Python', 'R语言']
        screenshot_keywords = ['见图', '下图', '如图', '截图', '软件界面', '输出结果']

        ref_section_start = False
        for i, line in enumerate(lines):
            if re.match(r'^\s*##?\s*参考文献', line):
                ref_section_start = True
            if ref_section_start:
                continue

            for software in software_names:
                if software in line:
                    context = ' '.join(lines[max(0, i-2):min(len(lines), i+3)])
                    if any(kw in context for kw in screenshot_keywords):
                        specs.has_software_screenshot = True
                        specs.issues.append(f"第{i+1}行存在{software}截图，应整理为规范表格")

        # 3. 统计表格
        table_pattern = r'\|[^\n]+\|\n\|[-:\s|]+\|\n(?:\|[^\n]+\|\n?)+'
        tables = re.findall(table_pattern, self.content)
        specs.table_count = len(tables)

        return specs

    def extract_abstract(self) -> str:
        """提取摘要"""
        if not self.content:
            return ""

        match = re.search(r'摘要[：:]?\s*([\s\S]{200,800})', self.content)
        if match:
            abstract = match.group(1)
            # 截取到关键词之前
            kw_match = re.search(r'关键词', abstract)
            if kw_match:
                abstract = abstract[:kw_match.start()]
            return abstract.strip()[:500]

        return ""

    def analyze(self) -> AnalysisResult:
        """执行完整分析"""
        print(f"📄 开始分析论文：{self.paper_path}")
        print()

        # 1. 转换文档
        content = self.convert_to_markdown()
        if not content:
            return AnalysisResult()

        print("✅ 文档转换成功")

        # 2. 提取学生信息
        student_info = self.extract_student_info()
        print(f"👤 学生姓名：{student_info.name or '未识别'}")
        print(f"🎓 学号：{student_info.student_id or '未识别'}")
        print(f"📚 班级/专业：{student_info.class_name or '未识别'}")
        print(f"📝 论文题目：{student_info.paper_title or '未识别'}")
        print()

        # 3. 检测论文类型
        paper_type = self.detect_paper_type()
        paper_type_cn = "实证性" if paper_type == 'empirical' else "学理性"
        print(f"📊 检测到论文类型：{paper_type_cn}论文")

        # 4. 统计字数
        word_count = self.count_words()
        print(f"📝 正文字数：{word_count}字 {'✅' if word_count >= 8000 else '❌'}")

        # 5. 统计参考文献
        ref_info = self.count_references()
        print(f"📚 参考文献：{ref_info['total']}篇 {'✅' if ref_info['total'] >= 15 else '❌'}")
        print(f"   外文文献：{ref_info['foreign']}篇 {'✅' if ref_info['foreign'] >= 3 else '❌'}")
        print(f"   期刊占比：{ref_info['journal_ratio']:.0%}")

        # 6. 分析结构
        structure_info = self.analyze_structure(paper_type)
        print(f"📖 章节数量：{len(structure_info.sections)}")

        # 7. 检查写作规范
        writing_specs = self.check_writing_specs()
        print(f"⚠️ 第一人称：{writing_specs.first_person_count}处")
        print(f"📊 表格数量：{writing_specs.table_count}")
        print(f"📋 软件截图：{'有' if writing_specs.has_software_screenshot else '无'}")

        if writing_specs.issues:
            for issue in writing_specs.issues:
                print(f"   - {issue}")

        print()

        # 8. 模型设定分析（仅实证性论文）
        model_spec_result = {}
        if paper_type == 'empirical':
            print("🔍 开始模型设定分析...")
            model_analyzer = ModelSpecAnalyzer(content, paper_type)
            ms_result = model_analyzer.analyze()
            model_spec_result = model_spec_to_dict(ms_result)

            print(f"   提取变量数：{len(ms_result.variables)}")
            print(f"   提取公式数：{len(ms_result.formulas)}")
            print(f"   提取因果链：{len(ms_result.causal_chains)}")

            if ms_result.over_control_issues:
                print(f"   ⚠️ 过度控制问题：{len(ms_result.over_control_issues)}处")
                for issue in ms_result.over_control_issues:
                    print(f"      - {issue.mediator_var}（{issue.severity}）")

            if ms_result.missing_control_issues:
                print(f"   ⚠️ 缺失控制变量：{len(ms_result.missing_control_issues)}处")

            print()

        # 9. 构建结果
        result = AnalysisResult(
            metadata={
                'file_path': self.paper_path,
                'file_name': os.path.basename(self.paper_path),
                'paper_type': paper_type,
                'paper_type_cn': paper_type_cn,
                'analysis_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            student_info=asdict(student_info),
            basic_stats={
                'word_count': word_count,
                'word_count_ok': word_count >= 8000,
                'ref_count': ref_info['total'],
                'ref_count_ok': ref_info['total'] >= 15,
                'foreign_ref_count': ref_info['foreign'],
                'foreign_ref_ok': ref_info['foreign'] >= 3,
                'journal_ratio': ref_info['journal_ratio'],
                'recent_5yr_ratio': ref_info['recent_5yr'] / ref_info['total'] if ref_info['total'] > 0 else 0,
            },
            structure={
                'sections': structure_info.sections,
                'has_required': structure_info.has_required,
                'issues': structure_info.issues,
            },
            writing_specs={
                'first_person_count': writing_specs.first_person_count,
                'first_person_locations': writing_specs.first_person_locations[:10],  # 只保留前10个
                'has_software_screenshot': writing_specs.has_software_screenshot,
                'table_count': writing_specs.table_count,
                'issues': writing_specs.issues,
            },
            abstract=self.extract_abstract(),
            reference_section=self.extract_reference_section()[:2000],  # 限制长度
            model_spec=model_spec_result,
        )

        return result

    def analyze_to_json(self) -> str:
        """分析并返回JSON格式结果"""
        result = self.analyze()
        return json.dumps(asdict(result), ensure_ascii=False, indent=2)


def main():
    """CLI入口"""
    import argparse

    parser = argparse.ArgumentParser(description='论文分析器')
    parser.add_argument('paper_path', help='论文文件路径（.docx格式）')
    parser.add_argument('--output', '-o', help='输出JSON文件路径')
    parser.add_argument('--pretty', '-p', action='store_true', help='美化输出')

    args = parser.parse_args()

    analyzer = PaperAnalyzer(args.paper_path)

    if args.output:
        result = analyzer.analyze()
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(asdict(result), f, ensure_ascii=False, indent=2)
        print(f"✅ 分析结果已保存至：{args.output}")
    else:
        result_json = analyzer.analyze_to_json()
        if args.pretty:
            print(result_json)
        else:
            # 简化输出
            result = analyzer.analyze()
            print(f"\n📊 分析完成")
            print(f"   论文类型：{result.metadata['paper_type_cn']}")
            print(f"   字数：{result.basic_stats['word_count']}字")
            print(f"   参考文献：{result.basic_stats['ref_count']}篇")
            print(f"   外文文献：{result.basic_stats['foreign_ref_count']}篇")


if __name__ == '__main__':
    main()
