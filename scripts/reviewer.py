# -*- coding: utf-8 -*-
"""
论文AI评价模块 - 基于统计数据和论文内容进行结构化评价

[DEPRECATED] 此模块已废弃，仅用于 .md 旧路径的向后兼容。
新流程（.docx 直接分析）请使用 main.py 中的 generate_review_prompt()。
"""

import json
import re
from pathlib import Path
from typing import Dict, Any, Optional


class Reviewer:
    """论文AI评价器"""

    def __init__(self, marked_md_path: str, stats_path: Optional[str] = None):
        """
        初始化评价器

        Args:
            marked_md_path: 标记好的markdown文件路径
            stats_path: 统计数据JSON文件路径（可选）
        """
        self.marked_md_path = Path(marked_md_path)
        self.stats_path = Path(stats_path) if stats_path else None

        # 读取内容
        self.content = self.marked_md_path.read_text(encoding='utf-8')
        self.stats = self._load_stats()

        # 提取各部分内容
        self.abstract = self._extract_section('摘要')
        self.body = self._extract_section('正文')
        self.references = self._extract_section('参考文献')

    def _load_stats(self) -> Dict[str, Any]:
        """加载统计数据"""
        if self.stats_path and self.stats_path.exists():
            return json.loads(self.stats_path.read_text(encoding='utf-8'))
        return {}

    def _extract_section(self, section_name: str) -> str:
        """
        从标记好的markdown中提取指定部分

        Args:
            section_name: 部分名称（摘要/正文/参考文献）

        Returns:
            提取的内容，如果没有找到则返回空字符串
        """
        start_tag = f"<!-- {section_name}开始 -->"
        end_tag = f"<!-- {section_name}结束 -->"

        start_idx = self.content.find(start_tag)
        end_idx = self.content.find(end_tag)

        if start_idx != -1 and end_idx != -1:
            # 内容在开始标记之后，结束标记之前
            start_idx += len(start_tag)
            return self.content[start_idx:end_idx].strip()

        return ""

    def extract_student_info(self) -> Dict[str, str]:
        """
        从markdown内容中提取学生信息

        Returns:
            学生信息字典
        """
        info = {}

        # 从文件名提取（文件名格式：专业+姓名+学号+题目+导师.docx）
        filename = self.marked_md_path.stem
        parts = filename.split('+')

        if len(parts) >= 5:
            info['class_name'] = parts[0]  # 专业
            info['name'] = parts[1]  # 姓名
            info['student_id'] = parts[2]  # 学号
            info['paper_title'] = parts[3]  # 论文题目
            info['advisor'] = parts[4]  # 指导教师

        return info

    def get_paper_type(self) -> str:
        """
        判断论文类型

        Returns:
            '实证性' 或 '学理性'
        """
        # 简单的关键词判断
        empirical_keywords = ['实证', '回归', '数据', '模型', '变量', '检验', '分析']
        body_lower = self.body.lower()

        empirical_score = sum(1 for kw in empirical_keywords if kw in body_lower)

        if empirical_score >= 2:
            return '实证性'
        return '学理性'

    def build_review_prompt(self, student_info: Dict[str, str]) -> str:
        """
        构建评价提示词

        Args:
            student_info: 学生信息字典

        Returns:
            完整的提示词
        """
        # 获取统计数据
        wc = self.stats.get('word_count', {})
        ref = self.stats.get('references', {})
        ws = self.stats.get('writing_specs', {})

        # 处理references字段名不一致问题（stats.py用journals，reviewer.py期望journal）
        ref_total = ref.get('total', 0)
        ref_journal = ref.get('journal', ref.get('journals', 0))
        ref_book = ref.get('book', ref.get('books', 0))

        # 判断论文类型
        paper_type = self.get_paper_type()

        prompt = f"""# 论文评价任务

## 基本信息
- 学生姓名：{student_info.get('name', '未知')}
- 学号：{student_info.get('student_id', '未知')}
- 专业：{student_info.get('class_name', '未知')}
- 论文题目：{student_info.get('paper_title', '未知')}
- 指导教师：{student_info.get('advisor', '未知')}
- 论文类型：{paper_type}

## 统计数据（Python精准统计，请以此为准）
- 标题字数：{wc.get('title', 0)}字
- 摘要字数：{wc.get('abstract', 0)}字
- 正文字数：{wc.get('body', 0)}字
- 参考文献总数：{ref_total}篇
- 外文参考文献：{ref.get('foreign', 0)}篇
- 期刊文献：{ref_journal}篇
- 书籍文献：{ref_book}篇
- 第一人称使用次数：{ws.get('first_person_count', 0)}处

## 论文内容

### 摘要
{self.abstract if self.abstract else '(未找到摘要)'}

### 正文（部分）
{self.body[:3000] if self.body else '(未找到正文)'}...

### 参考文献（部分）
{self.references[:2000] if self.references else '(未找到参考文献)'}...

---

## 评价要求

请基于以上信息，对论文进行6维度结构化评价。

### 评价维度（权重）

1. **选题与研究问题（15%）**
   - 题目是否具体明确（标题字数≤25字为佳）
   - 研究必要性是否充分论证
   - 研究价值和创新点是否明确

2. **参考文献与学术规范（15%）**
   - 数量是否≥15篇（硬性要求，不足则该维度0分）
   - 外文文献是否≥3篇
   - 期刊占比是否≥2/3
   - 格式是否规范

3. **内容创新性（15%）**
   - 研究视角是否新颖
   - 是否有文献评述（指出已有研究不足）
   - 是否明确边际贡献

4. **框架与逻辑结构（20%）**
   - 实证论文标准结构：引言→文献综述→理论分析→研究设计→实证分析→结论
   - 学理论文标准结构：引言→现状→问题→原因→对策→结语
   - 各章节逻辑是否清晰

5. **方法与论证严谨性（25%）**
   - 理论分析是否深入
   - 数据质量和变量定义是否清晰
   - **稳健性检验是否完整（实证论文强制要求，否则该维度0分）**
   - 是否讨论内生性问题

6. **语言与表达（10%）**
   - 是否使用第三人称（"本文"正确，"我/我们"错误）
   - 术语是否专业规范
   - 标点符号是否正确

### 一票否决项
- 参考文献<15篇 → 参考文献维度0分
- 实证论文无稳健性检验 → 方法与论证维度0分
- 实证论文存在明显过度控制 → 方法与论证维度扣5-10分

### 评分等级
- 90-100：优秀（优+）
- 85-89：良好上（优）
- 80-84：良好中（良+）
- 75-79：良好下（良）
- 70-74：中等上（中+）
- 65-69：中等（中）
- 60-64：中等下（中-）
- <60：不合格

---

## 输出格式

请按以下格式输出评价报告：

```
# 论文评价报告

## 基本信息
:[基本信息表格]

## 一、总体评价
### 1.1 综合评分：XX分（X等）
### 1.2 各维度得分
| 评价维度 | 权重 | 得分 | 关键问题 |
|---------|------|------|---------|
| 选题与研究问题 | 15% | | |
| 参考文献与学术规范 | 15% | | |
| 内容创新性 | 15% | | |
| 框架与逻辑结构 | 20% | | |
| 方法与论证严谨性 | 25% | | |
| 语言与表达 | 10% | | |

### 1.3 评价概述
:[100-200字的总体评价]

## 二、各维度详细评价

### 2.1 选题与研究问题（XX分）
**优点：**
- [列出优点，引用原文]

**问题：**
- [列出问题，引用原文]

### 2.2 参考文献与学术规范（XX分）
**优点：**
- [列出优点]

**问题：**
- [列出问题]

[继续其他维度...]

## 三、修改建议

### 高优先级（必须修改）
1. [建议1]
2. [建议2]

### 中优先级（建议修改）
1. [建议1]
2. [建议2]

### 低优先级（可选修改）
1. [建议1]

## 四、总结
:[50-100字的总结]
```

请开始评价：
"""
        return prompt

    def review(self, model_api_func=None) -> Dict[str, Any]:
        """
        执行论文评价

        Args:
            model_api_func: 模型API调用函数，签名为 func(prompt: str) -> str
                          如果为None，则返回提示词供外部使用

        Returns:
            评价结果字典，包含:
            - report: 评价报告文本
            - scores: 各维度得分
            - summary: 简要总结
        """
        student_info = self.extract_student_info()
        prompt = self.build_review_prompt(student_info)

        if model_api_func is None:
            # 返回提示词，不执行实际评价
            return {
                'prompt': prompt,
                'student_info': student_info,
                'stats': self.stats,
                'ready_for_evaluation': True
            }

        # 调用模型API
        report = model_api_func(prompt)

        # 解析得分（简单解析）
        scores = self._parse_scores(report)

        return {
            'report': report,
            'scores': scores,
            'student_info': student_info,
            'stats': self.stats
        }

    def _parse_scores(self, report: str) -> Dict[str, float]:
        """
        从评价报告中解析各维度得分

        Args:
            report: 评价报告文本

        Returns:
            各维度得分字典
        """
        scores = {}

        # 提取总分
        total_match = re.search(r'综合评分[：:]\s*(\d+)', report)
        if total_match:
            scores['total'] = int(total_match.group(1))

        # 提取各维度得分
        dimension_names = [
            '选题与研究问题',
            '参考文献与学术规范',
            '内容创新性',
            '框架与逻辑结构',
            '方法与论证严谨性',
            '语言与表达'
        ]

        for dim in dimension_names:
            # 匹配形如 "选题与研究问题（XX分）" 或 "选题与研究问题 | 15% | XX |"
            pattern = f'{dim}[（(](\\d+)'
            match = re.search(pattern, report)
            if match:
                scores[dim] = int(match.group(1))

        return scores


def load_reviewer_from_json(json_path: str) -> Reviewer:
    """
    从处理结果JSON加载Reviewer

    Args:
        json_path: 处理结果JSON文件路径

    Returns:
        Reviewer实例
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    marked_md = data.get('marked_md_path')
    stats = data.get('stats', {})

    if not marked_md:
        raise ValueError("JSON中未找到marked_md_path")

    reviewer = Reviewer(marked_md)
    return reviewer