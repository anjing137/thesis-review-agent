# -*- coding: utf-8 -*-
"""Render a formal, evidence-auditable thesis review as Markdown."""

from __future__ import annotations

from datetime import date

from . import criteria


GROUPS = [
    ("研究基础", ["D1", "D2", "D3"]),
    ("研究实施", ["D4", "D5", "D6", "D7"]),
    ("结论与表达", ["D8", "D9"]),
]


def _format_rule_value(value) -> str:
    if value is None:
        return "-"
    if isinstance(value, bool):
        return "是" if value else "否"
    return str(value)


def _table_text(value: str) -> str:
    return " ".join(str(value).split()).replace("|", "\\|")


def _evidence_location_map(evidence: dict) -> dict[str, str]:
    locations = {}
    for item in evidence.get("facts", []):
        locations[item["id"]] = f"统计信息：{item['name']}"
    for item in evidence.get("abstracts", []):
        language = "中文摘要" if item["language"] == "zh" else "英文摘要"
        locations[item["id"]] = language
    for item in evidence.get("paragraphs", []):
        locations[item["id"]] = item["section"]
    for item in evidence.get("references", []):
        locations[item["id"]] = "参考文献表"
    return locations


def _format_recommendation(item: dict) -> str:
    return (
        f"**修改位置：{item['location']}。** "
        f"{item['action']}。修改目的：{item['reason']}。"
    )


def render_report(review: dict, evidence: dict, score_result: dict, criteria_data: dict) -> str:
    paper = evidence["paper"]
    stats = evidence["statistics"]
    word_count = stats.get("word_count", {})
    dim_map = criteria.dimension_map(criteria_data)
    weights = criteria.get_weights(review["paper_type"], criteria_data)
    required_order = criteria.required_dimensions(review["paper_type"], criteria_data)
    required = set(required_order)
    today = date.today()

    lines = [
        "# 经济与管理类本科毕业论文评审报告",
        "",
        "## 基本信息",
        "",
        "| 项目 | 内容 |",
        "|---|---|",
        f"| 学生姓名 | {paper['student_name']} |",
        f"| 学号 | {paper['student_id']} |",
        f"| 专业 | {paper['major']} |",
        f"| 论文题目 | {paper['title']} |",
        f"| 指导教师 | {paper['advisor']} |",
        f"| 论文类型 | {score_result['paper_type_label']} |",
        f"| 正文字数 | {word_count.get('body', 0)}字 |",
        f"| 摘要字数 | {word_count.get('abstract', 0)}字 |",
        f"| 评审日期 | {today.year}年{today.month}月{today.day}日 |",
        "",
        "---",
        "",
        "## 一、总体评价",
        "",
        f"### 1.1 综合评定：{score_result['total_score']}分（{score_result['grade']}）",
        "",
        "### 1.2 分项评价",
        "",
        "| 评价维度 | 权重 | 得分 | 主要评价 |",
        "|---|---:|---:|---|",
    ]

    for dim_id in required_order:
        dim = dim_map[dim_id]
        issue = review["dimensions"][dim_id]["issues"][0]["text"]
        lines.append(
            f"| {dim['name']} | {weights[dim_id] * 100:g}% "
            f"| {score_result['dimension_scores'][dim_id]:g} "
            f"| {_table_text(issue)} |"
        )

    lines.extend([
        "",
        "### 1.3 总体评语",
        "",
        review["overall_evaluation"].strip(),
        "",
        "## 二、分项评价",
        "",
    ])

    section_number = 1
    for group_name, dim_ids in GROUPS:
        active = [dim_id for dim_id in dim_ids if dim_id in required]
        if not active:
            continue
        lines.extend([f"### 2.{section_number} {group_name}", ""])
        section_number += 1
        for dim_id in active:
            dim_review = review["dimensions"][dim_id]
            lines.append(
                f"#### {dim_map[dim_id]['name']}（{dim_review['score']:g}分）"
            )
            lines.extend(["", "**主要表现：**"])
            for item in dim_review["strengths"]:
                lines.append(f"- {item['text']}")
            lines.extend(["", "**主要问题：**"])
            for item in dim_review["issues"]:
                lines.append(f"- {item['text']}")
            lines.extend([
                "",
                f"**综合判断：** {dim_review['assessment'].strip()}",
                "",
            ])

    lines.extend(["## 三、修改建议", ""])
    priority_labels = {
        "high": "高优先级（必须修改）",
        "medium": "中优先级（建议修改）",
        "low": "低优先级（完善提升）",
    }
    for priority in ("high", "medium", "low"):
        lines.append(f"### {priority_labels[priority]}")
        items = review["recommendations"][priority]
        if items:
            for index, item in enumerate(items, 1):
                lines.append(f"{index}. {_format_recommendation(item)}")
        else:
            lines.append("无。")
        lines.append("")

    lines.extend([
        "## 四、评审结论",
        "",
        review["summary"].strip(),
        "",
        f"**评审意见：{review['final_decision'].strip()}**",
        "",
        "---",
        "",
        "## 附录一：规范性检查",
        "",
        "| 检查项目 | 检查结果 | 检测值 | 规范要求 |",
        "|---|---|---|---|",
    ])

    status_labels = {
        "pass": "符合要求",
        "warning": "需修改完善",
        "veto": "未达到硬性要求",
        "not_applicable": "不适用",
        "needs_llm_review": "需结合正文复核",
    }
    for item in evidence.get("hard_rule_checks", []):
        lines.append(
            f"| {item['name']} | {status_labels.get(item['status'], item['status'])} "
            f"| {_format_rule_value(item.get('value'))} | {item['requirement']} |"
        )

    locations = _evidence_location_map(evidence)
    lines.extend([
        "",
        "## 附录二：评价证据核验索引",
        "",
        "> 本附录用于复核评价依据，不作为面向学生的正文评语。",
        "",
        "| 评价维度 | 类型 | 评价要点 | 证据编号及位置 |",
        "|---|---|---|---|",
    ])
    for dim_id in required_order:
        dim_name = dim_map[dim_id]["name"]
        for item_type, key in (("主要表现", "strengths"), ("主要问题", "issues")):
            for item in review["dimensions"][dim_id][key]:
                refs = "；".join(
                    f"{ref}（{locations.get(ref, '位置未知')}）"
                    for ref in item["evidence"]
                )
                lines.append(
                    f"| {dim_name} | {item_type} | {_table_text(item['text'])} "
                    f"| {_table_text(refs)} |"
                )

    lines.append("")
    return "\n".join(lines)
