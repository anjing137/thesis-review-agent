# -*- coding: utf-8 -*-
"""Build traceable evidence from parsed thesis content."""

from __future__ import annotations

import re
from typing import Iterable

from . import criteria


def _split_paragraphs(text: str) -> Iterable[str]:
    for part in re.split(r"\n+", text or ""):
        normalized = re.sub(r"\s+", " ", part).strip()
        if normalized:
            yield normalized


def _section_paths(sections: list[dict]) -> list[tuple[str, str]]:
    stack: dict[int, str] = {}
    rows = []
    for section in sections:
        level = int(section.get("level") or 1)
        title = str(section.get("title") or "未命名章节").strip()
        stack[level] = title
        for deeper in [key for key in stack if key > level]:
            del stack[deeper]
        path = " > ".join(stack[key] for key in sorted(stack))
        for paragraph in _split_paragraphs(section.get("body", "")):
            rows.append((path, paragraph))
    return rows


def _rule_passes(rule: dict, value) -> bool:
    if "equals" in rule and value != rule["equals"]:
        return False
    if "min" in rule and value < rule["min"]:
        return False
    if "max" in rule and value > rule["max"]:
        return False
    return True


def _hard_rule_checks(
    xml_data: dict,
    stats: dict,
    paper_type: str,
    criteria_data: dict,
) -> list[dict]:
    word_count = stats.get("word_count", {})
    refs = stats.get("references", {})
    writing = stats.get("writing_specs", {})
    sections = xml_data.get("sections", [])
    section_titles = " ".join(str(s.get("title", "")) for s in sections)
    body_text = str(xml_data.get("body_text", ""))
    metric_values = {
        "title_length": word_count.get("title", 0),
        "abstract_length": word_count.get("abstract", 0),
        "body_length": word_count.get("body", 0),
        "references_total": refs.get("total", 0),
        "foreign_references": refs.get("foreign", 0),
        "first_person_count": writing.get("first_person_count", 0),
        "robustness_present": bool(
            re.search(
                r"稳健性检验|替换变量|更换样本|更换模型|缩尾处理",
                section_titles + body_text,
            )
        ),
    }

    checks = []
    for rule in criteria.hard_rules(criteria_data):
        applies_to = rule.get("applies_to", ["empirical", "theoretical"])
        if paper_type not in applies_to:
            status = "not_applicable"
            value = None
        elif rule.get("review") == "manual":
            status = "needs_llm_review"
            value = None
        else:
            metric = rule.get("metric")
            if metric not in metric_values:
                raise ValueError(f"硬规则 {rule['id']} 使用未知指标: {metric}")
            value = metric_values[metric]
            status = "pass" if _rule_passes(rule, value) else rule["failed_status"]
        checks.append({
            "id": rule["id"],
            "name": rule["name"],
            "dimension": rule["dimension"],
            "status": status,
            "value": value,
            "requirement": rule["requirement"],
            "spec_ref": rule["spec_ref"],
        })
    return checks


def build_evidence(
    xml_data: dict,
    stats: dict,
    paper_type: str,
    criteria_data: dict | None = None,
) -> dict:
    loaded_criteria = criteria_data or criteria.load_criteria()
    paragraphs = []
    for index, (section, text) in enumerate(_section_paths(xml_data.get("sections", [])), 1):
        paragraphs.append({
            "id": f"P{index:03d}",
            "section": section,
            "text": text,
        })

    references = [
        {"id": f"R{index:03d}", "text": text}
        for index, text in enumerate(_split_paragraphs(xml_data.get("reference_text", "")), 1)
    ]

    abstracts = []
    if xml_data.get("abstract"):
        abstracts.append({"id": "A001", "language": "zh", "text": xml_data["abstract"].strip()})
    if xml_data.get("english_abstract"):
        abstracts.append({"id": "A002", "language": "en", "text": xml_data["english_abstract"].strip()})

    word_count = stats.get("word_count", {})
    refs = stats.get("references", {})
    tables = stats.get("tables", {})
    writing = stats.get("writing_specs", {})
    facts = [
        {"id": "S001", "name": "标题字数", "value": word_count.get("title", 0)},
        {"id": "S002", "name": "摘要字数", "value": word_count.get("abstract", 0)},
        {"id": "S003", "name": "正文字数", "value": word_count.get("body", 0)},
        {"id": "S004", "name": "参考文献总数", "value": refs.get("total", 0)},
        {"id": "S005", "name": "外文文献数", "value": refs.get("foreign", 0)},
        {"id": "S006", "name": "期刊文献数", "value": refs.get("journals", 0)},
        {"id": "S007", "name": "近五年文献数", "value": refs.get("recent", 0)},
        {"id": "S008", "name": "原生表格数", "value": tables.get("native", 0)},
        {"id": "S009", "name": "截图表格数", "value": tables.get("screenshots", 0)},
        {"id": "S010", "name": "第一人称计数", "value": writing.get("first_person_count", 0)},
    ]

    return {
        "paper_type": paper_type,
        "paper": {
            "title": xml_data.get("title", "未知"),
            "student_name": xml_data.get("student_name", "未知"),
            "student_id": xml_data.get("student_id", "未知"),
            "major": xml_data.get("major", "未知"),
            "advisor": xml_data.get("advisor", "未知"),
        },
        "statistics": {
            "word_count": stats.get("word_count", {}),
            "references": stats.get("references", {}),
            "tables": stats.get("tables", {}),
            "images": stats.get("images", {}),
            "writing_specs": stats.get("writing_specs", {}),
        },
        "hard_rule_checks": _hard_rule_checks(xml_data, stats, paper_type, loaded_criteria),
        "facts": facts,
        "abstracts": abstracts,
        "paragraphs": paragraphs,
        "references": references,
    }


def evidence_ids(evidence: dict) -> set[str]:
    ids = set()
    for key in ("facts", "abstracts", "paragraphs", "references"):
        ids.update(item["id"] for item in evidence.get(key, []))
    return ids
