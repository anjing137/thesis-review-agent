# -*- coding: utf-8 -*-
"""Prompt helpers backed by criteria.yaml."""

from __future__ import annotations

from . import criteria


_CRITERIA = criteria.load_criteria()

# Compatibility exports for callers that still import these names.
DIMENSIONS = {
    dim["id"]: {
        "name": dim["name"],
        "empirical_weight": dim["empirical_weight"],
        "theoretical_weight": dim["theoretical_weight"],
        "criteria": dim["criteria"],
        "spec_ref": dim["spec_ref"],
    }
    for dim in criteria.dimensions(_CRITERIA)
}
GRADE_LEVELS = [
    (float(item["min"]), float(item["max"]), item["label"])
    for item in _CRITERIA["grade_levels"]
]
VETO_RULES = [
    rule for rule in criteria.hard_rules(_CRITERIA) if rule.get("failed_status") == "veto"
]
EMPIRICAL_KEYWORDS = criteria.empirical_keywords(_CRITERIA)


def get_grade_level(score: float) -> str:
    return criteria.grade_level(score, _CRITERIA)


def is_empirical_paper(body_content: str) -> bool:
    return criteria.classify_paper(body_content or "", _CRITERIA) == "empirical"


def check_veto_rules(stats: dict, is_empirical: bool = False) -> list:
    triggered = []
    refs = stats.get("references", {})
    reference_rule = next(rule for rule in VETO_RULES if rule["id"] == "H4")
    if refs.get("total", 0) < reference_rule["min"]:
        triggered.append({
            "rule": "H4",
            "dimension": reference_rule["dimension"],
            "action": "参考文献与学术规范维度记0分",
        })
    return triggered
