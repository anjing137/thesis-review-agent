# -*- coding: utf-8 -*-
"""Validate structured LLM review output."""

from __future__ import annotations

from . import criteria
from .evidence import evidence_ids


def _validate_evidence_items(items: list, label: str, valid_ids: set[str]) -> None:
    if not isinstance(items, list) or not items:
        raise ValueError(f"{label} 至少需要1条")
    for index, item in enumerate(items, 1):
        if not isinstance(item, dict) or not str(item.get("text", "")).strip():
            raise ValueError(f"{label} 第{index}条缺少 text")
        refs = item.get("evidence")
        if not isinstance(refs, list) or not refs:
            raise ValueError(f"{label} 第{index}条至少需要1个 evidence 编号")
        unknown = sorted(set(refs) - valid_ids)
        if unknown:
            raise ValueError(f"{label} 第{index}条引用未知证据: {', '.join(unknown)}")


def validate_review(review: dict, evidence: dict, criteria_data: dict | None = None) -> None:
    loaded = criteria_data or criteria.load_criteria()
    paper_type = review.get("paper_type")
    if paper_type != evidence.get("paper_type"):
        raise ValueError(
            f"评审论文类型 {paper_type!r} 与解析结果 {evidence.get('paper_type')!r} 不一致"
        )

    dimensions = review.get("dimensions")
    if not isinstance(dimensions, dict):
        raise ValueError("review.json 缺少 dimensions 对象")

    scores = {}
    valid_ids = evidence_ids(evidence)
    required_dimensions = set(criteria.required_dimensions(paper_type, loaded))
    unexpected_dimensions = sorted(set(dimensions) - required_dimensions)
    if unexpected_dimensions:
        raise ValueError(
            f"包含不适用或未知维度: {', '.join(unexpected_dimensions)}"
        )
    for dim_id in required_dimensions:
        dim_review = dimensions.get(dim_id)
        if not isinstance(dim_review, dict):
            raise ValueError(f"缺少 {dim_id} 评价")
        scores[dim_id] = dim_review.get("score")
        _validate_evidence_items(dim_review.get("strengths"), f"{dim_id}.strengths", valid_ids)
        _validate_evidence_items(dim_review.get("issues"), f"{dim_id}.issues", valid_ids)
        if not str(dim_review.get("assessment", "")).strip():
            raise ValueError(f"{dim_id} 缺少 assessment 综合判断")

    criteria.validate_score_payload(scores, paper_type, loaded)

    for field in ("overall_evaluation", "summary", "final_decision"):
        if not str(review.get(field, "")).strip():
            raise ValueError(f"review.json 缺少 {field}")

    recommendations = review.get("recommendations")
    if not isinstance(recommendations, dict):
        raise ValueError("review.json 缺少 recommendations")
    for priority in ("high", "medium", "low"):
        items = recommendations.get(priority)
        if not isinstance(items, list):
            raise ValueError(f"recommendations.{priority} 必须是列表")
        for index, item in enumerate(items, 1):
            if not isinstance(item, dict):
                raise ValueError(
                    f"recommendations.{priority} 第{index}条必须是对象"
                )
            for field in ("location", "action", "reason"):
                if not str(item.get(field, "")).strip():
                    raise ValueError(
                        f"recommendations.{priority} 第{index}条缺少 {field}"
                    )

    veto_dimensions = review.get("veto_dimensions", [])
    if not isinstance(veto_dimensions, list):
        raise ValueError("veto_dimensions 必须是列表")
    invalid_vetoes = sorted(set(veto_dimensions) - required_dimensions)
    if invalid_vetoes:
        raise ValueError(
            f"veto_dimensions 包含不适用或未知维度: {', '.join(invalid_vetoes)}"
        )
