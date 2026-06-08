# -*- coding: utf-8 -*-
"""Load and validate the thesis review criteria."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, List

try:
    import yaml
except ImportError as exc:  # pragma: no cover - dependency error is explicit.
    raise RuntimeError("需要安装 PyYAML 才能读取 criteria.yaml") from exc


SKILL_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CRITERIA_PATH = SKILL_ROOT / "criteria.yaml"


def load_criteria(path: str | Path | None = None) -> dict:
    criteria_path = Path(path) if path else DEFAULT_CRITERIA_PATH
    with open(criteria_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    _validate_criteria(data)
    return data


def _validate_criteria(data: dict) -> None:
    if not data or "dimensions" not in data:
        raise ValueError("criteria.yaml 缺少 dimensions")

    ids = [d["id"] for d in data["dimensions"]]
    if len(ids) != len(set(ids)):
        raise ValueError("criteria.yaml 存在重复维度 id")

    for paper_type in ("empirical", "theoretical"):
        weights = get_weights(paper_type, data)
        total = round(sum(weights.values()), 6)
        if total != 1.0:
            raise ValueError(f"{paper_type} 权重合计应为1.0，当前为 {total}")


def dimensions(data: dict | None = None) -> List[dict]:
    return list((data or load_criteria())["dimensions"])


def dimension_map(data: dict | None = None) -> Dict[str, dict]:
    return {d["id"]: d for d in dimensions(data)}


def get_weights(paper_type: str, data: dict | None = None) -> Dict[str, float]:
    if paper_type not in ("empirical", "theoretical"):
        raise ValueError(f"未知论文类型: {paper_type}")
    weight_key = "empirical_weight" if paper_type == "empirical" else "theoretical_weight"
    return {d["id"]: float(d[weight_key]) for d in dimensions(data)}


def required_dimensions(paper_type: str, data: dict | None = None) -> List[str]:
    loaded = data or load_criteria()
    if paper_type not in loaded["paper_types"]:
        raise ValueError(f"未知论文类型: {paper_type}")
    return list(loaded["paper_types"][paper_type]["required_dimensions"])


def paper_type_label(paper_type: str, data: dict | None = None) -> str:
    loaded = data or load_criteria()
    return loaded["paper_types"][paper_type]["label"]


def grade_level(score: float, data: dict | None = None) -> str:
    loaded = data or load_criteria()
    for item in loaded["grade_levels"]:
        if float(item["min"]) <= score <= float(item["max"]):
            return item["label"]
    return "不合格"


def empirical_keywords(data: dict | None = None) -> List[str]:
    loaded = data or load_criteria()
    return list(loaded.get("classification", {}).get("empirical", {}).get("keywords", []))


def classify_paper(body_text: str, data: dict | None = None) -> str:
    loaded = data or load_criteria()
    empirical = loaded.get("classification", {}).get("empirical", {})
    min_hits = int(empirical.get("min_keyword_hits", 3))
    normalized = body_text.lower()
    hits = sum(
        normalized.count(keyword.lower())
        for keyword in empirical.get("keywords", [])
    )
    return "empirical" if hits >= min_hits else "theoretical"


def hard_rules(data: dict | None = None) -> List[dict]:
    return list((data or load_criteria()).get("hard_rules", []))


def active_dimension_rows(paper_type: str, data: dict | None = None) -> List[dict]:
    loaded = data or load_criteria()
    required = set(required_dimensions(paper_type, loaded))
    return [d for d in dimensions(loaded) if d["id"] in required]


def format_dimension_label(dim_id: str, data: dict | None = None) -> str:
    dim = dimension_map(data).get(dim_id)
    if not dim:
        return dim_id
    return f"{dim_id}. {dim['name']}"


def validate_score_payload(scores: dict, paper_type: str, data: dict | None = None) -> None:
    loaded = data or load_criteria()
    required = set(required_dimensions(paper_type, loaded))
    valid = set(dimension_map(loaded))
    provided = set(scores)
    missing = sorted(required - provided)
    unknown = sorted(provided - valid)
    if missing:
        raise ValueError(f"缺少维度分数: {', '.join(missing)}")
    if unknown:
        raise ValueError(f"未知维度分数: {', '.join(unknown)}")
    for dim_id, score in scores.items():
        if not isinstance(score, (int, float)):
            raise ValueError(f"{dim_id} 分数必须是数字")
        if score < 0 or score > 100:
            raise ValueError(f"{dim_id} 分数必须在0-100之间，当前为 {score}")


def inactive_dimensions(paper_type: str, data: dict | None = None) -> Iterable[str]:
    loaded = data or load_criteria()
    required = set(required_dimensions(paper_type, loaded))
    for dim_id in dimension_map(loaded):
        if dim_id not in required:
            yield dim_id
