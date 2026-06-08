#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Score thesis reviews from LLM dimension scores.

The LLM assigns each active dimension a 0-100 score. Python only validates
those scores, applies explicit veto dimensions, and computes the weighted sum.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    from . import criteria
except ImportError:  # Allow direct CLI execution.
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from scripts import criteria


class Scorer:
    def __init__(self, criteria_path: str | Path | None = None):
        self.criteria = criteria.load_criteria(criteria_path)

    def score(
        self,
        llm_scores: dict,
        paper_type: str,
        veto_dimensions: list[str] | None = None,
    ) -> dict:
        criteria.validate_score_payload(llm_scores, paper_type, self.criteria)
        weights = criteria.get_weights(paper_type, self.criteria)
        dim_info = criteria.dimension_map(self.criteria)
        veto_dimensions = veto_dimensions or []
        invalid_veto = sorted(set(veto_dimensions) - set(dim_info))
        if invalid_veto:
            raise ValueError(f"未知一票否决维度: {', '.join(invalid_veto)}")

        adjusted_scores = {k: float(v) for k, v in llm_scores.items()}
        veto_applied = []
        for dim_id in veto_dimensions:
            if dim_id in adjusted_scores:
                veto_applied.append({
                    "dimension": dim_id,
                    "dimension_name": dim_info[dim_id]["name"],
                    "original_score": adjusted_scores[dim_id],
                })
                adjusted_scores[dim_id] = 0.0

        weighted = {}
        total = 0.0
        for dim_id, weight in weights.items():
            raw_score = adjusted_scores.get(dim_id, 0.0)
            contribution = raw_score * weight
            weighted[dim_id] = round(contribution, 2)
            total += contribution

        total = round(total, 2)
        return {
            "paper_type": paper_type,
            "paper_type_label": criteria.paper_type_label(paper_type, self.criteria),
            "total_score": total,
            "grade": criteria.grade_level(total, self.criteria),
            "dimension_scores": adjusted_scores,
            "dimension_weighted": weighted,
            "weights_used": weights,
            "veto_applied": veto_applied,
        }


def _parse_json(value: str, label: str):
    try:
        return json.loads(value)
    except json.JSONDecodeError as exc:
        raise SystemExit(f"{label} 不是合法 JSON: {exc}") from exc


def main() -> None:
    parser = argparse.ArgumentParser(description="LLM给分，Python按论文类型加权求和")
    parser.add_argument("--llm-scores", required=True, help='JSON，如 {"D1":85,"D2":90}')
    parser.add_argument("--paper-type", choices=("empirical", "theoretical"), required=True)
    parser.add_argument("--veto-dimensions", default="[]", help='JSON列表，如 ["D6"]')
    parser.add_argument("--criteria", default=None, help="criteria.yaml 路径")
    parser.add_argument("-o", "--output", help="输出 JSON 路径")
    args = parser.parse_args()

    llm_scores = _parse_json(args.llm_scores, "--llm-scores")
    veto_dimensions = _parse_json(args.veto_dimensions, "--veto-dimensions")
    result = Scorer(args.criteria).score(llm_scores, args.paper_type, veto_dimensions)
    payload = json.dumps(result, ensure_ascii=False, indent=2)

    if args.output:
        Path(args.output).parent.mkdir(parents=True, exist_ok=True)
        Path(args.output).write_text(payload, encoding="utf-8")
        print(f"[OK] 评分结果已保存: {args.output}")
    else:
        print(payload)


if __name__ == "__main__":
    main()
