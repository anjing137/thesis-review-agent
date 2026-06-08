import copy
import unittest

from scripts import criteria
from scripts.auto_scorer import Scorer
from scripts.evidence import build_evidence
from scripts.report_renderer import render_report
from scripts.review_schema import validate_review
from main import _automatic_veto_dimensions


class CoreTests(unittest.TestCase):
    def setUp(self):
        self.criteria = criteria.load_criteria()
        self.xml_data = {
            "title": "数字金融影响研究",
            "student_name": "测试学生",
            "student_id": "20260001",
            "major": "金融学",
            "advisor": "测试导师",
            "abstract": "摘要内容",
            "english_abstract": "Abstract text.",
            "body_text": "研究背景。\n稳健性检验采用替换变量方法。",
            "reference_text": "[1] 测试文献",
            "sections": [
                {"level": 1, "title": "引言", "body": "研究背景。"},
                {"level": 1, "title": "稳健性检验", "body": "采用替换变量方法。"},
            ],
        }
        self.stats = {
            "word_count": {"title": 8, "abstract": 320, "body": 9000},
            "references": {
                "total": 20,
                "foreign": 3,
                "journals": 18,
                "recent": 12,
            },
            "tables": {"total": 2, "native": 2, "screenshots": 0},
            "images": {},
            "writing_specs": {"first_person_count": 0},
        }
        self.evidence = build_evidence(self.xml_data, self.stats, "empirical")

    def _review(self):
        dimensions = {}
        for dim_id in criteria.required_dimensions("empirical", self.criteria):
            dimensions[dim_id] = {
                "score": 80,
                "strengths": [{"text": "存在明确证据", "evidence": ["P001"]}],
                "issues": [{"text": "仍可进一步完善", "evidence": ["P002"]}],
                "assessment": "本维度基本达到本科论文要求，但仍需针对上述问题修改。",
            }
        return {
            "paper_type": "empirical",
            "dimensions": dimensions,
            "veto_dimensions": [],
            "overall_evaluation": "总体达到本科论文基本要求。",
            "recommendations": {
                "high": [],
                "medium": [{
                    "location": "研究设计章节",
                    "action": "补充变量选择依据",
                    "reason": "增强模型设定与理论分析的一致性",
                }],
                "low": [],
            },
            "summary": "建议根据证据继续修改。",
            "final_decision": "总体达到本科毕业论文基本要求，建议修改完善后参加答辩。",
        }

    def test_weights_sum_to_one(self):
        self.assertAlmostEqual(sum(criteria.get_weights("empirical", self.criteria).values()), 1)
        self.assertAlmostEqual(sum(criteria.get_weights("theoretical", self.criteria).values()), 1)

    def test_scorer_rejects_invalid_score(self):
        scores = {dim_id: 80 for dim_id in criteria.required_dimensions("empirical", self.criteria)}
        scores["D1"] = 101
        with self.assertRaises(ValueError):
            Scorer().score(scores, "empirical")

    def test_scorer_rejects_missing_dimension(self):
        with self.assertRaises(ValueError):
            Scorer().score({"D1": 80}, "empirical")

    def test_decimal_grade_boundaries(self):
        scores = {dim_id: 89.5 for dim_id in criteria.required_dimensions("empirical", self.criteria)}
        result = Scorer().score(scores, "empirical")
        self.assertEqual(result["grade"], "良好上（优）")

    def test_review_requires_real_evidence(self):
        review = self._review()
        review["dimensions"]["D1"]["issues"][0]["evidence"] = ["P999"]
        with self.assertRaises(ValueError):
            validate_review(review, self.evidence, self.criteria)

    def test_review_rejects_inactive_or_unknown_dimension(self):
        review = self._review()
        review["dimensions"]["D10"] = review["dimensions"]["D1"]
        with self.assertRaises(ValueError):
            validate_review(review, self.evidence, self.criteria)

    def test_review_requires_actionable_recommendation(self):
        review = self._review()
        del review["recommendations"]["medium"][0]["location"]
        with self.assertRaises(ValueError):
            validate_review(review, self.evidence, self.criteria)

    def test_hard_rule_threshold_comes_from_criteria(self):
        custom = copy.deepcopy(self.criteria)
        h4 = next(rule for rule in custom["hard_rules"] if rule["id"] == "H4")
        h4["min"] = 25
        evidence = build_evidence(self.xml_data, self.stats, "empirical", custom)
        check = next(item for item in evidence["hard_rule_checks"] if item["id"] == "H4")
        self.assertEqual(check["status"], "veto")
        self.assertEqual(
            _automatic_veto_dimensions(
                evidence,
                criteria.required_dimensions("empirical", custom),
            ),
            ["D2"],
        )

    def test_paper_classification_comes_from_criteria(self):
        custom = copy.deepcopy(self.criteria)
        custom["classification"]["empirical"]["min_keyword_hits"] = 4
        self.assertEqual(criteria.classify_paper("OLS OLS OLS", custom), "theoretical")
        self.assertEqual(criteria.classify_paper("OLS OLS OLS OLS", custom), "empirical")

    def test_theoretical_branch_excludes_empirical_dimensions(self):
        evidence = build_evidence(
            self.xml_data,
            self.stats,
            "theoretical",
            self.criteria,
        )
        dimensions = {}
        for dim_id in criteria.required_dimensions("theoretical", self.criteria):
            dimensions[dim_id] = {
                "score": 80,
                "strengths": [{"text": "存在明确证据", "evidence": ["P001"]}],
                "issues": [{"text": "仍可进一步完善", "evidence": ["P002"]}],
                "assessment": "本维度基本达到本科论文要求。",
            }
        review = {
            "paper_type": "theoretical",
            "dimensions": dimensions,
            "veto_dimensions": [],
            "overall_evaluation": "总体达到本科论文基本要求。",
            "recommendations": {"high": [], "medium": [], "low": []},
            "summary": "建议继续修改。",
            "final_decision": "总体达到本科毕业论文基本要求，建议修改完善后参加答辩。",
        }
        validate_review(review, evidence, self.criteria)
        result = Scorer().score(
            {dim_id: item["score"] for dim_id, item in dimensions.items()},
            "theoretical",
        )
        report = render_report(review, evidence, result, self.criteria)
        self.assertEqual(result["total_score"], 80)
        self.assertNotIn("#### 实证分析", report)
        h7 = next(item for item in evidence["hard_rule_checks"] if item["id"] == "H7")
        self.assertEqual(h7["status"], "not_applicable")

    def test_report_uses_formal_body_and_evidence_appendix(self):
        review = self._review()
        validate_review(review, self.evidence, self.criteria)
        scores = {dim_id: item["score"] for dim_id, item in review["dimensions"].items()}
        score_result = Scorer().score(scores, "empirical")
        report = render_report(review, self.evidence, score_result, self.criteria)
        formal_body, appendix = report.split("## 附录一", 1)
        self.assertIn("# 经济与管理类本科毕业论文评审报告", formal_body)
        self.assertIn("#### 选题与研究问题（80分）", formal_body)
        self.assertIn("**综合判断：**", formal_body)
        self.assertIn("**评审意见：", formal_body)
        self.assertNotIn("LLM", formal_body)
        self.assertNotIn("D1.", formal_body)
        self.assertNotIn("P001", formal_body)
        self.assertIn("P001（引言）", appendix)


if __name__ == "__main__":
    unittest.main()
