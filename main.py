#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Thesis review agent CLI.

Pipeline:
1. Parse DOC/DOCX into review_data.json and evidence.json.
2. Generate an evidence-bound prompt for an LLM.
3. Validate structured LLM output.
4. Let Python apply configured weights and render the Markdown report.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
import traceback
from datetime import datetime
from pathlib import Path

from scripts import __version__
from scripts import criteria as criteria_config
from scripts.auto_scorer import Scorer
from scripts.converter import Converter
from scripts.evidence import build_evidence
from scripts.report_renderer import render_report
from scripts.review_schema import validate_review
from scripts.stats import stats_from_xml
from scripts.xml_analyzer import analyze_xml


def _ensure_docx(paper_path: Path) -> Path:
    if paper_path.suffix.lower() != ".doc":
        return paper_path
    try:
        converted = Converter()._convert_doc_to_docx(str(paper_path))
    except Exception as exc:
        raise RuntimeError(f".doc 转 .docx 失败: {exc}") from exc
    if not converted:
        raise RuntimeError(".doc 转 .docx 失败，请确认系统已安装可用的 Word/LibreOffice")
    return Path(converted)


def _validate_paper_path(paper_path: str | Path) -> Path:
    path = Path(paper_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {path}")
    if path.suffix.lower() not in (".doc", ".docx"):
        raise ValueError(f"仅支持 .doc 或 .docx，不支持: {path.suffix}")
    return _ensure_docx(path)


def _analyze(
    paper_path: str | Path,
    criteria_path: str | Path | None = None,
) -> tuple[Path, dict, dict, str, dict]:
    path = _validate_paper_path(paper_path)
    xml_data = analyze_xml(str(path))
    stats = stats_from_xml(xml_data)
    criteria_data = criteria_config.load_criteria(criteria_path)
    paper_type = criteria_config.classify_paper(
        xml_data.get("body_text", ""),
        criteria_data,
    )
    evidence = build_evidence(xml_data, stats, paper_type, criteria_data)
    return path, xml_data, stats, paper_type, evidence


def _dated_name(stem: str, suffix: str) -> str:
    return f"{stem}_{datetime.now().strftime('%Y%m%d')}_{suffix}"


def _write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _automatic_veto_dimensions(evidence: dict, required_dimensions: list[str]) -> list[str]:
    required = set(required_dimensions)
    return list(dict.fromkeys(
        item["dimension"]
        for item in evidence.get("hard_rule_checks", [])
        if item["status"] == "veto" and item["dimension"] in required
    ))


def _format_prompt(criteria_data: dict, evidence: dict) -> str:
    paper_type = evidence["paper_type"]
    dim_lines = []
    for dim in criteria_config.active_dimension_rows(paper_type, criteria_data):
        weight = dim[f"{paper_type}_weight"] * 100
        criteria_lines = "\n".join(f"   - {item}" for item in dim["criteria"])
        dim_lines.append(
            f"{dim['id']}. **{dim['name']}（{weight:g}%）**，规范依据：{dim['spec_ref']}\n"
            f"{criteria_lines}"
        )

    fact_lines = "\n".join(
        f"- [{item['id']}] {item['name']}：{item['value']}"
        for item in evidence.get("facts", [])
    )
    rule_lines = "\n".join(
        f"- {item['id']} {item['name']}：{item['status']}；检测值={item['value']}；"
        f"要求={item['requirement']}"
        for item in evidence.get("hard_rule_checks", [])
    )
    abstract_lines = "\n".join(
        f"[{item['id']}]（{item['language']}）{item['text']}"
        for item in evidence.get("abstracts", [])
    )
    paragraph_lines = "\n".join(
        f"[{item['id']}]（{item['section']}）{item['text']}"
        for item in evidence.get("paragraphs", [])
    )
    reference_lines = "\n".join(
        f"[{item['id']}] {item['text']}"
        for item in evidence.get("references", [])
    )

    required = criteria_config.required_dimensions(paper_type, criteria_data)
    dimension_schema = ",\n".join(
        f'''    "{dim_id}": {{
      "score": 0,
      "strengths": [{{"text": "具体、正式的主要表现", "evidence": ["P001"]}}],
      "issues": [{{"text": "具体、审慎的主要问题", "evidence": ["P002"]}}],
      "assessment": "本维度的综合判断"
    }}'''
        for dim_id in required
    )

    return f"""# 经济与管理类本科论文评审任务

你必须只依据本提示中的论文证据评价，不得凭关键词频次推断论文采用了某种方法。

## 核心约束

1. 论文事实只能来自编号证据：统计事实 Sxxx、摘要 Axxx、正文 Pxxx、参考文献 Rxxx。
2. 评价标准中的方法名称不属于论文证据。正文“提及某方法”也不自动等于“实际采用该方法”。
3. 每个维度给出0-100分，通常提供2-3条主要表现和2-4条主要问题；证据不足时至少各提供1条。
4. 每条优点和问题必须绑定真实证据编号；无法找到证据时不要作出该判断。
5. 每个维度的`assessment`应形成一段完整判断，概括完成度、主要短板及其影响。
6. 使用正式、审慎、可直接用于学院评审的书面语，避免“极差、凑数、完全错误”等情绪化措辞。
7. 修改建议必须分别写明修改位置、具体措施和修改目的，不得只写“进一步完善”。
8. `final_decision`应明确说明是否达到本科毕业论文基本要求，以及是否建议修改后参加答辩。
9. Python负责硬规则否决和加权求和，你不要计算综合分。
10. `veto_dimensions`仅填写需要结合正文人工判定的否决维度，Python预检查已判定的否决不要重复填写。
11. 输出只能是合法JSON，不要添加Markdown代码围栏或解释文字。

## 论文类型

{criteria_config.paper_type_label(paper_type, criteria_data)}（`{paper_type}`）

## 评价维度

{chr(10).join(dim_lines)}

## Python统计事实

{fact_lines}

## 硬规则预检查

{rule_lines}

## 摘要证据

{abstract_lines or "无"}

## 正文证据

{paragraph_lines or "无"}

## 参考文献证据

{reference_lines or "无"}

## 输出JSON结构

{{
  "paper_type": "{paper_type}",
  "dimensions": {{
{dimension_schema}
  }},
  "veto_dimensions": [],
  "overall_evaluation": "采用正式评审语体撰写的总体评语",
  "recommendations": {{
    "high": [{{
      "location": "具体章节或内容位置",
      "action": "必须采取的修改措施",
      "reason": "该修改解决的问题"
    }}],
    "medium": [{{
      "location": "具体章节或内容位置",
      "action": "建议采取的修改措施",
      "reason": "该修改带来的改进"
    }}],
    "low": [{{
      "location": "具体章节或内容位置",
      "action": "完善提升措施",
      "reason": "规范或表达层面的改进目标"
    }}]
  }},
  "summary": "对论文完成度、突出优点和主要不足的正式总结",
  "final_decision": "总体达到本科毕业论文基本要求，建议修改完善后参加答辩。"
}}
"""


def process_single_paper(
    paper_path: str | Path,
    output_dir: str | Path = "./reviews",
    criteria_path: str | Path | None = None,
) -> dict:
    path, xml_data, stats, paper_type, evidence = _analyze(paper_path, criteria_path)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    stem = path.stem

    data_file = output_path / _dated_name(stem, "review_data.json")
    evidence_file = output_path / _dated_name(stem, "evidence.json")
    _write_json(data_file, {
        "paper_path": str(path),
        "paper_type": paper_type,
        "xml_data": xml_data,
        "stats": stats,
    })
    _write_json(evidence_file, evidence)

    return {
        "success": True,
        "file": str(path),
        "paper_type": paper_type,
        "student_info": evidence["paper"],
        "stats": stats,
        "data_file": str(data_file),
        "evidence_file": str(evidence_file),
    }


def generate_review_prompt(
    paper_path: str | Path,
    output_dir: str | Path = "./reviews",
    criteria_path: str | Path | None = None,
) -> dict:
    criteria_data = criteria_config.load_criteria(criteria_path)
    path, xml_data, stats, paper_type, evidence = _analyze(paper_path, criteria_path)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    stem = path.stem

    prompt = _format_prompt(criteria_data, evidence)
    prompt_file = output_path / _dated_name(stem, "review_prompt.md")
    data_file = output_path / _dated_name(stem, "review_data.json")
    evidence_file = output_path / _dated_name(stem, "evidence.json")

    prompt_file.write_text(prompt, encoding="utf-8")
    _write_json(data_file, {
        "paper_path": str(path),
        "paper_type": paper_type,
        "xml_data": xml_data,
        "stats": stats,
    })
    _write_json(evidence_file, evidence)

    return {
        "student_info": evidence["paper"],
        "paper_type": paper_type,
        "prompt_file": str(prompt_file),
        "data_file": str(data_file),
        "evidence_file": str(evidence_file),
    }


def save_markdown_report(
    paper_path: str | Path,
    evaluation_result: str,
    output_dir: str | Path = "./reviews",
) -> str:
    path = _validate_paper_path(paper_path)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    report_file = output_path / _dated_name(path.stem, "评价报告.md")
    report_file.write_text(evaluation_result.strip() + "\n", encoding="utf-8")
    return str(report_file)


def generate_structured_report(
    paper_path: str | Path,
    review_json_path: str | Path,
    output_dir: str | Path = "./reviews",
    criteria_path: str | Path | None = None,
) -> dict:
    criteria_data = criteria_config.load_criteria(criteria_path)
    path, _xml_data, _stats, paper_type, evidence = _analyze(paper_path, criteria_path)
    review = json.loads(Path(review_json_path).read_text(encoding="utf-8"))
    validate_review(review, evidence, criteria_data)

    required_dimensions = criteria_config.required_dimensions(paper_type, criteria_data)
    scores = {
        dim_id: review["dimensions"][dim_id]["score"]
        for dim_id in required_dimensions
    }
    automatic_vetoes = _automatic_veto_dimensions(evidence, required_dimensions)
    review_vetoes = review.get("veto_dimensions", [])
    combined_vetoes = list(dict.fromkeys(automatic_vetoes + review_vetoes))
    score_result = Scorer(criteria_path).score(
        scores,
        paper_type,
        combined_vetoes,
    )
    score_result["automatic_veto_dimensions"] = automatic_vetoes
    score_result["review_veto_dimensions"] = review_vetoes
    report = render_report(review, evidence, score_result, criteria_data)

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    report_file = output_path / _dated_name(path.stem, "评价报告.md")
    score_file = output_path / _dated_name(path.stem, "score.json")
    evidence_file = output_path / _dated_name(path.stem, "evidence.json")
    report_file.write_text(report, encoding="utf-8")
    _write_json(score_file, score_result)
    _write_json(evidence_file, evidence)
    return {
        "report_file": str(report_file),
        "score_file": str(score_file),
        "evidence_file": str(evidence_file),
        "score": score_result,
    }


def generate_summary_csv(results: list[dict], output_path: str | Path) -> str:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "论文文件", "论文类型", "学生姓名", "学号", "专业", "论文题目",
            "正文字数", "参考文献数", "外文文献数", "表格数", "状态",
        ])
        writer.writeheader()
        for result in results:
            stats = result.get("stats", {})
            student_info = result.get("student_info", {})
            refs = stats.get("references", {})
            writer.writerow({
                "论文文件": Path(result.get("file", "")).name,
                "论文类型": result.get("paper_type", ""),
                "学生姓名": student_info.get("student_name", "未知"),
                "学号": student_info.get("student_id", "未知"),
                "专业": student_info.get("major", "未知"),
                "论文题目": student_info.get("title", "未知"),
                "正文字数": stats.get("word_count", {}).get("body", 0),
                "参考文献数": refs.get("total", 0),
                "外文文献数": refs.get("foreign", 0),
                "表格数": stats.get("tables", {}).get("total", 0),
                "状态": "已解析" if result.get("success") else result.get("error", "失败"),
            })
    return str(path)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="经济与管理类本科论文评审：证据提取、结构化评价校验、Python加权",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  python main.py 论文.docx --prompt -o ./reviews
  python main.py 论文.docx --report-json ./review.json -o ./reviews
  python main.py 论文.docx --report -e "旧版Markdown报告" -o ./reviews
  python main.py --batch ./papers --summary -o ./reviews
        """,
    )
    parser.add_argument("paper", nargs="?", help="论文文件路径（.doc/.docx）")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("-o", "--output", default="./reviews", help="输出目录")
    parser.add_argument("--criteria", default=None, help="自定义 criteria.yaml")
    parser.add_argument("--prompt", action="store_true", help="生成数据、证据和结构化评审prompt")
    parser.add_argument("--report-json", help="读取结构化review.json，校验、加权并生成报告")
    parser.add_argument("--report", action="store_true", help="兼容旧版：保存-e提供的Markdown报告")
    parser.add_argument("-e", "--evaluation", default="", help="旧版Markdown评价结果")
    parser.add_argument("--batch", help="批量解析目录")
    parser.add_argument("--summary", action="store_true", help="批量解析后生成summary.csv")
    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    try:
        criteria_config.load_criteria(args.criteria)

        if args.batch:
            batch_dir = Path(args.batch)
            files = sorted(
                list(batch_dir.glob("*.docx"))
                + list(batch_dir.glob("*.DOCX"))
                + list(batch_dir.glob("*.doc"))
                + list(batch_dir.glob("*.DOC"))
            )
            if not files:
                raise FileNotFoundError(f"目录中未找到.doc/.docx: {batch_dir}")
            results = []
            for file in files:
                try:
                    result = process_single_paper(file, args.output, args.criteria)
                    print(f"[OK] {file.name}")
                except Exception as exc:
                    result = {"success": False, "file": str(file), "error": str(exc)}
                    print(f"[失败] {file.name}: {exc}")
                results.append(result)
            if args.summary:
                summary = generate_summary_csv(results, Path(args.output) / "summary.csv")
                print(f"汇总已保存: {summary}")
            return

        if not args.paper:
            parser.print_help()
            return

        if args.report_json:
            result = generate_structured_report(
                args.paper, args.report_json, args.output, args.criteria
            )
            print(f"报告已生成: {result['report_file']}")
            print(f"评分明细: {result['score_file']}")
        elif args.report:
            if not args.evaluation:
                raise ValueError("--report 需要-e提供Markdown评价结果")
            report_file = save_markdown_report(args.paper, args.evaluation, args.output)
            print(f"报告已生成: {report_file}")
        elif args.prompt:
            result = generate_review_prompt(args.paper, args.output, args.criteria)
            print(f"论文类型: {result['paper_type']}")
            print(f"Prompt: {result['prompt_file']}")
            print(f"数据: {result['data_file']}")
            print(f"证据: {result['evidence_file']}")
        else:
            result = process_single_paper(args.paper, args.output, args.criteria)
            print(f"解析成功: {result['file']}")
            print(f"论文类型: {result['paper_type']}")
            print(f"数据: {result['data_file']}")
            print(f"证据: {result['evidence_file']}")
    except Exception as exc:
        print(f"[错误] {exc}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
