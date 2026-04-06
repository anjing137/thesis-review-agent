#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文评价 Agent - 主入口
支持 .doc / .docx 论文的自动分析与 AI 评价（.doc 会被自动转换为 .docx）。

[废弃说明]
旧版 .md 路径已废弃（marker/extractor/reviewer/stats_from_markdown 等模块）。
converter.py 中的 _convert_doc_to_docx() 被本文件调用以支持 .doc 格式。
其他废弃模块保留在 scripts/ 目录以供备份参考，但不再被调用。
"""

import argparse
import json
import sys
import traceback
from pathlib import Path

from scripts.xml_analyzer import analyze_xml
from scripts.stats import stats_from_xml
import scripts.prompts as prompts
from scripts.converter import Converter


# ============================================================================
# .doc → .docx 预处理
# ============================================================================

def _ensure_docx(paper_path: Path) -> Path:
    """
    如果文件是 .doc，转成临时目录的同名 .docx 并返回新路径。
    如果已是 .docx，直接返回原路径。
    """
    if paper_path.suffix.lower() == ".doc":
        try:
            converter = Converter()
            converted = converter._convert_doc_to_docx(str(paper_path))
            if converted is None:
                raise RuntimeError(
                    "Word COM 接口转换失败。请确保已安装 Microsoft Word。"
                )
            return Path(converted)
        except Exception as e:
            raise RuntimeError(
                f".doc 文件需要转换为 .docx 才能处理。转换失败: {e}"
            )
    return paper_path


# ============================================================================
# 核心数据解析
# ============================================================================

def process_single_paper(paper_path: str, output_dir: str = "./reviews", auto_mode: bool = False) -> dict:
    """
    解析单篇论文（支持 .doc / .docx，.doc 会自动转换为 .docx）。

    参数:
        paper_path: 论文文件路径
        output_dir: 输出目录
        auto_mode: 是否自动标注（调用 marker.py 的 auto_mark_paper）

    返回:
        dict，包含解析结果和统计信息
    """
    paper_path = Path(paper_path)
    if not paper_path.exists():
        return {"success": False, "error": f"文件不存在: {paper_path}"}

    # .doc → .docx 预处理
    paper_path = _ensure_docx(paper_path)

    stem = paper_path.stem  # 文件名（不含扩展名）

    # 1. XML 深度解析
    try:
        xml_data = analyze_xml(str(paper_path))
    except Exception as e:
        return {"success": False, "error": f"XML解析失败: {e}"}

    # 2. 生成统计数据
    try:
        stats_data = stats_from_xml(xml_data)
    except Exception as e:
        return {"success": False, "error": f"统计生成失败: {e}"}

    # 3. 论文标注（可选）
    '''
    if auto_mode:
        try:
            auto_mark_paper(str(paper_path))
        except Exception as e:
            print(f"[警告] 自动标注失败: {e}")
    '''
    
    # 4. 输出统计结果
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    result_file = output_path / f"{stem}_review_data.json"
    with open(result_file, "w", encoding="utf-8") as f:
        json.dump({
            "file": str(paper_path),
            "xml_data": xml_data,
            "stats": stats_data,
        }, f, ensure_ascii=False, indent=2)

    return {
        "success": True,
        "file": str(paper_path),
        "stats": stats_data,
        "result_file": str(result_file),
    }


# ============================================================================
# 评价 Prompt 生成
# ============================================================================

def generate_review_prompt(paper_path: str) -> dict:
    """
    为指定论文生成评价 prompt（支持 .doc / .docx，.doc 会自动转换为 .docx）。

    返回 dict：
        - student_info: 学生信息
        - review_prompt: 提示词文本
        - prompt_file: 保存路径
    """
    paper_path = Path(paper_path)
    if not paper_path.exists():
        raise FileNotFoundError(f"文件不存在: {paper_path}")

    # .doc → .docx 预处理
    paper_path = _ensure_docx(paper_path)

    stem = paper_path.stem

    # 1. 解析论文
    xml_data = analyze_xml(str(paper_path))
    stats = stats_from_xml(xml_data)

    # 2. 学生信息
    student_info = {
        "name": xml_data.get("student_name", "未知"),
        "student_id": xml_data.get("student_id", "未知"),
        "class_name": xml_data.get("major", "未知"),
        "paper_title": xml_data.get("title", "未知"),
        "advisor": xml_data.get("advisor", "未知"),
    }

    # 3. 构建评价维度文本
    dims_text = []
    for i, (dim, data) in enumerate(prompts.DIMENSIONS.items()):
        entry = f"{i+1}. **{dim}（{data['weight']*100:.0f}%）**"
        if data["criteria"]:
            entry += "\n" + "\n".join(f"   - {c}" for c in data["criteria"])
        dims_text.append(entry)
    dims_block = "\n".join(dims_text)

    # 4. 准备论文内容摘要
    _abstract = xml_data.get("abstract", "").strip()
    _body = xml_data.get("body_text", "").strip()
    _references = xml_data.get("reference_text", "").strip()

    abstract_section = f"\n\n## 中文摘要\n{_abstract}\n\n## 英文摘要\n{xml_data.get('english_abstract', '').strip()}" if _abstract else f"\n\n## 中文摘要\n（未找到）\n\n## 英文摘要\n{xml_data.get('english_abstract', '').strip() if xml_data.get('english_abstract') else '（未找到）'}"
    lang_review_section = f"\n\n## 语言与表达审核（摘要）\n{_abstract}\n\n{xml_data.get('english_abstract', '').strip()}"
    body_section = f"\n\n## 正文（前15000字）\n{_body[:15000] if _body else '(未找到正文)'}\n"
    ref_section = f"\n\n## 参考文献\n{_references[:5000] if _references else '(未找到参考文献)'}\n"

    # 5. 统计数字（供 prompt 使用）
    body_char_count = stats.get("word_count", {}).get("body", 0)
    abstract_char_count = stats.get("word_count", {}).get("abstract", 0)
    title_char_count = stats.get("word_count", {}).get("title", 0)

    # 6. 判断论文类型（使用 prompts.py 现成函数，从正文中提取关键词判断）
    body_for_check = (_body or "")[:5000]  # 只取前5000字加速判断
    is_empirical = prompts.is_empirical_paper(body_for_check)

    paper_type = "实证研究论文" if is_empirical else "非实证研究论文"

    # 6. 格式化评分等级
    grade_lines = [f"   {min_s}-{max_s}分：{level}" for min_s, max_s, level in prompts.GRADE_LEVELS]
    grade_block = "\n".join(grade_lines)

    # 7. 格式化一票否决规则
    veto_lines = [f"   - 条件：{r['condition']} → {r['dimension']}：{r['action']}" for r in prompts.VETO_RULES]
    veto_block = "\n".join(veto_lines)

    # 8. 完整提示词
    review_prompt = f"""## 论文信息
- 学生姓名：{student_info['name']}
- 学号：{student_info['student_id']}
- 专业：{student_info['class_name']}
- 论文题目：{student_info['paper_title']}
- 指导教师：{student_info['advisor']}
- 正文字数：{body_char_count} 字
- 摘要字数：{abstract_char_count} 字
- 标题字数：{title_char_count} 字

## 论文类型
{paper_type}

## 内容预览
{abstract_section}{body_section}{ref_section}

## 评价维度（请按以下维度逐项评分）

{dims_block}

{lang_review_section}

### 评分等级
{grade_block}

### 一票否决规则（请在评分前主动检查，满足条件必须执行对应扣分）
{veto_block}

## 输出要求
1. 首先判断论文类型（实证/学理）
2. 对每个维度按 0-100 评分，结合权重计算加权得分
3. **一票否决规则（硬性）**：仅"参考文献<15篇"时触发，该维度直接记0分；从参考文献章节核实数量
4. **实证论文特别关注**（在"方法与论证严谨性"维度中评价，不作为硬性否决）：
   - 稳健性检验是否完整；过度控制问题（如中介变量被纳入控制变量）；内生性处理
5. **特别关注**：
   - 表格是否规范（题目、表头、来源是否完整）
   - 参考文献是否≥15篇（不足则该维度0分）；外文文献是否≥3篇（0篇扣5分，1篇扣3分，2篇扣1分）
   - 文献时效性（近三年占比是否充足）
6. 给出综合评价、得分、等级，以及详细的优缺点和改进建议

## 输出格式（严格按以下Markdown格式输出，不要包含任何JSON或其他格式）

```markdown
# 论文评价报告

## 基本信息
| 项目 | 内容 |
|---|---|
| 学生姓名 | |
| 学号 | |
| 专业 | |
| 论文题目 | |
| 指导教师 | |
| 论文类型 | |
| 正文字数 | |
| 摘要字数 | |
| 标题字数 | |

---

## 一、总体评价
### 1.1 综合评分：XX分（X等）
### 1.2 各维度得分
| 评价维度 | 权重 | 得分 | 关键问题 |
|----------|------|------|----------|
| 选题与研究问题 | 15% | | |
| 参考文献与学术规范 | 15% | | |
| 内容创新性 | 15% | | |
| 框架与逻辑结构 | 20% | | |
| 方法与论证严谨性 | 25% | | |
| 语言与表达 | 10% | | |

### 1.3 评价概述
[100-200字的总体评价]

---

## 二、各维度详细评价

### 2.1 选题与研究问题（XX分）
**优点：**
-

**问题：**
-

### 2.2 参考文献与学术规范（XX分）
**优点：**
-

**问题：**
-

### 2.3 内容创新性（XX分）
**优点：**
-

**问题：**
-

### 2.4 框架与逻辑结构（XX分）
**优点：**
-

**问题：**
-

### 2.5 方法与论证严谨性（XX分）
**优点：**
-

**问题：**
-

### 2.6 语言与表达（XX分）
**优点：**
-

**问题：**
-

---

## 三、修改建议

### 高优先级（必须修改）
1.

### 中优先级（建议修改）
1.

### 低优先级（可选修改）
1.

---

## 四、总结
[50-100字的总结]
```

请开始评价：
"""

    # 7. 保存 prompt
    output_path = Path("./reviews")
    output_path.mkdir(parents=True, exist_ok=True)
    prompt_file = output_path / f"{stem}_review_prompt.md"
    with open(prompt_file, "w", encoding="utf-8") as f:
        f.write(review_prompt)

    # 7b. 同时保存 sidecar JSON（含 xml_data 和 stats），供 --report 复用以避免重复解析
    sidecar_file = output_path / f"{stem}_review_data.json"
    with open(sidecar_file, "w", encoding="utf-8") as f:
        json.dump({
            "paper_path": str(paper_path),
            "xml_data": xml_data,
            "stats": stats,
        }, f, ensure_ascii=False, indent=2)

    return {
        "student_info": student_info,
        "review_prompt": review_prompt,
        "prompt_file": str(prompt_file),
        "sidecar_file": str(sidecar_file),
    }


def generate_review_report(paper_path: str, evaluation_result: str, output_dir: str = "./reviews") -> str:
    """
    将 AI 评价结果（Markdown 文本）保存为报告文件。

    参数:
        paper_path:       论文路径（仅用于生成报告文件名）
        evaluation_result: AI 评价结果（Markdown 文本）
        output_dir:       输出目录
    """
    paper_path = Path(paper_path)
    paper_path = _ensure_docx(paper_path)
    stem = paper_path.stem

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    # evaluation_result 已经是完整 Markdown 报告，直接保存
    report_file = output_path / f"{stem}_评价报告.md"
    with open(report_file, "w", encoding="utf-8") as f:
        f.write(evaluation_result.strip())

    return str(report_file)


# ============================================================================
# 汇总统计（批量模式）
# ============================================================================

def generate_summary_csv(results: list, output_path: str = "./reviews/summary.csv"):
    """生成批量处理的汇总 CSV。"""
    import csv

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "论文文件", "学生姓名", "学号", "专业", "论文题目", "导师",
            "总字符数", "参考文献数", "外文文献数", "表格数", "图片数",
            "评价状态"
        ])
        writer.writeheader()
        for r in results:
            stats = r.get("stats", {})
            xml_in_stats = stats.get("_xml", {})
            refs = stats.get("references", {})
            writer.writerow({
                "论文文件": Path(r.get("file", "")).name,
                "学生姓名": xml_in_stats.get("student_name", "未知"),
                "学号": xml_in_stats.get("student_id", "未知"),
                "专业": xml_in_stats.get("major", "未知"),
                "论文题目": xml_in_stats.get("title", "未知"),
                "导师": xml_in_stats.get("advisor", "未知"),
                "总字符数": stats.get("word_count", {}).get("body", 0),
                "参考文献数": refs.get("total", 0),
                "外文文献数": refs.get("foreign", 0),
                "表格数": stats.get("tables", {}).get("total", 0),
                "图片数": xml_in_stats.get("media_image_count", 0),
                "评价状态": "已解析" if r.get("success") else f"失败: {r.get('error', '未知')}",
            })

    return str(output_path)


# ============================================================================
# 入口
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="论文评价 Agent（支持 .doc / .docx，.doc 会自动转换为 .docx）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  # 解析论文（生成统计数据）
  python main.py 论文.docx -o ./reviews
  python main.py 论文.doc   -o ./reviews   # .doc 自动转换为 .docx

  # 生成评价 prompt（供 AI 评价使用）
  python main.py 论文.docx --prompt -o ./reviews
  python main.py 论文.doc  --prompt -o ./reviews

  # 将 AI 评价结果生成为报告
  python main.py 论文.docx --report -e "AI评价Markdown结果" -o ./reviews

  # 自动标注论文格式问题
  python main.py 论文.docx --auto-mark -o ./reviews

  # 批量处理
  python main.py --batch ./papers/ -o ./reviews
        """
    )
    parser.add_argument("paper", nargs="?", help="论文文件路径（支持 .doc / .docx）")
    parser.add_argument("-o", "--output", default="./reviews", help="输出目录（默认: ./reviews）")
    parser.add_argument("--prompt", action="store_true", help="生成评价 prompt（生成后需人工调用 AI 评价）")
    parser.add_argument("--report", action="store_true", help="将 AI 评价结果生成为报告（需配合 -e 使用）")
    parser.add_argument("-e", "--evaluation", type=str, default="", help="AI 评价结果（Markdown 文本）")
    parser.add_argument("--auto-mark", action="store_true", help="自动标注论文格式问题")
    parser.add_argument("--batch", type=str, help="批量处理目录")
    parser.add_argument("--summary", action="store_true", help="批量处理后生成汇总 CSV（需配合 --batch）")

    args = parser.parse_args()

    # 批量模式
    if args.batch:
        batch_dir = Path(args.batch)
        docx_files = (
            list(batch_dir.glob("*.docx")) + list(batch_dir.glob("*.DOCX"))
            + list(batch_dir.glob("*.doc")) + list(batch_dir.glob("*.DOC"))
        )
        if not docx_files:
            print(f"[错误] 目录中未找到 .doc / .docx 文件: {batch_dir}")
            sys.exit(1)

        results = []
        for f in docx_files:
            print(f"处理: {f.name}")
            result = process_single_paper(str(f), output_dir=args.output, auto_mode=args.auto_mark)
            results.append(result)
            print(f"  -> {'成功' if result['success'] else '失败'}: {result.get('file', '')}")

        print(f"\n批量处理完成: {len(results)} 篇")
        success_count = sum(1 for r in results if r.get("success"))
        print(f"  成功: {success_count}")
        print(f"  失败: {len(results) - success_count}")

        if args.summary:
            csv_path = str(Path(args.output) / "summary.csv")
            generate_summary_csv(results, csv_path)
            print(f"汇总已保存: {csv_path}")
        return

    # 单篇模式
    if not args.paper:
        parser.print_help()
        return

    paper_path = Path(args.paper)
    if not paper_path.exists():
        print(f"[错误] 文件不存在: {paper_path}")
        sys.exit(1)

    if paper_path.suffix.lower() not in (".doc", ".docx"):
        print(f"[错误] 仅支持 .doc 或 .docx 文件，不支持: {paper_path.suffix}")
        sys.exit(1)

    # .doc → .docx 预处理（确保所有模式一致处理）
    paper_path = _ensure_docx(paper_path)

    try:
        if args.report:
            # 模式3：生成报告
            if not args.evaluation:
                print("[错误] --report 模式需要 -e 参数提供 AI 评价结果")
                sys.exit(1)
            report_file = generate_review_report(paper_path, args.evaluation, args.output)
            print(f"报告已生成: {report_file}")

        elif args.prompt:
            # 模式2：生成评价 prompt
            result = generate_review_prompt(paper_path)
            print(f"学生信息: {result['student_info']}")
            print(f"Prompt 已保存: {result['prompt_file']}")
            print("提示: 将 prompt 内容发送给 AI 模型进行评价，将返回的 Markdown 评价结果通过 -e 参数传入 --report 生成报告")

        else:
            # 模式1：解析论文
            result = process_single_paper(paper_path, output_dir=args.output, auto_mode=args.auto_mark)
            if result["success"]:
                stats = result.get("stats", {})
                xml_data_in_stats = stats.get("_xml", {})
                print(f"解析成功: {result['file']}")
                print(f"  学生: {xml_data_in_stats.get('student_name', '?')} ({xml_data_in_stats.get('student_id', '?')})")
                print(f"  专业: {xml_data_in_stats.get('major', '?')}")
                print(f"  题目: {xml_data_in_stats.get('title', '?')}")
                print(f"  导师: {xml_data_in_stats.get('advisor', '?')}")
                print(f"  正文字数: {stats.get('word_count', {}).get('body', '?')} 字")
                refs = stats.get("references", {})
                print(f"  参考文献: {refs.get('total', '?')} 篇（外文 {refs.get('foreign', '?')} 篇）")
                print(f"  表格: {stats.get('tables', {}).get('total', '?')} 个")
                print(f"结果已保存: {result.get('result_file', '')}")
            else:
                print(f"解析失败: {result.get('error', '未知错误')}")
                sys.exit(1)

    except Exception as e:
        print(f"[错误] {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
