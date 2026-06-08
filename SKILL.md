---
name: thesis-review-agent
description: 评审经济学、金融学和管理学本科论文。用于解析DOC/DOCX论文，按实证或学理类型执行9维度评价，生成可回查原文段落和参考文献编号的证据链，由LLM给各维度评分和理由，再由Python校验、加权并输出Markdown评价报告。
---

# Thesis Review Agent

## 核心原则

1. Python只提取事实、校验证据和计算加权总分，不替LLM给分。
2. LLM的每条优点和问题必须引用证据编号：
   - `Sxxx`：Python统计事实
   - `Axxx`：摘要
   - `Pxxx`：正文段落
   - `Rxxx`：参考文献
3. 不得把评价标准中的关键词当作论文事实。
4. 正文提及某方法不等于实际采用该方法，必须结合所在段落判断。
5. Word自动编号可能在XML解析时丢失，不评价编号是否连续。
6. 不估算重复率，不检查字体字号等版式细节。

## 工作流

### 1. 生成评审材料

```bash
python3 main.py 论文.docx --prompt -o ./reviews
```

输出：

- `*_review_data.json`：完整解析数据
- `*_evidence.json`：带稳定编号的证据
- `*_review_prompt.md`：要求LLM返回结构化JSON的提示词

### 2. LLM评审

把 `*_review_prompt.md` 交给当前LLM。LLM必须只返回合法JSON，并为每个维度提供：

- `score`：0-100分
- `strengths`：优点及证据编号
- `issues`：问题及证据编号
- `assessment`：采用正式评审语体撰写的维度综合判断

修改建议必须写明修改位置、具体措施和修改目的，并给出明确的最终评审意见。

保存为 `review.json`。

### 3. 校验、加权和生成报告

```bash
python3 main.py 论文.docx \
  --report-json ./review.json \
  -o ./reviews
```

Python会：

- 检查论文类型是否一致
- 检查必需维度是否齐全
- 拒绝负分、超过100分和未知维度
- 检查每条评价引用的证据编号是否存在
- 从 `criteria.yaml` 读取权重并计算总分
- 输出 `*_score.json` 和正式版 `*_评价报告.md`

正式报告采用双层结构：正文用于教师、学生和学院阅读，不显示内部技术编号；附录保留规范检查与证据核验索引。

## 评价体系

唯一规则源为 `criteria.yaml`：

- 实证论文：D1-D9全部评分，D5实证分析权重最高
- 学理论文：D5-D7不适用，D8结论与对策建议权重最高
- 参考文献数量、摘要字数、正文字数等硬规则也在该文件维护

修改维度、权重、等级或硬规则时，只改 `criteria.yaml`，不要在Python或模板中复制定义。

## 兼容模式

旧版Markdown报告仍可直接保存：

```bash
python3 main.py 论文.docx --report -e "Markdown报告" -o ./reviews
```

该模式不执行结构化证据校验和Python加权，仅用于兼容历史流程。

## 关键文件

- `criteria.yaml`：评价体系唯一规则源
- `main.py`：CLI与工作流编排
- `scripts/xml_analyzer.py`：DOCX解析
- `scripts/evidence.py`：证据编号和硬规则预检查
- `scripts/review_schema.py`：结构化评审校验
- `scripts/auto_scorer.py`：Python加权
- `scripts/report_renderer.py`：Markdown报告渲染

## 验证

```bash
python3 -m unittest discover -s tests -v
python3 /Users/anjing137/.codex/skills/.system/skill-creator/scripts/quick_validate.py .
```
