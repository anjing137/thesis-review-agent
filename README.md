# Thesis Review Agent

[![Release](https://img.shields.io/github/v/release/anjing137/thesis-review-agent)](https://github.com/anjing137/thesis-review-agent/releases)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

经济学、金融学和管理学本科论文评审 Skill。它将论文解析为可回查证据，由大模型进行内容评价，再由 Python 完成结构校验、硬规则检查、加权计分和正式报告渲染。

当前版本：**v0.6.0**

## 核心能力

- **证据约束评价**：每条主要表现和主要问题必须引用 `S/A/P/R` 证据编号，减少脱离原文的判断。
- **双论文类型**：自动区分实证研究论文与学理论文，使用不同的评价维度和权重。
- **9维度体系**：覆盖选题、文献、理论、设计、实证、稳健性、机制与异质性、结论建议、语言规范。
- **单一规则源**：维度、权重、等级、论文类型识别和硬规则统一维护在 `criteria.yaml`。
- **结构化校验**：拒绝缺失维度、越界分数、未知证据、不适用维度和不可执行的修改建议。
- **Python加权**：大模型只给维度分，Python 负责硬规则否决、加权总分和等级判定。
- **正式双层报告**：正文采用学院评审语体；规范检查与证据编号集中放入附录。

## 工作流程

```text
DOC/DOCX
  -> XML解析与统计
  -> evidence.json
  -> 结构化评审 prompt
  -> review.json
  -> Python校验与加权
  -> 正式 Markdown 评审报告
```

## 安装

```bash
git clone https://github.com/anjing137/thesis-review-agent.git
cd thesis-review-agent
python3 -m pip install -r requirements.txt
```

作为 Codex Skill 使用时，可放入：

```bash
~/.codex/skills/thesis-review-agent
```

`.docx` 可直接解析。处理旧版 `.doc` 文件时，系统还需要可用的 Pandoc、LibreOffice 或 Word 转换环境。

## 快速使用

### 1. 生成证据和评审提示词

```bash
python3 main.py 论文.docx --prompt -o ./reviews
```

输出：

- `*_review_data.json`：完整解析数据
- `*_evidence.json`：统计事实、摘要、正文段落和参考文献证据
- `*_review_prompt.md`：要求模型返回结构化 JSON 的提示词

### 2. 生成正式评审报告

将模型返回的合法 JSON 保存为 `review.json`：

```bash
python3 main.py 论文.docx \
  --report-json ./review.json \
  -o ./reviews
```

输出：

- `*_score.json`：权重、分项得分、总分和否决信息
- `*_evidence.json`：本次报告使用的证据包
- `*_评价报告.md`：正式评审正文及证据核验附录

### 3. 批量解析

```bash
python3 main.py --batch ./papers --summary -o ./reviews
```

### 4. 查看版本

```bash
python3 main.py --version
```

## 正式报告结构

1. 基本信息
2. 总体评价与分项评分
3. 各维度主要表现、主要问题和综合判断
4. 按优先级排列的可执行修改建议
5. 明确的最终评审意见
6. 规范性检查附录
7. 评价证据核验索引

正文不显示 `LLM`、`D1`、`H1`、`P001` 等内部技术标记，便于直接用于教师反馈；证据编号保留在附录中供复核。

## 评价配置

`criteria.yaml` 是唯一规则源，集中维护：

- 实证与学理论文的必需维度
- 各维度权重
- 评分等级
- 论文类型识别关键词及阈值
- 摘要、正文、参考文献、外文文献等硬规则

如需采用其他学院标准，复制并修改该文件后通过 `--criteria` 指定：

```bash
python3 main.py 论文.docx \
  --prompt \
  --criteria ./custom-criteria.yaml \
  -o ./reviews
```

## 项目结构

```text
thesis-review-agent/
├── agents/openai.yaml
├── criteria.yaml
├── main.py
├── SKILL.md
├── VERSION
├── scripts/
│   ├── auto_scorer.py
│   ├── criteria.py
│   ├── evidence.py
│   ├── report_renderer.py
│   ├── review_schema.py
│   ├── stats.py
│   └── xml_analyzer.py
└── tests/test_core.py
```

## 验证

```bash
python3 -m unittest discover -s tests -v
python3 -m compileall -q .
```

## 从 v0.3.1 升级

v0.6.0 是一次工作流升级：

- 原6维度体系调整为按论文类型生效的9维度体系。
- 参考文献数量规范统一为不少于18篇。
- `--report-json` 成为推荐报告生成路径。
- 旧版 `--report -e` 仅作为兼容模式保留，不执行证据校验和 Python 加权。
- 报告输入由自由 Markdown 改为带证据编号的结构化 JSON。

## 隐私说明

论文及评审产物可能包含姓名、学号和未公开研究内容。仓库不内置真实学生论文或真实评审报告，使用者应自行控制 `reviews/` 等输出目录的访问权限。

## License

[MIT](LICENSE)
