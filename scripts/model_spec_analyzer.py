#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模型设定分析器 - 提取并检测实证论文的模型设定

功能：
- 从论文中提取变量定义（被解释变量、解释变量、控制变量、中介变量）
- 提取因果链声明
- 检测"过度控制"（over-control）问题
- 检测变量角色混淆、遗漏控制变量等问题
"""

import re
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Optional, Set
from enum import Enum


class VariableRole(Enum):
    """变量角色分类"""
    DEPENDENT = "dependent"           # 被解释变量Y
    INDEPENDENT = "independent"        # 核心解释变量X
    MEDIATOR = "mediator"             # 中介变量M
    CONTROL = "control"               # 控制变量
    MODERATOR = "moderator"           # 调节变量
    UNKNOWN = "unknown"


class VariableType(Enum):
    """变量类型分类"""
    CONTINUOUS = "continuous"          # 连续变量
    DISCRETE = "discrete"             # 离散变量
    BINARY = "binary"                # 二值变量
    COUNT = "count"                   # 计数变量
    RATIO = "ratio"                   # 比例变量


@dataclass
class Variable:
    """单个变量定义"""
    name_en: str                      # 英文名 e.g., "innovation"
    name_cn: str                      # 中文名 e.g., "专利申请量"
    role: VariableRole = VariableRole.UNKNOWN
    var_type: VariableType = VariableType.CONTINUOUS
    definition_text: Optional[str] = None
    measurement: Optional[str] = None     # e.g., "企业年度专利申请数量"
    table_location: Optional[str] = None  # e.g., "表3-1"


@dataclass
class CausalChain:
    """因果链声明"""
    statement: str                    # 完整语句
    cause_var: str                   # 原因变量
    effect_var: str                  # 结果变量
    mediator_vars: List[str] = field(default_factory=list)  # 中介变量列表
    is_direct_effect: bool = False   # 是否为直接效应
    line_number: Optional[int] = None


@dataclass
class ModelFormula:
    """回归模型公式"""
    formula_text: str                # 完整公式文本
    dependent_var: str               # 被解释变量
    independent_vars: List[str] = field(default_factory=list)  # 解释变量
    control_vars: List[str] = field(default_factory=list)     # 控制变量
    mediator_vars: List[str] = field(default_factory=list)   # 中介变量
    table_location: Optional[str] = None
    model_type: Optional[str] = None  # e.g., "固定效应模型", "双向固定效应"
    section: Optional[str] = None     # 所在章节


@dataclass
class OverControlIssue:
    """过度控制问题"""
    mediator_var: str                 # 被过度控制的中介变量
    causal_chain: str                 # 相关因果链声明
    controlled_in_model: str          # 在哪个模型中被控制
    explanation: str                  # 解释
    severity: str = "high"            # high, medium, low


@dataclass
class VariableRoleIssue:
    """变量角色问题"""
    variable: str
    declared_role: str
    actual_role: str
    evidence: str
    severity: str = "medium"


@dataclass
class EndogeneityCheck:
    """内生性处理检查"""
    has_endogeneity: bool
    method_used: Optional[str]        # e.g., "IV-2SLS", "GMM", "滞后变量"
    is_adequate: bool
    issues: List[str] = field(default_factory=list)


@dataclass
class ModelSpecResult:
    """完整的模型设定分析结果"""
    # 提取的信息
    variables: List[Variable] = field(default_factory=list)
    causal_chains: List[CausalChain] = field(default_factory=list)
    formulas: List[ModelFormula] = field(default_factory=list)

    # 检测到的问题
    over_control_issues: List[OverControlIssue] = field(default_factory=list)
    variable_role_issues: List[VariableRoleIssue] = field(default_factory=list)
    missing_control_issues: List[str] = field(default_factory=list)

    # 验证结果
    endogeneity_check: Optional[EndogeneityCheck] = None

    # 评分（0-100）
    spec_completeness: int = 0        # 模型设定完整性
    spec_accuracy: int = 0           # 变量角色准确性
    causal_logic: int = 0            # 因果逻辑一致性

    # 原始提取的文本段落
    variable_definition_section: str = ""
    model_specification_section: str = ""
    hypothesis_section: str = ""

    # 元数据
    is_empirical: bool = True
    extraction_confidence: float = 0.0  # 0.0-1.0


class ModelSpecAnalyzer:
    """
    模型设定分析器

    从实证论文中提取并验证：
    - 变量定义和角色
    - 回归模型公式
    - 因果链声明
    - 过度控制问题
    - 内生性处理
    """

    # 已知中介变量关键词（按研究领域）
    KNOWN_MEDIATORS: Set[str] = {
        # 创新研究领域
        'rd', '研发投入', '研发支出', '研发强度', 'rd_investment',
        'innovation_input', '创新投入', '研发费用',
        'patent', '专利', '专利申请', '专利授权',
        'tech_progress', '技术进步', '技术水平',

        # 金融领域
        'financing', '融资约束', '融资', '资金约束', '融资难度',
        'cash_flow', '现金流', '经营性现金流', 'cashflow',
        'investment_efficiency', '投资效率', '投资水平',

        # 公司治理领域
        'roe', 'roa', '盈利能力', '利润', '经营绩效', '绩效',
        'board', '董事会', '公司治理', '治理结构', '股权集中度',
        'management', '管理层', '管理者',

        # 其他可能的
        '人力资本', '员工结构', '员工数量', 'employment',
    }

    # 常见被解释变量关键词
    COMMON_DEPENDENT_KEYWORDS: Set[str] = {
        'innovation', '创新', '专利', '新产品', '新产品收入',
        'performance', '绩效', '业绩', 'roa', 'roe', 'tfp',
        'risk', '风险', '债务', '违约', '信用风险',
        'investment', '投资', '研发', '研发投入',
        'growth', '增长', '营收增长', '营业收入增长',
        'value', '企业价值', '市值', '托宾Q',
    }

    # 典型控制变量关键词
    TYPICAL_CONTROL_KEYWORDS: Set[str] = {
        'size', '规模', '公司规模', 'asset', '总资产', 'ln_size',
        'age', '年龄', '企业年龄', '成立年限', 'firm_age',
        'leverage', '杠杆', '资产负债率', 'debt_ratio', 'lev',
        'tangibility', '有形资产', '固定资产', 'ppe',
        'cash', '现金', '现金持有', 'cash_holding',
        'growth', '增长', '营收增长', 'sales_growth',
        'year', '年份', '年度', 'time', 'year_fe',
        'industry', '行业', 'industry_fe', 'indutry',
        'roa', '盈利能力', '利润率',
    }

    # 因果链检测模式（优化：减少噪音）
    CAUSAL_PATTERNS: List[re.Pattern] = [
        # 假设声明模式（最可靠）: H1: X对Y有显著影响, H1: X显著提升Y
        re.compile(r'[Hh](\d+)\s*[:：]\s*([^:：]+?)\s*(?:对|影响|提升|降低|促进|抑制)?\s*([^:：\n]+?)(?:显著|正向|负向|具有)?'),
        # 中介路径（可靠）: A通过B影响C, A经由B作用于C
        re.compile(r'([^\s]+?)\s*(?:通过|经由|借助)\s*([^\s]+?)\s*(?:影响|作用于|促进|抑制)\s*([^\s，。；]+?)'),
    ]

    # 噪音词列表（匹配到这些词时过滤）
    NOISE_PATTERNS = [
        '研究', '分析', '本文', '本文', '论文', '理论', '方法', '数据', '结果',
        '影响', '作用', '关系', '相关', '效应', '结论', '发现', '表明', '认为',
        '发展', '推进', '促进', '完善', '提高', '加强', '深化', '拓展', '丰富',
        '全球', '国家', '社会', '经济', '市场', '企业', '居民', '消费', '数字',
    ]

    def __init__(self, content: str, paper_type: str = "empirical"):
        """
        初始化分析器

        Args:
            content: 论文全文（Markdown格式）
            paper_type: "empirical" or "theoretical"
        """
        self.content = content
        self.paper_type = paper_type
        self.variables: List[Variable] = []
        self.causal_chains: List[CausalChain] = []
        self.formulas: List[ModelFormula] = []
        self.over_control_issues: List[OverControlIssue] = []
        self.variable_role_issues: List[VariableRoleIssue] = []
        self.endogeneity_check: Optional[EndogeneityCheck] = None

    def analyze(self) -> ModelSpecResult:
        """
        执行完整的模型设定分析

        注意：对于实证性论文，本模块仅标记需要AI深度分析的信号，
        具体的变量提取、因果链识别、过度控制检测由AI在评价阶段完成。

        Returns:
            ModelSpecResult，包含元数据和需要AI分析的内容
        """
        result = ModelSpecResult()

        if self.paper_type != "empirical":
            result.is_empirical = False
            result.extraction_confidence = 1.0
            result.spec_completeness = 100
            result.spec_accuracy = 100
            result.causal_logic = 100
            return result

        # 对于实证性论文，标记为需要AI深度分析
        result.is_empirical = True
        result.spec_completeness = 0  # AI需要完成
        result.spec_accuracy = 0
        result.causal_logic = 0
        result.extraction_confidence = 0.5  # 需要AI判断

        # 检测内生性处理（这个可以用关键词简单检测）
        result.endogeneity_check = self.verify_endogeneity_handling()
        self.endogeneity_check = result.endogeneity_check

        return result

    def extract_variables(self) -> List[Variable]:
        """
        从论文中提取变量定义

        查找：
        - Markdown表格：变量 | 定义 / Variable | Definition
        - 文本模式：如"X定义为..." / "X is defined as..."
        - 变量列表

        Returns:
            Variable对象列表
        """
        variables = []
        var_table = self._find_section("变量定义")

        if var_table:
            table_vars = self._parse_variable_table(var_table)
            variables.extend(table_vars)

        # 如果没找到表格，尝试从文本中提取
        if not variables:
            text_vars = self._extract_variables_from_text()
            variables.extend(text_vars)

        # 分类变量角色
        self._classify_variable_roles(variables)

        return variables

    def _parse_variable_table(self, table_text: str) -> List[Variable]:
        """解析变量定义表格

        支持多种格式：
        - 标准markdown表格：| 变量 | 定义 |
        - 中文期刊表格：| 变量类型 | 变量名称 | 变量符号 | 变量定义 |
        - 多列表格（3列以上）
        """
        variables = []

        # 检测是否为多列表格（中文期刊格式）
        # 格式：| 变量类型 | 变量名称 | 变量符号 | 变量定义 | 量纲 |
        header_pattern = r'变量类型.*?变量名称.*?变量符号|变量.*?符号.*?定义|被解释变量.*?解释变量'
        is_multi_col = re.search(header_pattern, table_text)

        if is_multi_col:
            # 多列格式：提取符号列作为变量名
            rows = re.findall(r'\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|', table_text)
            for row in rows:
                if len(row) < 3:
                    continue
                var_type = row[0].strip()  # 变量类型（被解释/控制）
                var_name = row[1].strip()  # 变量名称
                var_symbol = row[2].strip()  # 变量符号

                # 跳过表头行
                if any(kw in var_name for kw in ['变量', '类型', '名称', '符号', 'Variable', 'Type']):
                    continue

                if len(var_name) < 1 and len(var_symbol) < 1:
                    continue

                # 确定变量角色
                role = VariableRole.UNKNOWN
                if '被解释' in var_type or '因变量' in var_type:
                    role = VariableRole.DEPENDENT
                elif '核心' in var_type or '自变量' in var_type or '解释' in var_type:
                    role = VariableRole.INDEPENDENT
                elif '控制' in var_type:
                    role = VariableRole.CONTROL
                elif '中介' in var_type:
                    role = VariableRole.MEDIATOR

                var = Variable(
                    name_en=self._normalize_var_name(var_symbol or var_name),
                    name_cn=var_symbol or var_name,
                    role=role,
                    definition_text=row[3].strip() if len(row) > 3 else '',
                    measurement=var_name,
                )
                variables.append(var)
        else:
            # 标准双列表格
            rows = re.findall(r'\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|', table_text)
            for row in rows[1:]:  # 跳过表头
                if len(row) < 2:
                    continue
                name_col = row[0].strip()
                def_col = row[1].strip()

                if len(name_col) < 2 or len(def_col) < 2:
                    continue
                if name_col in ['变量', 'Variable', '名称', '符号']:
                    continue

                var = Variable(
                    name_en=self._normalize_var_name(name_col),
                    name_cn=name_col,
                    definition_text=def_col,
                    measurement=def_col,
                )
                variables.append(var)

        return variables

    def _extract_variables_from_text(self) -> List[Variable]:
        """从文本中提取变量"""
        variables = []

        # 变量定义模式
        patterns = [
            r'([A-Za-z_]\w*)\s*[=:：]\s*([^，,。\n]+)',  # var = 定义
            r'([A-Za-z_]\w*)\s+[是为定义]\s+([^，,。\n]+)',  # var 是/为 定义
        ]

        for pattern in patterns:
            matches = re.finditer(pattern, self.content)
            for match in matches:
                name = match.group(1).strip()
                definition = match.group(2).strip()

                # 过滤太短的或明显不是变量名的
                if len(name) < 2 or len(definition) < 3:
                    continue

                var = Variable(
                    name_en=name,
                    name_cn=name,
                    definition_text=definition,
                )
                variables.append(var)

        return variables

    def _classify_variable_roles(self, variables: List[Variable]) -> None:
        """根据变量名和上下文分类变量角色"""
        # 收集因果链中的变量
        causal_vars = set()
        mediator_vars = set()
        for chain in self.causal_chains:
            causal_vars.add(chain.cause_var)
            causal_vars.add(chain.effect_var)
            mediator_vars.update(chain.mediator_vars)

        for var in variables:
            var_name_en = var.name_en
            var_name_lower = var.name_en.lower()
            var_name_cn = var.name_cn

            # 1. 检查是否在因果链中声明为中介
            if var_name_en in mediator_vars or any(m.lower() in var_name_lower for m in mediator_vars):
                var.role = VariableRole.MEDIATOR
                continue

            # 2. 检查是否为常见被解释变量
            if any(kw in var_name_lower or kw in var_name_cn for kw in self.COMMON_DEPENDENT_KEYWORDS):
                # 如果同时也是因果链的因变量，确认是被解释变量
                if var_name_en in causal_vars or any(var_name_en in c.cause_var for c in self.causal_chains):
                    var.role = VariableRole.DEPENDENT
                continue

            # 3. 检查是否为已知的中介变量关键词
            if var_name_lower in self.KNOWN_MEDIATORS or any(kw in var_name_lower for kw in ['研发', '融资', '现金', '人力']):
                # 需要结合因果链判断
                if var_name_en in mediator_vars:
                    var.role = VariableRole.MEDIATOR
                continue

            # 4. 检查是否为典型控制变量
            if var_name_lower in self.TYPICAL_CONTROL_KEYWORDS or any(kw in var_name_lower for kw in ['size', 'age', 'lev', 'roa']):
                var.role = VariableRole.CONTROL
                continue

    def extract_causal_chains(self) -> List[CausalChain]:
        """
        从理论/假设部分提取因果链声明

        Returns:
            CausalChain对象列表
        """
        chains = []

        # 从假设部分提取
        hypothesis_section = self._find_section("研究假设")
        if hypothesis_section:
            chains.extend(self._extract_chains_from_text(hypothesis_section))

        # 从理论分析部分提取
        theory_section = self._find_section("理论分析")
        if theory_section:
            chains.extend(self._extract_chains_from_text(theory_section))

        # 从中介效应检验部分提取
        mediator_section = self._find_section("中介效应")
        if mediator_section:
            chains.extend(self._extract_mediator_chains(mediator_section))

        # 去重
        unique_chains = []
        seen = set()
        for chain in chains:
            key = (chain.cause_var, chain.effect_var, tuple(chain.mediator_vars))
            if key not in seen:
                seen.add(key)
                unique_chains.append(chain)

        return unique_chains

    def _extract_chains_from_text(self, text: str) -> List[CausalChain]:
        """从文本中提取因果链（带噪音过滤）"""
        chains = []

        for pattern in self.CAUSAL_PATTERNS:
            matches = pattern.finditer(text)
            for match in matches:
                groups = match.groups()

                if len(groups) >= 2:
                    cause_var = groups[0].strip() if groups[0] else ""
                    effect_var = groups[-1].strip() if groups[-1] else ""

                    # 过滤噪音：检查是否包含噪音词
                    if self._is_noise_pair(cause_var, effect_var):
                        continue

                    # 过滤太短或太长的变量名
                    if len(cause_var) < 2 or len(effect_var) < 2:
                        continue
                    if len(cause_var) > 20 or len(effect_var) > 20:
                        continue

                    # 构建因果链
                    if len(groups) >= 3 and groups[1]:
                        mediator = groups[1].strip()
                        if mediator and len(mediator) >= 2 and len(mediator) < 20:
                            chain = CausalChain(
                                statement=match.group(0),
                                cause_var=cause_var,
                                effect_var=effect_var,
                                mediator_vars=[mediator],
                                is_direct_effect=False,
                            )
                        else:
                            chain = CausalChain(
                                statement=match.group(0),
                                cause_var=cause_var,
                                effect_var=effect_var,
                                is_direct_effect=True,
                            )
                    else:
                        chain = CausalChain(
                            statement=match.group(0),
                            cause_var=cause_var,
                            effect_var=effect_var,
                            is_direct_effect=True,
                        )
                    chains.append(chain)

        return chains

    def _is_noise_pair(self, cause: str, effect: str) -> bool:
        """检查因果变量对是否为噪音"""
        if not cause or not effect:
            return True

        cause_lower = cause.lower()
        effect_lower = effect.lower()

        # 如果任一变量包含噪音词，可能不是有效因果变量
        for noise in self.NOISE_PATTERNS:
            if noise in cause_lower or noise in effect_lower:
                # 但如果是假设H1/H2格式，豁免部分噪音
                if cause.startswith('H') or cause.startswith('h'):
                    return False
                return True

        return False

    def _extract_mediator_chains(self, text: str) -> List[CausalChain]:
        """从中介效应检验部分提取因果链"""
        chains = []

        # 提取"X通过M影响Y"模式
        pattern = re.compile(r'([\w\u4e00-\u9fff]+)\s*(?:通过|经由)\s*([\w\u4e00-\u9fff]+)\s*(?:影响|作用于)\s*([\w\u4e00-\u9fff]+)')
        matches = pattern.finditer(text)

        for match in matches:
            chain = CausalChain(
                statement=match.group(0),
                cause_var=match.group(1),
                effect_var=match.group(3),
                mediator_vars=[match.group(2)],
                is_direct_effect=False,
            )
            chains.append(chain)

        return chains

    def extract_model_formulas(self) -> List[ModelFormula]:
        """
        提取回归模型公式

        Returns:
            ModelFormula对象列表
        """
        formulas = []
        model_section = self._find_section("模型设定")

        if not model_section:
            # 尝试在整个内容中查找
            model_section = self.content

        # 匹配回归公式
        # 模式1: Y = β₀ + β₁·X₁ + ...
        eq_patterns = [
            r'([A-Za-z_]\w*)\s*=\s*β₀?\s*[\+]?\s*(?:β\d?[·•]?\s*)?([A-Za-z_]\w*)',
            r'innovation\s*=\s*β',  # 特殊：包含innovation的公式
        ]

        for pattern_str in eq_patterns:
            pattern = re.compile(pattern_str)
            matches = pattern.finditer(model_section)

            for match in matches:
                formula = ModelFormula(
                    formula_text=match.group(0),
                    dependent_var=match.group(1) if len(match.groups()) >= 1 else "",
                )
                formulas.append(formula)

        # 尝试从文本中提取控制变量
        control_var_pattern = r'(?:控制变量|Control variables?)[:：]\s*([^\n]+)'
        control_matches = re.finditer(control_var_pattern, model_section, re.IGNORECASE)
        for cmatch in control_matches:
            control_text = cmatch.group(1)
            # 分割并清理控制变量
            c_vars = re.split(r'[,，、]', control_text)
            for f in formulas:
                for v in c_vars:
                    v_clean = v.strip()
                    if v_clean and len(v_clean) < 20:
                        if v_clean not in f.control_vars:
                            f.control_vars.append(v_clean)

        # 如果找到了变量定义，也检查控制变量列表
        if self.variables:
            for var in self.variables:
                if var.role == VariableRole.CONTROL:
                    for f in formulas:
                        if var.name_en not in f.control_vars and var.name_cn not in f.control_vars:
                            f.control_vars.append(var.name_en)

        return formulas

    def detect_over_control(self) -> List[OverControlIssue]:
        """
        检测过度控制问题

        当中介变量被作为控制变量放入模型时，会阻断中介路径，
        导致对直接效应的估计偏误。

        Returns:
            OverControlIssue列表
        """
        issues = []

        # 1. 找出所有声明的中介变量
        mediators_in_chains = set()
        for chain in self.causal_chains:
            mediators_in_chains.update(chain.mediator_vars)

        if not mediators_in_chains:
            return issues

        # 2. 检查每个公式
        for formula in self.formulas:
            control_vars_lower = [v.lower() for v in formula.control_vars]

            for mediator in mediators_in_chains:
                mediator_lower = mediator.lower()

                # 检查中介变量是否在控制变量中
                if any(mediator_lower in ctrl or ctrl in mediator_lower
                       for ctrl in control_vars_lower):
                    # 找到了过度控制问题
                    issue = OverControlIssue(
                        mediator_var=mediator,
                        causal_chain=self._get_causal_chain_text(mediator),
                        controlled_in_model=formula.formula_text or "主回归模型",
                        severity="high",
                        explanation=(
                            f"变量'{mediator}'在因果链中作为中介变量被识别，"
                            f"但在研究直接效应时被作为控制变量控制。\n"
                            f"这会阻断中介路径，导致对直接效应的估计偏误。\n"
                            f"建议：\n"
                            f"1. 从主回归中移除该变量，估计总效应\n"
                            f"2. 或单独进行中介效应分析（Bootstrap检验）"
                        )
                    )
                    issues.append(issue)

        return issues

    def _get_causal_chain_text(self, mediator: str) -> str:
        """获取包含特定中介变量的因果链文本"""
        for chain in self.causal_chains:
            if mediator in chain.mediator_vars:
                if chain.mediator_vars:
                    return f"{chain.cause_var} → {chain.mediator_vars[0]} → {chain.effect_var}"
                else:
                    return f"{chain.cause_var} → {chain.effect_var}"
        return ""

    def check_missing_controls(self) -> List[str]:
        """
        检查是否遗漏了重要的控制变量

        Returns:
            警告信息列表
        """
        issues = []

        if not self.variables:
            return issues

        control_var_names = set()
        for var in self.variables:
            if var.role == VariableRole.CONTROL:
                control_var_names.add(var.name_en.lower())

        # 检查常见控制变量
        typical_controls = {
            'size': '企业规模（总资产或对数）',
            'lev': '资产负债率',
            'age': '企业年龄',
            'roa': '盈利能力',
        }

        for ctrl, name in typical_controls.items():
            if ctrl not in control_var_names and not any(ctrl in v for v in control_var_names):
                issues.append(f"建议增加{name}作为控制变量")

        # 检查是否缺少行业/年份固定效应
        has_year_fe = any('year' in var.name_en.lower() or '时间' in var.name_cn for var in self.variables)
        has_industry_fe = any('industry' in var.name_en.lower() or '行业' in var.name_cn for var in self.variables)

        if not has_year_fe:
            issues.append("建议控制年份固定效应（year fixed effects）")
        if not has_industry_fe:
            issues.append("建议控制行业固定效应（industry fixed effects）")

        return issues

    def verify_endogeneity_handling(self) -> EndogeneityCheck:
        """
        验证内生性问题是否被妥善处理

        Returns:
            EndogeneityCheck对象
        """
        check = EndogeneityCheck(
            has_endogeneity=False,
            method_used=None,
            is_adequate=True,
            issues=[]
        )

        # 检测内生性讨论关键词
        endogeneity_keywords = {
            'iv': '工具变量法（IV）',
            '2sls': '两阶段最小二乘法（2SLS）',
            'gmm': '广义矩估计（GMM）',
            'did': '双重差分（DID）',
            'rd': '断点回归（RD）',
            '滞后': '滞后变量',
            'lag': '滞后变量',
            '内生性': '内生性讨论',
        }

        # 搜索内生性相关文本
        found_methods = []
        for keyword, method_name in endogeneity_keywords.items():
            if keyword in self.content.lower():
                found_methods.append(method_name)

        if found_methods:
            check.has_endogeneity = True
            check.method_used = '、'.join(found_methods)

            # 检查方法是否充分
            if '内生性讨论' in found_methods and len(found_methods) == 1:
                check.is_adequate = False
                check.issues.append("仅讨论了内生性问题但未采用适当的计量方法处理")

        return check

    def _find_section(self, section_name: str) -> str:
        """查找并提取论文中特定章节的内容

        支持多种格式：
        - ## 标题 (标准markdown二级标题)
        - # 标题 (标准markdown一级标题)
        - []{#anchor}**标题** (Pandoc转换格式1)
        - **标题** (加粗标题)
        - （二）标题 (编号+标题)
        """
        # 首先，查找章节标题的位置
        # 格式1: ## 标题
        pattern1 = rf'(?:^|\n)##\s+{re.escape(section_name)}(?:\s*\n|$)'
        # 格式2: # 标题
        pattern2 = rf'(?:^|\n)#\s+{re.escape(section_name)}(?:\s*\n|$)'
        # 格式3: **（编号）标题** 或 **标题** - 需要匹配**之间的内容包含section_name
        pattern3 = rf'\*\*(?:[^*]|\*(?!\*))*?{re.escape(section_name)}(?:[^*]|\*(?!\*))*?\*\*'
        # 格式4: 标题直接出现在某行（作为关键词搜索，兜底）
        pattern4 = rf'(?:^|\n)[^\n]*?{re.escape(section_name)}[^\n]*?(?:\n|$)'

        patterns = [pattern1, pattern2, pattern3, pattern4]

        section_pos = -1
        section_end = -1
        best_match = None

        for pattern in patterns:
            match = re.search(pattern, self.content)
            if match:
                if section_pos == -1 or match.start() < section_pos:
                    section_pos = match.start()
                    section_end = match.end()
                    best_match = match.group(0)

        if section_pos == -1:
            return ""

        # 找到了标题，现在提取标题后面的内容
        # 跳过标题行，查找下一个章节标题
        next_section_pos = section_end

        # 查找下一个章节标题（多种格式）
        next_patterns = [
            r'\n#{1,2}\s+[^\n]+',  # ## 新标题 或 # 新标题
            r'\n\[\]\{[^}]+\}\s*\*\*(?:[^*]|\*(?!\*))*?\*\*',  # []{#anchor}**新标题**
            r'\n\*\*(?:[^*]|\*(?!\*))*?\*\*',  # **新标题**
        ]

        next_section_end = len(self.content)
        for npat in next_patterns:
            nmatch = re.search(npat, self.content[next_section_pos:])
            if nmatch:
                next_section_end = min(next_section_end, next_section_pos + nmatch.start())

        section_content = self.content[next_section_pos:next_section_end].strip()

        # 限制长度
        return section_content[:5000]

    def _normalize_var_name(self, name: str) -> str:
        """标准化变量名"""
        # 移除常见前缀/后缀
        name = name.strip()
        name = re.sub(r'\s*\([^)]*\)', '', name)  # 移除括号内容
        name = re.sub(r'[（）()【】\[\]]', '', name)
        name = name.strip()

        # 转为英文变量名（如果是中文）
        if re.search(r'[\u4e00-\u9fff]', name):
            # 常见映射
            mapping = {
                '创新': 'innovation', '专利': 'patent', '耐心资本': 'patient_capital',
                '企业创新': 'innovation', '研发投入': 'rd', '研发支出': 'rd',
                '企业规模': 'size', '规模': 'size', '资产负债率': 'lev',
                '营业收入': 'revenue', '营收': 'revenue', '企业年龄': 'firm_age',
                '年龄': 'age', '融资约束': 'sa', '战略型股权': 'patient_equity',
            }
            for cn, en in mapping.items():
                if cn in name:
                    return en

        return name.lower()

    def _calculate_completeness(self) -> int:
        """计算模型设定完整性评分（0-100）"""
        score = 0

        # 有变量定义 +20
        if self.variables:
            score += 20

        # 有因果链 +20
        if self.causal_chains:
            score += 20

        # 有模型公式 +20
        if self.formulas:
            score += 20

        # 有内生性讨论 +20
        if self.content and ('内生性' in self.content or 'endogeneity' in self.content.lower()):
            score += 20
        elif self.content and ('稳健性' in self.content or 'robustness' in self.content.lower()):
            score += 10

        return min(score, 100)

    def _calculate_accuracy(self) -> int:
        """计算变量角色准确性评分（0-100）"""
        if not self.variables:
            return 0

        score = 100
        issues = len(self.over_control_issues) + len(self.variable_role_issues)

        # 每个问题扣10分
        score -= issues * 10

        return max(score, 0)

    def _calculate_causal_logic(self) -> int:
        """计算因果逻辑一致性评分（0-100）"""
        if not self.causal_chains:
            return 50  # 没有因果链给一半分

        score = 100

        # 过度控制问题严重影响因果逻辑
        high_severity_issues = sum(1 for i in self.over_control_issues if i.severity == 'high')
        score -= high_severity_issues * 30

        # 中介变量问题
        medium_severity = sum(1 for i in self.over_control_issues if i.severity == 'medium')
        score -= medium_severity * 15

        return max(score, 0)

    def _calculate_confidence(self) -> float:
        """计算提取置信度（0.0-1.0）"""
        confidence = 0.0

        # 基于提取到的内容
        if self.variables:
            confidence += 0.3
        if self.causal_chains:
            confidence += 0.3
        if self.formulas:
            confidence += 0.2

        # 基于问题数量调整
        if self.over_control_issues:
            confidence -= 0.1 * len(self.over_control_issues)

        return max(0.0, min(1.0, confidence))


def model_spec_to_dict(result: ModelSpecResult) -> Dict:
    """将ModelSpecResult转换为字典，用于JSON序列化"""
    if result is None:
        return {}

    return {
        'variables': [
            {
                'name_en': v.name_en,
                'name_cn': v.name_cn,
                'role': v.role.value if v.role else 'unknown',
                'var_type': v.var_type.value if v.var_type else 'continuous',
                'measurement': v.measurement,
            }
            for v in result.variables
        ],
        'causal_chains': [
            {
                'statement': c.statement,
                'cause_var': c.cause_var,
                'effect_var': c.effect_var,
                'mediator_vars': c.mediator_vars,
                'is_direct_effect': c.is_direct_effect,
            }
            for c in result.causal_chains
        ],
        'formulas': [
            {
                'formula_text': f.formula_text,
                'dependent_var': f.dependent_var,
                'independent_vars': f.independent_vars,
                'control_vars': f.control_vars,
                'mediator_vars': f.mediator_vars,
                'model_type': f.model_type,
            }
            for f in result.formulas
        ],
        'over_control_issues': [
            {
                'mediator_var': i.mediator_var,
                'causal_chain': i.causal_chain,
                'controlled_in_model': i.controlled_in_model,
                'severity': i.severity,
                'explanation': i.explanation,
            }
            for i in result.over_control_issues
        ],
        'variable_role_issues': [
            {
                'variable': i.variable,
                'declared_role': i.declared_role,
                'actual_role': i.actual_role,
                'severity': i.severity,
            }
            for i in result.variable_role_issues
        ],
        'missing_control_issues': result.missing_control_issues,
        'endogeneity_check': {
            'has_endogeneity': result.endogeneity_check.has_endogeneity if result.endogeneity_check else False,
            'method_used': result.endogeneity_check.method_used if result.endogeneity_check else None,
            'is_adequate': result.endogeneity_check.is_adequate if result.endogeneity_check else True,
            'issues': result.endogeneity_check.issues if result.endogeneity_check else [],
        } if result.endogeneity_check else None,
        'spec_completeness': result.spec_completeness,
        'spec_accuracy': result.spec_accuracy,
        'causal_logic': result.causal_logic,
        'extraction_confidence': result.extraction_confidence,
    }
