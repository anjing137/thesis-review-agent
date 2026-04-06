# -*- coding: utf-8 -*-
"""
论文评价提示词模板
"""

# 评价维度配置（共6个，与main.py generate_review_prompt()动态渲染保持一致）
DIMENSIONS = {
    '选题与研究问题': {
        'weight': 0.15,
        'max_score': 15,
        'keywords': ['选题', '研究问题', '题目', '研究价值', '创新点'],
        'criteria': [
            '题目是否具体明确（标题≤20字）',
            '题名与内容是否一致：标题应准确反映论文研究对象、方法或核心结论，不得出现题名与正文内容明显不符',
            '引言是否开门见山、不绕圈子（缘由300字左右、问题提出300字左右、目的意义400字左右），总篇幅600–1000字',
            '研究必要性是否充分论证',
            '研究价值和创新点是否明确',
            '引言是否不与摘要雷同',
            '引言中是否存在"国内首创、填补空白"等夸张自我评语'
        ]
    },
    '参考文献与学术规范': {
        'weight': 0.15,
        'max_score': 15,
        'keywords': ['参考文献', '文献', '学术规范'],
        'criteria': [
            '数量是否≥15篇（硬性要求，不足则该维度0分）',
            '外文文献是否≥3篇；期刊占比是否≥2/3',
            '近3年文献占比是否充足（若近3年占比低于1/3应发出"文献时效性偏弱"提醒）',
            '格式是否规范'
        ]
    },
    '内容创新性': {
        'weight': 0.15,
        'max_score': 15,
        'keywords': ['创新', '创新性', '边际贡献', '文献评述'],
        'criteria': [
            '研究视角是否新颖；是否有文献评述（指出已有研究不足）',
            '是否明确边际贡献',
            '文献综述是否具有系统性：是否全面梳理了该领域前人研究，突出重点'
        ]
    },
    '框架与逻辑结构': {
        'weight': 0.20,
        'max_score': 20,
        'keywords': ['框架', '结构', '逻辑', '章节'],
        'criteria': [
            '实证论文标准结构：引言→文献综述→理论分析→研究设计→实证分析→结论',
            '学理论文标准结构：引言→现状→问题→原因→对策→结语',
            '各章节逻辑是否清晰',
            '正文字数≥8000字；分章字数是否符合指导要点要求（引言三要素各约300/300/400字、文献综述1000字、理论分析≥2000字、实证分析2000字、结论1000字）',
            '论文是否包含致谢部分'
        ]
    },
    '方法与论证严谨性': {
        'weight': 0.25,
        'max_score': 25,
        'keywords': ['方法', '论证', '稳健性', '内生性', '检验'],
        'criteria': [
            '理论分析是否深入',
            '数据质量和变量定义是否清晰',
            '稳健性检验是否完整（实证论文强制要求，否则该维度扣分）',
            '是否讨论内生性问题'
        ]
    },
    '语言与表达': {
        'weight': 0.10,
        'max_score': 10,
        'keywords': ['语言', '表达', '第三人称', '术语', '标点'],
        'criteria': [
            '是否使用第三人称（"本文"正确，"我/我们"错误）',
            '术语是否专业规范；标点符号是否正确',
            '摘要（中英文）和正文的语言表达是否准确、流畅'
        ]
    }
}

# 一票否决项
# 注：稳健性检验、过度控制等问题由 AI 在"方法与论证严谨性"维度自主评价，
# 通过 comment 和 suggestions 体现，不作为硬性否决规则。
VETO_RULES = [
    {'condition': '参考文献数量不足15篇', 'dimension': '参考文献与学术规范', 'action': '该维度0分'},
    {'condition': '外文参考文献不足3篇（0篇扣5分，1篇扣3分，2篇扣1分）', 'dimension': '参考文献与学术规范', 'action': '按缺少数目扣分'},
]

# 评分等级
# 评分等级（闭区间：lo <= score <= hi）
GRADE_LEVELS = [
    (90, 100, '优秀（优+）'),
    (85, 89, '良好上（优）'),
    (80, 84, '良好中（良+）'),
    (75, 79, '良好下（良）'),
    (70, 74, '中等上（中+）'),
    (65, 69, '中等（中）'),
    (60, 64, '中等下（中-）'),
    (0, 59, '不合格'),
]

# 论文类型
PAPER_TYPES = ['实证性', '学理性']

# 实证论文关键词
EMPIRICAL_KEYWORDS = ['实证', '回归', '数据', '模型', '变量', '检验', '分析', '面板', 'OLS', 'GDP']


def get_dimension_score(total_score: float, dimension: str) -> float:
    """
    计算某维度得分

    Args:
        total_score: 总分
        dimension: 维度名称

    Returns:
        该维度的绝对得分
    """
    weight = DIMENSIONS.get(dimension, {}).get('weight', 0)
    return total_score * weight


def check_veto_rules(stats: dict, is_empirical: bool = False) -> list:
    """
    检查一票否决项

    Args:
        stats: 统计数据
        is_empirical: 是否为实证论文

    Returns:
        触发的一票否决项列表
    """
    triggered = []
    ref_count = stats.get('references', {}).get('total', 0)

    if ref_count < 15:
        triggered.append({
            'rule': '参考文献<15篇',
            'dimension': '参考文献与学术规范',
            'action': '该维度0分'
        })

    if is_empirical:
        # 检查稳健性检验（这里需要AI辅助判断）
        pass

    return triggered


def get_grade_level(score: float) -> str:
    """
    根据得分获取等级

    Args:
        score: 总分

    Returns:
        等级描述
    """
    for min_score, max_score, level in GRADE_LEVELS:
        if min_score <= score <= max_score:
            return level
    return '不合格'


def is_empirical_paper(body_content: str) -> bool:
    """
    判断是否为实证论文

    Args:
        body_content: 正文内容

    Returns:
        True if 实证性, False if 学理性
    """
    body_lower = body_content.lower()
    score = sum(1 for kw in EMPIRICAL_KEYWORDS if kw in body_lower)
    return score >= 2