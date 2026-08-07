#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
招标报告价值评分排序脚本
用法:
    python3 tools/score_reports.py [meta.json路径] [--top N] [--detail]

原理:
    价值无法直接测量, 只能把可观察特征量化后加权求和。本脚本对每条公告提取:
      1. 金额特征  : 正文中的"总投资金额/合同估算金额" (万元), log缩放
      2. 类型特征  : EPC/施工总承包/全过程咨询/监理等招标类型
      3. 资料完整度: 是否带工程量清单、图纸等附件
      4. 时间特征  : 投标截止时间距今天数 (临近截止=更急迫) 与发布时间新鲜度
      5. 关键词特征: 命中用户业务方向的词表
      6. 信息量    : 正文长度 (描述越详细, 可投标性越高)
    加权求和得到总分, 从高到低排序输出。

设计要点:
    - 不依赖绝对路径: meta.json 中的路径失效时, 自动按 meta.json 所在目录重拼
    - 权重可通过命令行调, 各特征先归一化到 0~1 再加权, 避免金额大数碾压
"""
import json
import re
import sys
import os
from datetime import datetime, date


# ---------------- 特征提取 ----------------

MONEY_PATTERNS = [
    r'合同估算金额[：:]\s*([\d,]+\.?\d*)\s*万元',
    r'工程总投资金额[：:]\s*([\d,]+\.?\d*)\s*万元',
    r'总投资[：:]\s*([\d,]+\.?\d*)\s*万元',
    r'工程概算投资额[：:]\s*([\d,]+\.?\d*)\s*万元',
    r'建安工程费[：:]\s*([\d,]+\.?\d*)\s*万元',
]

TYPE_KEYWORDS = [
    (10, ['EPC', '工程总承包', '设计施工', '设计采购施工']),   # 总包大单
    (8,  ['施工总承包', '工程施工', '施工招标', '标段施工']),   # 纯施工
    (7,  ['全过程咨询', '项目管理']),                            # 咨询类
    (5,  ['监理']),
    (4,  ['勘察设计', '初步设计', '施工图设计']),
]

DEADLINE_PATTERNS = [
    # [^。；]{0,30}? 允许中间的逗号/括号(如"（投标截止时间，下同）为")，但不过句号分号
    r'投标截止时间[^。；]{0,30}?(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})日?',
    r'投标文件递交的截止时间[^。；]{0,30}?(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})日?',
    r'开标时间[：:]\s*(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})日?',
]

# 按业务方向扩展此表; 命中即加分, 无关方向词不放进来
INTEREST_KEYWORDS = [
    '老旧小区', '学校', '医院', '产业园', '数据中心', '智慧',
    '市政', '道路', '污水', '给排水', '绿化', '装修', '托育',
]


def extract_money(text):
    """返回公告中出现的最大金额(万元), 无则 None"""
    amounts = []
    for pat in MONEY_PATTERNS:
        for m in re.finditer(pat, text):
            try:
                amounts.append(float(m.group(1).replace(',', '')))
            except ValueError:
                pass
    return max(amounts) if amounts else None


def extract_type_score(text):
    for score, kws in TYPE_KEYWORDS:
        for kw in kws:
            if kw in text:
                return score
    return 2  # 未知类型保底


def extract_deadline(text):
    for pat in DEADLINE_PATTERNS:
        m = re.search(pat, text)
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                continue
    return None


def extract_keywords(text):
    return [kw for kw in INTEREST_KEYWORDS if kw in text]


# ---------------- 归一化与打分 ----------------

def normalize(values):
    """列表归一化到 0~1; 全相等时返回全 0 避免除零"""
    lo, hi = min(values), max(values)
    if hi == lo:
        return [0.0] * len(values)
    return [(v - lo) / (hi - lo) for v in values]


def score_report(meta, md_text, today):
    att_names = ' '.join(a.get('name', '') for a in meta.get('attachments', []))
    full = md_text + ' ' + att_names + ' ' + meta.get('title', '')
    f = {}  # 原始特征
    f['money'] = extract_money(full) or 0.0
    f['type'] = extract_type_score(full)
    f['has_qingdan'] = 1.0 if '工程量清单' in att_names else 0.0
    f['has_attachment'] = 1.0 if meta.get('attachments') else 0.0
    f['length'] = len(md_text)
    f['keywords'] = len(extract_keywords(full))
    dl = extract_deadline(md_text)
    f['days_left'] = (dl - today).days if dl else None
    f['freshness'] = (today - date.fromisoformat(meta.get('date', '1970-01-01'))).days
    f['deadline_str'] = dl.isoformat() if dl else '未写明'
    f['matched_kws'] = extract_keywords(full)
    return f


def main():
    meta_path = sys.argv[1] if len(sys.argv) > 1 else 'output/zbgg_data/meta.json'
    top_n = None
    detail = False
    args = sys.argv[2:]
    while args:
        a = args.pop(0)
        if a == '--top':
            top_n = int(args.pop(0))
        elif a == '--detail':
            detail = True

    base = os.path.dirname(os.path.abspath(meta_path))
    with open(meta_path, encoding='utf-8') as fp:
        reports = json.load(fp)

    today = date.today()
    features = []
    for r in reports:
        md_p = r.get('md', '')
        if not os.path.isfile(md_p):
            cand = os.path.join(base, os.path.basename(md_p))
            md_p = cand if os.path.isfile(cand) else None  # 路径失效时按 meta 目录重拼
        md_text = ''
        if md_p:
            try:
                md_text = open(md_p, encoding='utf-8').read()
            except OSError:
                pass
        features.append(score_report(r, md_text, today))

    # 各维度跨报告归一化, 缺失(未写明截止)的 days_left 单独处理
    n = len(features)
    if n == 0:
        print('meta.json 为空, 无报告可评分')
        return

    norm = {}
    norm['money'] = normalize([f['money'] for f in features])
    norm['type'] = normalize([f['type'] for f in features])
    norm['attach'] = normalize([f['has_qingdan'] + f['has_attachment'] for f in features])
    norm['length'] = normalize([f['length'] for f in features])
    norm['kw'] = normalize([f['keywords'] for f in features])

    dl = [f['days_left'] for f in features if f['days_left'] is not None]
    norm_dl = normalize(dl) if dl else []
    dl_idx = 0
    # 权重: 金额>类型>资料>紧急度>关键词>信息量, 可自行调整
    W = dict(money=0.30, type=0.20, attach=0.15, urgency=0.15, kw=0.10, length=0.10)

    rows = []
    for i, (r, f) in enumerate(zip(reports, features)):
        s = 0.0
        s += W['money'] * norm['money'][i]
        s += W['type'] * norm['type'][i]
        s += W['attach'] * norm['attach'][i]
        if f['days_left'] is not None:
            u = norm_dl[dl_idx] if dl else 0.0
            dl_idx += 1
        else:
            u = 0.0  # 未写明截止日期的紧急度按 0
        s += W['urgency'] * u
        s += W['kw'] * norm['kw'][i]
        s += W['length'] * norm['length'][i]
        rows.append((s, f, r))

    rows.sort(key=lambda x: -x[0])
    if top_n:
        rows = rows[:top_n]

    print(f"{'排名':<3}{'评分':<7}{'金额(万)':<11}{'类型分':<6}{'截止日':<12}标题")
    print('-' * 90)
    for rank, (s, f, r) in enumerate(rows, 1):
        money = f'{f["money"]:,.1f}' if f['money'] else '--'
        print(f"{rank:<4}{s:<7.3f}{money:<12}{f['type']:<7}{f['deadline_str']:<12}{r['title']}")
        if detail:
            print(f"      关键词: {f['matched_kws']} | 距离发布 {f['freshness']} 天 | 正文 {f['length']} 字")
    print('-' * 90)
    print(f'共 {len(reports)} 条, 权重: {W}')


if __name__ == '__main__':
    main()
