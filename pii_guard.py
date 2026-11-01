# -*- coding: utf-8 -*-
"""pii_guard —— 中文书稿的隐私信息脱敏与落盘前自检

为什么单独成一个模块：书稿里混着真实的手机号、邮箱、身份证号是常态
（尤其是回忆录、家史、访谈整理这类材料）。排版工具一旦把它们印进 PDF
或 docx，就等于公开。这个模块提供两步保护：

    1. mask()  —— 把联系方式替换成占位符（可留白名单）
    2. audit() —— 落盘前再扫一遍，发现漏网的就让构建失败

用法：

    from pii_guard import mask, audit

    clean, n = mask(text, whitelist={'13800000000'})
    leaks = audit(clean, whitelist={'13800000000'})
    assert not leaks, f'仍含隐私信息：{leaks}'
"""

import re

__all__ = ['mask', 'audit', 'EMAIL', 'MOBILE', 'LANDLINE', 'ID_CARD', 'PATTERNS']

# 邮箱：dotted+local@domain.tld，兼容全角 @ 与中文输入法常见的 ＠
EMAIL = re.compile(r'[0-9A-Za-z_\-\.+]+[@＠][0-9A-Za-z\-]+(?:[.。＠@][0-9A-Za-z\-]+)+')

# 手机号：11 位，允许中间夹空格或短横（139-1234-5678 / 139 1234 5678）
MOBILE = re.compile(r'(?<!\d)1(?:[ \-]?\d){10}(?!\d)')

# 座机：区号 3-4 位 + 7-8 位，允许短横
LANDLINE = re.compile(r'(?<!\d)0\d{2,3}[ \-]?\d{7,8}(?!\d)')

# 身份证：18 位（末位可能是 X）
ID_CARD = re.compile(r'(?<!\d)\d{17}[\dXx](?!\d)')

PATTERNS = (('id_card', ID_CARD), ('email', EMAIL),
            ('mobile', MOBILE), ('landline', LANDLINE))

DEFAULT_PLACEHOLDER = '〔联系方式已隐去〕'


def _norm(token):
    """归一化后再比对白名单：去掉空格与短横，全角＠转半角"""
    return token.replace(' ', '').replace('-', '').replace('＠', '@')


def mask(text, whitelist=(), placeholder=DEFAULT_PLACEHOLDER):
    """把所有联系方式替换为占位符。

    Args:
        text: 原始文本
        whitelist: 允许保留的号码/邮箱集合（例如作者授权公开的联系方式）
        placeholder: 替换文案

    Returns:
        (处理后文本, 替换次数)
    """
    wl = {_norm(w) for w in whitelist}
    count = 0

    def _sub(m):
        nonlocal count
        token = m.group(0)
        if _norm(token) in wl:
            return token
        count += 1
        return placeholder

    out = text
    for _, pat in PATTERNS:
        out = pat.sub(_sub, out)
    return out, count


def audit(text, whitelist=()):
    """扫描文本里仍然存在的联系方式，返回泄漏清单（已去重、已排除白名单）。

    构建脚本应在保存前调用它并 assert 为空——这是最后一道闸。
    """
    wl = {_norm(w) for w in whitelist}
    leaks = set()
    for kind, pat in PATTERNS:
        for m in pat.finditer(text):
            token = m.group(0)
            if _norm(token) not in wl:
                leaks.add(f'{kind}:{token}')
    return sorted(leaks)


if __name__ == '__main__':
    demo = ('联系作者：zhangsan@example.com 或 138-0013-8000，'
            '座机 010-12345678，身份证 110101199003071234。'
            '授权公开的号码 13900001111 应被保留。')
    print('原文：', demo)
    out, n = mask(demo, whitelist={'13900001111'})
    print(f'脱敏（替换 {n} 处）：', out)
    print('残留检查：', audit(out, whitelist={'13900001111'}) or '无')
