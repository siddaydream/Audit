# -*- coding: utf-8 -*-
"""名称与代码归一化：全角/半角括号、空白等干扰项统一。"""
import re


def norm_name(value) -> str:
    """企业名称归一化：全角括号转半角、去所有空白。"""
    if value is None:
        return ""
    s = str(value).replace("（", "(").replace("）", ")").replace("\u3000", " ")
    return re.sub(r"\s+", "", s).strip()


def norm_code(value) -> str:
    """统一社会信用代码归一化：去空白、转大写。"""
    if value is None:
        return ""
    return str(value).strip().upper()


def is_uscc(value) -> bool:
    """是否为 18 位统一社会信用代码。

    身份证号同为 18 位，需排除：含字母的必是信用代码；
    纯数字的按身份证特征排除（前2位省级行政区码 + 第7-14位为出生日期）。"""
    code = norm_code(value)
    if not re.fullmatch(r"[0-9A-Z]{18}", code):
        return False
    if any(c.isalpha() for c in code):
        return True
    if re.match(r"^(1[1-9]|[2-7]\d|81|82)\d{4}(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])", code):
        return False
    return True


def parse_industry_codes(value):
    """从"国标行业"文本中解析小类码，如 '污水处理及其再生利用(D4620)' -> ['4620']，
    '…水污染治理（N7721）' -> ['7721']。返回去重后的 4 位码列表。"""
    if value is None:
        return []
    codes = re.findall(r"[A-Z](\d{4})", str(value))
    return list(dict.fromkeys(codes))
