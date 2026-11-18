# -*- coding: utf-8 -*-
"""证据采集：绿色能效明细匹配、环保关键名单匹配。"""
from collections import Counter
from dataclasses import dataclass, field

from .normalize import norm_name, norm_code, is_uscc


@dataclass
class GreenStat:
    total: int = 0
    env_count: int = 0                      # 绿色目录名称落在环境类关键词上的记录数
    types: Counter = field(default_factory=Counter)
    env_catalogs: Counter = field(default_factory=Counter)


@dataclass
class EnvListHit:
    name: str = ""
    list_type: str = ""
    tags: list = field(default_factory=list)


def build_green_index(paths, env_keywords):
    """扫描绿色能效明细（Sheet1），返回三个索引：
    by_code: {USCC: GreenStat}, by_name: {归一化企业名称/标准主体名称: GreenStat}"""
    import openpyxl

    by_code, by_name = {}, {}
    for path in paths:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.worksheets[0]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            name = norm_name(row[0])
            std_name = norm_name(row[6]) if len(row) > 6 else ""
            catalog = str(row[21]) if len(row) > 21 and row[21] is not None else ""   # 绿色目录名称
            gtype = str(row[15]) if len(row) > 15 and row[15] is not None else ""     # 绿色类型
            is_env = any(k in catalog for k in env_keywords)
            def bump(index, key):
                if not key:
                    return
                st = index.setdefault(key, GreenStat())
                st.total += 1
                if gtype:
                    st.types[gtype] += 1
                if is_env:
                    st.env_count += 1
                    if catalog:
                        st.env_catalogs[catalog] += 1
            code = norm_code(row[1]) if len(row) > 1 and row[1] is not None else ""
            if is_uscc(code):
                bump(by_code, code)
            for nm in {name, std_name}:
                bump(by_name, nm)
        wb.close()
    return by_code, by_name


def lookup_green(by_code, by_name, *, code: str = "", names=()) -> GreenStat:
    """按信用代码优先、名称其次查绿色线索，合并唯一命中。"""
    merged = GreenStat()
    code_n = norm_code(code)
    hit = by_code.get(code_n) if is_uscc(code_n) else None
    if hit is None:
        name_hits = [by_name[n] for n in map(norm_name, names) if n in by_name]
        if len({id(h) for h in name_hits}) == 1:
            hit = name_hits[0]
    if hit:
        merged = hit
    return merged


def build_env_list_index(path, sheet="总名单1017"):
    """环保关键名单总表索引：按信用代码、按名称。"""
    import openpyxl

    by_code, by_name = {}, {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if sheet not in wb.sheetnames:
        sheet = wb.worksheets[0].title
    ws = wb[sheet]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None:
            continue
        name = str(row[1]).strip()
        code = norm_code(row[2]) if len(row) > 2 and row[2] is not None else ""
        list_type = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""
        tags = [str(t).strip() for t in row[6:17] if t is not None and str(t).strip()]
        hit = EnvListHit(name=name, list_type=list_type, tags=tags)
        if is_uscc(code):
            by_code[code] = hit
        by_name.setdefault(norm_name(name), hit)
    wb.close()
    return by_code, by_name


def lookup_env_list(by_code, by_name, *, code: str = "", names=()):
    code_n = norm_code(code)
    if is_uscc(code_n) and code_n in by_code:
        return by_code[code_n]
    for name in map(norm_name, names):
        if name in by_name:
            return by_name[name]
    return None
