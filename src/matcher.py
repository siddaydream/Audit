# -*- coding: utf-8 -*-
"""表格读取与两级匹配：企业级（平台四列回填）、借据号级（上期结论沿用）。"""
from dataclasses import dataclass, field
from typing import Optional

from .normalize import norm_name, norm_code, is_uscc

# 表头关键词 -> 字段名（按"表头包含关键词"识别列，容忍列序变动）
COLUMN_HINTS = {
    "seq": "序号",
    "date": "数据日期",
    "bank": "银行名称",
    "claim_industry": "行业代码及名称",
    "borrower": "借款人名称",
    "entity": "经营主体名称",
    "id_number": "统一社会信用代码",
    "receipt": "贷款借据号",
    "prev_result": "是否发现存在不符合贴息条件情况",
    "prev_reason": "简要描述审核发现的问题",
    "status": "审核状态",
    "plat_yjt": "国标行业（预警通）",
    "plat_scope": "经营范围（企查查）",
    "plat_qcc_ind": "企查查行业",
    "plat_qcc_gb": "国标行业（企查查）",
}

PLATFORM_KEYS = ["plat_yjt", "plat_scope", "plat_qcc_ind", "plat_qcc_gb"]


@dataclass
class Loan:
    row: int                       # 工作表中的行号
    seq: str = ""
    date: str = ""
    bank: str = ""
    claim_industry: str = ""
    borrower: str = ""
    entity: str = ""
    id_number: str = ""
    receipt: str = ""
    prev_result: str = ""
    prev_reason: str = ""
    status: str = ""
    plat: dict = field(default_factory=dict)  # 四列平台数据

    @property
    def entity_key(self) -> str:
        """企业级主键：优先统一社会信用代码，否则归一化经营主体名称。"""
        if is_uscc(self.id_number):
            return norm_code(self.id_number)
        return norm_name(self.entity)

    @property
    def entity_name_norm(self) -> str:
        return norm_name(self.entity)


def _match_header(header_row) -> dict:
    """把表头行映射成 {字段名: 列号(1-based)}。表头中的换行与空白先归一。"""
    mapping = {}
    for idx, cell in enumerate(header_row, start=1):
        text = norm_name(cell)
        if not text:
            continue
        for key, hint in COLUMN_HINTS.items():
            if key not in mapping and norm_name(hint) in text:
                mapping[key] = idx
    return mapping


def load_form(path: str, sheet_name: Optional[str] = None):
    """读取申请表。返回 (loans: list[Loan], header_map, ws_max_row)。

    约定：第 1 行为表名，第 2 行为表头，数据从第 3 行起，"合计"行跳过。
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]

    rows = ws.iter_rows(values_only=False)
    loans = []
    header_map = {}
    row_iter = ws.iter_rows()
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx == 2:
            header_map = _match_header(row)
        if r_idx < 3:
            continue
        first = row[header_map.get("seq", 1) - 1] if row and len(row) >= 1 else None
        if first is None or norm_name(first) in ("", "合计"):
            continue
        def cell(key):
            col = header_map.get(key)
            if col is None or col > len(row):
                return ""
            v = row[col - 1]
            return "" if v is None else str(v).strip()
        loan = Loan(
            row=r_idx,
            seq=cell("seq"),
            date=cell("date"),
            bank=cell("bank"),
            claim_industry=cell("claim_industry"),
            borrower=cell("borrower"),
            entity=cell("entity"),
            id_number=cell("id_number"),
            receipt=cell("receipt"),
            prev_result=cell("prev_result"),
            prev_reason=cell("prev_reason"),
            status=cell("status"),
            plat={k: cell(k) for k in PLATFORM_KEYS},
        )
        loans.append(loan)
    wb.close()
    if not header_map:
        raise ValueError(f"未在 {path} 第2行识别到表头，请检查表格格式")
    return loans, header_map


def match_by_receipt(current_loans, prev_loans):
    """借据号精确匹配：{借据号: 上期Loan}。"""
    return {p.receipt: p for p in prev_loans if p.receipt}


def consistent_fields(cur: Loan, prev: Loan) -> bool:
    """沿用前的三项一致性校验：数据日期、借款人名称、经营主体名称。"""
    return (
        norm_name(cur.date) == norm_name(prev.date)
        and norm_name(cur.borrower) == norm_name(prev.borrower)
        and norm_name(cur.entity) == norm_name(prev.entity)
    )


def build_enterprise_registry(loans):
    """企业注册表：{entity_key: {names:set, codes:set, loans:[Loan], plat:dict}}。

    两遍合并：有统一社会信用代码的行先按代码建键；
    无代码（个人借款、主体为法人）的行按归一化名称并入既有键或新建名称键，
    保证同一企业只占一个键。每个 Loan 上回写 reg_key 指向所属键。"""
    registry = {}

    def _attach(key, loan):
        if key not in registry:
            registry[key] = {"names": set(), "codes": set(), "loans": [], "plat": {}}
        reg = registry[key]
        if loan.entity_name_norm:
            reg["names"].add(loan.entity_name_norm)
        if is_uscc(loan.id_number):
            reg["codes"].add(norm_code(loan.id_number))
        reg["loans"].append(loan)
        for k, v in loan.plat.items():
            if v and not reg["plat"].get(k):
                reg["plat"][k] = v
        loan.reg_key = key

    for loan in loans:
        if is_uscc(loan.id_number):
            _attach(norm_code(loan.id_number), loan)
    for loan in loans:
        if is_uscc(loan.id_number):
            continue
        nn = loan.entity_name_norm
        target = None
        if nn:
            for key, reg in registry.items():
                if nn in reg["names"]:
                    target = key
                    break
        _attach(target if target is not None else nn, loan)
    return registry


def resolve_enterprise(registry, *, code: str = "", names=()):
    """按信用代码或名称在企业注册表中查找，返回 (registry_key 或 None, 是否多候选)。"""
    code_n = norm_code(code)
    if code_n and is_uscc(code_n) and code_n in registry:
        return code_n, False
    hit_keys = set()
    for name in names:
        nn = norm_name(name)
        if not nn:
            continue
        for key, reg in registry.items():
            if nn in reg["names"]:
                hit_keys.add(key)
    if len(hit_keys) == 1:
        return hit_keys.pop(), False
    if len(hit_keys) > 1:
        return None, True   # 多候选，交人工
    return None, False
