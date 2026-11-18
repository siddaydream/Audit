# -*- coding: utf-8 -*-
"""阶段1 预填：本期表 + 上期表 -> 中间表（四列回填、新增留空高亮）+ 新增企业清单 + 沿用数据sidecar。"""
import json
import os
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from .matcher import (PLATFORM_KEYS, build_enterprise_registry, consistent_fields,
                      load_form, match_by_receipt)
from .normalize import norm_name

BLANK_FILL = PatternFill("solid", fgColor="FFF2CC")
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")

PLATFORM_HEADERS = {
    "plat_yjt": "国标行业（预警通）",
    "plat_scope": "经营范围（企查查）",
    "plat_qcc_ind": "企查查行业",
    "plat_qcc_gb": "国标行业（企查查）",
}


def prefill(current_path: str, prev_path: str, out_dir: str):
    today = date.today().strftime("%Y%m%d")
    cur_loans, cur_header = load_form(current_path)
    prev_loans, prev_header = load_form(prev_path)

    prev_registry = build_enterprise_registry(prev_loans)
    cur_registry = build_enterprise_registry(cur_loans)

    # 企业级回填决策：key -> 四列数据 / 多候选告警
    fill_map, ambiguous = {}, []
    for key, reg in cur_registry.items():
        names = reg["names"] | {norm_name(l.borrower) for l in reg["loans"] if norm_name(l.borrower)}
        code = next(iter(reg["codes"])) if reg["codes"] else ""
        # 信用代码优先
        if code and code in prev_registry and prev_registry[code]["plat"]:
            fill_map[key] = prev_registry[code]["plat"]
            continue
        # 名称匹配
        name_keys = {k for k, p in prev_registry.items()
                     if names & p["names"]}
        plats = {json.dumps(prev_registry[k]["plat"], ensure_ascii=False, sort_keys=True)
                 for k in name_keys if prev_registry[k]["plat"]}
        if len(plats) == 1 and name_keys:
            fill_map[key] = prev_registry[next(iter(name_keys))]["plat"]
        elif len(plats) > 1:
            ambiguous.append(sorted(reg["names"])[0])

    # 借据号级沿用信息（sidecar）
    receipt_map = match_by_receipt(cur_loans, prev_loans)
    carry = []
    for loan in cur_loans:
        prev = receipt_map.get(loan.receipt)
        carry.append({
            "row": loan.row, "seq": loan.seq, "receipt": loan.receipt,
            "prev_result": prev.prev_result if prev else "",
            "prev_reason": prev.prev_reason if prev else "",
            "consistent": bool(prev and consistent_fields(loan, prev)),
        })
    # 上期有、本期没有的借据（如已退回贷款），供"换借据再报"预警
    current_receipts = {l.receipt for l in cur_loans if l.receipt}
    prev_only = []
    for p in prev_loans:
        if p.receipt and p.receipt not in current_receipts:
            prev_only.append({
                "entity": p.entity, "entity_norm": norm_name(p.entity),
                "id_number": p.id_number, "receipt": p.receipt,
                "prev_result": p.prev_result, "status": p.status,
            })

    # ---- 写中间表（原表整表复制 + 追加四列）----
    base = os.path.splitext(os.path.basename(current_path))[0]
    base = norm_stem(base)
    mid_path = os.path.join(out_dir, f"{today}-中间表-{base}.xlsx")
    wb = openpyxl.load_workbook(current_path)
    ws = wb.worksheets[0]
    start_col = ws.max_column + 1
    for offset, key in enumerate(PLATFORM_KEYS):
        col = start_col + offset
        cell = ws.cell(2, col, PLATFORM_HEADERS[key])
        cell.font = Font(bold=True)
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_row2 = [norm_name(c.value) for c in ws[2]]
    for loan in cur_loans:
        reg_key = getattr(loan, "reg_key", None) or loan.entity_key
        plat = fill_map.get(reg_key, {})
        for offset, key in enumerate(PLATFORM_KEYS):
            col = start_col + offset
            cell = ws.cell(loan.row, col)
            value = plat.get(key, "")
            if value:
                cell.value = value
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                cell.fill = BLANK_FILL   # 新增企业留空高亮
    wb.save(mid_path)

    # sidecar：沿用数据（与中间表同名的 .carry.json，audit 自动读取）
    with open(os.path.splitext(mid_path)[0] + ".carry.json", "w", encoding="utf-8") as f:
        json.dump({"carry": carry, "prev_only": prev_only}, f, ensure_ascii=False, indent=1)

    # ---- 新增企业清单（只读，供查找）----
    new_path = os.path.join(out_dir, f"{today}-新增企业清单.xlsx")
    out = openpyxl.Workbook()
    sh = out.active
    sh.title = "新增企业清单"
    headers = ["经营主体名称", "借款人名称", "中间表行号", "贷款序号", "贷款借据号",
               "上期是否有该企业", "备注"]
    sh.append(headers)
    for c in sh[1]:
        c.font = Font(bold=True)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for key, reg in cur_registry.items():
        if key in fill_map:
            continue
        display = reg["loans"][0].entity
        borrowers = "、".join(dict.fromkeys(l.borrower for l in reg["loans"] if l.borrower))
        rows = ",".join(str(l.row) for l in reg["loans"])
        seqs = ",".join(l.seq for l in reg["loans"])
        receipts = "\n".join(l.receipt for l in reg["loans"])
        note = []
        if display in ambiguous or any(n in ambiguous for n in reg["names"]):
            note.append("上期存在疑似同名多候选，请人工核对")
        sh.append([display, borrowers, rows, seqs, receipts,
                   "是（其他借据）" if _prev_has_entity(reg, prev_registry) else "否",
                   "；".join(note)])
    for col, w in zip("ABCDEFG", [36, 30, 12, 12, 30, 14, 36]):
        sh.column_dimensions[col].width = w
    sh.freeze_panes = "A2"
    out.save(new_path)

    n_new = len(cur_registry) - len(fill_map)
    print(f"[预填完成] 中间表: {mid_path}")
    print(f"[预填完成] 新增企业清单: {new_path}（新增企业 {n_new} 家，已回填 {len(fill_map)} 家）")
    if ambiguous:
        print(f"[注意] 疑似同名多候选 {len(ambiguous)} 家，已在清单备注，请人工核对")
    return mid_path, new_path


def _prev_has_entity(reg, prev_registry):
    names = reg["names"]
    return any(names & preg["names"] or (reg["codes"] & preg["codes"])
               for preg in prev_registry.values())


def norm_stem(stem: str) -> str:
    """去掉文件名开头的 8 位日期前缀，避免产出文件双重前缀。"""
    return stem[9:] if len(stem) > 9 and stem[:8].isdigit() and stem[8] in "-_" else stem
