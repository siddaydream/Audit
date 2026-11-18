# -*- coding: utf-8 -*-
"""共用审核管线：命令行（run.py）与图形界面（gui.py）调用同一套流程。"""
import json
import os
from collections import defaultdict

from . import engine, evidence, matcher, report
from . import prefill as prefill_mod
from . import validator
from .paths import DATA_DIR, OUT_DIR
from .normalize import is_uscc, norm_code, norm_name


def log_print(msg):
    print(msg)


def run_prefill(current_path, prev_path, out_dir=OUT_DIR, log=log_print):
    """阶段1 预填，返回 (中间表路径, 新增企业清单路径)。"""
    mid_path, new_path = prefill_mod.prefill(current_path, prev_path, out_dir)
    return mid_path, new_path


def collect_new_entities(new_path, log=log_print):
    """读取新增企业清单，返回 [(经营主体, 借款人, 行号)]，并在日志中列出。"""
    import openpyxl
    if not os.path.exists(new_path):
        return []
    wb = openpyxl.load_workbook(new_path, data_only=True, read_only=True)
    sh = wb.worksheets[0]
    items = []
    for r in sh.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        items.append((str(r[0]), str(r[1] or ""), str(r[2] or "")))
    wb.close()
    if items:
        log(f"请打开中间表，为以下 {len(items)} 家新增企业填写四列（黄色格子）：")
        for name, borrower, rows in items:
            log(f"  · {name}（中间表第 {rows} 行）")
    else:
        log("无新增企业，中间表四列已全部回填，可直接进入审核。")
    return items


def build_carry(current_path, prev_path, log=log_print):
    """结论沿用数据：优先读取中间表同名 sidecar，否则从 --prev 表重建。
    返回 (carry_by_receipt, prev_only)。"""
    loans, _ = matcher.load_form(current_path)
    carry_by_receipt, prev_only = {}, []
    sidecar = os.path.splitext(current_path)[0] + ".carry.json"
    if os.path.exists(sidecar):
        with open(sidecar, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            carry, prev_only = data.get("carry", []), data.get("prev_only", [])
        else:
            carry = data
        for rec in carry:
            if rec.get("receipt"):
                carry_by_receipt[rec["receipt"]] = rec
        log(f"沿用数据来源: {os.path.basename(sidecar)}")
    elif prev_path:
        prev_loans, _ = matcher.load_form(prev_path)
        prev_by_receipt = matcher.match_by_receipt(loans, prev_loans)
        for receipt, prev_loan in prev_by_receipt.items():
            carry_by_receipt[receipt] = {
                "receipt": receipt,
                "prev_result": prev_loan.prev_result,
                "consistent": matcher.consistent_fields(
                    next(l for l in loans if l.receipt == receipt), prev_loan),
            }
        current_receipts = {l.receipt for l in loans if l.receipt}
        for p in prev_loans:
            if p.receipt and p.receipt not in current_receipts:
                prev_only.append({"entity": p.entity, "entity_norm": norm_name(p.entity),
                                  "id_number": p.id_number, "receipt": p.receipt,
                                  "prev_result": p.prev_result, "status": p.status})
        log(f"沿用数据来源: {prev_path}")
    else:
        log("未提供上期数据（sidecar 或上期表），全部按新增借据走证据判定")
    return carry_by_receipt, prev_only


def build_reborrows(loans, registry, prev_only):
    """换借据再报预警：上期有"是"结论且本期未出现的旧借据，其企业在本期有新借据时标记。"""
    def _find(entity_norm, id_number):
        code = norm_code(id_number)
        if is_uscc(code):
            for k, reg in registry.items():
                if code in reg["codes"]:
                    return k
        if entity_norm:
            for k, reg in registry.items():
                if entity_norm in reg["names"]:
                    return k
        return None

    reborrows = {}
    for po in prev_only:
        if po.get("prev_result") != "是":
            continue
        key = _find(po.get("entity_norm", ""), po.get("id_number", ""))
        if key and key not in reborrows:
            note = "上期同企业借据判“是”"
            if po.get("status"):
                note += f"（{po['status']}）"
            reborrows[key] = note
    return reborrows


def run_audit(current_path, prev_path=None, out_dir=OUT_DIR, strict=False,
              green_detail=False, log=log_print):
    """阶段3 审核，返回输出文件路径字典。"""
    import time
    started = time.time()
    loans, header_map = matcher.load_form(current_path)
    validator.validate_form(loans, header_map, current_path)
    rules = engine.load_rules()
    manifest = validator.load_manifest(strict=strict)

    carry_by_receipt, prev_only = build_carry(current_path, prev_path, log=log)
    cur_registry = matcher.build_enterprise_registry(loans)
    reborrows = build_reborrows(loans, cur_registry, prev_only)

    green_paths = [os.path.join(DATA_DIR, manifest["files"][r]["path"])
                   for r in ("green_1", "green_2") if r in manifest["files"]]
    t0 = time.time()
    green_by_code, green_by_name = evidence.build_green_index(
        green_paths, rules["green_env_keywords"])
    log(f"绿色明细索引完成: {len(green_by_code)} 信用代码 / {len(green_by_name)} 名称 "
        f"({time.time()-t0:.1f}s)")
    env_by_code, env_by_name = evidence.build_env_list_index(
        os.path.join(DATA_DIR, manifest["files"]["env_list"]["path"]))

    outcomes = []
    for loan in loans:
        reg_key = getattr(loan, "reg_key", None) or loan.entity_key
        plat = cur_registry[reg_key]["plat"]
        rec = carry_by_receipt.get(loan.receipt, {})
        names = {loan.entity_name_norm, norm_name(loan.borrower)}
        green = evidence.lookup_green(green_by_code, green_by_name,
                                      code=loan.id_number, names=names)
        env_hit = evidence.lookup_env_list(env_by_code, env_by_name,
                                           code=loan.id_number, names=names)
        outcomes.append(engine.audit_loan(
            loan, plat, rec.get("prev_result", ""), rec.get("consistent", False),
            green, env_hit, rules, reborrows))

    base = prefill_mod.norm_stem(os.path.splitext(os.path.basename(current_path))[0])
    base = base.replace("中间表-", "")
    outputs = {
        "result": report.write_audit_workbook(current_path, loans, outcomes, out_dir, base),
        "review": report.write_review_list(loans, outcomes, out_dir)[0],
        "keywords": report.write_keyword_detail(loans, outcomes, out_dir),
        "log": report.write_evidence_log(loans, outcomes, out_dir),
    }
    if green_detail:
        outputs["green_detail"] = _collect_green_detail(
            green_paths, loans, cur_registry, out_dir)

    dist = defaultdict(int)
    for oc in outcomes:
        dist[oc.result] += 1
    log(f"审核完成，结果分布: {dict(dist)}（共 {len(outcomes)} 行，含待补数据行）")
    for key, path in outputs.items():
        log(f"产出 [{key}]: {path}")
    log(f"耗时 {time.time()-started:.1f}s")
    return outputs, outcomes


def _collect_green_detail(green_paths, loans, registry, out_dir):
    """二次扫描绿色明细，按本期表企业首现顺序输出命中行（含原始44列）。"""
    import openpyxl

    code2key, name2keys = {}, defaultdict(set)
    for key, reg in registry.items():
        for c in reg["codes"]:
            code2key[c] = key
        for n in reg["names"]:
            name2keys[n].add(key)
    name2key = {n: ks.pop() for n, ks in name2keys.items() if len(ks) == 1}

    order = []
    for loan in loans:
        if loan.entity_key not in order:
            order.append(loan.entity_key)
    blocks = {}
    for key, reg in registry.items():
        blocks[key] = {
            "entity": reg["loans"][0].entity,
            "seqs": ",".join(l.seq for l in reg["loans"]),
            "borrowers": "/".join(dict.fromkeys(l.borrower for l in reg["loans"])),
            "rows": [],
        }
    for path in green_paths:
        tag = "明细1" if "1" in os.path.basename(path).split("-")[-1] else "明细2"
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.worksheets[0]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            code = norm_code(row[1]) if len(row) > 1 and row[1] is not None else ""
            key = code2key.get(code) if is_uscc(code) else None
            how = "统一社会信用代码" if key else ""
            if not key:
                key = name2key.get(norm_name(row[0]))
                how = "企业名称" if key else ""
                if not key and len(row) > 6:
                    key = name2key.get(norm_name(row[6]))
                    how = "标准主体名称" if key else ""
            if key:
                blocks[key]["rows"].append((row, tag, how))
        wb.close()
    ordered = [blocks[k] for k in order if blocks[k]["rows"]]
    return report.write_green_detail(ordered, out_dir)
