# -*- coding: utf-8 -*-
"""阶段3 输出层：审核结果表、绿色匹配明细、不确定复核清单、关键词命中明细、证据日志。"""
import csv
import json
import os
from collections import Counter, defaultdict
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from .engine import AuditOutcome
from .reason import build_reason, review_point

HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
RESULT_HEADERS = ["审核结果（是否发现不符合贴息条件情况）", "审核理由"]


def _today() -> str:
    return date.today().strftime("%Y%m%d")


def write_audit_workbook(mid_path, loans, outcomes, out_dir, base_name):
    """在中间表基础上追加审核结果两列，输出日期前缀结果表。"""
    result_path = os.path.join(out_dir, f"{_today()}-{base_name}-审核结果.xlsx")
    wb = openpyxl.load_workbook(mid_path)
    ws = wb.worksheets[0]
    start_col = ws.max_column + 1
    for offset, title in enumerate(RESULT_HEADERS):
        cell = ws.cell(2, start_col + offset, title)
        cell.font = Font(bold=True)
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for loan, outcome in zip(loans, outcomes):
        c1 = ws.cell(loan.row, start_col, outcome.result)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2 = ws.cell(loan.row, start_col + 1, build_reason(outcome))
        c2.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions[openpyxl.utils.get_column_letter(start_col)].width = 22
    ws.column_dimensions[openpyxl.utils.get_column_letter(start_col + 1)].width = 90
    wb.save(result_path)
    return result_path


def write_review_list(loans, outcomes, out_dir):
    path = os.path.join(out_dir, f"{_today()}-不确定复核清单.xlsx")
    wb = openpyxl.Workbook()
    sh = wb.active
    sh.title = "不确定复核清单"
    sh.append(["序号", "经营主体名称", "借款人名称", "贷款借据号", "审核结果", "复核要点", "预警标记"])
    for c in sh[1]:
        c.font = Font(bold=True)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    rows = 0
    for loan, oc in zip(loans, outcomes):
        if oc.result != "不确定" and not oc.review_flag:
            continue
        sh.append([loan.seq, loan.entity, loan.borrower, loan.receipt, oc.result,
                   review_point(oc), oc.review_flag])
        rows += 1
    for col, w in zip("ABCDEFG", [8, 34, 26, 26, 10, 80, 40]):
        sh.column_dimensions[col].width = w
    sh.freeze_panes = "A2"
    wb.save(path)
    return path, rows


def write_keyword_detail(loans, outcomes, out_dir):
    path = os.path.join(out_dir, f"{_today()}-关键词命中明细.xlsx")
    wb = openpyxl.Workbook()
    sh = wb.active
    sh.title = "关键词命中明细"
    sh.append(["序号", "经营主体名称", "预警通行业码", "企查查行业码", "目录内命中",
               "排除码命中", "强命中(小类:词)", "弱命中", "C42疑义命中",
               "绿色总条数", "环境类条数", "环保名单", "规则", "结果"])
    for c in sh[1]:
        c.font = Font(bold=True)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for loan, oc in zip(loans, outcomes):
        ev = oc.evidence
        strong = "；".join(f"{cat}:{'/'.join(ws)}" for cat, ws in ev.strong_hits.items())
        excluded = "；".join(f"{k} {v[:20]}" for k, v in ev.excluded_hits.items())
        sh.append([loan.seq, loan.entity, "/".join(ev.codes_yjt), "/".join(ev.codes_qcc),
                   "/".join(ev.eligible_hits), excluded, strong, "/".join(ev.weak_hits),
                   "/".join(ev.c42_hits), ev.green_total, ev.green_env, ev.env_list,
                   oc.rule, oc.result])
    for col, w in zip("ABCDEFGHIJKLMN", [8, 32, 12, 12, 12, 24, 40, 24, 20, 10, 10, 24, 16, 8]):
        sh.column_dimensions[col].width = w
    sh.freeze_panes = "A2"
    wb.save(path)
    return path


def write_evidence_log(loans, outcomes, out_dir):
    path = os.path.join(out_dir, f"{_today()}-证据日志.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["序号", "经营主体", "借据号", "规则", "结果", "理由",
                         "预警通码", "企查查码", "绿色总数", "环境类", "名单"])
        for loan, oc in zip(loans, outcomes):
            ev = oc.evidence
            writer.writerow([loan.seq, loan.entity, loan.receipt, oc.rule, oc.result,
                             build_reason(oc), "/".join(ev.codes_yjt), "/".join(ev.codes_qcc),
                             ev.green_total, ev.green_env, ev.env_list])
    return path


def write_green_detail(green_rows, out_dir):
    """绿色匹配明细：按本期表企业首现顺序，前 6 列索引列 + 绿色明细原始 44 列。"""
    path = os.path.join(out_dir, f"{_today()}-绿色匹配明细.xlsx")
    wb = openpyxl.Workbook()
    sh = wb.active
    sh.title = "匹配明细"
    header_written = False
    for block in green_rows:
        if not header_written and block["rows"]:
            front = ["大表序号", "借款人名称", "经营主体名称", "匹配方式", "来源文件"]
            sh.append(front + list(block["rows"][0][0]))
            for c in sh[1]:
                c.font = Font(bold=True)
                c.fill = HDR_FILL
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            header_written = True
        for row_tuple, source_tag, match_how in block["rows"]:
            sh.append([block["seqs"], block["borrowers"], block["entity"], match_how, source_tag]
                      + list(row_tuple))
    if header_written:
        for i, w in enumerate([10, 24, 30, 14, 8] + [16] * 44, start=1):
            sh.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        sh.freeze_panes = "F2"
    wb.save(path)
    return path
