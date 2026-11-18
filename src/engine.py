# -*- coding: utf-8 -*-
"""判定引擎：按 P1沿用 / P2冲突复核 / P3证据判定 的优先级输出 是/否/不确定 + 证据结构。"""
import json
import os
import re
from dataclasses import dataclass, field

from .normalize import parse_industry_codes, norm_name

from .paths import CONFIG_DIR  # noqa: F401/E501  (路径统一由 src/paths.py 解析)


@dataclass
class AuditEvidence:
    seq: str = ""
    entity: str = ""
    codes_yjt: list = field(default_factory=list)
    codes_qcc: list = field(default_factory=list)
    eligible_hits: list = field(default_factory=list)
    excluded_hits: dict = field(default_factory=dict)
    strong_hits: dict = field(default_factory=dict)     # {小类: [命中词]}
    weak_hits: list = field(default_factory=list)
    c42_hits: list = field(default_factory=list)
    claim_mismatch: bool = False                        # E列申报行业与两平台登记不符
    green_total: int = 0
    green_env: int = 0
    green_types: dict = field(default_factory=dict)
    env_list: str = ""                                  # 名单类型+标签描述
    prev_result: str = ""
    prev_receipt_matched: bool = False
    rule: str = ""


@dataclass
class AuditOutcome:
    result: str            # 是 / 否 / 不确定
    rule: str              # P1 / P2 / P3-xxx
    evidence: AuditEvidence
    review_flag: str = ""  # 换借据再报等复核标记


def load_rules(config_dir: str = CONFIG_DIR):
    with open(os.path.join(config_dir, "directory.json"), encoding="utf-8") as f:
        directory = json.load(f)
    with open(os.path.join(config_dir, "exclusions.json"), encoding="utf-8") as f:
        exclusions = json.load(f)
    with open(os.path.join(config_dir, "scope_keywords.json"), encoding="utf-8") as f:
        keywords = json.load(f)
    return {
        "eligible": {k: v for k, v in directory["eligible"].items()},
        "exclusions": exclusions,
        "strong": keywords["strong"],
        "weak": keywords["weak"],
        "exclude_c42": keywords.get("exclude_c42", []),
        "green_env_keywords": keywords["green_env_keywords"],
        "green_env_threshold": keywords.get("green_env_threshold", 5),
    }


def _scope_hits(scope: str, rules) -> tuple:
    strong = {}
    for cat, words in rules["strong"].items():
        hits = [w for w in words if w in scope]
        if hits:
            strong[cat] = hits
    weak = [w for w in rules["weak"] if w in scope]
    c42 = [w for w in rules["exclude_c42"] if w in scope]
    return strong, weak, c42


def _env_list_desc(hit) -> str:
    if not hit:
        return ""
    tags = "、".join(hit.tags) if hit.tags else ""
    return f"{hit.list_type}（{tags}）" if tags else hit.list_type


def audit_loan(loan, plat, prev_result, prev_consistent, green_stat, env_hit, rules, reborrows):
    """对单笔贷款执行判定。

    :param loan: 本期 Loan
    :param plat: 该企业四列平台数据 dict
    :param prev_result: 借据号匹配到的上期结论（"是"/"否"/""）
    :param prev_consistent: 借据号匹配且三项字段一致
    :param green_stat: GreenStat
    :param env_hit: EnvListHit 或 None
    :param rules: load_rules() 结果
    :param reborrows: {entity_key: 上期该企业结论摘要} 换借据再报信息
    """
    ev = AuditEvidence(seq=loan.seq, entity=loan.entity)
    ev.codes_yjt = parse_industry_codes(plat.get("plat_yjt", ""))
    ev.codes_qcc = parse_industry_codes(plat.get("plat_qcc_gb", ""))
    all_codes = list(dict.fromkeys(ev.codes_yjt + ev.codes_qcc))
    ev.eligible_hits = [c for c in all_codes if c in rules["eligible"]]
    ev.excluded_hits = {c: rules["exclusions"][c] for c in all_codes if c in rules["exclusions"]}
    claim_codes = parse_industry_codes(loan.claim_industry)
    if claim_codes and all(c not in all_codes for c in claim_codes):
        ev.claim_mismatch = True

    scope = plat.get("plat_scope", "") or ""
    ev.strong_hits, ev.weak_hits, ev.c42_hits = _scope_hits(scope, rules)
    ev.green_total = green_stat.total if green_stat else 0
    ev.green_env = green_stat.env_count if green_stat else 0
    ev.green_types = dict(green_stat.types) if green_stat else {}
    ev.env_list = _env_list_desc(env_hit)

    ev.prev_receipt_matched = bool(prev_consistent)
    ev.prev_result = prev_result if prev_consistent else ""

    review = ""
    reg_key = getattr(loan, "reg_key", None) or loan.entity_key
    if reg_key in reborrows and not ev.prev_result:
        review = f"换借据再报：{reborrows[reg_key]}"

    # ---- P1 沿用 ----
    if ev.prev_result in ("是", "否"):
        return AuditOutcome(ev.prev_result, "P1-沿用", ev, review)

    green_pass = ev.green_env >= rules["green_env_threshold"]
    has_plat_data = bool(ev.codes_yjt or ev.codes_qcc or scope)

    # ---- 无平台数据（四列未填）：不判定 ----
    if not has_plat_data:
        return AuditOutcome("不确定", "P0-待补平台数据", ev, review or "四列平台数据未填写")

    # ---- 排除码处理 ----
    if ev.excluded_hits and ev.eligible_hits:
        return AuditOutcome("不确定", "P2-两平台冲突", ev, review)
    if ev.excluded_hits:
        if ev.strong_hits:
            return AuditOutcome("不确定", "P2-排除码冲突", ev, review)
        return AuditOutcome("是", "P3-排除码", ev, review)

    # ---- P3 证据判定 ----
    if ev.eligible_hits:
        return AuditOutcome("否", "P3-行业目录内", ev, review)
    if ev.strong_hits and ev.c42_hits:
        return AuditOutcome("不确定", "P3-固废加工存疑", ev, review)
    if ev.strong_hits:
        return AuditOutcome("否", "P3-经营范围涉及", ev, review)
    if ev.weak_hits and (green_pass or ev.env_list):
        return AuditOutcome("否", "P3-弱表述加佐证", ev, review)
    if ev.weak_hits:
        return AuditOutcome("不确定", "P3-弱表述存疑", ev, review)
    if green_pass and ev.env_list:
        return AuditOutcome("否", "P3-绿色加名单", ev, review)
    if green_pass:
        return AuditOutcome("不确定", "P3-仅绿色线索", ev, review)
    if ev.env_list:
        return AuditOutcome("不确定", "P3-仅名单", ev, review)
    return AuditOutcome("是", "P3-无目录证据", ev, review)


def _fields_consistent(cur, prev) -> bool:
    return (
        norm_name(cur.date) == norm_name(prev.date)
        and norm_name(cur.borrower) == norm_name(prev.borrower)
        and norm_name(cur.entity) == norm_name(prev.entity)
    )
