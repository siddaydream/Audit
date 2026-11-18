# -*- coding: utf-8 -*-
"""理由文本生成：沿用行简写，新增行按命中规则写证据详述。"""
from .engine import AuditOutcome


def _code_str(ev) -> str:
    parts = []
    if ev.codes_yjt:
        parts.append("预警通" + "/".join(ev.codes_yjt))
    if ev.codes_qcc:
        parts.append("企查查" + "/".join(ev.codes_qcc))
    return "、".join(parts) if parts else "（未解析到行业码）"


def _green_str(ev) -> str:
    if ev.green_total == 0:
        return "无绿色经营线索"
    top = "、".join(f"{k}x{v}" for k, v in sorted(ev.green_types.items(), key=lambda x: -x[1])[:3])
    return f"绿色线索{ev.green_total}条（{top}），其中环境类{ev.green_env}条"


def _list_str(ev) -> str:
    return f"在环保名单：{ev.env_list}" if ev.env_list else "不在环保名单"


def _mismatch_note(ev) -> str:
    return "；E列申报行业与登记行业不符" if ev.claim_mismatch else ""


def build_reason(outcome: AuditOutcome) -> str:
    ev = outcome.evidence
    rule = outcome.rule

    if rule == "P1-沿用":
        return f"上一期判{ev.prev_result}"

    if rule == "P0-待补平台数据":
        return "四列平台数据未填写，待人工补充后重跑审核"

    if rule == "P2-两平台冲突":
        return (f"两平台国标行业冲突：{_code_str(ev)}，其一落在排除码（"
                f"{'；'.join(ev.excluded_hits.values())}），需人工核实实际业务")

    if rule == "P2-排除码冲突":
        excl = "；".join(ev.excluded_hits.values())
        return (f"国标行业为{_code_str(ev)}，按分类注释{excl}；"
                f"但经营范围含目录活动（{ _strong_str(ev)}），需人工核实主营业务")

    if rule == "P3-排除码":
        excl = "；".join(ev.excluded_hits.values())
        return (f"1、国标行业为{_code_str(ev)}，按分类注释{excl}，不属于贴息目录；"
                f"经营范围未涉及其他目录活动{_mismatch_note(ev)}。2、{_green_str(ev)}，{_list_str(ev)}。")

    if rule == "P3-行业目录内":
        cats = "、".join(ev.eligible_hits)
        return (f"国标行业为{_code_str(ev)}，属目录内（{cats}）；经营范围涉及环保相关业务；"
                f"{_green_str(ev)}，{_list_str(ev)}{_mismatch_note(ev)}。")

    if rule == "P3-经营范围涉及":
        return (f"登记行业{_code_str(ev)}不在目录，但经营范围明确含目录活动（{_strong_str(ev)}）；"
                f"{_green_str(ev)}，{_list_str(ev)}{_mismatch_note(ev)}。")

    if rule == "P3-固废加工存疑":
        return (f"登记行业{_code_str(ev)}不在目录；经营范围含目录表述（{_strong_str(ev)}），"
                f"但同时含{'、'.join(ev.c42_hits)}（按注释属废弃资源加工/冶炼口径），"
                f"实际是否从事目录内7723活动需人工核实。")

    if rule == "P3-弱表述加佐证":
        return (f"登记行业{_code_str(ev)}不在目录，经营范围含{'、'.join(ev.weak_hits)}等目录相关表述，"
                f"且有绿色线索或名单佐证（环境类绿色线索{ev.green_env}条；{_list_str(ev)}），判定未发现不符合。")

    if rule == "P3-弱表述存疑":
        return (f"登记行业{_code_str(ev)}不在目录；经营范围仅含{'、'.join(ev.weak_hits)}等泛表述，"
                f"{_green_str(ev)}，{_list_str(ev)}，实际是否从事目录活动需人工核实{_mismatch_note(ev)}。")

    if rule == "P3-仅绿色线索":
        return (f"登记行业{_code_str(ev)}不在目录，经营范围未见明确目录活动表述，"
                f"但环境类绿色线索{ev.green_env}条（共{ev.green_total}条），可能实际从事目录活动，"
                f"建议核实绿色招投标中标角色后认定。")

    if rule == "P3-仅名单":
        return (f"登记行业{_code_str(ev)}不在目录，经营范围未见明确目录活动表述，{_list_str(ev)}，"
                f"实际业务需人工核实。")

    if rule == "P3-绿色加名单":
        return (f"登记行业{_code_str(ev)}不在目录，但{_list_str(ev)}且环境类绿色线索{ev.green_env}条，判定未发现不符合。")

    # P3-无目录证据
    return (f"1、预警通/企查查国标行业为{_code_str(ev)}，不属于贴息目录；经营范围未涉及目录内活动"
            f"{_mismatch_note(ev)}。2、{_green_str(ev)}，{_list_str(ev)}。")


def _strong_str(ev) -> str:
    parts = [f"{cat}({'、'.join(words)})" for cat, words in ev.strong_hits.items()]
    return "；".join(parts) if parts else "无"


def review_point(outcome: AuditOutcome) -> str:
    """不确定复核清单中的复核要点。"""
    reason = build_reason(outcome)
    return f"[{outcome.rule}] {reason}"
