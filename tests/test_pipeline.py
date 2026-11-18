# -*- coding: utf-8 -*-
"""单元测试：用合成数据覆盖 预填回填/新增留空 与 判定引擎各规则分支。
运行：python -m unittest discover tests  （在项目根目录）"""
import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

from src import engine, matcher


def build_form(path, rows, with_platform=False, with_prev_xy=False):
    """rows: list of dict(seq,date,borrower,entity,id,receipt,plat_yjt?,...,X?,Y?)"""
    headers = ["序号", "银行名称", "数据日期", "所属行业领域", "行业代码及名称", "借款人名称",
               "经营主体名称", "借款人或经营场所所在省份", "统一社会信用代码/身份证号码",
               "贷款借据号", "贷款发放\n金额（万元）", "贷款利率（%）", "贷款合同签订日期\n（具体到日）",
               "贷款发放日\n（具体到日）", "贷款到期日\n（具体到日）", "贷款期限\n（个月）",
               "贴息期限\n（个月）", "贷款用途", "符合条件的贴息贷款金额\n（万元）",
               "是否为本月新增", "是否申请服务消费与养老再贷款", "审核的省级行业管理部门名称",
               "备注", "是否发现存在不符合贴息条件情况", "简要描述审核发现的问题",
               "报送日期", "审批日期", "审核状态"]
    if with_platform:
        headers += ["国标行业（预警通）", "经营范围（企查查）", "企查查行业", "国标行业（企查查）"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["服务业经营主体贷款贴息审核申请表"])
    ws.append(headers)
    for r in rows:
        base = [r.get("seq", ""), "测试银行", r.get("date", "2026年3季度"), "健康",
                r.get("claim", "772环境治理业"), r.get("borrower", ""), r.get("entity", ""),
                "上海市", r.get("id", ""), r.get("receipt", ""), 100, 3.0, "2026-01-01",
                "2026-01-01", "2027-01-01", 12, 12, "测试", 100, "否", "否", "市生态环境局",
                "", r.get("X", ""), r.get("Y", ""), "", "", ""]
        if with_platform:
            base += [r.get("plat_yjt", ""), r.get("plat_scope", ""), r.get("qcc_ind", ""),
                     r.get("qcc_gb", "")]
        ws.append(base)
    wb.save(path)


def build_green(path, rows):
    headers = ["企业名称", "统一社会信用代码", "类型", "注册资本", "币种", "成立日期",
               "标准主体名称", "登记机关", "一级行业", "二级行业", "三级行业", "四级行业",
               "一级目录", "二级目录", "三级目录", "绿色类型", "绿色产品类型", "绿色评级",
               "打分", "经营范围", "绿色目录代码", "绿色目录名称"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(headers)
    for r in rows:
        row = [""] * 22
        row[0] = r.get("name", "")
        row[1] = r.get("code", "")
        row[6] = r.get("name", "")
        row[15] = r.get("gtype", "绿色专利")
        row[21] = r.get("catalog", "")
        ws.append(row)
    wb.save(path)


def build_env_list(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "总名单1017"
    ws.append(["序号", "单位名称", "统一社会信用代码", "分类结果", "来源", "名单类型"] +
              [f"标签{i}" for i in range(11)])
    for i, r in enumerate(rows, start=1):
        ws.append([i, r.get("name", ""), r.get("code", ""), "", 4, r.get("type", "")] +
                  [r.get("tag", "")] + [""] * 10)
    wb.save(path)


RULES = engine.load_rules(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config"))


class TestEngine(unittest.TestCase):
    def _loan(self, **kw):
        return matcher.Loan(row=3, seq="1", date="2026年3季度", borrower=kw.get("entity", "甲公司"),
                            entity=kw.get("entity", "甲公司"), id_number=kw.get("id", "91310000TEST0000AA"),
                            receipt=kw.get("receipt", "R001"), claim_industry=kw.get("claim", ""))

    def _plat(self, yjt="", scope="", qcc_gb="", qcc_ind=""):
        return {"plat_yjt": yjt, "plat_scope": scope, "plat_qcc_ind": qcc_ind, "plat_qcc_gb": qcc_gb}

    def _green(self, env_count):
        from src.evidence import GreenStat
        st = GreenStat()
        st.total = env_count
        st.env_count = env_count
        return st

    def test_p1_carryover(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("水污染治理(N7721)"), "否", True,
                               None, None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("否", "P1-沿用"))

    def test_industry_eligible(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("水污染治理(N7721)", "环保业务"), "", True,
                               None, None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("否", "P3-行业目录内"))

    def test_exclusion_code(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("污水处理及其再生利用(D4620)", "污水净化处理与再生利用",
                                                qcc_gb="污水处理及其再生利用（D4620）"), "", True,
                               None, None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("是", "P3-排除码"))

    def test_exclusion_conflict(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("污水处理及其再生利用(D4620)", "固体废物治理",
                                                qcc_gb="水污染治理（N7721）"), "", True,
                               None, None, RULES, {})
        self.assertEqual(oc.result, "不确定")

    def test_platform_conflict(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("污水处理及其再生利用(D4620)",
                                                qcc_gb="水污染治理（N7721）"), "", True,
                               None, None, RULES, {})
        self.assertEqual(oc.rule, "P2-两平台冲突")

    def test_strong_scope(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("其他科技推广服务业(M7590)",
                                                "技术开发；水污染治理；设备销售"), "", True,
                               None, None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("否", "P3-经营范围涉及"))

    def test_c42_conflict(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("其他未列明专业技术服务业(M7499)",
                                                "固体废物治理；金属废料和碎屑加工处理"), "", True,
                               None, None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("不确定", "P3-固废加工存疑"))

    def test_weak_only_uncertain(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("其他未列明专业技术服务业(M7499)",
                                                "技术开发；水环境污染防治服务"), "", True,
                               None, None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("不确定", "P3-弱表述存疑"))

    def test_weak_with_green_pass(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("其他未列明专业技术服务业(M7499)",
                                                "技术开发；土壤生态修复"), "", True,
                               self._green(10), None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("否", "P3-弱表述加佐证"))

    def test_green_only_uncertain(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("应用软件开发(I6513)", "软件开发"), "", True,
                               self._green(30), None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("不确定", "P3-仅绿色线索"))

    def test_no_evidence_fail(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("其他未列明建筑业(E5090)",
                                                "新型建筑材料制造；园林绿化工程施工"), "", True,
                               self._green(1), None, RULES, {})
        self.assertEqual((oc.result, oc.rule), ("是", "P3-无目录证据"))

    def test_missing_platform_data(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat(), "", False, None, None, RULES, {})
        self.assertEqual(oc.result, "不确定")

    def test_reborrow_flag(self):
        loan = self._loan()
        oc = engine.audit_loan(loan, self._plat("水污染治理(N7721)"), "", False,
                               None, None, RULES, {loan.entity_key: "上期同企业借据判“是”"})
        self.assertIn("换借据再报", oc.review_flag)


class TestPrefill(unittest.TestCase):
    def test_prefill_fill_and_blank(self):
        from src import prefill as pf
        from src.normalize import norm_name
        with tempfile.TemporaryDirectory() as td:
            prev = os.path.join(td, "prev.xlsx")
            cur = os.path.join(td, "cur.xlsx")
            build_form(prev, [
                dict(seq="1", entity="老企业环保有限公司", id="91310000AAAAAAAA1A",
                     receipt="P001", plat_yjt="水污染治理(N7721)", plat_scope="水污染治理",
                     qcc_gb="水污染治理（N7721）", qcc_ind="环保", X="否", Y="/"),
            ], with_platform=True)
            build_form(cur, [
                dict(seq="1", entity="老企业环保有限公司", id="91310000AAAAAAAA1A", receipt="N001"),
                dict(seq="2", entity="新企业（上海）环保科技有限公司", id="91310000BBBBBBBB2B",
                     receipt="N002"),
            ])
            mid, new = pf.prefill(cur, prev, td)
            wb = openpyxl.load_workbook(mid)
            ws = wb.worksheets[0]
            # 老企业四列已回填
            self.assertEqual(ws.cell(3, 29).value, "水污染治理(N7721)")
            self.assertEqual(ws.cell(3, 30).value, "水污染治理")
            # 新企业留空且高亮
            self.assertIsNone(ws.cell(4, 29).value)
            self.assertEqual(ws.cell(4, 29).fill.fgColor.rgb, "00FFF2CC")
            # sidecar 存在且记录沿用
            with open(os.path.splitext(mid)[0] + ".carry.json", encoding="utf-8") as f:
                carry = json.load(f)
            self.assertEqual(len(carry), 2)
            # 新增清单有 1 行
            wb2 = openpyxl.load_workbook(new)
            sh = wb2.worksheets[0]
            self.assertEqual(sh.max_row, 2)
            self.assertIn("新企业", sh.cell(2, 1).value)

    def test_fullwidth_halwidth_name_match(self):
        from src.normalize import norm_name
        self.assertEqual(norm_name("堡森（上海）新材料"), norm_name("堡森(上海)新材料"))


if __name__ == "__main__":
    unittest.main()


class TestUSCC(unittest.TestCase):
    def test_id_card_excluded(self):
        from src.normalize import is_uscc
        self.assertFalse(is_uscc("321283198004230027"))   # 身份证
        self.assertFalse(is_uscc("330323196304060024"))   # 身份证
        self.assertTrue(is_uscc("91310120669421165L"))    # 信用代码
        self.assertTrue(is_uscc("91310000747640452F"))
        self.assertTrue(is_uscc("91441322MA4WCPT28C"))
