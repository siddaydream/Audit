# -*- coding: utf-8 -*-
"""服务业经营主体贷款贴息审核助手 —— 命令行入口（图形界面见 gui.py）。

用法：
  python run.py init-data                                 # 首次：登记 data/ 下的参考库
  python run.py prefill --current 本期表.xlsx --prev 上期产出.xlsx
  python run.py audit   --current 中间表.xlsx [--prev 上期产出.xlsx]
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src import pipeline, validator  # noqa: E402
from src.paths import OUT_DIR  # noqa: E402


def cmd_init_data(_args):
    manifest = validator.init_manifest()
    print("[init-data] 参考库登记完成:")
    for role, entry in manifest["files"].items():
        print(f"  {role:10s} {entry['path']}")
    need = {"green_1", "green_2", "env_list", "policy", "gb_notes"}
    missing = need - set(manifest["files"])
    if missing:
        print(f"[init-data] 注意：缺少角色 {missing}，请确认五份参考文件均已放入 data/ 目录")


def cmd_prefill(args):
    mid_path, new_path = pipeline.run_prefill(args.current, args.prev, args.outdir or OUT_DIR)
    pipeline.collect_new_entities(new_path)


def cmd_audit(args):
    pipeline.run_audit(args.current, prev_path=args.prev, out_dir=args.outdir or OUT_DIR,
                       strict=args.strict, green_detail=args.green_detail)


def main():
    parser = argparse.ArgumentParser(description="服务业经营主体贷款贴息审核助手")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("init-data", help="扫描 data/ 生成参考库 manifest.json")

    p2 = sub.add_parser("prefill", help="阶段1：生成中间表与新增企业清单")
    p2.add_argument("--current", required=True, help="本期申请表（可无平台四列）")
    p2.add_argument("--prev", required=True, help="上期产出表（含平台四列与结论）")
    p2.add_argument("--outdir", default=None)

    p3 = sub.add_parser("audit", help="阶段3：输出审核结果表等")
    p3.add_argument("--current", required=True, help="填写完成的中间表")
    p3.add_argument("--prev", default=None, help="上期产出表（无 sidecar 时用于结论沿用）")
    p3.add_argument("--outdir", default=None)
    p3.add_argument("--strict", action="store_true", help="严格校验参考库哈希")
    p3.add_argument("--green-detail", action="store_true",
                    help="额外输出绿色匹配明细（需二次扫描绿色明细，耗时增加）")

    args = parser.parse_args()
    {"init-data": cmd_init_data, "prefill": cmd_prefill, "audit": cmd_audit}[args.command](args)


if __name__ == "__main__":
    main()
