# -*- coding: utf-8 -*-
"""输入校验与参考库 manifest 管理。"""
import hashlib
import json
import os

from .bundled_data import restore_bundled_data
from .paths import DATA_DIR  # noqa: F401/E501  (路径统一由 src/paths.py 解析)

# 参考库角色 -> 文件名识别关键词（init-data 自动探测用）
FILE_ROLES = {
    "green_1": "绿色能效",
    "green_2": "绿色能效",   # 两份绿色明细同名模式，按文件名排序取前两个
    "env_list": "环保关键名单",
    "policy": "政策审核要求",
    "gb_notes": "国民经济行业分类",
}


def sha256_of(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def detect_role(filename: str) -> str:
    for role, kw in FILE_ROLES.items():
        if kw in filename and filename.lower().endswith((".xlsx", ".xls")):
            return role
    return ""


def init_manifest(data_dir: str = DATA_DIR) -> dict:
    """扫描 data/ 目录，识别五份参考文件并生成 manifest.json。"""
    restore_bundled_data(data_dir)
    greens = []
    manifest = {"files": {}, "created": _today()}
    for fname in sorted(os.listdir(data_dir)):
        if not fname.lower().endswith((".xlsx", ".xls")) or fname.startswith("~$"):
            continue
        role = detect_role(fname)
        path = os.path.join(data_dir, fname)
        entry = {"path": fname, "sha256": sha256_of(path)}
        if role in ("green_1", "green_2"):
            greens.append(entry)
        elif role:
            manifest["files"][role] = entry
    for i, entry in enumerate(greens[:2], start=1):
        manifest["files"][f"green_{i}"] = entry
    with open(os.path.join(data_dir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return manifest


def load_manifest(data_dir: str = DATA_DIR, strict: bool = False) -> dict:
    restore_bundled_data(data_dir)
    path = os.path.join(data_dir, "manifest.json")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"未找到 {path}。请先把五份参考文件放入 data/ 目录并运行: python run.py init-data")
    with open(path, encoding="utf-8") as f:
        manifest = json.load(f)
    missing = []
    for role, entry in manifest["files"].items():
        fpath = os.path.join(data_dir, entry["path"])
        if not os.path.exists(fpath):
            missing.append(f"{role}: {entry['path']}")
        elif strict and sha256_of(fpath) != entry["sha256"]:
            missing.append(f"{role}: {entry['path']} (内容与登记哈希不符，可能已更新)")
    if missing:
        msg = "参考库校验未通过:\n  " + "\n  ".join(missing)
        if strict:
            raise RuntimeError(msg)
        print("[提示] " + msg)
    return manifest


def validate_form(loans, header_map, source: str):
    """申请表必备列校验。"""
    required = ["seq", "date", "borrower", "entity", "id_number", "receipt"]
    missing = [COLUMN_HINTS_FULL[k] for k in required if k not in header_map]
    if missing:
        raise ValueError(f"{source} 缺少必备列: {','.join(missing)}")
    if not loans:
        raise ValueError(f"{source} 未读取到数据行")
    empty_receipt = sum(1 for l in loans if not l.receipt)
    if empty_receipt:
        print(f"[提示] {source} 有 {empty_receipt} 行贷款借据号为空，这些行无法沿用上期结论")


COLUMN_HINTS_FULL = {
    "seq": "序号", "date": "数据日期", "borrower": "借款人名称", "entity": "经营主体名称",
    "id_number": "统一社会信用代码/身份证号码", "receipt": "贷款借据号",
}


def _today() -> str:
    import datetime
    return datetime.date.today().isoformat()
