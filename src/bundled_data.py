"""从仓库内的分块还原大型参考表；不覆盖已有参考库。"""
import hashlib
import json
import os
from pathlib import Path
import tempfile


def restore_bundled_data(data_dir):
    root = Path(data_dir)
    bundle = root / "bundled"
    index = bundle / "index.json"
    if not index.exists():
        return
    for entry in json.loads(index.read_text(encoding="utf-8")):
        name = entry["path"]
        if Path(name).name != name:
            raise ValueError("参考库文件名无效")
        target = root / name
        if target.exists():
            continue
        digest = hashlib.sha256()
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(dir=root, delete=False) as out:
                temp_path = Path(out.name)
                for part in entry["parts"]:
                    if Path(part).name != part:
                        raise ValueError("参考库分块名称无效")
                    with (bundle / part).open("rb") as source:
                        for chunk in iter(lambda: source.read(1024 * 1024), b""):
                            digest.update(chunk)
                            out.write(chunk)
            if digest.hexdigest() != entry["sha256"]:
                raise ValueError(f"参考库分块校验失败，请重新下载完整项目：{name}")
            os.replace(temp_path, target)
            print(f"[参考库] 已自动还原：{name}")
        finally:
            if temp_path is not None and temp_path.exists():
                temp_path.unlink()
