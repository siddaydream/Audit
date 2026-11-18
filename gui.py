# -*- coding: utf-8 -*-
"""贴息审核助手 —— 图形界面（全中文、弹窗选文件、双击运行）。

源码运行:  python3 gui.py
打包:      pyinstaller --noconfirm --windowed --name 贴息审核助手 gui.py
           （可执行文件需放在项目文件夹内，与 config/ data/ output/ 同级）
"""
import json
import os
import queue
import subprocess
import sys
import threading
import traceback

def base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, base_dir())

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from src import pipeline, validator
from src.paths import OUT_DIR

CFG_FILE = os.path.join(base_dir(), "gui_config.json")
XLSX = [("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]


def open_path(path, reveal_only=False):
    """用系统默认方式打开文件或文件夹。"""
    if not path or not os.path.exists(path):
        return
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path] if reveal_only else ["open", path])
        elif os.name == "nt":
            if reveal_only:
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
            else:
                os.startfile(path)  # noqa
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


class QueueLog:
    """把 pipeline 的日志写入队列，供界面刷新。"""
    def __init__(self, q):
        self.q = q

    def __call__(self, msg):
        self.q.put(str(msg))


class RedirectStdout:
    """把 print 输出转入日志队列（工作线程内生效）。"""
    def __init__(self, log):
        self.log = log

    def write(self, text):
        if text.strip():
            self.log(text.strip())

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title("贴息审核助手")
        root.geometry("860x620")
        self.q = queue.Queue()
        self.working = False
        self.mid_path = ""
        self.result_path = ""
        self.cfg = self.load_cfg()

        style = ttk.Style(root)
        if "aqua" in style.theme_names():
            style.theme_use("aqua")

        # ---- 顶部进度 ----
        top = ttk.Frame(root, padding=(10, 8))
        top.pack(fill="x")
        self.step_var = tk.StringVar(value="流程：① 预填（自动） → ② 填新增企业四列（人工） → ③ 审核（自动）")
        ttk.Label(top, textvariable=self.step_var, font=("", 12, "bold")).pack(side="left")

        body = ttk.Frame(root, padding=(10, 0))
        body.pack(fill="both", expand=True)

        # ---- 阶段① ----
        f1 = ttk.LabelFrame(body, text=" 第①步 预填（本期表 + 上期表 → 中间表） ", padding=8)
        f1.pack(fill="x", pady=4)
        self.cur_var = tk.StringVar(value=self.cfg.get("current", ""))
        self.prev_var = tk.StringVar(value=self.cfg.get("prev", ""))
        ttk.Label(f1, text="本期表：").grid(row=0, column=0, sticky="w")
        ttk.Entry(f1, textvariable=self.cur_var).grid(row=0, column=1, sticky="we", padx=4)
        ttk.Button(f1, text="选择…", width=8, command=lambda: self.pick("current")).grid(row=0, column=2)
        ttk.Label(f1, text="上期表：").grid(row=1, column=0, sticky="w")
        ttk.Entry(f1, textvariable=self.prev_var).grid(row=1, column=1, sticky="we", padx=4, pady=2)
        ttk.Button(f1, text="选择…", width=8, command=lambda: self.pick("prev")).grid(row=1, column=2)
        self.btn_prefill = ttk.Button(f1, text="开始预填", command=self.on_prefill)
        self.btn_prefill.grid(row=0, column=3, rowspan=2, padx=10)
        f1.columnconfigure(1, weight=1)

        # ---- 阶段② ----
        f2 = ttk.LabelFrame(body, text=" 第②步 人工填列（只填新增企业） ", padding=8)
        f2.pack(fill="x", pady=4)
        self.hint_var = tk.StringVar(value="运行预填后，这里会列出需要手工填写四列的新增企业。")
        ttk.Label(f2, textvariable=self.hint_var, wraplength=760, justify="left").pack(anchor="w")
        ttk.Button(f2, text="打开中间表", command=lambda: open_path(self.mid_path)).pack(anchor="w", pady=2)

        # ---- 阶段③ ----
        f3 = ttk.LabelFrame(body, text=" 第③步 审核（中间表 → 审核结果表） ", padding=8)
        f3.pack(fill="x", pady=4)
        self.mid_var = tk.StringVar(value="")
        ttk.Label(f3, text="中间表：").grid(row=0, column=0, sticky="w")
        ttk.Entry(f3, textvariable=self.mid_var).grid(row=0, column=1, sticky="we", padx=4)
        ttk.Button(f3, text="选择…", width=8, command=lambda: self.pick("mid")).grid(row=0, column=2)
        self.btn_audit = ttk.Button(f3, text="开始审核（约20秒）", command=self.on_audit)
        self.btn_audit.grid(row=0, column=3, padx=10)
        f3.columnconfigure(1, weight=1)

        # ---- 工具行 ----
        tools = ttk.Frame(body)
        tools.pack(fill="x", pady=2)
        ttk.Button(tools, text="登记参考库（首次使用点我）", command=self.on_init_data).pack(side="left")
        ttk.Button(tools, text="打开输出文件夹", command=lambda: open_path(OUT_DIR)).pack(side="left", padx=6)
        self.btn_open_result = ttk.Button(tools, text="打开审核结果表", command=lambda: open_path(self.result_path), state="disabled")
        self.btn_open_result.pack(side="left")

        # ---- 日志 ----
        lf = ttk.LabelFrame(body, text=" 运行日志 ", padding=4)
        lf.pack(fill="both", expand=True, pady=6)
        self.log_box = scrolledtext.ScrolledText(lf, height=10, state="disabled", font=("Menlo" if sys.platform == "darwin" else "Consolas", 12))
        self.log_box.pack(fill="both", expand=True)

        root.after(100, self.poll_queue)

    # ---------- 基础 ----------
    def load_cfg(self):
        try:
            with open(CFG_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def save_cfg(self):
        self.cfg.update({"current": self.cur_var.get(), "prev": self.prev_var.get(),
                         "mid": self.mid_var.get()})
        try:
            with open(CFG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", str(msg) + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def qlog(self, msg):
        """供工作线程使用的日志函数。"""
        self.q.put(str(msg))

    def poll_queue(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if isinstance(msg, tuple):
                    kind, value = msg
                    if kind == "__HINT__":
                        self.hint_var.set(value)
                    elif kind == "__RESULT__":
                        self.result_path = value
                        self.btn_open_result.configure(state="normal")
                        self.step_var.set("流程完成：结果表已生成，请复核“不确定”行后定稿")
                    elif kind == "__ERROR__":
                        messagebox.showerror("运行出错", value[-800:])
                elif msg == "__DONE__":
                    self.set_busy(False)
                else:
                    self.log(msg)
        except queue.Empty:
            pass
        self.root.after(100, self.poll_queue)

    def pick(self, which):
        path = filedialog.askopenfilename(title="选择文件", filetypes=XLSX,
                                          initialdir=base_dir())
        if not path:
            return
        if which == "current":
            self.cur_var.set(path)
        elif which == "prev":
            self.prev_var.set(path)
        elif which == "mid":
            self.mid_var.set(path)
        self.save_cfg()

    def set_busy(self, busy):
        self.working = busy
        state = "disabled" if busy else "normal"
        for b in (self.btn_prefill, self.btn_audit):
            b.configure(state=state)

    def run_worker(self, fn):
        def wrapped():
            old_stdout = sys.stdout
            sys.stdout = RedirectStdout(QueueLog(self.q))
            try:
                fn()
            except Exception:
                err = traceback.format_exc()
                self.qlog("【出错】" + err)
                self.q.put(("__ERROR__", err))
            finally:
                sys.stdout = old_stdout
                self.q.put("__DONE__")
        self.set_busy(True)
        threading.Thread(target=wrapped, daemon=True).start()

    # ---------- 动作 ----------
    def on_prefill(self):
        cur, prev = self.cur_var.get().strip(), self.prev_var.get().strip()
        if not cur or not os.path.exists(cur):
            messagebox.showwarning("提示", "请先选择本期表")
            return
        if not prev or not os.path.exists(prev):
            messagebox.showwarning("提示", "请先选择上期表（上期审核产出）")
            return
        self.save_cfg()
        self.log(f"—— 开始预填 ——\n本期表: {cur}\n上期表: {prev}")
        self.hint_var.set("预填运行中…")

        def job():
            mid, new = pipeline.run_prefill(cur, prev, OUT_DIR, log=self.qlog)
            self.mid_path = mid
            items = pipeline.collect_new_entities(new, log=self.qlog)
            if items:
                names = "、".join(n for n, _, _ in items[:6]) + ("…" if len(items) > 6 else "")
                self.q.put(("__HINT__", f"待人工填写 {len(items)} 家新增企业：{names}。"
                                       f"请打开中间表填好黄色格子（四列）并保存，再回到第③步点“开始审核”。"))
            else:
                self.q.put(("__HINT__", "无新增企业，四列已全部回填，直接点第③步“开始审核”。"))

        self.run_worker(job)

    def on_audit(self):
        mid = self.mid_var.get().strip() or self.mid_path
        if not mid or not os.path.exists(mid):
            messagebox.showwarning("提示", "请先选择填写完成的中间表")
            return
        if mid.endswith(".xlsx") and not os.path.exists(os.path.splitext(mid)[0] + ".carry.json"):
            self.log("【提示】未找到中间表旁边的 .carry.json 沿用数据文件，"
                     "若该中间表不是本工具预填生成的，结论沿用将不可用。")
        self.log(f"—— 开始审核 ——\n中间表: {mid}")

        def job():
            outputs, _ = pipeline.run_audit(mid, out_dir=OUT_DIR, log=self.qlog)
            self.result_path = outputs["result"]
            self.q.put(("__RESULT__", outputs["result"]))

        self.run_worker(job)

    def on_init_data(self):
        self.log("—— 登记参考库 ——")

        def job():
            manifest = validator.init_manifest()
            for role, entry in manifest["files"].items():
                self.qlog(f"  {role}: {entry['path']}")
            missing = {"green_1", "green_2", "env_list", "policy", "gb_notes"} - set(manifest["files"])
            if missing:
                self.qlog(f"【注意】缺少: {missing}，请把五份参考文件放入 data/ 文件夹后重新登记。")

        self.run_worker(job)


def main():
    # 调试/自检钩子：SUBSIDY_SELFTEST=中间表路径 运行打包版全流程自检
    if os.environ.get("SUBSIDY_SELFTEST"):
        from src import paths, pipeline
        with open("/tmp/subsidy_diag.txt", "w", encoding="utf-8") as f:
            f.write(f"exe={sys.executable}\n")
            f.write(f"frozen={getattr(sys, 'frozen', None)}\n")
            f.write(f"cwd={os.getcwd()}\n")
            f.write(f"base={paths.base_dir()}\n")
            f.write(f"cfg={paths.CONFIG_DIR} exists={os.path.isdir(paths.CONFIG_DIR)}\n")
            f.write(f"data={paths.DATA_DIR} exists={os.path.isdir(paths.DATA_DIR)}\n")
        outputs, _ = pipeline.run_audit(os.environ["SUBSIDY_SELFTEST"])
        with open("/tmp/subsidy_diag.txt", "a", encoding="utf-8") as f:
            f.write(f"result={outputs['result']}\n")
        sys.exit(0)
    root = tk.Tk()
    app = App(root)  # noqa: F841
    root.mainloop()


if __name__ == "__main__":
    main()
