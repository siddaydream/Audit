# Audit · 贴息审核助手

本地运行的服务业经营主体贷款贴息审核辅助程序，支持 Windows、macOS 和 Linux。先回填上期企业信息，再人工补齐新增企业的平台数据，最后生成审核结果及复核清单。

**[下载完整 ZIP](https://github.com/siddaydream/Audit/archive/refs/heads/main.zip)** · [详细使用说明](docs/使用说明.md) · [参考库说明](data/README.md)

## 下载与安装

下载 ZIP 后完整解压，进入 `Audit-main` 文件夹。安装 Python 3.10 或更新版本；图形界面还需要 Tkinter。请保留整个目录结构。

### Windows

在项目文件夹打开 PowerShell：

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe gui.py
```

以后启动只需运行最后一行。

### macOS / Linux

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python gui.py
```

如提示缺少 Tkinter，请为当前 Python 安装对应的 Tkinter 支持。

## 五份参考库已包含

无需额外复制参考表或做首次登记。三份较小的 Excel 位于 `data/`；两份绿色能效明细保存在 `data/bundled/`，首次审核时自动还原为原始 Excel，并校验 SHA256。已有同名表格不会被覆盖。

下载包约 63 MB。请完整下载，不要只下载 Python 文件，也不要删除 `data/bundled/`。依赖安装需要联网，审核在本地进行。

## 每期使用

1. 选择上期表和本期原始表，运行预填。
2. 在中间表中补齐黄色空白的平台数据，保存。
3. 运行审核，在 `output/` 查看结果和复核清单。

表格约定：第一张工作表，第 1 行表名、第 2 行列名、第 3 行起为数据。本期原始表不需要提前添加平台四列。

**现有版本限制：** 上期结论的读取列名与程序输出列名不一致；本期已有平台四列时也会重复追加。滚动审核前请阅读[表格要求和处理办法](docs/使用说明.md#第一步的表格要求当前版本)。本次迁移没有改变判定规则。工具只提供辅助初判，最终结果应由人工复核。

## 给其他设备上的 AI

> 下载 https://github.com/siddaydream/Audit/archive/refs/heads/main.zip 并完整解压。阅读根目录 README.md，按当前操作系统安装依赖、启动 gui.py。五份参考库已随包提供，大型明细会在首次审核时自动还原，不需要另行寻找数据文件。

## 目录

| 路径 | 用途 |
|---|---|
| `gui.py` / `run.py` | 图形界面 / 命令行入口 |
| `src/` | 程序逻辑 |
| `config/` | 审核规则配置 |
| `data/` | 五份参考库、登记清单及分块 |
| `input/` / `output/` | 每期输入、输出，本地保存 |
| `docs/` | 详细使用说明与已知限制 |
| `tests/` | 自动化测试 |

## 验证与许可

```bash
python -m unittest discover tests
```

程序保留原有 MIT 许可证。公开下载不改变参考数据的原有权利及许可范围；每期业务输入表和审核结果不随仓库分发。
