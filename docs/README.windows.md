# Windows 运行说明

Windows 原生启动支持网页建卡、XLSX/CSV/TSV 导入、PDF 导出与 Excel 导出。使用 Python 3.12（64 位）；仅 XLSX 导出需要已安装、完成首次启动的桌面版 Microsoft Excel。PDF 导出不依赖 Office。WPS 不作为原生 Excel COM 后端，没有 Excel 时可选择 README 中的 Docker 方案。

## 安装与启动

在项目根目录执行；已有 .venv 时复用，不要重复创建：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r docs/requirements.windows.lock.txt
.\.venv\Scripts\python.exe -m coc7_card.launcher
```

Windows 锁定文件复用主运行依赖版本并补充 pywin32，不再使用缺少 lxml 的历史依赖列表。仓库没有 run.bat，不需要执行 PowerShell 激活脚本。

启动器保留已绑定的监听 socket，优先使用 http://127.0.0.1:8765，端口占用时选择空闲端口；以终端显示地址为准。服务就绪后打开浏览器，按 Ctrl+C 停止。原生启动不会读取 Compose 的 .env 文件。

## Excel 与资源使用

Excel 导出在独立、隐藏的实例中完成，不连接用户已有的 Excel 窗口。一个 Python 服务进程同时只执行一项 Excel 导出，其余请求等待，最多等待 120 秒。此等待期限不等于正在运行的 Excel COM 调用具有强制超时。保持单进程运行，避免每个 worker 各自加载目录和启动 Office。

填写时暂停该实例的自动重算，填完后全量重算，并在保存前恢复自动计算。导出文件仍支持直接修改年份与币种后自动换算。数值格式覆盖完整合并区域；汇率沿用模板中的通用格式，避免部分语言版本拒绝英文 General 格式字符串。用户文字通过 COM 按字面量写入，保留前导零及以等号开头的姓名、背景等内容。

目录读取按需要的区域顺序扫描，不加载整本工作簿；Excel 校验逐段扫描 XML 内的公式及计算缓存，不再同时建立两份完整的 openpyxl 工作簿。校验继续检查工作表数量/顺序、模板兼容公式、断裂引用和公式错误。原模板的哈希保护保留。

年度汇率覆盖 1920—2026 年，2026 年为 1—8 月均值，详见 [年度汇率说明](EXCHANGE_RATES.md)。PDF 使用随附的两页 1920s 原底版，不依赖 Excel。草稿仍存于浏览器 sessionStorage，导出文件由浏览器下载。

## 验证

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -m pytest -q
```

安装了桌面版 Excel 后，启用真实 COM 测试：

```powershell
$env:COC7_TEST_EXCEL = "1"
.\.venv\Scripts\python.exe -m pytest -q
```

真实测试会创建并关闭独立 Excel 实例，覆盖头像、自定义职业、文字保真、资产换算、导出后再导入、重新打开后的自动计算、下拉菜单与并发隔离。不会保存用户的模板或关闭用户的 Excel 窗口。Linux 的进程组超时测试在 Windows 上跳过；未安装 LibreOffice 时相应重算测试也跳过。Node.js 不在 PATH 时可将 COC7_NODE 设为 node.exe 的完整路径。

内存基准脚本在每次新进程中运行：

```powershell
.\.venv\Scripts\python.exe scripts/benchmark_memory.py catalog
.\.venv\Scripts\python.exe scripts/benchmark_memory.py verify --workbook "已导出的有效角色.xlsx"
```

使用填过有效人物的导出文件测校验，空白模板中的既有除零状态不能作为有效人物导出。脚本使用 tracemalloc 测量 Python 分配峰值，并记录进程峰值工作集；后者包含跟踪开销，不代表无跟踪时的日常占用，也不包含 Excel 子进程。

本次检查结果与优化前后数值见 [Windows 兼容性与内存验证记录](WINDOWS_MEMORY_VERIFICATION.md)。
