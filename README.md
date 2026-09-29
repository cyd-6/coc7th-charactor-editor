# COC7 调查员车卡器

可自行部署的中文 COC7 建卡工具：网页填写人物、导入表格、检查规则与技能点数、按年度换算资产，并导出 XLSX 或两页 1920s 版式 PDF。支持手机布局及浅色／深色主题，日常使用和导出无需外网。

## 实现逻辑

项目由 **Python 3.12 + FastAPI + Uvicorn** 提供服务，前端是原生 HTML、CSS、JavaScript，无需 Node.js 构建。数据流程为：

**网页填写／表格导入 → 统一人物草稿 → 规则计算与校验 → 模板导出 → 浏览器下载。**

| 模块 | 职责 |
| --- | --- |
| `static/`、`app.py` | 网页交互、会话草稿；后端提供静态资源及计算、校验、导入、导出接口。 |
| `coc7_card/catalog.py` | 从模板读取职业、技能等目录，加载年度汇率并校验模板版本。 |
| `coc7_card/` 下的 `web.py`、`models.py`、`rules.py` | 将请求转换为统一的 `CharacterDraft`，计算派生属性与技能预算，校验导出条件。 |
| `coc7_card/importers/` | 解析完整调查员 XLSX 或简化属性表；只读取单元格及已保存的公式结果，不执行上传公式。网页先预览，再确认回填。 |
| `coc7_card/currency.py`、`assets/rates/annual.json` | 按年份和实际币制查找报价，以美元为输入基准换算；缺失报价不插补。覆盖 1920—2026 年，2026 年为 1—8 月均值。 |
| `coc7_card/exporters/`、`assets/templates/` | 后端重新构建并校验草稿、查找汇率，使用随附模板导出。 |

两种导出各自处理排版与计算：

- **XLSX**：使用 `COC7空白卡CY26.3.xlsx`。Linux／Docker 局部修改模板副本，让 LibreOffice Calc 重算临时副本，再仅回写公式缓存，保留原有 14 个工作表、公式和模板结构；Windows 原生运行通过 Excel COM 填写副本并重算。`factory.py` 自动选择平台实现。
- **PDF**：在 `1920sCha.pdf` 原底版上叠加中文人物资料和头像，生成两页静态 PDF 与预览，不依赖 Office。内容超出两页容量时明确报错。两种导出均不改写原始模板。

## Docker 启动


```bash
git clone https://github.com/cyd-6/coc7th-charactor-editor.git
cd coc7th-charactor-editor
docker info
docker compose version
docker compose up -d --build
docker compose ps
```

`coc7` 状态显示 `healthy` 后，打开 **[http://127.0.0.1:8765/](http://127.0.0.1:8765/)**。首次构建需要联网下载依赖；`-d` 表示后台运行，完成后可以关闭终端。

[健康检查](http://127.0.0.1:8765/api/health) 应返回 `{"status":"ok","service":"coc7-investigator-builder"}`。若仍在启动，稍后再查看；失败时执行 `docker compose logs --tail=100 coc7`。

## Windows 直接运行 Python


```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r docs/requirements.windows.lock.txt
.\.venv\Scripts\python.exe -m coc7_card.launcher
```

已有本机 Python 3.12 虚拟环境时跳过第一行。依赖会自动安装 `pywin32`；直接使用虚拟环境解释器即可，无需激活脚本，见 [Python venv 文档](https://docs.python.org/3.12/library/venv.html)。

启动器优先使用 `127.0.0.1:8765`，占用时自动选择空闲端口，并在服务就绪后打开浏览器；以终端显示的地址为准。保持终端运行，按 `Ctrl+C` 停止。以后只需执行最后一行即可启动。

当前仓库不附带 `run.bat`。Windows 安装、内存优化与真实 Excel 测试见 [Windows 运行说明](docs/README.windows.md)。网页和 PDF 导出不依赖 Excel；Windows 原生 XLSX 导出需要桌面版 Microsoft Excel，没有 Excel 时可使用 Docker。

## 配置、启停与更新

以下命令均在项目根目录执行。

| Docker 操作 | 命令 |
| --- | --- |
| 日常开启 | `docker compose up -d` |
| 停止／重启 | `docker compose stop` / `docker compose restart coc7` |
| 查看状态／日志 | `docker compose ps` / `docker compose logs --tail=100 -f coc7` |
| 停止并移除容器 | `docker compose down` |
| 代码或依赖更新后重建 | `docker compose up -d --build` |

Git 用户保存自己的修改后，用 `git pull --ff-only` 获取更新并重建；ZIP 用户下载新版并保留 `.env`。单独 `restart` 不会应用新源码或配置。Linux／macOS／WSL 可用 `bash scripts/start-container.sh` 备份旧镜像并更新，健康检查失败时尝试恢复旧镜像（需要支持 `up --wait` 的 Compose）。

**Docker 默认仅供本机访问，无需创建 `.env`。** 需要改端口、开放局域网或增加重算时间时，将 `.env.example` 复制为 `.env`，已有文件直接编辑：

```dotenv
COC7_BIND_IP=127.0.0.1
COC7_PORT=8765
TZ=Asia/Hong_Kong
COC7_EXCEL_TIMEOUT=120
```

- 端口冲突：将 `COC7_PORT` 改为 `8766` 等空闲端口。
- 局域网访问：将 `COC7_BIND_IP` 改为 `0.0.0.0`，放行防火墙端口，其他设备访问 `http://服务器局域网IP:配置端口/`。本工具没有登录系统，适合可信网络。
- XLSX 重算超时：增大 `COC7_EXCEL_TIMEOUT`，单位为正整数秒，作用于 Linux／Docker 的 LibreOffice 重算。

保存后执行 `docker compose up -d` 并检查状态。上述 `.env` 由 Compose 读取，原生 Python 启动不会自动使用它。

启动报错时：找不到 `docker`／`compose` 则补装组件；无法连接 Docker 则启动引擎；Linux 的 `docker.sock` 权限错误可使用 `sudo docker …`；找不到配置文件则进入项目根目录。

## 使用与保存

依次填写 **人物信息 → 属性 → 职业 → 技能点 → 背景资产 → 武器物品 → 检查导出**，修正阻断错误后下载文件。也可从现有表格开始：

- **完整导入**：支持 CY26.3、布局兼容的 CY26.2 调查员卡或本工具导出的 `.xlsx`，先预览再确认，并可在当前页面撤销。
- **仅属性导入**：在属性页下载简化模板，支持 XLSX／CSV／TSV；只更新有数值的属性，空白保留原值，`0` 有效，不覆盖人物资料或删除已有技能加点。两类表格均限 12 MB。

**草稿仅保存在当前标签页的 `sessionStorage`，请及时导出保存。** 刷新通常可恢复草稿，头像需重选；关闭标签页后不应依赖会话保存。项目没有账号、数据库或服务端人物存档，文件由浏览器下载。规则计算与导出由你部署的服务处理，远程部署时数据会发送到该服务器。

## Linux 源码运行与验证

正式模板版本和文件名统一记录在 `assets/templates/manifest.json`，应用、测试和模板维护脚本共用此配置。CY26.3 使用五项精简换算页，主表资产币种可下拉选择。更换模板后需重启已运行的服务；Docker 使用重新构建的镜像，不会自动载入宿主机文件。

准备 Python 3.12 和 LibreOffice Calc（XLSX 导出需要，`libreoffice` 或 `soffice` 须在 PATH 中），在项目根目录执行；已有 `.venv` 时复用：

```bash
python3.12 -m venv .venv
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m uvicorn app:app --host 127.0.0.1 --port 8765
```

启动后打开同一首页地址，按 `Ctrl+C` 停止；端口占用时修改 `--port`。测试执行 `.venv/bin/python -m pytest -q`，差异检查执行 `git diff --check`。

`requirements.txt` 定义跨平台依赖范围；`requirements.lock.txt` 锁定运行依赖，供 Docker 使用；`docs/requirements.windows.lock.txt` 在相同版本上加入 Windows COM 依赖；`requirements-dev.txt` 加入测试工具和按平台安装的 `pywin32`。Node.js 用于脚本检查和部分测试。缺少 LibreOffice／Node.js 会分别跳过 LibreOffice XLSX／浏览器金额对照检查，跳过不等于通过。Windows 真实 Excel 测试需设置 `COC7_TEST_EXCEL=1` 后运行，详见 Windows 说明。

汇率来源与数据维护见 [年度汇率说明](docs/EXCHANGE_RATES.md)。项目代码采用 [MIT 许可证](LICENSE)；模板、字体、PDF 底版及其他第三方素材遵循各自许可，见 [第三方说明](THIRD_PARTY_NOTICES.txt)。
