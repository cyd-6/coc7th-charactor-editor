# COC7 调查员车卡器 · Docker 版

中文 COC7 调查员建卡工具，网页名称为「coc7车卡器」，支持浏览器填写、规则检查、Excel 导出和两页 1920s 原版式 PDF 导出。镜像包含 Python、LibreOffice Calc、中文字体和角色卡模板，宿主机无需安装 Python 或 Microsoft Office。

网页默认以表单和操作为主。职业资料、经历包说明和非阻断的完整性提醒可按需展开；错误、规则警告、点数预算及汇率来源摘要直接显示。

字号随窗口宽度自动调整：正文和常规输入框在小屏保持 16px，1920px 宽度约为 19px，2560px 约为 22px，最大 24px（浏览器默认字号下）。右上角「外观」可选跟随系统、浅色或深色；默认跟随系统，选择单独保存在当前浏览器，不随调查员草稿清空。主题仅改变网页外观，PDF 预览保留纸张原色。

## 启动

安装并启动 Docker Engine 或 Docker Desktop，以及 Docker Compose v2 插件。在本项目目录执行：

```bash
docker compose up -d --build
```

浏览器访问 **http://localhost:8765**。首次构建需要联网下载基础镜像、系统软件和 Python 依赖；构建完成后，正常使用和导出不需要外部网络。

```bash
# 查看状态；启动后应显示 healthy
docker compose ps

# 查看日志
docker compose logs -f coc7

# 停止并移除本项目的容器
docker compose down

# 修改代码或依赖后重新构建
docker compose up -d --build
```

容器直接运行 Uvicorn，不调用 Windows 的 `.bat`、`.exe` 或桌面启动器。容器内监听 `0.0.0.0:8765`；Compose 默认仅将端口发布到宿主机 `127.0.0.1`。

## 源码与版本管理

项目仓库：[cyd-6/coc7th-charactor-editor](https://github.com/cyd-6/coc7th-charactor-editor)。首次获取源码：

```bash
git clone https://github.com/cyd-6/coc7th-charactor-editor.git
cd coc7th-charactor-editor
docker compose up -d --build
```

Git 保存应用源码、测试、使用说明、依赖锁定文件和运行所需的模板、字体、年度汇率数据。虚拟环境、本机配置、依赖目录、导出样张、测试产物及发布压缩包不进入版本历史；`.env.example` 作为配置示例保留。

项目代码采用 [MIT 许可证](LICENSE)。角色卡模板、PDF 底版、字体、汇率来源和第三方组件仍遵循各自的版权与许可，不因仓库许可证而重新授权，详见 [第三方说明](THIRD_PARTY_NOTICES.txt)。

## 修改端口或供局域网访问

复制配置示例：

```bash
cp .env.example .env
```

编辑 `.env`：

```dotenv
COC7_BIND_IP=127.0.0.1
COC7_PORT=8765
TZ=Asia/Hong_Kong
COC7_EXCEL_TIMEOUT=120
```

- 端口被占用时修改 `COC7_PORT`，例如 `8766`，再执行 `docker compose up -d`。
- 在 NAS 或可信局域网中使用时，将 `COC7_BIND_IP` 改为 `0.0.0.0`，再通过 `http://服务器IP:8765` 访问。本工具没有登录系统。
- 低性能设备导出超时时，可以增加 `COC7_EXCEL_TIMEOUT`，单位为秒。

不使用 Compose 也可以：

```bash
docker build -t coc7-card:local .
docker run -d --name coc7-card --init --restart unless-stopped \
  -p 127.0.0.1:8765:8765 coc7-card:local
```

## Excel 与 PDF

在「调查员信息」标题旁点击 **导入 Excel**，选择原始 CY26.2 调查员卡或本工具导出的 `.xlsx`（最大 12 MB）。先查看姓名、职业、资料数量和导入提示，确认后回填网页；取消不会修改当前调查员，导入后可用「撤销导入」恢复导入前的资料与头像（刷新后不再保留撤销记录）。上传文件原样保留。

导入读取表格中实际填写的身份、日期、属性、技能及专攻、经历包、背景、年度币种及美元输入、武器和物品，并识别肖像区域的内嵌头像。成长点数与经历包点分别保留；旧版合并增长栏放入成长点数并提示按原记录拆分。无法唯一恢复的职业可选技能会提示补选，不会清除已填点数。表格自定义的技能基础值会保留，并参与后续计算与导出。

完整调查员导入仅支持上述布局的 XLSX，不支持通用表格、CSV、旧版 XLS 或带宏文件。导入只读取已有单元格与保存的公式结果，不运行上传表格中的公式；缺少计算缓存时会提示先在 Excel／LibreOffice 中重算保存。表格未记录的网页专属内容无法恢复，非零资产明细金额会写入「资产说明」。未知年份／币种需要重新选择。纸卡中的当前 HP／MP／SAN 等游玩状态不替代网页的建卡派生值。

**仅导入属性：**在「最终属性」页下载 **简化表模板**，填写后点击 **导入属性**。支持 XLSX、CSV、TSV（最大 12 MB），横向属性表头加一行人物数值，或纵向「属性／数值」两列。识别力量 STR、体质 CON、体型 SIZ、敏捷 DEX、外貌 APP、智力 INT、意志 POW、教育 EDU、幸运 Luck，中文、英文缩写及中英组合均可。数值为 0—300 的整数，空白项保留网页原值，0 是有效输入。姓名、年龄、技能、资产等其他字段不参与导入。

属性导入先预览原值与新值，确认后仅更新有值的属性，并重新计算派生值、动态技能基础值及点数预算。已有职业、技能加点、经历包、背景、资产和头像保持原样；超出新预算的加点由原有规则检查提示，不自动删减。当前页面可「撤销属性导入」，只恢复本次改变的属性，保留其他资料的后续编辑；刷新、导入完整调查员或清空会话后撤销记录失效。重复属性、多人物数据、多个属性工作表或非法数值会整体拒绝；含公式的属性仅使用已保存的结果，无缓存或公式错误须先重算保存。接口为 `POST /api/import/attributes`，上传字段为 `workbook`。

**Excel：**Linux 版直接修改模板副本中的角色数据和必要的格式设置，保留 14 个工作表、原有公式、下拉选项、图表、批注和团务记录区域。LibreOffice Calc 对临时副本重算，程序只将计算结果缓存写回导出文件，因此不会将 LibreOffice 重写后的模板直接交付。导出后检查工作表顺序、已知断裂引用、公式错误与源模板哈希。

导出使用项目随附的已升级模板，或由维护流程生成并验证符合相同结构与公式要求的模板。兼容修补和固定资产换算公式由模板维护阶段负责；导出时不再自动修补替代模板或搬运旧模板的资产明细金额，已知模板校验失败会明确拒绝导出。此限制不改变旧版调查员卡的导入支持。

技能表分别填写 **成长点数** 与 **经历包点**：左侧技能的 L、M 列和右侧技能的 AH、AI 列依次对应两者。成长点数只写入网页的额外增量（`extra_final`），经历包点写入经历分配（`experience_points`）；普通／困难／极难成功率同时包含两项，经历包剩余额度只减去经历包点。两列可在 XLSX 中独立修改并重算。原有「成长表（测试）」继续作为模组成长记录，未增加自动回填功能。旧导出中已经合并的点数无法可靠拆分，请从保留独立字段的网页草稿重新导出。

每次重算使用独立临时目录和 LibreOffice 用户配置。同一服务进程的 Excel 导出顺序执行，以控制内存使用；超时会终止该次重算进程，临时文件随请求结束清理。头像作为 PNG 嵌入工作簿。

公式由 LibreOffice Calc 计算，数值和格式仍可能与特定版本的 Microsoft Excel 存在兼容性差异。需要固定纸面效果时使用 PDF。修改模板或新增特殊公式后，应重新运行导出测试。

**PDF：**使用用户提供的 `assets/templates/1920sCha.pdf` 作为原始底版，保留两页尺寸（668.692 × 848.692 pt）、原装饰、四列技能表、状态刻度、武器、背景、装备、资产和同伴区域，包含原件的印刷辅助标记。程序在原有空白栏中叠加中文人物数据，普通／困难／极难数值分别填写，头像放入原头像框；PDF 保持静态并提供两页预览，不依赖 Office。

原卡只有 60 个技能位置：已有普通技能填入同名位置，专攻技能与额外技能使用可用空白格，超过栏位的已填写技能转入背面补充区；未填写且未修改的可选技能不占空白格。原版没有独立栏位的个人故事、关键连接、经历包、时代日期、武器备注等带明确前缀放入背面两个补充区域，联系人自由文本使用同伴框。汇率年份、实际币名、口径、观测期间和来源保留在资产区。超长内容在两页内无法完整容纳时会明确提示修正，不静默截断；原始 PDF 不会被改写。

**Windows 源码运行：**保留原 Excel COM 导出实现；在 Windows 上直接运行 Python 源码时，Excel 导出仍需要桌面版 Microsoft Excel。Docker 使用 Linux 实现。

## 数据保存

- 草稿保存在当前浏览器标签页的 `sessionStorage`，没有账号、数据库或服务端草稿目录。
- 刷新当前标签页通常可以恢复草稿；关闭标签页会结束该会话。头像需要重新选择，完成后请及时导出。
- 规则计算和导出会将角色资料提交给你部署的服务。如果部署在其他机器上，处理发生在那台机器的容器中。
- 下载文件保存在浏览器设置的下载目录。容器不保存已导出的角色文件，因此不需要数据卷。
- 更换访问端口、域名或浏览器会使用不同的浏览器存储空间。

故事年份与汇率支持 **1920—2026 年**。数据来自用户提供的 `historical_exchange_rates_1920_2026.xlsx`，按当年实际币制精确查找，不再插值，也不混用旧的单日报价。合并代理和重复欧元后，有 1,061 条有效报价、60 条缺失记录；缺失币种不可选，2026 年明确标为 1—8 月均值。

在「背景资产」中选择币种，可以看到来源名称、统计口径、期间和报价提示；展开「查看数据来源与换算说明」可查看原始资料链接、引用文件和原始报价的转换说明。更改年份导致币种失效时，需要重新选择币种，已输入的美元金额会保留。

原始 Excel 模板也可独立换算：打开「货币汇率」，修改黄色的 **K1 年份、K2 币种**，人物卡的币名、金额与来源随之更新。K9:K11 是信用评级自动计算的美元金额，也可输入明确的美元金额覆盖；K13:K17 是资产明细的美元金额。缺失汇率显示提示并留空换算金额。年度数据及来源位于第 35 行起，工作簿不依赖宏或外部链接。

信用评级对应的美元基准沿用原有规则（2026 年使用现代倍率，其余年份使用既有历史基准），不据汇率推算通胀或购买力。更多数据口径、维护与回滚说明见 [年度汇率说明](docs/EXCHANGE_RATES.md)。[原版功能说明](docs/README.windows.md) 保留为历史资料，其中旧的汇率范围、插值方式及 Windows Office 要求不适用于本次 Docker 版本。

## Linux 本地开发与验证

安装 Python 3.12 和 LibreOffice Calc。在 Debian / Ubuntu 上安装系统组件后运行：

```bash
sudo apt-get update
sudo apt-get install python3-venv libreoffice-calc nodejs
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m pytest -q
.venv/bin/python -m uvicorn app:app --host 127.0.0.1 --port 8765
```

测试覆盖网页与目录加载、属性计算、PDF 页数和中文文本、历史及现代币种、自定义职业、头像、Excel 公式缓存、模板结构、临时进程超时和并发导出的角色隔离。没有 LibreOffice Calc 时，真实 Excel 导出测试会跳过；应在装有 Calc 的环境完成验证。

`requirements.txt` 定义兼容范围，`requirements.lock.txt` 固定经过验证的 Linux 运行依赖。Docker 使用锁定文件；更新依赖后应同步锁定文件并重新测试。

草稿校验及导出接口对年龄、技能点数和经历包数值中的非空非法数字返回 HTTP 422；省略字段、`null` 和空字符串沿用各字段的默认值。

## 文件结构

```text
Dockerfile                  Linux 镜像、中文字体与健康检查
compose.yaml                一条命令启动与端口配置
.env.example                可选部署参数
app.py                      Web 接口
coc7_card/exporters/
  excel.py                  原 Windows Excel 导出与共用字段映射
  excel_linux.py            Linux 重算、缓存回写及验证
  xlsx_template.py          保留原模板结构的 OOXML 编辑
  pdf.py                    两页 PDF
static/                     浏览器界面
assets/                     模板、字体、年度汇率数据及原始数据文件
scripts/                    汇率导入、模板生成及容器启动维护工具
tests/                      接口及真实导出测试
```

容器使用普通用户运行，并提供 `/api/health` 健康检查。Windows 便携运行环境、开发虚拟环境和测试产物不会进入镜像构建上下文。

## Docker 环境问题

- `docker: 'compose' is not a docker command`：安装 Compose v2 插件；也可使用上面的 `docker build` / `docker run` 命令。
- `permission denied ... docker.sock`：当前用户没有访问 Docker 守护进程的权限。由本机管理员配置 Docker 使用权限，或在允许的环境下使用 `sudo docker ...`。
- `Cannot connect to the Docker daemon`：启动 Docker Engine 或 Docker Desktop 后重试。
- 端口冲突：修改 `.env` 中的 `COC7_PORT`，重建容器后使用新地址访问。

部署用法参考 [Docker 官方 Python 指南](https://docs.docker.com/guides/python/)；LibreOffice 的无界面运行参数参考 [官方参数说明](https://help.libreoffice.org/latest/zh-CN/text/shared/guide/start_parameters.html)。

第三方组件信息见 [THIRD_PARTY_NOTICES.txt](THIRD_PARTY_NOTICES.txt)。用户提供的角色卡模板沿用其原有许可。
