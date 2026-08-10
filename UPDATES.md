---

## v2.4.1 — 2026-08-10

### 修复：临时公司录入中文退格残留
- 根因：PTY/终端下 `input()` 的中文退格按字节而非完整 UTF-8 字符删除，
  删一个汉字会残留半个字节（表现为乱码、删不干净）
- 新增 `smart_input()` 自定义输入（仅 tty 生效，管道/重定向回退内置 input）：
  - raw 模式逐字符读取，退格按 UTF-8 首字节边界弹出完整字符（`_pop_utf8_char`）
  - 按终端显示列宽擦除（中文 2 列，`_disp_width`），不再整行重绘避免输出阻塞
  - 模块级 `_input_cache` 缓存多读字节：连续多字段录入时预写输入不丢失
- `ask_temp_company()` 四个字段全部改用 `smart_input()`
- 新增依赖导入：termios / tty / unicodedata（Python 标准库，无需 pip 安装）

### 验证
- PTY 模拟：输入「甲乙丙」按两次退格 → 得「甲」（DEL 0x7f 与 BS 0x08 均通过）
- ASCII 退格正常；中文+ASCII 混合列宽计算正确
- ask_temp_company 全流程：一次性预写 名称/关键词(含退格)/预算/区域/结束，
  结果字段全部正确、无残留、无卡死

---
---

## v2.4.0 — 2026-08-10

### 新增：临时公司交互录入（不落盘，直接喂 AI）
- 爬取完成、生成 HTML 报表前，交互询问是否录入临时公司：
  - 四行提示：公司名称 / 关键词 / 预算范围 / 关注区域，每项可回车跳过
  - 输入完一家自动询问下一家；公司名称直接回车 = 结束录入
  - 临时公司不写入 company_profiles/*.json，仅本次运行内存中存在，
    与现有公司画像一起作为 AI 价值报告的视角
- 解析辅助：
  - `_parse_keywords()`：逗号/空格/顿号分隔 → {关键词: 1}
  - `_parse_budget()`：`100-5000` / `5000至100` / `100 5000` → [100, 5000]，无法解析返回 None
  - `_parse_regions()`：逗号/空格分隔 → 区域列表
- `ask_temp_company()`：返回临时画像列表（可为空）；`main()` 中
  `profiles = load_profiles() + temps` 传入 `generate_html_report()`

### 验证
- 解析函数单测：关键词/预算（含反序至）/区域，均正确
- 交互流程模拟：录入 1 家后名称回车结束 → 返回 1 家；直接回车 → 空列表
- 真实 API 全链路：7 视图（6 现有 + 1 临时），临时公司视图含 AI 动态价值报告
- 确认临时公司未写入 company_profiles 目录

---
---

## v2.3.0 — 2026-08-10

### 变更：删除规则评分与图表，价值判断只由 AI 动态生成
- 删除全部加权数值评分代码：`value_components()` / `score_from_md()` /
  `budget_score()` / `region_score()` / `qual_mismatch()` /
  TYPE_VALUE / INTEREST_KEYWORDS / QUAL_LEVEL 权重逻辑
- 删除全部 SVG 图表代码：`make_value_charts()` 及其辅助常量
  （CHART_COLORS / COMPONENT_NAMES / COMPONENT_MAX / _short_title / _match_notes），
  HTML 报表不再包含图表卡片与相关 CSS
- 价值判断完全交给 DeepSeek（deepseek-v4-flash）动态生成：
  - 每个公司视角一份 AI 报告（总体判断 + 每条公告 评分/等级/理由 + 建议）
  - 表格按 AI 报告分数降序排列；AI 失败时保持原序号顺序，不影响报表
  - 爬取主流程不再计算 score，meta.json 不再写入 score 字段（历史数据中的 score 字段被忽略）
- 优化：`llm_value_report()` 新增 `base_rows` 参数复用已解析的 金额/工期/附件，
  避免对每份 md 重复读取与正则解析；`render_llm_report_html()` 容错非法 score、去除 reason 换行
- 公司画像保留 keywords/exclude/qualifications/budget_range/regions 字段，
  作为 LLM 的业务背景描述（`_profile_desc()`），不再参与数值计算

### 验证
- 6 个公司视角 × 5 条公告全部生成 AI 报告（真实 API 调用）
- 报表无任何图表残留（无 value-charts / 价值总分排行 / chart-card）
- 默认视角表格首行 = AI 评分最高的城口排水管网项目（index 5），排序与 AI 判断一致
- 无 key 时报表正常回退：表格按原顺序展示，无 AI 卡片，不报错

---
---

## v2.2.0 — 2026-08-10

### 新增：大模型动态价值报告（DeepSeek）
- 价值报告改为由 DeepSeek（deepseek-v4-flash）动态生成，不再只靠规则加权数值：
  - `llm_chat()`：调用 DeepSeek chat/completions（base_url https://api.deepseek.com/v1），
    强制 JSON 输出，失败重试 2 次，401 时直接放弃
  - `llm_value_report()`：把每条公告的 标题/编号/金额/工期/附件名/正文前 300 字 + 公司画像
    打包发送给模型，返回 {summary, items[{index, score, verdict, reason}], notes}
  - `render_llm_report_html()`：渲染为 HTML 卡片，每个公司视角视图块顶部显示
    （总体判断 + 每条公告的 评分/跟进等级/理由 + 整体建议），按分数降序
- 保留原规则评分图表（金额30/类型15/附件15/关键词20/地区10/资质10）作为参考：
  LLM 调用失败或无 key 时自动回退到规则评分，报表不受影响
- API key 读取：优先环境变量 `DEEPSEEK_API_KEY`，其次本地 `llm_key.txt`；
  llm_key.txt 已加入 .gitignore，不会上传到 GitHub

### 验证
- 用现有 5 条真实公告 + 6 个公司视角实测：6 个视图全部生成 AI 卡片
- 默认视角：排水管网项目 85 分「重点跟进」，理由匹配产业园/给排水/污水/市政关键词
- 公路交通视角：整体判定为非主业、区域不符，建议忽略，与画像逻辑一致
- LLM 调用成功（含 reasoning 字段，模型 deepseek-v4-flash 可用）

---
---

## v2.1.1 — 2026-08-07

### 修复：备份脚本未含公司画像
- `backup.sh` 增加 `cp -rv company_profiles "$DIR/"`，备份目录现包含 5 家公司画像 JSON

### 验证
- `./backup.sh` 新备份目录含 company_profiles/ 全部 5 个 JSON

---

## v2.1.0 — 2026-08-07

### 新增：5 家虚拟公司画像
- 新增 3 家公司（与既有 2 家组成 5 家不同定位）：
  - 电力能源工程公司（电力/配电/光伏/能源，关注潼南/黔江/渝北）
  - 智能化弱电工程公司（智能化/数据中心/智慧/安防，小预算区间 100-5000 万）
  - 公路交通工程公司（道路/桥梁/隧道/公路，大预算区间 300-50000 万）
- 既有：渝发建设（房建总承包）、市政环境工程公司
- 下拉切换体验优化：
  - JS 用 localStorage 记住上次选择的公司，刷新页面保持
  - 视图块注释明确：默认视角在首位，自定义公司按 company_profiles 文件顺序
- upload_github.py FILES 清单补入 3 个新画像 JSON

### 验证
- 6 个视角（默认+5 家）对同一批报告榜首完全不同：
  默认→小学；公路→足球场地(97.3)；市政→大足道路(94.7)；智能化→阿里EPC(68.7)；
  渝发→阿里EPC(79.2)；电力→烟草公司(61.6)
- HTML 6 view/6 option/12 SVG，localStorage 逻辑嵌入
- Edge 无头截图渲染正常

---

## v2.0.0 — 2026-08-07

### 新增：多公司画像价值判断（方案A：单文件下拉切换）
- `company_profiles/*.json`：每家公司一个画像，包含
  - keywords（业务关键词加权）/ exclude_keywords（排除词，命中大幅降权）
  - qualifications（资质，匹配正文资质要求，等级高于公司→0分）
  - budget_range（预算区间，超出上限→0，低于下限→低分）
  - regions（关注区域，建设地点匹配）/ weights（六维权重）
- `load_profiles()`：读取画像目录，未配置时用 DEFAULT_PROFILE（综合视角）
- `value_components()` 改为按画像打分：金额30/类型15/附件15/关键词20/地区10/资质10
  - `budget_score()` 预算区间匹配；`region_score()` 地区匹配；
    `qual_mismatch()` 资质匹配（等级对比：一级>二级>三级）
- `generate_html_report()`：每个画像生成独立视图块（画像说明+图表+按该公司价值降序的表格），
  顶部下拉切换，前端 JS 仅做显示切换，无外部 CDN、单文件
- 图表升级：六维堆叠（新增地区/资质段）、匹配明细列表（金额/地区/资质/关键词命中原因）
- 示例画像：company_profiles/渝发建设.json（房建）、市政环境.json（市政管网）

### 验证
- 3 家视角对同一批 11 份报告排序完全不同（市政→大足道路第一；房建→阿里EPC第一）
- Edge 无头截图确认六维图表颜色渲染正常、下拉 3 个选项、视图块 3 个

---

## v1.11.0 — 2026-08-07

### 修复：zip 与报表混入历史报告
- 根因：`make_summary_zip` 打包整个 docx 目录、`generate_html_report` 读 meta.json 全量，
  与用户本次选择的爬取范围无关 → 每次运行都包含历史累积的全部报告
- `make_summary_zip(zip_path, records)`：改为只打包传入的本次记录（docx + 附件）
- `generate_html_report(records, html_path)`：改为只渲染传入的本次记录，不再读 meta.json
- 主流程新增 `cur_meta`（本次范围结果集）：历史已爬的 URL 从 meta 复用，新爬的实时加入；
  zip、报表、完成提示均基于 `cur_meta`
- 完成提示区分：`本次 N 份公告（历史累计 M 份）`
- 清理库存：历史 11 份报告归档至 `backups/archive_zbgg_data_*`，output 恢复干净

### 验证
- 模拟选前 3 份：zip 仅 3 个文件夹、报表仅 3 行、2 个 SVG 图表正常
- 假数据契约测试：make_summary_zip / generate_html_report 新签名行为正确

---

## v1.10.0 — 2026-08-07

### 新增：HTML 报表价值图表（不改动原表格）
- 新增 `make_value_charts()`：纯 SVG 生成两张价值图表，不依赖外部 CDN（离线可用）
  - 图1 价值总分排行：横向条形图，按分数降序，≥75 深橙 / 60-75 橙 / <60 灰
  - 图2 价值构成分析：堆叠条形图，金额/类型/附件/关键词/信息量五维按满分归一
- `generate_html_report()`：原表格结构、列、顺序保持与 v1.8 完全一致；
  图表以独立卡片追加在表格上方，`make_value_charts()` 无数据时返回空串不影响报表
- 新增 `COMPONENT_NAMES` / `COMPONENT_MAX` / `_short_title()` 辅助常量与函数

### 验证
- 现有 8 份报告生成报表：2 个 SVG XML 合法、原表 7 列不变
- Edge 无头截图渲染确认：图表条颜色分布正常、表格正常显示
- 图表排序与 value_components 评分一致（黔江老旧小区改造居首）

---

## v1.9.0 — 2026-08-07

### 新增：报告价值评分
- `value_components()`：从标题+正文+附件名计算价值分（0-100）
  - 金额 40 分（log 缩放，100万→20，1亿→40，避免大数碾压）
  - 招标类型 20 分（EPC/总承包 > 施工 > 咨询 > 监理 > 默认）
  - 附件完整度 15 分（含工程量清单/图纸=资料全）
  - 业务关键词 15 分（`INTEREST_KEYWORDS` 按需增删，命中越多越高）
  - 信息量 10 分（正文长度）
- `score_from_md()`：从 md 文件+附件名重算分（兼容旧 meta.json 无 score 字段）
- `generate_html_report()`：报表按价值分降序，新增"价值分"列（条形图+分量明细）
- 爬取时 `meta.json` 每条记录新增 `score` 字段
- `backup.sh`：修正备份文件名为实际源码（spider_newest.py + tools/*.py）

### 验证
- 现有 8 份报告评分排序合理；HTML 报表生成、排序、条形图正常
- 新增独立工具 `tools/score_reports.py`（离线批量评分）

# spider-government 更新报告

本文件记录每次代码更新的内容，与 git 提交一一对应。

---

## v1.8.1 — 2026-08-07

### 修改
- 主程序文件重命名：`zbgg_crawler.py` → `spider_newest.py`
- 同步更新：docstring、README、.vscode/tasks.json、.vscode/launch.json、tools/upload_github.py 清单

### 验证
- 重命名后导入/运行正常

---

## v1.8.0 — 2026-08-07

### 通用化改造
- Edge 路径自动探测：`tools/start_edge_cdp.ps1` 从注册表 + 常见路径自动定位 msedge.exe，找不到报 `EDGE_NOT_FOUND`（不再写死安装路径）
- 桌面路径自动发现：`desktop_path()` 优先 PowerShell 获取，失败则扫描 `/mnt/c/Users/*/Desktop`，兜底输出目录（不再写死用户名）
- 附件临时目录：`windows_temp_dir()` 用 Windows TEMP（PowerShell 获取），兜底项目 output（不再写死 `/mnt/c/tmp_zbgg_dl`）
- 依赖清单：新增 `requirements.txt` + `install.sh` 一键安装
- README 重写：环境要求（Windows + WSL2 + Edge）、安装、运行、输出说明
- PowerShell 脚本注释统一转英文，规避编码解析风险

### 验证
- 通用化后前 2 份报告端到端通过：详情渲染、3 个 PDF 附件、汇总包复制桌面、HTML 报表生成

---

## v1.7.0 — 2026-08-07

### 新增
- HTML 统计报表：
  - `parse_md_fields()` 从 md 前几行解析 标题/项目编号/信息时间
  - `extract_key_info()` 从正文正则提取 投资金额（优先总投资金额，其次合同估算金额）、工期（工期要求/总工期/服务期限）
  - `generate_html_report()` 生成表格报表：序号、名称（带链接）、项目编号、信息时间、投资金额、工期、附件
  - `open_html()` 用默认浏览器自动打开
- 主流程末尾自动生成 `output/zbgg_report.html` 并打开

### 验证
- 前 10 份报告实测：报表 10 行 × 7 列，投资金额（1411.42 万元、6299.39 万元等）与工期（136 日历天、450 日历天等）提取正确，浏览器自动打开
- 修正两个提取边界：`总工期140日历天`（无冒号）、`总投资金额：含税3491.39万元`（含税前缀），修正后 10/10 份均有完整数据

---

## v1.6.0 — 2026-08-07

### 新增
- 交互式爬取范围询问 `ask_scope()`：
  - 选项 1：按日期范围（YYYY-MM-DD 起止，自动交换大小）
  - 选项 2：按序号范围（第 N 份 ~ 第 M 份）
  - 选项 3：全部（近一个月，回车默认）
  - 输入错误/越界时自动回退爬取全部
- 运行 `python3 zbgg_crawler.py`（不带 --limit）即弹出询问

### 验证
- 5 种场景实测：日期范围命中正确、序号范围命中正确、默认全部、日期自动交换、序号越界回退

---

## v1.5.0 — 2026-08-07

### 新增（自动化健壮性）
- Edge CDP 生命周期管理：
  - `edge_alive()` 检查 CDP 存活
  - `start_edge()` 自动启动 Edge（Windows 侧）
  - `ensure_edge()` 主流程启动时自动确保 Edge 存活，无需手动运行前置脚本
- 失败自动恢复：
  - `fetch_detail()` 失败自动重启 Edge 并重试（默认 3 次）
  - `download_attachment()` 失败自动重试（默认 3 次），并校验 PDF 头
- CLI 参数：`python3 zbgg_crawler.py --limit N` 仅处理前 N 条，便于局部验证

### 验证
- 前 4 份报告完整运行：4 份 docx + 6 个 PDF 附件全部生成，PDF 头均有效（%PDF）
- 汇总包"招标报告汇总.zip"含 4 个报告文件夹，自动复制到桌面

---

## v1.4.2 — 2026-08-07

### 新增
- 汇总打包：`make_summary_zip()` 把近一个月所有报告文件夹（每份 = docx + PDF 附件）打进一个"招标报告汇总.zip"
- 自动复制到桌面：`desktop_path()` 通过 PowerShell 获取 Windows 桌面路径，打包后自动复制

### 修改
- 主流程末尾改为生成单个汇总 zip（替代 v1.4.1 每份报告一个 zip 的方式）

### 验证
- 以现有 3 份报告实测：zip 含 3 个报告文件夹（第一份含 2 个 PDF 附件，其余含 docx），已自动复制到桌面

---

## v1.4.1 — 2026-08-07

### 修改
- 打包方式改为每份报告一个独立 zip：
  - `make_report_zip()`：zip 内一个以报告标题命名的文件夹，含该报告的 docx + PDF 附件
  - 主流程末尾遍历 docx 目录，逐份生成 `output/zbgg_zips/<报告标题>.zip`
  - 替代之前"所有报告一个总 zip"的 `make_full_zip()` 方式

### 验证
- 以第一份公告（烟草公司潼南项目）实测：zip 内含报告标题文件夹，含 docx + 2 个 PDF（招标公告.pdf、工程量清单.pdf），已保存至桌面

---

## v1.4.0 — 2026-08-07

### 新增
- PDF 附件爬取：
  - `extract_detail()` 解析 `.ewb-blue-a` 的 `downloadAttach` 链接（AttachGuid/FileCode/ClientGuid + 文件名）
  - `is_pdf()` 按扩展名判断附件类型，仅下载 `.pdf`
  - `download_attachment()` 用 Edge CDP 导航触发下载（先导航 cqggzy 根域种 cookie，再导航附件 URL，浏览器原生处理 521 反爬），校验 PDF 头（%PDF）后移入附件目录
  - `download_pdf_attachments()` 批量下载，每份公告一个子目录
- 打包：`make_full_zip()` docx 在 zip 根目录，PDF 附件按公告名目录组织

### 修复
- 附件下载域 ggzydl.cqggzy.com 有 521 反爬：requests 直连失败（CORS/JS challenge）→ 改用 Edge CDP 导航下载（Browser.setDownloadBehavior + 先种 cookie）

### 验证
- 以第一份公告（烟草公司潼南项目）实测：识别 4 个附件（2 PDF + 1 rar + 1 cqzf），仅下载 2 个 PDF（招标公告.pdf 261KB、工程量清单.pdf 12MB）
- 完整 zip 打包成功（6.7MB），已复制到桌面

---

## v1.3.2 — 2026-08-07

### 新增
- 表格提取：
  - 块内表格（联系方式表，位于 titleStr 块内、无 content class）
  - 独立 table 块（投标保证金账号表）
  - `paragraphs` 新增 `table` 类型，保存行/列数据
- `to_docx()` 支持表格输出：Table Grid 样式，自动列宽（取最大行宽）

### 修复
- 之前仅提取 titleStr + content，"7.联系方式"后的表格信息缺失 → 补充两类表格提取

### 验证
- 以第一份公告（烟草公司潼南项目）实测：提取 2 个表格（联系方式 11 行、保证金账号 10 行），docx 打开表格正常显示

---

## v1.3.1 — 2026-08-07

### 新增
- 正文小标题提取：按 `.app-detail` 内子块顺序遍历，识别 `titleStr`（小标题）与 `content`（正文）
- 段落类型标记：`paragraphs` 每项含 `type`（heading/body），小标题加粗、不缩进输出
- 跳过页面隐藏块（class=hide，如"3.政府采购工程"折叠内容）

### 修复
- 之前正文仅取 `.content`，小标题（"1.招标条件"等）缺失 → 改为按文档顺序同时提取两类

### 验证
- 以第一份公告（烟草公司潼南项目）实测：提取 7 个小标题（1.招标条件 … 7.联系方式），与网页一致；docx 中小标题加粗、正文缩进排版正常

---

## v1.3.0 — 2026-08-07

### 新增
- 正文富文本解析：`_parse_runs()` 递归提取段落 run（保留下划线、分段）
- `extract_detail()` 返回 `paragraphs`（富文本段落列表）+ 纯文本 `body`
- `to_docx()` 按网页排版生成：
  - 项目标题：居中、加粗、18pt（网页 h1 24px）
  - 项目编号：居中、18pt
  - 信息时间：居中、14pt
  - 正文：宋体 12pt、左对齐、首行缩进 2 字符（24pt）、行距 150%、保留下划线
- 中文字体设置：eastAsia 属性指向宋体

### 验证
- 以第一份公告（烟草公司潼南项目）实测：17 段正文，下划线信息保留（如项目名、招标人等），docx 打开排版与网页一致

---

## v1.2.0 — 2026-08-07

### 新增
- `extract_detail()` 从详情页 `.detail-wrapper` 提取三个元字段：
  - 项目标题（第一个 h1）
  - 项目编号（形如"项目编号：xxx"的 h1）
  - 信息时间（"信息时间：yyyy-mm-dd"文本）
- 输出格式调整为：项目标题 + 项目编号 + 信息时间 + 正文
- `to_docx()` 首行（项目标题）加粗显示

### 修改
- 引入 `re` 模块（信息时间正则提取）

### 验证
- 以第一份公告（烟草公司潼南项目）实测：编号 50000120260806025160101、时间 2026-08-07 提取正确，docx 打开显示三行元信息 + 正文

---

## v1.1.1 — 2026-08-07

### 修改
- 输出格式：删除 `# 标题`、`公告日期`、`来源`、分隔线等头部行，md/docx 直接从正文开始
- `to_docx()` 相应简化：正文逐段输出，不再生成标题段落

### 验证
- 以第一份公告（烟草公司潼南项目）实测，docx 打开后正文直接呈现

---

## v1.1.0 — 2026-08-07

### 新增
- 定向爬虫 `zbgg_crawler.py`：重庆市政府采购招标公告（zbgg 栏目）近一个月内容
  - 列表页翻页收集：`collect_items()` 自动翻页至近一个月边界
  - 详情页反爬绕过：cqggzy.com 有 JS cookie 挑战（521），用 Windows Edge 无头渲染
  - 正文提取：多个 `.content` 分块拼接，去脚本/按钮噪声
  - 输出：每份公告 md + docx（python-docx），最后打包 zip
  - 元数据 `meta.json`：日期/标题/URL/文件路径索引
- 依赖补充：python-docx 1.2.0（本地 whl 安装，规避 PyPI 网络超时）

### 修复
- `make_zip()` 打包路径带目录前缀 → 改为相对文件夹名直接写 zip

### 验证
- 详情页 Edge 渲染实测通过（521 反爬绕过）
- extract_detail 提取正文 1307 字符，docx 生成 38KB
- zip 打包逻辑实测通过

---

## v1.0.0 — 2026-08-07

### 新增
- 通用爬虫脚手架核心 `spider.py`
  - 请求层：`fetch()` 带重试（默认 3 次）、超时、请求间隔限速
  - 解析层：CSS 选择器驱动，`item_selector` 定位条目容器，`fields` 定义字段映射
  - 字段选择器支持属性提取：`a@href`、`img@src`
  - 去重：按 `dedup_field`，未配置时用内容 hash
  - 输出：JSON / Markdown 双格式
- 入口 `main.py`：读取配置逐条规则执行
- 配置模板 `config.json`：示例规则（example.com 验证用）
- 备份机制 `backup.sh`：每次改动前备份源码到 `backups/`，时间戳命名，可选备注，自动打 zip
- 更新报告机制：本文件

### 修复
- `pick()` 只查后代节点，字段选择器与条目容器同标签时取不到值 → 增加节点自身匹配兜底

### 验证
- `python3 main.py` 跑通，example.com 解析 2 条记录并输出 JSON
