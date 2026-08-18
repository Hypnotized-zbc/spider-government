# Agent 轻量化改造 Demo：自然语言指令驱动招标爬虫

> 收尾任务 Day1 要求：让 Agent 作为**调度者**调用你写好的工具，完成对 Agent 工具调用机制的真实理解。
> 本目录不改动 `spider_newest.py` 任何原有代码，仅新增一个调度层（约 450 行），原有工程保持原样可运行。

## 一、改造点（做了什么）

| 层 | 说明 |
|---|---|
| 原有能力（不动） | `spider_newest.py` 的 collect_items / fetch_detail / extract_detail / to_markdown / to_docx / download_pdf_attachments / make_summary_zip / extract_key_info / generate_html_report / llm_chat / load_profiles 等，全部原样复用 |
| 新增调度层 | `demo/agent_nl/agent_nl.py`：自然语言指令 → 结构化动作计划 → 逐个调用上述函数 |
| 意图解析 | 一次 DeepSeek 调用（`response_format: json_object` + 严格 JSON 契约：intent / scope / filters / export_excel / summary / html_report / profile）；无 key 或失败时自动降级为内置规则解析器（正则启发式），保证 Demo 离线可演示 |
| 新增能力（少量） | ① 按区域/关键词过滤已有或新爬公告（复用 extract_key_info 补充投资额/工期/建设地点）；② Excel 导出（openpyxl，缺失时降级 CSV，Excel 可直接打开）；③ 简要分析摘要（LLM 生成，无 key 时输出本地统计） |
| 工具调用轨迹 | 执行过程打印每个被调用的原函数及其入参/结果，直观展示"Agent 调度工具" |

**不重写原则**：爬取主循环与 `spider_newest.main()` 逐行对应，仅去掉交互询问（ask_scope / ask_temp_company），由解析出的 scope / filters 替代人工输入。

## 二、运行方式

```bash
# 环境：复用项目 venv（requests / beautifulsoup4 / python-docx / websocket-client）
# 额外依赖（可选，仅 Excel 导出需要）：
pip install openpyxl

# 1) 自然语言指令（LLM 解析，需配置 DeepSeek key：llm_key.txt 或 DEEPSEEK_API_KEY）
python3 demo/agent_nl/agent_nl.py "爬取近7天教育类招标并导出Excel、生成分析摘要"

# 1b) 只爬最近 N 份（不再爬满近一个月；"近N份/最近N份/前N份/最新N份"均支持）
python3 demo/agent_nl/agent_nl.py "爬取近5份招标公告并导出Excel"

# 2) 只分析已有数据（不触发爬取；演示最安全）
python3 demo/agent_nl/agent_nl.py --no-crawl "分析已有数据中市政管网类招标，导出Excel并生成简要分析摘要"

# 3) 离线演示（不调 LLM，规则解析 + 本地统计摘要）
python3 demo/agent_nl/agent_nl.py --offline --demo 2

# 4) 内置演示指令
python3 demo/agent_nl/agent_nl.py --demo 1     # 爬取近7天教育类…
python3 demo/agent_nl/agent_nl.py --demo 2     # 分析已有数据中市政管网类…
python3 demo/agent_nl/agent_nl.py --demo 3     # 导出全部已有公告 Excel 统计表

# 5) 交互模式：不带指令参数，回车退出
python3 demo/agent_nl/agent_nl.py
```

说明：
- 指令含「爬取/近N天/近N份」→ 走 `crawl_and_analyze`，会真实调用 Edge CDP 爬取（需 Windows 侧 Edge，同原工程）；
  「近N天」限定日期范围，「近N份」限定公告份数（列表按新→旧，取最新 N 份，不会爬满近一个月）；
- 想先看效果又不想动网络 → 加 `--no-crawl` 或说「已有/历史数据」；
- `--limit N` 限制爬取条数（对应原工程 `--limit` 的测试语义）。

## 三、效果演示（基于 output/zbgg_data 已有 6 条真实公告）

```bash
python3 demo/agent_nl/agent_nl.py --offline --demo 2
```

输出示意（工具调用轨迹 + 结果）：

```
指令: 分析已有数据中市政管网类招标，导出Excel并生成简要分析摘要
意图解析（规则解析器（--offline））: {"intent": "analyze_existing", "scope": {...}, "filters": {"region": null, "keywords": ["管网","污水","给排水","道路","市政","排水"]}, "export_excel": true, "summary": true, ...}

=== Agent 调度执行（工具调用轨迹）===
  [计划] {...}
  [调度] load_meta() → 6 条已有记录
  [调度] filter_records() region=None keywords=['管网','污水','给排水','道路','市政','排水'] → 3 条
  [调度] export_excel() → demo/agent_nl/output/招标筛选结果_20260817_xxxxxx.xlsx
  [调度] llm_chat(summary) → 共 3 条公告；总投资约 8.3 亿元；…

=== 执行结果 ===
命中 3 条公告：
  - 2026-08-07 城口县特色产业园片区排水管网更新改造工程(高燕产业园区)
  - 2026-08-07 井口老旧街区改造监理
  - 2026-08-07 黔江区城东街道南海城片区老旧小区改造项目（二期）工程总承包

分析摘要：
（…总体判断 / 数量与投资规模 / 跟进建议 / 风险提示…）

Excel 导出：demo/agent_nl/output/招标筛选结果_20260817_xxxxxx.xlsx
```

其他演示命令：

```bash
# 教育类（命中 开州区歇马小学综合楼新建项目第二标段）
python3 demo/agent_nl/agent_nl.py --offline --demo 2 换成指令：
python3 demo/agent_nl/agent_nl.py --no-crawl --offline "分析已有数据中教育类招标，导出Excel并生成摘要"

# 区域过滤（城口县）
python3 demo/agent_nl/agent_nl.py --no-crawl --offline "分析已有数据中城口县招标，导出Excel"

# 带 DeepSeek key 时（LLM 解析 + LLM 摘要，效果见上方摘要段）
python3 demo/agent_nl/agent_nl.py --no-crawl "分析已有数据中市政管网类招标，导出Excel并生成简要分析摘要"
```

## 四、目录结构

```
spider_framework/
├── spider_newest.py          # 原有工程（未改动）
├── demo/
│   └── agent_nl/
│       ├── agent_nl.py       # 自然语言指令入口（调度层，本次新增）
│       ├── README.md         # 本文件
│       ├── requirements.txt  # 额外依赖（openpyxl，可选）
│       └── output/           # 演示输出（Excel / HTML 报表，不入库）
```

## 五、设计要点（对应考核点）

1. **不推倒重来**：spider_newest.py 零改动，通过 `sys.path` 引入并复用其全部函数；
2. **自然语言驱动**：指令 → LLM JSON 契约解析（无 key 自动降级规则解析），覆盖"爬取范围 / 区域 / 业务关键词 / 是否导出 Excel / 是否生成摘要"；
3. **工具调用机制的真实呈现**：执行过程打印 `[调度] 工具名(入参) → 结果` 轨迹，就是 Agent 工具调用的最小实现——LLM 只负责"决定做什么"，执行仍走你写好的确定性工具；
4. **容错**：LLM 解析失败 / 无 key / openpyxl 缺失均有降级路径，任何环境都能跑通演示。
