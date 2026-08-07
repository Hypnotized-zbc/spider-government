# spider-government 更新报告

本文件记录每次代码更新的内容，与 git 提交一一对应。

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
