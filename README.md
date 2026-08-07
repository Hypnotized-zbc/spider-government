# spider-government

通用爬虫脚手架，配置驱动，开箱即用。

## 环境要求

- Python 3.10+
- requests
- beautifulsoup4

```bash
pip install requests beautifulsoup4
```

## 快速开始

```bash
python3 main.py                # 使用默认 config.json
python3 main.py config.json    # 指定配置
```

## 项目结构

```
spider_framework/
├── spider.py      # 核心爬虫类（请求/解析/去重/输出）
├── main.py        # 入口，逐规则执行
├── config.json    # 规则配置
├── backup.sh      # 改动前备份脚本
├── UPDATES.md     # 更新报告
├── output/        # 爬取结果输出（不入库）
└── backups/       # 源码备份（不入库）
```

## 配置说明

`config.json` 中每条 `rule` 定义一个爬取任务：

| 字段 | 说明 |
|------|------|
| `name` | 规则名 |
| `urls` | 目标 URL 列表 |
| `item_selector` | 条目容器 CSS 选择器 |
| `fields` | 字段名 → CSS 选择器，支持 `a@href` 属性语法 |
| `required` | 必填字段，缺失则该条目丢弃 |
| `dedup_field` | 去重依据字段，缺省用内容 hash |
| `fmt` | 输出格式 json / md |

## 扩展方式

- 新站点：在 `config.json` 加一条 `rule`
- 特殊解析：子类化 `Spider`，重写 `parse_item` / `extract`
- 自定义输出：重写 `save` / `write_md`

## 更新记录

见 [UPDATES.md](UPDATES.md)
