# spider-government

重庆市政府采购招标公告爬虫。爬取近一个月内的招标公告，生成 docx + PDF 附件汇总压缩包，并输出 HTML 统计报表。

## 环境要求（必须）

- Windows 10/11 + WSL2（Ubuntu 22.04/24.04）—— 程序调用 powershell.exe、wslpath，访问 /mnt/c
- Windows 侧 Microsoft Edge（无头渲染详情页、下载 PDF 附件的核心通道）
- Python 3.10+（WSL 内）

## 安装

```bash
./install.sh          # 创建 venv 并安装依赖
```

或手动：

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 运行

```bash
source venv/bin/activate
python3 spider_newest.py              # 交互询问爬取范围（日期/序号/全部）
python3 spider_newest.py --limit 10   # 只处理前 10 份（测试用）
```

运行过程自动完成：
1. 自动启动 Windows Edge（无头 + CDP），无需手动操作
2. 翻页收集近一个月公告列表
3. 逐份渲染详情页，提取标题/编号/信息时间/正文/小标题/表格
4. 下载每份公告的 PDF 附件（招标公告.pdf、工程量清单.pdf 等）
5. 每份报告生成 docx
6. 所有报告打包为"招标报告汇总.zip"，自动复制到 Windows 桌面
7. 生成 HTML 统计报表（名称/编号/信息时间/投资金额/工期/附件）并自动打开

## 项目结构

```
spider_framework/
├── spider_newest.py      # 主程序（爬取/解析/下载/打包/报表）
├── main.py              # 通用爬虫脚手架入口（次要）
├── spider.py            # 通用爬虫脚手架核心（次要）
├── config.json          # 脚手架规则配置
├── requirements.txt     # Python 依赖
├── install.sh           # 一键安装
├── backup.sh            # 改动前备份脚本
├── UPDATES.md           # 更新报告
├── tools/               # PowerShell 工具脚本（Edge CDP 生命周期）
│   ├── start_edge_cdp.ps1   # 自动探测 Edge 并启动 CDP
│   ├── fetch_via_edge.ps1   # CDP 渲染详情页
│   └── download_v2.ps1      # CDP 下载附件
├── output/              # 爬取结果（不入库）
└── backups/             # 源码备份（不入库）
```

## 输出说明

| 输出 | 位置 |
|------|------|
| docx 报告 | output/zbgg_data/docx/ |
| PDF 附件 | output/zbgg_data/attachments/ |
| 汇总压缩包 | output/zbgg_zips/招标报告汇总.zip + Windows 桌面 |
| HTML 报表 | output/zbgg_report.html（自动打开） |

## 注意事项

- 目标站点有 JS 反爬（521），程序用 Edge 无头渲染绕过，首次运行较慢属正常
- 需要网络能访问 cq.gov.cn、cqggzy.com、ggzydl.cqggzy.com
- 桌面/临时目录均自动发现，不依赖固定用户名

## AI 价值报告（DeepSeek，可选）

HTML 报表顶部的「AI 动态价值报告」由 DeepSeek 大模型（deepseek-v4-flash）动态生成，
按公司画像判断每条公告的跟进价值（评分/等级/理由），失败时自动回退到内置规则评分。

配置 API key（二选一）：

```bash
export DEEPSEEK_API_KEY=sk-xxxx        # 方式1：环境变量
echo sk-xxxx > llm_key.txt             # 方式2：项目根目录本地文件（已 gitignore，不会上传）
```

未配置 key 时跳过 AI 报告，仅显示规则评分图表，不影响其余功能。

## 更新记录

见 [UPDATES.md](UPDATES.md)
