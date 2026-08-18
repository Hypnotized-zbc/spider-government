# spider-government

重庆市政府采购招标公告爬虫。爬取近一个月内的招标公告，生成 docx + PDF 附件汇总压缩包，并输出 HTML 统计报表；支持按公司画像做 DeepSeek 动态价值分析，并提供自然语言指令入口（Agent 轻量化改造 Demo）。

## 功能一览

- **全流程爬取**：列表页自动翻页 → 详情页渲染 → 提取标题 / 编号 / 信息时间 / 正文 / 小标题 / 表格
- **文档产出**：每份公告生成排版还原的 docx + 下载 PDF 附件（招标公告 / 工程量清单等）
- **汇总与报表**：全部报告打包「招标报告汇总.zip」自动复制到 Windows 桌面；生成 HTML 统计报表（名称 / 编号 / 信息时间 / 投资金额 / 工期 / 附件）并自动打开
- **交互式范围选择**：按日期范围 / 按序号范围 / 全部（近一个月），越界自动回退
- **公司画像价值判断**：5 家虚拟公司画像（渝发建设 / 市政环境 / 电力能源 / 智能化弱电 / 公路交通）+ 临时公司交互录入（不落盘）
- **DeepSeek 动态价值报告**：以公司画像为视角，对每条公告输出 0–100 评分 / 跟进等级 / 理由 / 建议，报表按 AI 评分降序
- **自然语言指令入口**（`demo/agent_nl/`）：输入「爬取近7天教育类招标并导出Excel、生成分析摘要」等指令即可驱动整个流程，详见 [demo/agent_nl/README.md](demo/agent_nl/README.md)

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

爬取完成后还会询问是否录入临时公司（公司名称/关键词/预算/关注区域，每项可回车跳过，名称回车结束）——临时公司不保存到 company_profiles，仅本次运行作为 AI 价值报告的公司视角。

运行过程自动完成：

1. 自动启动 Windows Edge（无头 + CDP），无需手动操作
2. 翻页收集近一个月公告列表
3. 逐份渲染详情页，提取标题/编号/信息时间/正文/小标题/表格
4. 下载每份公告的 PDF 附件（招标公告.pdf、工程量清单.pdf 等）
5. 每份报告生成 docx
6. 所有报告打包为「招标报告汇总.zip」，自动复制到 Windows 桌面
7. 生成 HTML 统计报表（名称/编号/信息时间/投资金额/工期/附件）并自动打开

### 自然语言指令入口（Agent 轻量化改造 Demo）

```bash
# 分析已有数据（演示最安全，不触发爬取）
python3 demo/agent_nl/agent_nl.py --offline --demo 2

# 爬取近 5 份公告并导出 Excel（需 Edge 可用）
python3 demo/agent_nl/agent_nl.py "爬取近5份招标公告并导出Excel"

# 完整指令示例与原理说明见 demo/agent_nl/README.md
```

## 部署方式

本项目是 Windows + WSL2 环境下的本地/单机工具，部署形态即环境搭建：

1. **WSL2 侧**：`./install.sh` 创建 venv 并安装依赖；
2. **Windows 侧**：安装 Microsoft Edge（无头渲染依赖，标准路径即可，程序自动探测）；
3. **网络**：需能访问 cq.gov.cn、cqggzy.com、ggzydl.cqggzy.com；
4. **定时运行（可选）**：Windows 计划任务调用 WSL 命令，参考 `run_ak_crawler.bat` 的写法（WSL 执行 + 日志落盘）。

> 已知限制：Edge CDP 方案绑定 Windows 环境；后续优化方向为迁移 Playwright 实现纯 Linux 部署（见 UPDATES.md）。

## 项目结构

```
spider_framework/
├── spider_newest.py      # 主程序（爬取/解析/下载/打包/报表/AI 价值报告）
├── company_profiles/     # 5 家公司画像 JSON（渝发建设/市政环境/电力能源/智能化弱电/公路交通）
├── tools/                # 工具脚本
│   ├── start_edge_cdp.ps1   # 自动探测 Edge 并启动 CDP
│   ├── fetch_via_edge.ps1   # CDP 渲染详情页
│   ├── download_v2.ps1      # CDP 下载附件
│   ├── score_reports.py     # 离线批量评分（历史工具，评分逻辑已由 AI 报告替代）
│   └── upload_github.py     # GitHub 上传脚本
├── demo/agent_nl/        # Agent 轻量化改造：自然语言指令入口（独立 README）
├── install.sh            # 一键安装
├── backup.sh             # 改动前备份脚本
├── requirements.txt      # Python 依赖
├── README.md             # 本文件
└── UPDATES.md            # 更新报告（与 git 提交一一对应）
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

## AI 价值报告（DeepSeek）

HTML 报表中的「AI 动态价值报告」由 DeepSeek 大模型（deepseek-v4-flash）动态生成，按公司画像判断每条公告的跟进价值（评分/等级/理由），是唯一的价值判断来源；调用失败时报表仍正常展示基础表格（按原顺序）。

配置 API key（二选一）：

```bash
export DEEPSEEK_API_KEY=sk-xxxx        # 方式1：环境变量
echo sk-xxxx > llm_key.txt             # 方式2：项目根目录本地文件（已 gitignore，不会上传）
```

未配置 key 时跳过 AI 报告，报表只显示基础表格，不影响其余功能。

## 更新记录

见 [UPDATES.md](UPDATES.md)
