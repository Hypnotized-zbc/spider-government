#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
agent_nl.py —— 政府招标爬虫的自然语言指令入口（Agent 轻量化改造 Demo）
========================================================================

改造点：不改动 spider_newest.py 任何原有代码，仅在其上新增一个"调度层"：
自然语言指令 → LLM 解析为结构化动作计划（JSON 契约）→ 按计划逐个调用原有函数。

原有能力（全部复用，不重写）：
  collect_items / fetch_detail / extract_detail / to_markdown / to_docx /
  download_pdf_attachments / make_summary_zip / extract_key_info /
  generate_html_report / llm_chat / load_profiles / load_llm_key ...

支持指令示例：
  爬取近7天教育类招标并导出Excel、生成分析摘要
  分析已有数据中城口市政管网类招标，导出Excel并生成简要摘要
  导出全部已有公告的Excel统计表

离线演示（无 DeepSeek key 或 --offline）：自动降级为内置规则解析器 + 本地统计摘要。

用法：
  python3 agent_nl.py "爬取近7天教育类招标并导出Excel、生成分析摘要"
  python3 agent_nl.py --offline --demo 2
  python3 agent_nl.py --no-crawl "分析已有数据中市政类招标，导出Excel"
"""
import argparse
import datetime
import json
import re
import sys
from pathlib import Path

# ---- 引入原有工程（保持其自身路径解析不变） ----
PROJ = Path(__file__).resolve().parents[2]  # spider_framework/
sys.path.insert(0, str(PROJ))
import spider_newest as sn  # noqa: E402

DEMO_DIR = Path(__file__).resolve().parent
OUT_DIR = DEMO_DIR / "output"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# =====================================================================
# 一、意图解析：自然语言 → 结构化动作计划
# =====================================================================

PLAN_SYSTEM = """你是"重庆市政府采购招标爬虫"工具集的调度解析器。用户会用自然语言下达指令，你需要把指令解析成一个 JSON 动作计划，只输出 JSON，不要任何多余文字。

可用能力（工具）：
- crawl_and_analyze：爬取新公告（列表翻页 → 详情渲染 → md/docx → PDF 附件 → 汇总打包）
- analyze_existing：直接分析 output/zbgg_data/meta.json 里已有的公告数据
- export_excel：把筛选结果导出为 Excel
- summary：用大模型生成简要分析摘要
- html_report：调用原工程 generate_html_report 生成 HTML 统计报表

JSON 契约（严格遵循，缺失字段用 null / false）：
{
  "intent": "crawl_and_analyze" 或 "analyze_existing",
  "scope": {"days": null 或 数字, "count": null 或 数字, "date_start": null 或 "YYYY-MM-DD", "date_end": null 或 "YYYY-MM-DD"},
  "filters": {"region": null 或 地区名, "keywords": [] 或 关键词数组},
  "export_excel": true 或 false,
  "summary": true 或 false,
  "html_report": true 或 false,
  "profile": null 或 公司画像名
}

解析规则：
- 出现"爬取/抓取/更新/近N天/近N份" → intent=crawl_and_analyze
- "近N天" → scope.days=N；"近N份/最近N份/前N份/最新N份" → scope.count=N（N 是公告份数）
- 出现"已有/历史/存量/已爬"或用户没提爬取 → intent=analyze_existing
- "某地/地区名"（如 城口、沙坪坝、渝北）→ filters.region
- "教育类/学校/小学/医疗/医院/市政/管网/污水/道路/电力/光伏/智能化/安防"等业务类型 → 映射为 filters.keywords 具体词
- "导出/Excel/xlsx" → export_excel=true
- "摘要/分析/总结/报告" → summary=true
- "报表" → html_report=true"""


def _plan_skeleton() -> dict:
    return {
        "intent": "analyze_existing",
        "scope": {"days": None, "count": None, "date_start": None, "date_end": None},
        "filters": {"region": None, "keywords": []},
        "export_excel": False,
        "summary": True,
        "html_report": False,
        "profile": None,
    }


def parse_plan_llm(instruction: str) -> dict | None:
    """用 DeepSeek 解析指令（强制 JSON 输出）。无 key 或失败返回 None，由规则解析器兜底。"""
    key = sn.load_llm_key()
    if not key:
        return None
    text = sn.llm_chat(PLAN_SYSTEM, instruction)
    if not text:
        return None
    try:
        data = json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, re.S)
        if not m:
            return None
        try:
            data = json.loads(m.group(0))
        except Exception:
            return None
    plan = _plan_skeleton()
    plan.update({k: v for k, v in data.items() if k in plan})
    if isinstance(plan["scope"], dict):
        plan["scope"] = {**plan["scope"], **(data.get("scope") or {})}
    plan["intent"] = "crawl_and_analyze" if plan["intent"] == "crawl_and_analyze" else "analyze_existing"
    return plan


CATEGORY_KEYWORDS = {
    "教育": ["小学", "学校", "教学楼", "教育", "校园", "综合楼"],
    "医疗": ["医院", "卫生院", "医疗", "门诊"],
    "市政": ["管网", "污水", "给排水", "道路", "市政", "排水"],
    "电力": ["电力", "光伏", "储能", "配电", "充电桩"],
    "智能化": ["智能化", "弱电", "安防", "数据中心", "监控"],
    "房建": ["老旧小区", "住宅", "保障房", "安置房"],
    "监理": ["监理"],
}


def parse_plan_fallback(instruction: str) -> dict:
    """离线规则解析器：无 key / --offline / LLM 失败时兜底，保证 Demo 可演示。"""
    plan = _plan_skeleton()
    text = instruction.strip()

    if re.search(r"已有|历史|存量|已爬|已经爬", text):
        plan["intent"] = "analyze_existing"  # 明确指向已有数据时优先，避免"近N份"误判为爬取
    elif re.search(r"爬取|抓取|更新|近\d+天|近\d+份", text):
        plan["intent"] = "crawl_and_analyze"
    m = re.search(r"近\s*(\d+)\s*天", text)
    if m:
        plan["scope"]["days"] = int(m.group(1))
    m = re.search(r"(?:近|最近|最新|前)\s*(\d+)\s*份", text)
    if m:
        plan["scope"]["count"] = int(m.group(1))
    for pat in (r"(\d{4}-\d{2}-\d{2})\s*(?:到|至|~|—)\s*(\d{4}-\d{2}-\d{2})",
                r"从\s*(\d{4}-\d{2}-\d{2})\s*(?:到|至)\s*(\d{4}-\d{2}-\d{2})"):
        m = re.search(pat, text)
        if m:
            plan["scope"]["date_start"], plan["scope"]["date_end"] = m.group(1), m.group(2)
            break
    # 区域：优先匹配重庆已知区县名（避免把"市政/数据中市"这类误当地区）
    REGION_NAMES = ("渝北|渝中|江北|南岸|沙坪坝|九龙坡|大渡口|北碚|巴南|两江新区|"
                    "高新区|万州|涪陵|黔江|永川|江津|合川|大足|綦江|璧山|铜梁|潼南|"
                    "荣昌|开州|梁平|城口|丰都|垫江|忠县|云阳|奉节|巫山|巫溪|石柱|"
                    "秀山|酉阳|彭水|武隆|长寿|南川")
    m = re.search(rf"(?:{REGION_NAMES})(?:区|县|市)?", text)
    if m:
        plan["filters"]["region"] = m.group(0)
    else:
        m = re.search(r"([\u4e00-\u9fa5]{2,3}?(?:区|县))", text)
        if m:
            plan["filters"]["region"] = m.group(1)
    matched = False
    for cat, words in CATEGORY_KEYWORDS.items():
        if cat in text or any(w in text for w in words):
            plan["filters"]["keywords"] = words
            matched = True
            break
    if not matched:
        plan["filters"]["keywords"] = []
    plan["export_excel"] = bool(re.search(r"导出|excel|xlsx|表格", text, re.I))
    plan["summary"] = bool(re.search(r"摘要|分析|总结", text))
    plan["html_report"] = bool(re.search(r"报表", text))
    return plan


def parse_plan(instruction: str, offline: bool) -> tuple[dict, str]:
    if offline:
        return parse_plan_fallback(instruction), "规则解析器（--offline）"
    plan = parse_plan_llm(instruction)
    if plan is not None:
        return plan, "DeepSeek（JSON 契约）"
    plan = parse_plan_fallback(instruction)
    return plan, "规则解析器（LLM 不可用，兜底）"

# =====================================================================
# 二、调度执行：按计划调用原有函数
# =====================================================================

def log_tool(name: str, detail: str = "") -> None:
    print(f"  [调度] {name} {detail}", flush=True)


def load_meta() -> list[dict]:
    if sn.META_PATH.exists():
        return json.loads(sn.META_PATH.read_text(encoding="utf-8"))
    return []


def scope_filter(items: list[dict], scope: dict) -> list[dict]:
    """按 scope.days / date_start~date_end 过滤条目（复用原有 ask_scope 的日期语义）。"""
    if scope.get("days"):
        d1 = datetime.date.today() - datetime.timedelta(days=int(scope["days"]))
        d2 = datetime.date.today()
        return [it for it in items if d1 <= it["date_obj"] <= d2]
    if scope.get("date_start"):
        try:
            d1 = datetime.date.fromisoformat(scope["date_start"])
            d2 = datetime.date.fromisoformat(scope.get("date_end") or scope["date_start"])
            if d1 > d2:
                d1, d2 = d2, d1
            return [it for it in items if d1 <= it["date_obj"] <= d2]
        except ValueError:
            return items
    return items


def run_crawl(scope: dict, limit: int | None = None) -> list[dict]:
    """非交互式爬取：按计划的范围爬取新公告，返回本次结果集 cur_meta。
    逻辑与 spider_newest.main() 一致，仅去掉交互询问（ask_scope/ask_temp_company）。"""
    for d in (sn.DATA_DIR, sn.MD_DIR, sn.DOCX_DIR, sn.ATTACH_DIR):
        d.mkdir(parents=True, exist_ok=True)
    if not sn.ensure_edge():
        raise RuntimeError("Edge CDP 无法启动，终止爬取")

    meta = load_meta()
    done = {m["url"] for m in meta}
    failed = []
    if sn.FAILED_PATH.exists():
        failed = json.loads(sn.FAILED_PATH.read_text(encoding="utf-8"))

    items = sn.collect_items()
    seen, dedup = set(), []
    for it in items:
        key = (it["title"], it["date"])
        if key not in seen:
            seen.add(key)
            dedup.append(it)
    items = scope_filter(dedup, scope)
    if scope.get("count"):
        items = items[:int(scope["count"])]
    if limit:
        items = items[:limit]
    log_tool(f"collect_items() → {len(items)} 条（近一月内，按范围筛选 days={scope.get('days')} count={scope.get('count')}）")

    cur_meta = []
    for idx, it in enumerate(items, 1):
        if it["url"] in done:
            for m in meta:
                if m["url"] == it["url"]:
                    cur_meta.append(m)
                    break
            continue
        print(f"  [{idx}/{len(items)}] {it['date']} {it['title'][:40]}", flush=True)
        html = sn.fetch_detail(it["url"])
        if not html:
            failed.append(it)
            continue
        detail = sn.extract_detail(html)
        if not detail["body"]:
            failed.append(it)
            continue
        name = f"{it['date']}_{sn.safe_name(detail['title'])}"
        md_path = sn.MD_DIR / f"{name}.md"
        docx_path = sn.DOCX_DIR / f"{name}.docx"
        md_path.write_text(sn.to_markdown(it, detail), encoding="utf-8")
        sn.to_docx(md_path, docx_path, detail)
        att_dir = sn.ATTACH_DIR / name
        att_dir.mkdir(parents=True, exist_ok=True)
        pdfs = sn.download_pdf_attachments(detail.get("attachments", []), att_dir)
        rec = {"date": it["date"], "title": detail["title"], "url": it["url"],
               "md": str(md_path), "docx": str(docx_path), "attachments": pdfs}
        meta.append(rec)
        cur_meta.append(rec)
        done.add(it["url"])
        if len(meta) % 10 == 0:
            sn.META_PATH.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        time_sleep(0.5)

    sn.META_PATH.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    # 失败条目剔除不可 JSON 序列化的 date_obj（原工程同一隐患：failed 含 date 对象会崩）
    sn.FAILED_PATH.write_text(json.dumps(
        [{k: v for k, v in f.items() if k != "date_obj"} for f in failed],
        ensure_ascii=False, indent=2), encoding="utf-8")

    if cur_meta:
        zip_path = sn.OUT_DIR / "zbgg_zips" / "招标报告汇总.zip"
        sn.make_summary_zip(zip_path, cur_meta)
        log_tool("make_summary_zip() → " + str(zip_path))
    return cur_meta


def time_sleep(s: float) -> None:
    import time
    time.sleep(s)


def enrich_record(rec: dict) -> dict:
    """补充字段：投资金额 / 工期 / 建设地点（复用 extract_key_info + 正则）。"""
    out = dict(rec)
    body = ""
    if rec.get("md"):
        try:
            body = Path(rec["md"]).read_text(encoding="utf-8")
        except Exception:
            body = ""
    info = sn.extract_key_info(body)
    out["investment"] = info.get("investment", "")
    out["duration"] = info.get("duration", "")
    m = re.search(r"建设地点[：:]\s*([^。\n；]+)", body)
    out["region"] = m.group(1).strip() if m else ""
    out["_body"] = body
    return out


def filter_records(records: list[dict], filters: dict) -> list[dict]:
    """按 region / keywords 过滤，命中信息写入字段，便于导出与摘要。"""
    region = filters.get("region")
    keywords = filters.get("keywords") or []
    out = []
    for rec in records:
        rec = enrich_record(rec)
        text = rec["title"] + "\n" + rec["_body"]
        hit_region = bool(region and region in text)
        hit_kw = [k for k in keywords if k in text]
        if region and not hit_region:
            continue
        if keywords and not hit_kw:
            continue
        rec["hit_region"] = region if hit_region else ""
        rec["hit_keywords"] = "、".join(hit_kw)
        out.append(rec)
    return out


def export_excel(records: list[dict], path: Path) -> str | None:
    """导出 Excel（优先 openpyxl，缺失时降级 CSV，Excel 可直接打开）。"""
    headers = ["序号", "信息时间", "项目标题", "项目编号", "投资金额", "工期",
               "建设地点", "命中区域", "命中关键词", "附件"]
    rows = []
    for i, r in enumerate(records, 1):
        m = re.search(r"项目编号[：:]\s*(\S+)", r.get("_body", ""))
        atts = "、".join(a["name"] for a in r.get("attachments", []))
        rows.append([i, r.get("date", ""), r.get("title", ""), m.group(1) if m else "",
                     r.get("investment", ""), r.get("duration", ""), r.get("region", ""),
                     r.get("hit_region", ""), r.get("hit_keywords", ""), atts])
    try:
        import openpyxl  # noqa: PLC0415
    except ImportError:
        csv_path = path.with_suffix(".csv")
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            import csv
            w = csv.writer(f)
            w.writerow(headers)
            w.writerows(rows)
        return str(csv_path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "招标筛选结果"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    for col, width in zip("ABCDEFGHIJ", (6, 12, 42, 24, 14, 22, 18, 10, 22, 30)):
        ws.column_dimensions[col].width = width
    wb.save(path)
    return str(path)


def _local_stats(records: list[dict], plan: dict) -> str:
    """无 key 或 LLM 失败时的本地统计摘要（保证任何环境都有输出）。"""
    total_inv = 0.0
    for r in records:
        v = re.search(r"([\d.]+)", r.get("investment", ""))
        if v:
            total_inv += float(v.group(1))
    kws = "、".join(plan.get("filters", {}).get("keywords") or []) or "无"
    return (f"共 {len(records)} 条公告；识别到总投资约 {total_inv:.1f} 万元；"
            f"附件齐全 {sum(1 for r in records if r.get('attachments'))} 份；"
            f"命中关键词：{kws}。")


def gen_summary(records: list[dict], plan: dict) -> str:
    """生成简要分析摘要：LLM 优先，无 key 或 LLM 失败时回退本地统计。"""
    if not records:
        return "筛选结果为空，无摘要。"
    rows = []
    for r in records:
        rows.append({
            "title": r.get("title", ""),
            "date": r.get("date", ""),
            "investment": r.get("investment", "-"),
            "duration": r.get("duration", "-"),
            "attachments": "、".join(a["name"] for a in r.get("attachments", [])),
            "body": r.get("_body", "")[:300],
        })
    if not sn.load_llm_key():
        return _local_stats(records, plan) + "（未配置 DeepSeek key，跳过 AI 摘要）"
    system = ("你是资深工程招投标分析师。根据给定公告列表输出简洁的中文分析摘要，"
              "只输出摘要正文，不要任何前后缀。")
    user = (f"共 {len(rows)} 条公告（JSON）：\n{json.dumps(rows, ensure_ascii=False)}\n\n"
            "请输出：1) 总体判断（一句话）；2) 数量与投资规模统计；3) 类型/区域分布；"
            "4) 每条公告的跟进建议（重点跟进/值得关注/一般/忽略 + 一句话理由）；"
            "5) 风险提示。全文 300 字以内，无空话。")
    text = sn.llm_chat(system, user)
    if text:
        return text
    return _local_stats(records, plan) + "（LLM 摘要失败，回退本地统计）"


def run_html_report(records: list[dict], plan: dict, out_path: Path) -> str:
    """调用原工程 generate_html_report 生成 HTML 报表（含 DeepSeek 价值卡片，若有 key）。"""
    profiles = []
    prof_name = plan.get("profile")
    if prof_name:
        for p in sn.load_profiles():
            if prof_name in (p.get("name", ""), Path(p.get("_file", "")).stem):
                profiles.append(p)
                break
    if not profiles:
        profiles = [sn.DEFAULT_PROFILE]
    sn.generate_html_report(records, out_path, profiles=profiles)
    return str(out_path)


def dispatch(plan: dict, offline: bool, no_crawl: bool = False, limit: int | None = None) -> None:
    print("\n=== Agent 调度执行（工具调用轨迹）===", flush=True)
    print(f"[计划] {json.dumps(plan, ensure_ascii=False)}", flush=True)

    if plan["intent"] == "crawl_and_analyze" and not no_crawl:
        records = run_crawl(plan["scope"], limit=limit)
        print(f"  爬取完成：本次 {len(records)} 份（历史累计 {len(load_meta())} 份）", flush=True)
    else:
        records = load_meta()
        log_tool("load_meta() → " + str(len(records)) + " 条已有记录")
        if not records:
            print("meta.json 无数据。请先运行 spider_newest.py 爬取，或用『爬取…』指令。", flush=True)
            return

    filtered = filter_records(records, plan["filters"])
    log_tool("filter_records()", f"region={plan['filters'].get('region')} keywords={plan['filters'].get('keywords')} → {len(filtered)} 条")

    if plan["scope"].get("count"):
        filtered = sorted(filtered, key=lambda r: r.get("date", ""), reverse=True)[:int(plan["scope"]["count"])]
        log_tool(f"scope.count 截取 → 最近 {len(filtered)} 条")

    if not filtered:
        print("筛选后无匹配公告，终止。", flush=True)
        return

    if plan.get("html_report"):
        html_path = OUT_DIR / f"筛选报表_{datetime.date.today()}.html"
        p = run_html_report(filtered, plan, html_path)
        log_tool("generate_html_report() → " + p)

    excel_path = None
    if plan.get("export_excel"):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = export_excel(filtered, OUT_DIR / f"招标筛选结果_{ts}.xlsx")
        log_tool("export_excel() → " + str(excel_path))

    summary = None
    if plan.get("summary"):
        summary = gen_summary(filtered, plan)
        log_tool("llm_chat(summary) → " + (summary[:40] + "…" if len(summary) > 40 else summary))

    print("\n=== 执行结果 ===", flush=True)
    print(f"命中 {len(filtered)} 条公告：", flush=True)
    for r in filtered:
        print(f"  - {r['date']} {r['title']}", flush=True)
    if excel_path:
        print(f"Excel 导出：{excel_path}", flush=True)
    if summary:
        print(f"\n分析摘要：\n{summary}", flush=True)
    print("\n（注：预测/分析仅供技术演示，不构成投资建议）", flush=True)


# =====================================================================
# 三、入口
# =====================================================================

DEMO_INSTRUCTIONS = [
    "爬取近7天教育类招标并导出Excel、生成分析摘要",
    "分析已有数据中市政管网类招标，导出Excel并生成简要分析摘要",
    "导出全部已有公告的Excel统计表，并生成分析摘要",
]


def main() -> None:
    ap = argparse.ArgumentParser(description="政府招标爬虫 · 自然语言指令入口（Agent 轻量改造）")
    ap.add_argument("instruction", nargs="?", default=None, help="自然语言指令；缺省进入交互输入")
    ap.add_argument("-i", "--instruction", dest="i", default=None, help="同上（-i 形式）")
    ap.add_argument("--offline", action="store_true", help="强制使用规则解析器（不调用 LLM）")
    ap.add_argument("--no-crawl", action="store_true", help="即使指令含“爬取”也只在已有数据上执行")
    ap.add_argument("--demo", type=int, choices=[1, 2, 3], help="运行内置演示指令 N")
    ap.add_argument("--limit", type=int, default=None, help="爬取时最多处理前 N 条")
    args = ap.parse_args()

    if args.demo:
        instruction = DEMO_INSTRUCTIONS[args.demo - 1]
    elif args.i:
        instruction = args.i
    elif args.instruction:
        instruction = args.instruction
    else:
        instruction = input("请输入自然语言指令（回车退出）：").strip()
        if not instruction:
            return

    plan, parser = parse_plan(instruction, offline=args.offline)
    print(f"指令: {instruction}", flush=True)
    print(f"意图解析（{parser}）: {json.dumps(plan, ensure_ascii=False)}", flush=True)

    dispatch(plan, offline=args.offline, no_crawl=args.no_crawl, limit=args.limit)


if __name__ == "__main__":
    main()
