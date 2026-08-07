# -*- coding: utf-8 -*-
"""
通用爬虫脚手架
==============
核心类 Spider：请求 -> 解析 -> 输出，规则由 config.json 驱动。

扩展方式（后续按需填）：
  1. 新站点：在 config.json 里加一条 rule
  2. 新字段：在 rule.fields 里加选择器
  3. 特殊解析：子类化 Spider，重写 parse_item / extract
  4. 自定义输出：重写 save / 加 pipeline
"""

import hashlib
import json
import logging
import time
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("spider")


class Spider:
    """通用爬虫基类。"""

    def __init__(self, config: dict):
        self.name = config.get("name", "spider")
        self.output_dir = Path(config.get("output_dir", "output"))
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # 请求参数
        self.timeout = config.get("timeout", 15)
        self.retries = config.get("retries", 3)
        self.delay = config.get("delay", 1.0)          # 每次请求间隔（秒），限速用
        self.headers = config.get("headers", {})
        self.headers.setdefault("User-Agent",
                                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Spider")

        self.session = requests.Session()
        self.seen_urls = set()      # 本轮去重
        self.items = []             # 解析结果汇总

    # ---------- 请求层 ----------

    def fetch(self, url: str) -> str | None:
        """带重试的 GET，返回 HTML 文本；失败返回 None。"""
        for attempt in range(1, self.retries + 1):
            try:
                resp = self.session.get(url, headers=self.headers,
                                        timeout=self.timeout)
                resp.raise_for_status()
                resp.encoding = resp.apparent_encoding or resp.encoding
                return resp.text
            except requests.RequestException as e:
                log.warning("请求失败 %s (第 %d 次): %s", url, attempt, e)
                time.sleep(self.delay * attempt)
        return None

    # ---------- 解析层 ----------

    def parse(self, html: str, rule: dict) -> list[dict]:
        """按 rule 解析页面，返回条目列表。"""
        soup = BeautifulSoup(html, "html.parser")
        items = []

        # 列表容器：rule.item_selector 定位每条记录的根节点
        if rule.get("item_selector"):
            nodes = soup.select(rule["item_selector"])
        else:
            nodes = [soup]

        for node in nodes:
            item = self.extract(node, rule)
            if item:
                items.append(item)
        return items

    def extract(self, node, rule: dict) -> dict | None:
        """从单个节点提取字段。fields: {字段名: CSS选择器}"""
        item = {}
        for field, selector in rule.get("fields", {}).items():
            item[field] = self.pick(node, selector)
        # 校验必填字段
        for required in rule.get("required", []):
            if not item.get(required):
                return None
        return item

    @staticmethod
    def pick(node, selector: str) -> str:
        """提取单个字段文本。selector 支持属性: 'a@href'、'img@src'。
        先查后代，查不到时测试节点自身是否匹配（字段与容器同标签的场景）。"""
        if "@" in selector:
            css, attr = selector.rsplit("@", 1)
            if not css:  # 形如 '@href'：直接取节点自身属性
                return node.get(attr, "").strip()
            el = Spider._first_match(node, css)
            return el.get(attr, "").strip() if el else ""
        el = Spider._first_match(node, selector)
        return el.get_text(strip=True) if el else ""

    @staticmethod
    def _first_match(node, css: str):
        """返回 node 后代中第一个匹配，或 node 自身。"""
        el = node.select_one(css)
        if el is not None:
            return el
        # 后代无匹配：用临时 soup 测自身
        try:
            tmp = BeautifulSoup(str(node), "html.parser")
            return tmp.select_one(css)
        except Exception:
            return None

    # ---------- 去重 ----------

    def dedup_key(self, item: dict, rule: dict) -> str:
        """默认用 rule.dedup_field 的值做去重；没有则用内容 hash。"""
        field = rule.get("dedup_field")
        if field and item.get(field):
            return str(item[field])
        return hashlib.md5(json.dumps(item, ensure_ascii=False).encode()).hexdigest()

    # ---------- 输出层 ----------

    def save(self, fmt: str = "json") -> Path:
        """保存结果。fmt: json / md"""
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = self.output_dir / f"{self.name}_{ts}.{fmt}"
        if fmt == "json":
            path.write_text(json.dumps(self.items, ensure_ascii=False,
                                       indent=2), encoding="utf-8")
        elif fmt == "md":
            self.write_md(path)
        log.info("已保存 %d 条 -> %s", len(self.items), path)
        return path

    def write_md(self, path: Path):
        """默认 markdown 输出，子类可覆盖。"""
        lines = [f"# {self.name}", ""]
        for item in self.items:
            for k, v in item.items():
                lines.append(f"**{k}**: {v}")
            lines.append("---")
        path.write_text("\n".join(lines), encoding="utf-8")

    # ---------- 主流程 ----------

    def run(self, rule: dict, fmt: str = "json") -> list[dict]:
        """执行一条规则。"""
        self.items = []
        for url in rule["urls"]:
            if url in self.seen_urls:
                continue
            self.seen_urls.add(url)
            log.info("抓取 %s", url)
            html = self.fetch(url)
            if html is None:
                continue
            items = self.parse(html, rule)
            for item in items:
                key = self.dedup_key(item, rule)
                if key in self.seen_urls:
                    continue
                self.seen_urls.add(key)
                self.items.append(item)
            time.sleep(self.delay)

        if self.items:
            self.save(fmt)
        else:
            log.warning("没有解析到任何条目")
        return self.items


def load_config(path: str = "config.json") -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)
