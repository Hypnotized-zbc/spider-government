# -*- coding: utf-8 -*-
"""爬虫入口：python3 main.py [config.json]"""
import sys

from spider import Spider, load_config


def main():
    config_path = sys.argv[1] if len(sys.argv) > 1 else "config.json"
    config = load_config(config_path)
    spider = Spider(config)

    for rule in config["rules"]:
        fmt = rule.get("fmt", config.get("fmt", "json"))
        items = spider.run(rule, fmt=fmt)
        print(f"[{rule.get('name', '?')}] 共 {len(items)} 条")


if __name__ == "__main__":
    main()
