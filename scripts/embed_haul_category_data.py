#!/usr/bin/env python3
"""Embed the reviewed Haul ASIN category mapping in the single-file dashboard."""

import argparse
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("mapping", type=Path, nargs="?", default=Path("reports/1店_可售ASIN品类_2026-10-05.json"))
    parser.add_argument("html", type=Path, nargs="?", default=Path("货盘动销分析_WK29-WK30_可下载版.html"))
    args = parser.parse_args()
    mapping = json.loads(args.mapping.read_text(encoding="utf-8"))
    if not mapping or not all(isinstance(value, list) and len(value) == 2 for value in mapping.values()):
        raise ValueError("Invalid category mapping")
    start = "var HAUL_CATEGORY_ASINS=/* HAUL_CATEGORY_MAP_START */"
    end = "/* HAUL_CATEGORY_MAP_END */;"
    html = args.html.read_text(encoding="utf-8")
    if html.count(start) != 1 or html.count(end) != 1:
        raise ValueError("Missing or duplicate HAUL_CATEGORY_MAP markers")
    before, rest = html.split(start, 1)
    _, after = rest.split(end, 1)
    payload = json.dumps(mapping, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    args.html.write_text(before + start + payload + end + after, encoding="utf-8")
    print(f"Embedded {len(mapping)} ASIN classifications in {args.html}")


if __name__ == "__main__":
    main()
