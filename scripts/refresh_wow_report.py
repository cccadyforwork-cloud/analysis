"""Embed a two-period Business Report comparison for the current sellable store Listing."""

import argparse
import csv
import hashlib
import json
import re
from collections import defaultdict
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_HTML = ROOT / "货盘动销分析_WK29-WK30_可下载版.html"


def embedded(html, name):
    match = re.search(r"^var " + name + r"\s*=\s*([^\n;]+);$", html, re.M)
    if not match:
        raise ValueError(f"Missing embedded {name}")
    return json.loads(match.group(1))


def replace_var(html, name, value):
    payload = json.dumps(value, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    pattern = r"^var " + name + r"\s*=\s*[^\n;]+;$"
    updated, count = re.subn(pattern, lambda _: f"var {name} = {payload};", html, count=1, flags=re.M)
    if count != 1:
        raise ValueError(f"Could not replace {name}")
    return updated


def amount(value):
    cleaned = re.sub(r"[^\d.\-]", "", str(value or ""))
    return float(cleaned or 0)


def report(path):
    required = {"（子）ASIN", "标题", "会话数 - 总计", "已订购商品数量", "已订购商品销售额"}
    records = defaultdict(lambda: {"n": "", "g": 0, "q": 0, "u": 0, "s": 0.0})
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers = {header.strip() for header in reader.fieldnames or []}
        if not required.issubset(headers) or not any(header.startswith("页面浏览量 - 总计") for header in headers):
            raise ValueError(f"{path.name}: missing required Business Report columns")
        views_key = next(header for header in reader.fieldnames if header.strip() == "页面浏览量 - 总计")
        for row in reader:
            asin = (row["（子）ASIN"] or "").strip().upper()
            if not re.fullmatch(r"[A-Z0-9]{10}", asin):
                continue
            item = records[asin]
            item["n"] = item["n"] or (row["标题"] or "").strip()
            item["g"] += round(amount(row[views_key]))
            item["q"] += round(amount(row["会话数 - 总计"]))
            item["u"] += round(amount(row["已订购商品数量"]))
            item["s"] += amount(row["已订购商品销售额"])
    if not records:
        raise ValueError(f"{path.name}: no valid ASINs")
    return dict(records)


def weekly_total(label, rows, listing):
    asins = [asin for asin in rows if asin in listing]
    values = [rows[asin] for asin in asins]
    units = sum(item["u"] for item in values)
    sales = round(sum(item["s"] for item in values), 2)
    active = len(asins)
    awas = sum(item["u"] > 0 for item in values)
    return {"label": label, "gv": sum(item["g"] for item in values), "sessions": sum(item["q"] for item in values),
            "units": units, "sales": sales, "asp": round(sales / units, 4) if units else 0,
            "awas": awas, "active": active, "rate": round(awas / active * 100, 2) if active else 0}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("previous", type=Path)
    parser.add_argument("current", type=Path)
    parser.add_argument("--html", type=Path, default=DEFAULT_HTML)
    parser.add_argument("--previous-start", required=True)
    parser.add_argument("--previous-end", required=True)
    parser.add_argument("--current-start", required=True)
    parser.add_argument("--current-end", required=True)
    parser.add_argument("--store-id", default="1店")
    args = parser.parse_args()
    previous_start, previous_end = date.fromisoformat(args.previous_start), date.fromisoformat(args.previous_end)
    current_start, current_end = date.fromisoformat(args.current_start), date.fromisoformat(args.current_end)
    if previous_end >= current_start or (previous_end - previous_start) != (current_end - current_start):
        raise ValueError("Periods must be ordered, nonoverlapping, and have the same number of days")
    previous, current = report(args.previous), report(args.current)
    html = args.html.read_text(encoding="utf-8")
    scope = json.dumps([args.store_id, "US"], ensure_ascii=False, separators=(",", ":"))
    listing = embedded(html, "OPS_STORE_LISTINGS").get(scope)
    if not listing:
        raise ValueError(f"No embedded Listing for {scope}")
    sellable = {asin: row for asin, row in listing["rows"].items() if float(row["i"]) > 0}
    cur_asins = sorted(current.keys() & sellable.keys())
    prev_asins = sorted(previous.keys() & sellable.keys())
    records = []
    for asin in sorted(set(cur_asins) | set(prev_asins)):
        c, p = current.get(asin), previous.get(asin)
        records.append({"a": asin, "n": sellable[asin]["n"] or (c or p)["n"],
                        "cg": c["g"] if c else 0, "cq": c["q"] if c else 0,
                        "cu": c["u"] if c else 0, "cs": round(c["s"], 2) if c else 0,
                        "pg": p["g"] if p else 0, "pq": p["q"] if p else 0,
                        "pu": p["u"] if p else 0, "ps": round(p["s"], 2) if p else 0,
                        "cp": c is not None, "pp": p is not None})
    current_label = f"9月第4周 ({current_start.month}/{current_start.day}–{current_end.month}/{current_end.day})"
    previous_label = f"9月第3周 ({previous_start.month}/{previous_start.day}–{previous_end.month}/{previous_end.day})"
    revision = hashlib.sha256(args.previous.read_bytes() + args.current.read_bytes() + listing["date"].encode()).hexdigest()[:16]
    payload = {"curLabel": current_label, "prevLabel": previous_label,
               "curActive": len(cur_asins), "prevActive": len(prev_asins),
               "curAsins": cur_asins, "prevAsins": prev_asins, "records": records,
               "savedDate": date.today().isoformat(), "storeId": args.store_id, "marketplace": "US",
               "listingDate": listing["date"], "sourceFiles": [args.previous.name, args.current.name],
               "excludedPrev": len(previous) - len(prev_asins), "excludedCur": len(current) - len(cur_asins),
               "periodStartPrev": str(previous_start), "periodEndPrev": str(previous_end),
               "periodStartCur": str(current_start), "periodEndCur": str(current_end)}
    weeks = [weekly_total(previous_label, previous, sellable), weekly_total(current_label, current, sellable)]
    html = replace_var(html, "SAVED_REPORT", payload)
    html = replace_var(html, "EMBEDDED_WEEKS", weeks)
    html = replace_var(html, "WOW_REPORT_REVISION", revision)
    args.html.write_text(html, encoding="utf-8")
    print(json.dumps({"revision": revision, "listingSellable": len(sellable), "previous": weeks[0],
                      "current": weeks[1], "excludedPrev": payload["excludedPrev"],
                      "excludedCur": payload["excludedCur"], "union": len(records)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
