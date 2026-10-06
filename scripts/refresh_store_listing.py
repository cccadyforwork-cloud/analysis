"""Embed one store's Listing and reconcile it with the embedded Business Report."""

import argparse
import csv
import json
import math
import re
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_HTML = ROOT / "货盘动销分析_WK29-WK30_可下载版.html"


def embedded(html, name):
    match = re.search(r"^var " + name + r"\s*=\s*(\{.*\});$", html, re.M)
    if not match:
        raise ValueError(f"Missing embedded {name}")
    return json.loads(match.group(1))


def listing_rows(workbook):
    sheet = load_workbook(workbook, read_only=True, data_only=True).active
    values = iter(sheet.values)
    headers = [str(value or "").strip() for value in next(values)]
    required = {"ASIN", "MSKU", "父ASIN", "品名", "FBA可售", "负责人1（业绩归属人）"}
    if not required.issubset(headers):
        raise ValueError(f"Missing Listing columns: {', '.join(sorted(required - set(headers)))}")
    rows = {}
    for number, values_row in enumerate(values, 2):
        row = dict(zip(headers, values_row))
        asin = str(row["ASIN"] or "").strip().upper()
        if not asin:
            continue
        if not re.fullmatch(r"[A-Z0-9]{10}", asin):
            raise ValueError(f"Invalid ASIN on row {number}: {asin}")
        if asin in rows:
            raise ValueError(f"Duplicate ASIN on row {number}: {asin}")
        try:
            stock = float(row["FBA可售"] or 0)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid FBA可售 on row {number}: {asin}") from exc
        if not math.isfinite(stock) or stock < 0:
            raise ValueError(f"Invalid FBA可售 on row {number}: {asin}")
        try:
            sales7 = float(row.get("7日销量") or 0)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid 7日销量 on row {number}: {asin}") from exc
        if not math.isfinite(sales7) or sales7 < 0:
            raise ValueError(f"Invalid 7日销量 on row {number}: {asin}")
        rows[asin] = {
            "i": int(stock) if stock.is_integer() else stock,
            "d7": int(sales7) if sales7.is_integer() else sales7,
            "o": str(row["负责人1（业绩归属人）"] or "").strip(),
            "n": str(row["品名"] or "").strip(),
            "m": str(row["MSKU"] or "").strip(),
            "p": str(row["父ASIN"] or "").strip().upper(),
        }
    return rows


def write_csv(path, header, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--html", type=Path, default=DEFAULT_HTML)
    parser.add_argument("--store-id", default="1店")
    parser.add_argument("--marketplace", default="")
    parser.add_argument("--date", help="Listing snapshot date YYYY-MM-DD; inferred from filename by default")
    args = parser.parse_args()

    name_date = re.search(r"Listing(20\d{2})(\d{2})(\d{2})", args.workbook.name)
    date = args.date or ("-".join(name_date.groups()) if name_date else "")
    if not re.fullmatch(r"20\d{2}-\d{2}-\d{2}", date):
        raise ValueError("Specify --date YYYY-MM-DD")
    html = args.html.read_text(encoding="utf-8")
    report = embedded(html, "OPS_SEED_REPORT")
    if (report["storeId"], report.get("marketplace", "")) != (args.store_id, args.marketplace):
        raise ValueError("The embedded Business Report has a different store or marketplace")
    names = embedded(html, "PRODUCT_NAMES")
    listings = embedded(html, "OPS_STORE_LISTINGS")
    rows = listing_rows(args.workbook)
    scope = json.dumps([args.store_id, args.marketplace], ensure_ascii=False, separators=(",", ":"))
    listings[scope] = {"name": args.workbook.name, "date": date, "storeId": args.store_id,
                       "marketplace": args.marketplace, "rows": rows}
    payload = json.dumps(listings, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    pattern = r"^var OPS_STORE_LISTINGS=\{.*\};$"
    if not re.search(pattern, html, flags=re.M):
        raise ValueError("Could not replace OPS_STORE_LISTINGS")
    updated = re.sub(pattern, lambda _: "var OPS_STORE_LISTINGS=" + payload + ";", html, count=1, flags=re.M)
    if updated != html:
        args.html.write_text(updated, encoding="utf-8")

    report_rows = {row["asin"]: row for row in report["rows"]}
    period = f"{report['periodStart']} 至 {report['periodEnd']}"
    match_rows = []
    for asin in sorted(report_rows.keys() | rows.keys()):
        business = report_rows.get(asin)
        listing = rows.get(asin)
        if business and listing:
            status = "可售已匹配" if listing["i"] > 0 else "Listing库存0_已排除"
            action = "进入行动筛选" if listing["i"] > 0 else "不可售，无需处理"
        elif business:
            status, action = "仅业务报告_不可售已排除", "不可售，无需处理"
        else:
            status, action = "仅Listing_报告未出现", "不据此判断零销量"
        match_rows.append([status, asin, args.store_id, (listing["o"] or "未分配") if listing else "",
                           (listing["n"] or names.get(asin, "")) if listing else names.get(asin, business.get("title", "")),
                           listing["m"] if listing else "", listing["p"] if listing else "",
                           listing["i"] if listing else "", business["sessions"] if business else "",
                           business["gv"] if business else "", business["units"] if business else "",
                           period if business else "", date if listing else "", action])
    match_path = ROOT / "reports" / f"{args.store_id}_业务报告与可售Listing匹配_{date}.csv"
    write_csv(match_path,
              ["匹配状态", "ASIN", "店铺", "负责人", "品名", "MSKU", "父ASIN", "FBA可售", "近31天Sessions",
               "近31天GV", "近31天Units", "业务报告周期", "Listing日期", "处理建议"], match_rows)

    available = [row for row in report["rows"] if row["asin"] in rows and rows[row["asin"]]["i"] > 0]
    sold = sorted((row["units"] for row in available if row["units"] > 0), reverse=True)
    cutoff = max(5, sold[math.ceil(len(sold) * 0.2) - 1]) if sold else math.inf
    candidates = []
    for business in available:
        units, sessions = business["units"], business["sessions"]
        if units == 0:
            strategy, action = "零销量（不限曝光）", "本店Listing已匹配；按流量排查链接"
        elif sessions >= 10 and units <= 5:
            strategy, action = "有曝光低转化", "排查价格、主图、评价、文案"
        elif units >= cutoff:
            strategy, action = "头部销量候选", "核实本店双周走势及安全库存，再规划变体"
        else:
            continue
        listing = rows[business["asin"]]
        candidates.append((strategy, business, listing, action))
    order = {"头部销量候选": 0, "有曝光低转化": 1, "零销量（不限曝光）": 2}
    candidates.sort(key=lambda item: (order[item[0]], -item[1]["units"] if item[0] == "头部销量候选"
                                      else -item[1]["sessions"], item[1]["asin"]))
    candidate_path = ROOT / "reports" / f"{args.store_id}_运营候选_{report['periodStart']}_{report['periodEnd']}.csv"
    write_csv(candidate_path,
              ["优化路径", "店铺", "站点", "负责人（本店Listing）", "ASIN", "品名", "父ASIN", "近31天Sessions",
               "近31天GV", "近31天Units", "Units/Sessions", f"本店FBA可售({date})", "判断与下一步", "报告周期"],
              ([strategy, args.store_id, args.marketplace or "未指定", listing["o"] or "未分配", business["asin"],
                listing["n"] or names.get(business["asin"], business.get("title", "")), listing["p"],
                business["sessions"], business["gv"], business["units"],
                f"{business['units'] / business['sessions']:.2%}" if business["sessions"] else "",
                listing["i"], action, period] for strategy, business, listing, action in candidates))
    print(f"Listing {len(rows)}, report {len(report_rows)}, matched {len(report_rows.keys() & rows.keys())}, "
          f"eligible {len(available)}, report-only {len(report_rows.keys() - rows.keys())}, "
          f"listing-only {len(rows.keys() - report_rows.keys())}")
    print("Candidates:", dict(Counter(item[0] for item in candidates)))
    print(match_path)
    print(candidate_path)


if __name__ == "__main__":
    main()
