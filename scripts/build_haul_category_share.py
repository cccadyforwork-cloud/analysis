#!/usr/bin/env python3
"""Classify sellable ASINs in a Listing export and write an auditable mapping.

Usage: python3 scripts/build_haul_category_share.py Listing.xlsx
The taxonomy follows Haul商品一级二级分类文档.md. Rules use product titles
before seller names because some seller names in the export are inaccurate.
"""

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl


TAXONOMY = {
    "女装": "服装 上衣 连衣裙 下装 内衣与睡衣 泳装 鞋履 包袋与钱包 配饰".split(),
    "家居": "家居装饰 收纳整理 墙面艺术与标牌 床品与浴室用品 派对与季节装饰 清洁与地板护理".split(),
    "男装": "服装 上衣 下装 内衣与袜子 手表 包袋与收纳包 配饰 鞋履".split(),
    "厨房与餐饮": ["厨房用具与小工具", "餐饮与宴客用品", "炊具与烘焙用品", "收纳整理", "厨房与餐桌布艺", "咖啡、茶与小家电"],
    "珠宝首饰": "女士首饰 男士首饰 耳环 项链 手链 戒指 身体饰品".split(),
    "美妆": "彩妆 护肤 美发与造型 临时纹身与身体彩绘".split(),
    "家装工具": "五金与安全 照明 管道与卫浴 电工电料 墙面涂装与建筑用品".split(),
    "电子产品": "充电器与数据线 电脑配件 表带 音频".split(),
    "玩具与游戏": "动作人偶与收藏玩具 游戏与拼图 美术手工 解压与感官玩具 积木套装与玩具车 服装扮演 娃娃与玩具套装".split(),
    "办公文具": "卡片与文具 标签与徽章 桌面用品 笔记本与计划本 书写与绘画".split(),
    "运动户外": "健身与训练 团队运动 水上运动 冬季运动 狩猎与钓鱼".split(),
    "庭院园艺": "种子与植物 花园雕像与户外装饰 花盆与植物养护 园艺工具与浇灌 花园标牌与庭院艺术 露台烧烤与户外生活".split(),
    "宠物用品": "狗狗用品 猫咪用品 小宠与鱼类用品".split(),
}


def has(text, pattern):
    return bool(re.search(pattern, text, re.I))


def category(title, name, msku):
    t = str(title or "").lower()
    lead = t[:145]
    n = str(name or "")
    s = str(msku or "").lower()

    def hit(pattern):
        return has(lead, pattern)

    # The supplied taxonomy has no vehicle category. Use the nearest product
    # function so these sellable items still contribute to the category share.
    if hit(r"car detail brush|car interior soft brush"):
        return "家居", "清洁与地板护理"
    if hit(r"steering wheel charm"):
        return "家居", "家居装饰"
    if hit(r"shift knob|gear shifter|motorcycle handlebar|valve stem caps"):
        return "家装工具", "五金与安全"
    if hit(r"window door track cleaning|window groove cleaning"):
        return "家居", "清洁与地板护理"
    if hit(r"jewelry storage box|jewelry organizer|jewelry tray"):
        return "家居", "收纳整理"
    if hit(r"toothpaste squeezer|toothbrush cover|toothbrush holder|cotton swab holder|soap dish"):
        return "家居", "床品与浴室用品"
    if hit(r"makeup bag|cosmetic bag|toiletry pouch"):
        return "女装", "包袋与钱包"
    if hit(r"shoe laundry|laundry bag|broom hair cover|damp duster|cleaning sponge|shoe care eraser"):
        return "家居", "清洁与地板护理"
    if hit(r"precision tweezers|finger cots|sanding polishing sponge|multi.angle.*ruler|tape measure"):
        return "家装工具", "五金与安全"
    if hit(r"voltage tester|nail holding safety plier|ratchet screwdriver|utility magnets?"):
        return "家装工具", "电工电料" if hit(r"voltage tester") else "五金与安全"
    if hit(r"phone stand|cell phone strap|phone chain|phone charm|phone grip|phone mount|smartphone holder"):
        return "电子产品", "电脑配件"
    if hit(r"carabiner clip|locking carabiner"):
        return "运动户外", "狩猎与钓鱼"
    if hit(r"can covers?|soda can lids?"):
        return "厨房与餐饮", "餐饮与宴客用品"
    if hit(r"can opener|jar opener|egg cracker"):
        return "厨房与餐饮", "厨房用具与小工具"
    if hit(r"fitness training log book|workout planner"):
        return "办公文具", "笔记本与计划本"
    if hit(r"powder sugar shaker"):
        return "厨房与餐饮", "咖啡、茶与小家电"
    if hit(r"glow stones.*garden"):
        return "庭院园艺", "花园雕像与户外装饰"
    if hit(r"scrapbooking|craft beads|painting sponges for kids"):
        return "玩具与游戏", "美术手工"

    # Title-first exceptions where the seller name is ambiguous or demonstrably wrong.
    if hit(r"dog rope|dog toy|dog ball|dog collar|dog leash|dog harness|dog poop|dog waste|dog vest|puppy.*(toy|teething|vest|collar)|pet neck collar|pet feeding|pet grooming|pet bath|pet hair brush|pet waste|pet toy.*dog|rope tug toy|rope chew toy|squeaky ball.*dog|toy.*for .*dogs|pet blanket"):
        return "宠物用品", "狗狗用品"
    if hit(r"catnip|cat toy|cat collar|cat feather|cat wand|cat chew|cat scratch|cat mouse|cat wobble|cat butterfly|cat parrot|cat string|pet toy.*cat"):
        return "宠物用品", "猫咪用品"
    if hit(r"aquarium|fish tank|bird swing|parakeet|bird cage"):
        return "宠物用品", "小宠与鱼类用品"
    if hit(r"water shoes|aqua shoes|aqua socks|beach shoes|swim shoes|surf shoes|waterproof dry bag"):
        return "运动户外", "水上运动"
    if hit(r"bikini|women.*swimsuit|women.*swimwear"):
        return "女装", "泳装"
    if hit(r"\bmen'?s swim trunk|board shorts"):
        return "男装", "下装"
    if hit(r"\bmen'?s.*(short tights|compression shorts|athletic shorts)"):
        return "男装", "下装"
    if hit(r"women.?s.*(panties|underwear|thong)|cotton underwear|silicone foot mask"):
        return "女装", "内衣与睡衣"
    if hit(r"\bmen'?s.*socks|\bmen'?s.*boat socks"):
        return "男装", "内衣与袜子"
    if hit(r"yoga grip socks|women.?s.*socks|unisex.*socks|no show liner socks"):
        return "女装", "内衣与睡衣"
    if hit(r"flip flops|slippers|sandals"):
        return "女装", "鞋履"
    if hit(r"\bmen'?s.*wallet|\bmen wallet|pop-up wallet"):
        return "男装", "包袋与收纳包"
    if hit(r"coin purse|ladies purse|wallet|cosmetic bag|makeup pouch|toiletry bag|tote bag|grocery bag|shopping bag|beach tote"):
        return "女装", "包袋与钱包"
    if hit(r"scarf|scrunchie|hair clip|hair tie|shawl|clothing accessor|eyeglass|sunglass|shoe lace|shoelace|brooch|enamel pin|waist buckle|waist cinch|beanie|zipper locks?"):
        return "女装", "配饰"
    if hit(r"carabiner.*(camp|climb|outdoor)|heavy duty locking carabiner"):
        return "运动户外", "狩猎与钓鱼"
    if hit(r"blank resin keychain|diy engrav.*keychain|craft keychain"):
        return "玩具与游戏", "美术手工"
    if hit(r"keychain|key chain|key ring charm"):
        return "女装", "配饰"
    if hit(r"arm sleeves|cycling mask|sports face cover|sports glasses lanyard|ankle support|wrist wrap|wrist rest|compression sleeve|kinesiology tape|yoga knee pad|resistance band|exercise band|fitness|gym workout|weightlifting"):
        return "运动户外", "健身与训练"
    if hit(r"tennis|badminton|pickleball|racket grip"):
        return "运动户外", "团队运动"
    if hit(r"carabiner|tent rope|rope tightener|camping|emergency blanket|survival blanket"):
        return "运动户外", "狩猎与钓鱼"

    if hit(r"jewelry making|jewellery making|spacer beads|loose beads|acrylic beads|natural stone beads|waxed thread"):
        return "玩具与游戏", "美术手工"
    if hit(r"anklet|foot chain"):
        return "珠宝首饰", "身体饰品"
    if hit(r"bracelet|crystal heart chain") and not hit(r"jewelry storage|jewelry organizer|bracelet making"):
        return "珠宝首饰", "手链"
    if hit(r"earrings?|ear studs?"):
        return "珠宝首饰", "耳环"
    if hit(r"necklace|pendant chain"):
        return "珠宝首饰", "项链"
    if hit(r"ring sizer|finger ring|fashion ring|statement ring|adjustable ring|jewelry ring"):
        return "珠宝首饰", "戒指"

    if hit(r"makeup|foundation|cosmetic brush|powder puff|beauty sponge|false eyelash|eyelashes|eyeliner|eyebrow|nail file|nail clipper|toenail clipper|false nail|press on nail|nail tips|toe separator|manicure|gua sha|facial massage|face applicator|dermaplane"):
        return "美妆", "彩妆"
    if hit(r"hair curler|hair comb|hair brush|hair styling|hair bun|hair shaper|hair roller|hair claw"):
        return "美妆", "美发与造型"
    if hit(r"foot scrub|foot care|foot file|callus remover|heel pad|heel cushion|moisturizing heel|bath sponge|bath loofah|body scrub|back scrubber|exfoliating net|back scratcher|earwax|ear pick|dental floss|floss pick"):
        return "美妆", "护肤"

    if hit(r"plant protection|plant blanket|plant clip|plant tie|plant hook|repotting mat|potting mat|transplanting pad|plant propagation|hydroponic|flower arranging grid|flower grid"):
        return "庭院园艺", "花盆与植物养护"
    if hit(r"garden tool|weeder|weed puller|garden sharpener|garden leaf bag|yard waste bag|hose attachment|spray gun|watering|drip irrigation|garden sprayer"):
        return "庭院园艺", "园艺工具与浇灌"
    if hit(r"garden flag|yard banner|plant stake|flower pot pick|garden decoration|garden decor|glow stones|outdoor garden hanging|wind chime"):
        return "庭院园艺", "花园雕像与户外装饰"
    if hit(r"artificial peony|faux flower|fake flower|seed packet"):
        return "庭院园艺", "种子与植物"

    if hit(r"recipe card|stationery|envelope|bookmark|greeting card|business card holder sleeve|card protector"):
        return "办公文具", "卡片与文具"
    if hit(r"label roll|write-on label|adhesive label|enamel pin|badge"):
        return "办公文具", "标签与徽章"
    if hit(r"pencil sharpener|pencil eraser|eraser set for writing|stampers?|self inking stamp|acrylic paint brush|watercolor brush|detailing paint brush|art paint brush|paint sponge|ruler.*(desk|office)|writing correction"):
        return "办公文具", "书写与绘画"
    if hit(r"planner|notebook|log book|coloring book|calendar"):
        return "办公文具", "笔记本与计划本"
    if hit(r"desktop organizer|desk tool storage|cable organizer|cord organizer|cable strap|desk holder"):
        return "办公文具", "桌面用品"

    if hit(r"playing cards|poker deck|magic ruler|scratch art kit|puzzle|game card|board game"):
        return "玩具与游戏", "游戏与拼图"
    if hit(r"figurine toy|animal figurines toys|action figure|collectible toy|toy car"):
        return "玩具与游戏", "动作人偶与收藏玩具"
    if hit(r"fidget|stress relief toy|sensory toy"):
        return "玩具与游戏", "解压与感官玩具"
    if hit(r"craft beads|craft supplies|diy beads|scrapbook|crochet|sewing|bookbinding|wrapping paper cutter|metal cutting die|resin keychain|blank keychain|keychain kit|drawing stencils?|rhinestone tape|felt flowers?"):
        return "玩具与游戏", "美术手工"

    if hit(r"solder|wire connector|wire stripper|voltage tester|circuit tester|electrician|electrical terminal|heat shrink"):
        return "家装工具", "电工电料"
    if hit(r"caulk tape|door seal|door sweep|sink strainer|drain hair catcher|drain cover|toilet seal|bathroom base seal"):
        return "家装工具", "管道与卫浴"
    if hit(r"paint brush.*(wood|wall|stain)|paint scraper|razor blade scraper|razor scraper|window screen repair|wall patch|adhesive tape|nano tape"):
        return "家装工具", "墙面涂装与建筑用品"
    if hit(r"ratchet|screwdriver|drill bit|driver bit|bit holder|pliers|magnetic pick up|work gloves|cut resistant gloves|bracket|door latch|nail holding|hole punch plier|casters?|magnet stick|tool kit|tape measure|angle ruler|marking ruler|utility magnet|wall hook"):
        return "家装工具", "五金与安全"
    if hit(r"led light|flashlight|lamp|light strip"):
        return "家装工具", "照明"

    if hit(r"charger|charging cable|usb cable|data cable"):
        return "电子产品", "充电器与数据线"
    if hit(r"computer|tablet stand|phone stand|cell phone holder|smartphone holder|phone grip|phone mount|phone lanyard|phone charm|phone chain|cell phone strap"):
        return "电子产品", "电脑配件"
    if hit(r"watch band|smart watch strap"):
        return "电子产品", "表带"
    if hit(r"headphone|earphone|speaker|audio"):
        return "电子产品", "音频"

    if hit(r"coffee scoop|coffee spoon|tea spoon|pepper grinder|spice mill|coffee grinder|coffee.*clip"):
        return "厨房与餐饮", "咖啡、茶与小家电"
    if hit(r"tumbler|shaker bottle|drinking straw|water bottle|drinking cup|coffee cup|tea cup|\bmug\b|taco holder|bottle opener|beer opener|coaster|placemat|table mat|spoon.*dessert|serving stand|soda can lids?"):
        return "厨房与餐饮", "餐饮与宴客用品"
    if hit(r"ice cube|egg ring|egg poacher|baking paper|parchment paper|baking mat|griddle|burger press|steamer rack|whisk|egg beater|dough mixer|potato masher|spatula|\boven\b|cookware|bakeware"):
        return "厨房与餐饮", "炊具与烘焙用品"
    if hit(r"spice jar|salt and pepper shaker|condiment bottle|lunch box|food storage|freezer container|refrigerator organizer|fridge organizer|can cover|bowl cover|sauce bottle|soap holder.*kitchen"):
        return "厨房与餐饮", "收纳整理"
    if hit(r"kitchen towel|dish cloth|dish towel|washcloth|kitchen linen"):
        return "厨房与餐饮", "厨房与餐桌布艺"
    if hit(r"garlic press|garlic peeler|can opener|jar opener|egg cracker|egg timer|fruit slicer|fruit baller|watermelon scoop|vegetable peeler|grater|zester|herb stripper|food chopper|kitchen chopper|food tongs|kitchen tongs|strainer|filter bag|food filter|kitchen gadget|kitchen utensil|knife sharpener|oil separator|fat skimmer|kitchen scraper"):
        return "厨房与餐饮", "厨房用具与小工具"

    if hit(r"halloween|ghost figurine|pumpkin decor|party favor|party decoration|seasonal event|candy bag"):
        return "家居", "派对与季节装饰"
    if hit(r"wall panel|wall art|wall decal|wall sticker|wall sign"):
        return "家居", "墙面艺术与标牌"
    if hit(r"pillow cover|cushion case|throw blanket|bath towel|hand towel|bathroom|toilet seat cover|shower mat|soap dish|toothbrush holder|shower sponge"):
        return "家居", "床品与浴室用品"
    if hit(r"cleaning brush|car detail brush|duster sponge|rust remover|lint roller|lint remover|floor protector|rug corner gripper|chair leg cover|mop|broom|shoe wash bag|laundry wash bag|shoe cleaning|cleaning sponge|cleaning eraser|scrub eraser|spray bottle.*sanitizer"):
        return "家居", "清洁与地板护理"
    if hit(r"decorative refrigerator magnet holder|decorative.*vase|figurine|ornament|decor|vase|fridge magnet|refrigerator magnet|suncatcher|mirror wind sequin|window prism|hanging ornament|flower ribbon|bouquet wrap|fishtail ribbon"):
        return "家居", "家居装饰"
    if hit(r"storage box|organizer|holder|shelf|storage rack|hanging storage|storage caddy|hat hook|hat hanger|pill box|pill organizer|pill cutter|medicine case|jewelry storage|eyeglass holder"):
        return "家居", "收纳整理"

    # Use a few specific seller names only when the English title is missing.
    if not t:
        if has(n, r"溯溪|涉水鞋") or has(s, r"^qs-"):
            return "运动户外", "水上运动"
        if has(n, r"拖鞋"):
            return "女装", "鞋履"
    return "待确认", "待确认"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("listing", type=Path)
    parser.add_argument("--output", type=Path, default=Path("reports/1店_可售ASIN品类_2026-10-05.csv"))
    parser.add_argument("--json", type=Path, default=Path("reports/1店_可售ASIN品类_2026-10-05.json"))
    args = parser.parse_args()
    sheet = openpyxl.load_workbook(args.listing, read_only=True, data_only=True).active
    values = sheet.values
    header = {name: i for i, name in enumerate(next(values))}
    rows = []
    for raw in values:
        asin = str(raw[header["ASIN"]] or "").strip().upper()
        stock = int(float(raw[header["FBA可售"]] or 0))
        if not asin or stock <= 0:
            continue
        item = {
            "ASIN": asin,
            "父ASIN": str(raw[header["父ASIN"]] or "").strip().upper(),
            "品名": str(raw[header["品名"]] or "").strip(),
            "标题": str(raw[header["标题"]] or "").strip(),
            "MSKU": str(raw[header["MSKU"]] or "").strip(),
            "FBA可售": stock,
        }
        item["一级分类"], item["二级分类"] = category(item["标题"], item["品名"], item["MSKU"])
        if item["一级分类"] != "待确认":
            assert item["二级分类"] in TAXONOMY[item["一级分类"]], item
        rows.append(item)
    if len({r["ASIN"] for r in rows}) != len(rows):
        raise ValueError("The Listing contains duplicate sellable ASINs")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    args.json.write_text(json.dumps({r["ASIN"]: [r["一级分类"], r["二级分类"]] for r in rows}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    counts = Counter(r["一级分类"] for r in rows)
    print(f"可售 ASIN {len(rows)}；一级分类 {dict(counts)}")
    for row in rows:
        if row["一级分类"] == "待确认":
            print("待确认", row["ASIN"], row["品名"], row["标题"][:95])


if __name__ == "__main__":
    main()
