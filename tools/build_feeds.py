#!/usr/bin/env python3
"""Fill Amazon CELLULAR_PHONE_CASE feed templates (series 215 and 223).

Reads SEO / image / ASIN / price workbooks from the repo root and writes the
Template sheet rows (row 7+) straight into the .xlsm XML, so every dropdown,
conditional format and hidden sheet of the Amazon template stays untouched.

usage: python3 tools/build_feeds.py [--include-high]
"""
import html, re, shutil, sys, zipfile, os, collections
import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INCLUDE_HIGH = "--include-high" in sys.argv
SHEET = "xl/worksheets/sheet5.xml"        # Template
SST = "xl/sharedStrings.xml"
NCOLS = 308                               # A..KV

# ---------- fixed values requested by the account manager ----------
BRAND = "MOBILIUS"
COUNTRY = "China"
WARRANTY = "6 Months"
BATT = "No"
FULFILL = "Fulfillment by Merchant (Default)"
QTY = "1"
ITEM_TYPE_KW = ("Electronics > Cell Phones & Accessories > Cases, Holsters & Sleeves > "
                "Basic Cases (cell-phone-basic-cases)")
ITEM_W, ITEM_H, ITEM_D, UNIT = 64, 48, 25, "Millimeters"   # case: width, height, thickness (mm)
PKG_ADD = 3                                                 # package is 3 mm bigger
ITEM_WEIGHT_G, PKG_WEIGHT_G = 30, 50
FEATURES = ["Magnetic", "Wireless Charging Compatible", "Slim Fit", "Lightweight", "Scratch resistant"]

def rows(path, sheet):
    wb = openpyxl.load_workbook(os.path.join(ROOT, path), read_only=True, data_only=True)
    return list(wb[sheet].iter_rows(values_only=True))

def money(v): return f"{float(v):.2f}"
def num(v):
    v = float(v); return str(int(v)) if v == int(v) else str(v)

def colour(series, title, case_color):
    if series == "215": return "Clear"
    pre = re.match(r"MOBILIUS (.*?) Case for", title).group(1)
    return {"Clear Matte": "Clear", "Black Smoky Matte": "Black",
            "Gray Frosted Matte": "Gray", "Orange": "Orange"}[pre]

def models(model_cell):
    out = []
    for p in model_cell.split("/"):
        p = p.strip()
        if not p.startswith("Apple "): p = "Apple " + p
        out.append(p)
    return out

def build_series(series, img_file, img_sheet, seo_sheet, template, asin_map):
    img = rows(img_file, img_sheet); seo = rows(img_file, seo_sheet); pr = rows("Amazon_prices_by_series(1).xlsx", series)
    seo = {r[0]: r for r in seo[1:] if r[0]}
    price = {r[0]: r for r in pr[14:] if r[0]}          # header is row 14
    seen, recs, skipped_high, dups = set(), [], [], 0
    for r in img[1:]:
        sku = r[1]
        if not sku: continue
        if sku in seen: dups += 1; continue
        seen.add(sku)
        s = seo[sku]; p = price[sku]
        if s[10] == "HIGH" and not INCLUDE_HIGH:
            skipped_high.append((sku, s[1], s[11])); continue
        urls = [u for u in r[3:11 if series == "215" else 11] if u]
        d = {"A": sku, "B": "CELLULAR_PHONE_CASE"}
        asin = asin_map.get(sku)
        d["C"] = "Edit (Partial Update)" if asin else "Create or Replace (Full Update)"
        d["G"], d["H"] = s[1], s[2]
        d["I"] = BRAND
        d["J"] = "ASIN" if asin else "GTIN Exempt"
        if asin: d["K"] = asin
        d["L"] = ITEM_TYPE_KW
        d["P"] = sku                                    # Model Number
        d["R"] = BRAND                                  # Manufacturer
        for i, u in enumerate(urls):                    # U = main, V.. = other images
            d[L(CI("U") + i)] = u
        d["AE"] = s[8]
        for i in range(5): d[L(CI("AF") + i)] = s[3 + i]
        d["AK"] = s[9]
        for i, f in enumerate(FEATURES): d[L(CI("AL") + i)] = f
        d["AR"] = "Thermoplastic Polyurethane"
        d["AW"] = "1"
        d["BB"] = colour(series, s[1], r[14] if series == "223" else None)
        d["BD"] = sku                                   # Part Number
        d["BE"] = "Thermoplastic Polyurethane"
        ms = models(r[2])
        ms = [m for m in ms if m in template["DB"]]          # drop models missing from the dropdown (e.g. iPhone 17e)
        for i, m in enumerate(ms[:5]): d[L(CI("DB") + i)] = m
        bm = [m[6:] for m in ms if m[6:] in template["BM"]]
        for i, m in enumerate(bm[:3]): d[L(CI("BM") + i)] = m
        d["BP"] = "Basic Case"
        d["BU"] = "iPhone"
        d["BZ"] = "Smooth" if series == "215" else "Matte"
        d["CR"], d["CT"] = num(ITEM_D), UNIT             # thickness
        d["DI"] = "MagSafe"
        d["DP"], d["DQ"] = num(ITEM_WEIGHT_G), "Grams"
        d["DS"] = "New"
        d["DU"] = money(p[1])
        d["ES"], d["ET"] = FULFILL, QTY
        d["EX"], d["EZ"], d["FA"] = money(p[2]), money(p[3]), money(p[4])
        d["FB"], d["FC"], d["FD"] = money(p[5]), p[6], p[7]
        d["FG"], d["FH"], d["FI"] = money(p[8]), money(p[9]), money(p[10])
        d["FL"], d["FM"], d["FN"], d["FO"], d["FP"] = p[11], num(p[12]), money(p[13]), num(p[14]), money(p[15])
        d["FX"], d["FY"] = num(ITEM_D), UNIT             # length  (= depth/thickness)
        d["FZ"], d["GA"] = num(ITEM_W), UNIT             # width
        d["GB"], d["GC"] = num(ITEM_H), UNIT             # height
        d["GD"], d["GE"] = num(ITEM_D + PKG_ADD), UNIT
        d["GF"], d["GG"] = num(ITEM_W + PKG_ADD), UNIT
        d["GH"], d["GI"] = num(ITEM_H + PKG_ADD), UNIT
        d["GJ"], d["GK"] = num(PKG_WEIGHT_G), "Grams"
        d["GO"], d["GP"] = COUNTRY, WARRANTY
        d["GS"], d["GT"] = BATT, BATT
        recs.append((d, asin, s[10]))
    return recs, skipped_high, dups

def load_template_lists(xlsm):
    """dropdown values we validate against (static lists only)"""
    import pickle
    wb = openpyxl.load_workbook(xlsm, read_only=True)
    dl = list(wb["Dropdown Lists"].iter_rows(min_row=1, max_row=14170, values_only=True))
    def col(c, a, b): return {dl[r - 1][CI(c) - 1] for r in range(a, b + 1) if dl[r - 1][CI(c) - 1] is not None}
    return {"BM": col("AD", 4, 1363), "DB": col("EP", 4, 14170)}

def write_feed(series, tpl_name, recs):
    src = os.path.join(ROOT, tpl_name)
    tmp = src + ".tmp"
    zin = zipfile.ZipFile(src)
    sheet = zin.read(SHEET).decode("utf8"); sst = zin.read(SST).decode("utf8")
    head = re.search(r"<sst [^>]*>", sst).group(0)
    cnt = int(re.search(r'count="(\d+)"', head).group(1)); uniq = int(re.search(r'uniqueCount="(\d+)"', head).group(1))
    # default style per column from <cols>
    style = {}
    for a, b, s in re.findall(r'<col min="(\d+)" max="(\d+)"[^>]*?style="(\d+)"', sheet):
        for c in range(int(a), int(b) + 1): style[L(c)] = s
    new, idx, used = [], {}, 0
    def sid(v):
        nonlocal uniq
        if v not in idx:
            idx[v] = uniq; uniq += 1
            new.append('<si><t xml:space="preserve">%s</t></si>' % html.escape(v, quote=False))
        return idx[v]
    order = [L(i) for i in range(1, NCOLS + 1)]
    out = []
    for n, (d, _, _) in enumerate(recs, start=7):
        cells = []
        for c in order:
            if c in d and d[c] not in (None, ""):
                cells.append('<c r="%s%d" s="%s" t="s"><v>%d</v></c>' % (c, n, style.get(c, "53"), sid(str(d[c])))); used += 1
        out.append('<row r="%d" spans="1:%d">%s</row>' % (n, NCOLS, "".join(cells)))
    last = 6 + len(recs)
    assert "</sheetData>" in sheet and not re.search(r'<row r="7"', sheet), "template already filled"
    sheet = sheet.replace("</sheetData>", "".join(out) + "</sheetData>", 1)
    sheet = sheet.replace('<dimension ref="A1:KV6"/>', f'<dimension ref="A1:KV{last}"/>', 1)
    sst = sst.replace(head, re.sub(r'count="\d+"', f'count="{cnt + used}"', re.sub(r'uniqueCount="\d+"', f'uniqueCount="{uniq}"', head)), 1)
    sst = sst.replace("</sst>", "".join(new) + "</sst>")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == SHEET: data = sheet.encode("utf8")
            elif it.filename == SST: data = sst.encode("utf8")
            zo.writestr(it, data, compress_type=zipfile.ZIP_DEFLATED)
    zin.close(); shutil.move(tmp, src)

def main():
    asin_rows = rows("215 Amazon_ASIN_SKU.xlsx", "ASIN SKU")
    asin_map = {r[1]: r[0] for r in asin_rows[1:] if r[0] and r[1]}
    tpl = load_template_lists(os.path.join(ROOT, "223 CELLULAR_PHONE_CASE.xlsm"))
    cfg = {"215": ("215 CELLULAR_PHONE_CASE.xlsm", "215_Amazon_29.09_SEO.xlsx", "Sheet", "SEO 215", asin_map),
           "223": ("223 CELLULAR_PHONE_CASE.xlsm", "223_Amazon_with_prints_SEO.xlsx", "Sheet", "SEO 223", {})}
    report = []
    for ser, (tname, f, s1, s2, am) in cfg.items():
        recs, high, dups = build_series(ser, f, s1, s2, tpl, am)
        bad = [(d["A"], c, d[c]) for d, _, _ in recs for c in ("BM", "BN", "BO", "DB", "DC", "DD", "DE", "DF")
               if c in d and d[c] not in tpl["BM" if c[0] == "B" else "DB"]]
        assert not bad, bad[:5]
        write_feed(ser, tname, recs)
        n_asin = sum(1 for _, a, _ in recs if a)
        report.append(f"{ser}: rows={len(recs)} edit(ASIN)={n_asin} create={len(recs)-n_asin} "
                      f"skipped_HIGH={len(high)} dup_rows_dropped={dups}")
        if high:
            with open(os.path.join(ROOT, f"{ser}_excluded_HIGH_policy_risk.csv"), "w", encoding="utf-8-sig") as fh:
                fh.write("SKU;Title;Policy type\n")
                for sku, t, pt in high: fh.write(f"{sku};{t};{pt}\n")
    print("\n".join(report))

if __name__ == "__main__":
    main()
