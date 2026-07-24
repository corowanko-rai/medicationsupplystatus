#!/usr/bin/env python3
"""Excel -> 単一HTML検索ツール 変換モジュール"""
import pandas as pd, json, re, datetime, os

TEMPLATE_HEAD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_head.html")
TEMPLATE_TAIL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_tail.html")

def clean(v):
    if pd.isna(v): return ""
    s = str(v).strip()
    return "" if s in ("　", "-", "nan") else s

def strip_prefix(v):
    s = clean(v)
    s = re.sub(r'^[①-⑳]', '', s)
    s = re.sub(r'^[0-9]+[：:.]\s*', '', s)
    s = re.sub(r'^[ア-ンA-Za-z0-9１-９]+[．.]\s*', '', s)
    return s.strip()

def serial_to_date(v):
    if pd.isna(v): return ""
    try:
        n = int(float(v))
        if 20000 < n < 60000:
            return (datetime.date(1899,12,30)+datetime.timedelta(days=n)).strftime('%Y/%m/%d')
    except (ValueError, TypeError): pass
    return clean(v)

def to_half(s):
    # U+30FC (ー) is a long-vowel LETTER, never folded to '-'.
    s = s.translate(str.maketrans(
        'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
        'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９．％　',
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        'abcdefghijklmnopqrstuvwxyz0123456789.% '))
    return re.sub(r'[－‐‑‒–—―ｰ]', '-', s)

SC = {'通常出荷':0, '限定出荷（自社の事情）':1, '限定出荷（他社品の影響）':1,
      '限定出荷（その他）':1, '供給停止':2}

def find_header_row(path, max_scan=10):
    """Locate the header row (it has moved before; don't hardcode)."""
    probe = pd.read_excel(path, header=None, nrows=max_scan)
    for i in range(len(probe)):
        joined = "".join(str(x) for x in probe.iloc[i].tolist())
        if "薬剤区分" in joined and "品名" in joined:
            return i
    return 1

def build(xlsx_path, out_path, as_of=None, source_label="", source_url=""):
    hdr = find_header_row(xlsx_path)
    df = pd.read_excel(xlsx_path, header=hdr)
    c = list(df.columns)
    if len(c) < 21:
        raise ValueError(f"列数が想定と異なります（{len(c)}列）。様式変更の可能性があります。")

    dicts = {k: {} for k in ['st','vol','rsn','out','cls','m','k','note']}
    def idx(key, val):
        d = dicts[key]
        if val not in d: d[val] = len(d)
        return d[val]

    rows = []
    for _, r in df.iterrows():
        name = clean(r[c[5]])
        if not name: continue
        st = strip_prefix(r[c[11]])
        ing, mk, sp, yj = clean(r[c[2]]), clean(r[c[6]]), clean(r[c[3]]), clean(r[c[4]])
        rows.append([
            name, ing, idx('m',mk), sp, yj,
            idx('k', strip_prefix(r[c[0]])),
            idx('cls', clean(r[c[1]]).replace('\n','')),
            idx('st', st), SC.get(st, 3),
            idx('vol', strip_prefix(r[c[16]])),
            idx('rsn', strip_prefix(r[c[13]])),
            idx('out', strip_prefix(r[c[14]])),
            idx('note', clean(r[c[15]])),
            serial_to_date(r[c[12]]),
            1 if clean(r[c[20]]) == 'New' else 0,
            to_half(f"{name} {ing} {mk} {sp} {yj}").lower(),
        ])
    if not rows:
        raise ValueError("有効なデータ行が0件です。")

    data = {
        "date": as_of or datetime.date.today().isoformat(),
        "n": len(rows),
        "src": source_label, "srcurl": source_url,
        "gen": datetime.datetime.now().strftime("%Y/%m/%d %H:%M"),
        "d": {k: [s for s,_ in sorted(v.items(), key=lambda x: x[1])] for k,v in dicts.items()},
        "r": rows,
    }
    head = open(TEMPLATE_HEAD, encoding="utf-8").read()
    tail = open(TEMPLATE_TAIL, encoding="utf-8").read()
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(head)
        json.dump(data, f, ensure_ascii=False, separators=(',',':'))
        f.write(tail)
    return len(rows)

if __name__ == "__main__":
    import sys
    n = build(sys.argv[1], sys.argv[2] if len(sys.argv)>2 else "out.html")
    print(f"{n} items")
