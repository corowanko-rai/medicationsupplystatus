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
    """全角英数・記号（U+FF01〜U+FF5E）をASCIIへ畳む。
    テンプレート側のJavaScript HALF() と必ず同じ規則にすること。
    U+30FC (ー) は長音符という「文字」なのでハイフンに畳まない。"""
    s = "".join(
        chr(ord(ch) - 0xFEE0) if 0xFF01 <= ord(ch) <= 0xFF5E else ch
        for ch in s
    )
    s = s.replace("\u3000", " ")                    # 全角スペース
    return re.sub(r"[\u2010-\u2015\uFF70]", "-", s)  # ダッシュ類と半角カナ長音

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

def load_prices(path):
    """prices.json を読む。無い/壊れていても薬価なしで続行する。"""
    if not path or not os.path.exists(path):
        return None
    try:
        d = json.load(open(path, encoding="utf-8"))
        if d.get("exact"):
            return d
    except Exception:
        pass
    return None


def load_kiso(path):
    """kiso.json（変更調剤が認められる基礎的医薬品の品名集合）を読む。"""
    if not path or not os.path.exists(path):
        return None
    try:
        d = json.load(open(path, encoding="utf-8"))
        ns = d.get("names")
        if ns:
            return set(ns)
    except Exception:
        pass
    return None


def lookup_price(pr, yj):
    """YJコードから薬価を引く。
    ① 12桁完全一致（銘柄別収載品）
    ② 先頭9桁で統一名収載品に紐付け（後発品の多くはこちら）"""
    if not pr or not yj:
        return None
    v = pr["exact"].get(yj)
    if v is not None:
        return v
    return pr.get("uni", {}).get(yj[:9])


def build(xlsx_path, out_path, as_of=None, source_label="", source_url="",
          prev_snapshot=None, snapshot_out=None, prices_path=None,
          kiso_path=None):
    """prev_snapshot: {YJコード: sc} from the previous edition, for 悪化/改善 detection.
    snapshot_out: path to write this edition's snapshot for the next run."""
    hdr = find_header_row(xlsx_path)
    df = pd.read_excel(xlsx_path, header=hdr)
    c = list(df.columns)
    if len(c) < 21:
        raise ValueError(f"列数が想定と異なります（{len(c)}列）。様式変更の可能性があります。")

    dicts = {k: {} for k in ['st','vol','rsn','out','cls','m','k','note','pc','i']}
    def idx(key, val):
        d = dicts[key]
        if val not in d: d[val] = len(d)
        return d[val]

    pr = load_prices(prices_path)
    kiso = load_kiso(kiso_path)
    rows = []
    snap = {}
    n_price = 0
    n_kiso = 0
    n_kchg = 0
    for _, r in df.iterrows():
        name = clean(r[c[5]])
        if not name: continue
        st = strip_prefix(r[c[11]])
        ing, mk, sp, yj = clean(r[c[2]]), clean(r[c[6]]), clean(r[c[3]]), clean(r[c[4]])
        sc = SC.get(st, 3)
        # 悪化/改善: compare severity against the previous edition (0<1<2)
        chg = 0
        if prev_snapshot is not None and yj in prev_snapshot:
            old_sc = prev_snapshot[yj]
            if sc in (0,1,2) and old_sc in (0,1,2):
                if   sc > old_sc: chg = 1   # 悪化
                elif sc < old_sc: chg = 2   # 改善
        elif prev_snapshot:
            chg = 3                          # 新規掲載
        snap[yj] = sc
        price = lookup_price(pr, yj)
        if price is not None:
            n_price += 1
        # ⑨基礎的医薬品（1：対象）
        kb = 1 if clean(r[c[8]]).startswith("1") else 0
        n_kiso += kb
        # 変更調剤が認められる基礎的医薬品（品名で突合）
        kc = 1 if (kiso is not None and name in kiso) else 0
        n_kchg += kc
        rows.append([
            name, idx('i', ing), idx('m',mk), sp, yj,
            idx('k', strip_prefix(r[c[0]])),
            idx('cls', clean(r[c[1]]).replace('\n','')),
            idx('st', st), sc,
            idx('vol', strip_prefix(r[c[16]])),
            idx('rsn', strip_prefix(r[c[13]])),
            idx('out', strip_prefix(r[c[14]])),
            idx('note', clean(r[c[15]])),
            serial_to_date(r[c[12]]),
            1 if clean(r[c[20]]) == 'New' else 0,
            to_half(name).lower(),   # [15] 品名の検索用
            to_half(ing).lower(),    # [16] 成分名の検索用
            idx('pc', strip_prefix(r[c[7]])),  # [17] ⑧製品区分
            chg,                                # [18] 0=変化なし 1=悪化 2=改善 3=新規
            price,                              # [19] 薬価（円）／不明はnull
            kb,                                 # [20] 1=⑨基礎的医薬品
            kc,                                 # [21] 1=変更調剤が認められる基礎的医薬品
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
    if snapshot_out:
        with open(snapshot_out, "w", encoding="utf-8") as f:
            json.dump({"date": data["date"], "sc": snap}, f, separators=(',',':'))

    data["price"] = {
        "available": pr is not None,
        "as_of": (pr or {}).get("as_of", ""),
        "matched": n_price,
    }

    data["kiso"] = {
        "basic": n_kiso,
        "swap_available": kiso is not None,
        "swap": n_kchg,
    }

    data["chg"] = {
        "worse":    sum(1 for r in rows if r[18] == 1),
        "better":   sum(1 for r in rows if r[18] == 2),
        "new":      sum(1 for r in rows if r[18] == 3),
        "compared": prev_snapshot is not None and len(prev_snapshot) > 0,
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
