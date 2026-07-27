#!/usr/bin/env python3
"""Excel -> 単一HTML検索ツール 変換モジュール"""
import pandas as pd, json, re, datetime, os, unicodedata

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


def norm_name(s):
    """品名の表記揺れを吸収する。全角/半角、カギ括弧、空白を無視して比較する。"""
    s = unicodedata.normalize("NFKC", str(s))
    for a, b in (("「", ""), ("」", ""), ("（", "("), ("）", ")"), ("・", "")):
        s = s.replace(a, b)
    return re.sub(r"\s+", "", s).lower()


JP_DATE = re.compile(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日")
ISO_DATE = re.compile(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})")
# 貼り付けデータに混ざる見出し・状態表記
DISC_NOISE = {"販売中止", "予定", "販売中止予定", "供給停止", "限定出荷", "通常出荷",
              "経過措置", "告知日", "実施日", "回収", "出荷調整"}
DISC_LABEL = ("販売会社", "製造会社", "薬価", "包装薬価", "備考", "剤形", "規格")
# 包装欄の判定。「バラシクロビル」を「バラ」で誤判定しないよう、数字を伴う場合のみ
PKG_PATTERNS = [
    re.compile(r"^(PTP|ＰＴＰ)\s*\d"),
    re.compile(r"^バラ\s*\d"),
    re.compile(r"^\d"),
    re.compile(r"^(瓶|びん|袋|箱|缶|ボトル|キット|アンプル|バイアル|シリンジ|管|筒)\s*\d"),
    re.compile(r"[（(]\s*\d+\s*(mg|g|mL|ml|μg|IU|単位)\s*/"),
    re.compile(r"^\s*[×x]\s*\d"),
]


def _is_pkg_line(s):
    return any(p.search(s) for p in PKG_PATTERNS)


def _to_iso(s):
    m = JP_DATE.search(s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = ISO_DATE.search(s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return ""


def parse_discontinued_text(text):
    """DrugShortage.JP からコピーした文章をそのまま解釈する。

    想定する形（改行位置や空行は問わない）:
        告知日:
        2026年7月24日
        実施日:
        2026年10月1日 予定
        販売中止
        カルベジロール錠２０ｍｇ「ＪＧ」
        PTP100錠(20mg/錠 PTP 10錠×10)

    薬剤名だけを並べた行や、`薬剤名, 実施日, メモ` 形式も受け付ける。
    """
    recs = []
    rec = {"n": "", "d": "", "name": "", "m": ""}
    want = None

    def push(r):
        if r["name"]:
            recs.append(r)

    for raw in text.splitlines():
        ln = raw.strip()
        if not ln or ln.startswith("#"):
            continue

        # ラベル行・包装行はカンマ判定より先に落とす
        # （「包装薬価: 1,080.00円」をカンマ区切りと誤解しないため）
        if ln.startswith(DISC_LABEL) or ln in DISC_NOISE:
            continue

        # 「薬剤名, 日付, メモ」形式（カンマ区切り）はそのまま1件として扱う
        if "," in ln and not ln.startswith(("告知日", "実施日")):
            parts = [x.strip() for x in ln.split(",")]
            if parts[0] and not _is_pkg_line(parts[0]) and parts[0] not in DISC_NOISE:
                push(rec)
                rec = {"n": "", "d": "", "name": "", "m": ""}
                d1 = _to_iso(parts[1]) if len(parts) > 1 else ""
                recs.append({
                    "name": parts[0],
                    "n": "",
                    "d": d1,
                    "m": (parts[2] if len(parts) > 2 else
                          (parts[1] if len(parts) > 1 and not d1 else "")),
                })
                want = None
                continue

        if ln.startswith("告知日"):
            push(rec)
            rec = {"n": "", "d": "", "name": "", "m": ""}
            want = None
            if JP_DATE.search(ln) or ISO_DATE.search(ln):
                rec["n"] = _to_iso(ln)
            else:
                want = "n"
            continue

        if ln.startswith("実施日"):
            want = None
            if JP_DATE.search(ln) or ISO_DATE.search(ln):
                rec["d"] = _to_iso(ln)
            else:
                want = "d"
            continue

        if JP_DATE.search(ln) or ISO_DATE.search(ln):
            if want:
                rec[want] = _to_iso(ln)
                want = None
            continue

        if ln in DISC_NOISE or ln.startswith(DISC_LABEL) or _is_pkg_line(ln):
            continue

        if not rec["name"]:
            rec["name"] = ln
        else:
            # 名前が埋まっている状態で別の名前行が来た＝次の品目
            # （包装や会社名は上で除外済みなので、ここに来るのは薬剤名）
            push(rec)
            rec = {"n": "", "d": "", "name": ln, "m": ""}

    push(rec)
    return recs


def norm_name(s):
    """品名の表記揺れを吸収する。全角/半角、カギ括弧、空白を無視して比較する。"""
    s = unicodedata.normalize("NFKC", str(s))
    for a, b in (("「", ""), ("」", ""), ("（", "("), ("）", ")"), ("・", "")):
        s = s.replace(a, b)
    return re.sub(r"\s+", "", s).lower()


def load_discontinued(path, name_index=None):
    """販売中止の登録を読む。DSJPからの貼り付けをそのまま解釈できる。
    同名の品目が複数ある場合は取り違えを防ぐため登録せず警告を出す。"""
    if not path or not os.path.exists(path):
        return {}
    try:
        text = open(path, encoding="utf-8").read()
    except Exception as e:
        print(f"  警告: 販売中止ファイルを読めませんでした: {e}")
        return {}

    out = {}
    ambiguous, notfound = [], []
    for rec in parse_discontinued_text(text):
        key = rec["name"]
        info = {"n": rec["n"], "d": rec["d"], "m": rec.get("m", "")}
        # YJコード直接指定
        if re.fullmatch(r"[0-9A-Za-z]{6,12}", key) and re.search(r"\d", key):
            out[key.upper()] = info
            continue
        if not name_index:
            notfound.append(key)
            continue
        hits = name_index.get(norm_name(key))
        if not hits:
            notfound.append(key)
        elif len(hits) > 1:
            ambiguous.append((key, hits))
        else:
            out[hits[0]] = info

    for k in notfound:
        print(f"  警告: 販売中止「{k}」は該当する薬剤が見つかりません（表記をご確認ください）")
    for k, hits in ambiguous:
        print(f"  警告: 販売中止「{k}」は同名が {len(hits)} 件あります。"
              f"YJコードで指定してください → {', '.join(hits)}")
    return out


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
          kiso_path=None, disc_path=None):
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
    # 薬剤名 → YJコード の索引（販売中止を名前で登録できるようにするため）
    name_index = {}
    for _, r in df.iterrows():
        nm = clean(r[c[5]])
        yjc = clean(r[c[4]])
        if nm and yjc:
            name_index.setdefault(norm_name(nm), []).append(yjc)
    disc = load_discontinued(disc_path, name_index)
    rows = []
    snap = {}
    n_price = 0
    n_kiso = 0
    n_kchg = 0
    n_disc = 0
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
        # 販売中止（手動登録）。供給状況とは別軸の情報として持つ
        dc = disc.get(yj)
        if dc:
            n_disc += 1
            dcv = [dc["n"], dc["d"], dc.get("m", "")]
        else:
            dcv = 0
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
            dcv,                                # [22] 販売中止 [告知日,実施日,メモ] / 0
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

    data["disc"] = {"count": n_disc, "registered": len(disc)}

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
