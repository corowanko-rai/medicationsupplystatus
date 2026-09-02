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
            # 局方コードは毎行照会するので集合にしておく
            d["_jpset"] = set(d.get("jpharm") or [])
            return d
    except Exception:
        pass
    return None


# ---- 細かい剤形の判定 --------------------------------------------
# 供給状況Excelには「内用薬/外用薬/注射薬」の3分類しかないため、
# 品名のキーワードを優先し、決まらない分だけYJコード8桁目の
# 剤形記号で補う。記号の意味は剤形区分ごとに異なる点に注意
# （M＝内用薬ではカプセル、外用薬では軟膏）。
FORM_RULES = {
  "内用薬": [
    ("OD錠", r"OD錠|口腔内崩壊"),
    ("チュアブル", r"チュアブル|かみ砕|咀嚼"),
    ("ドライシロップ", r"ドライシロップ|DS\d|ＤＳ"),
    ("シロップ", r"シロップ"),
    ("内用液", r"内用液|経口液|内服液|内服ゼリー|経口ゼリー"),
    ("細粒", r"細粒"),
    ("顆粒", r"顆粒|グラニュール"),
    ("散", r"散\b|散剤|散\d|散$|原末|末$"),
    ("カプセル", r"カプセル"),
    ("錠", r"錠"),
  ],
  "外用薬": [
    ("眼軟膏", r"眼軟膏"),
    ("軟膏", r"軟膏"),
    ("クリーム", r"クリーム"),
    ("ゲル", r"ゲル|ジェル"),
    ("ローション", r"ローション|乳液"),
    ("テープ", r"テープ|パッチ"),
    ("パップ", r"パップ|湿布"),
    ("貼付剤", r"貼付"),
    ("坐剤", r"坐剤|座薬|坐薬|ホスコ"),
    ("点眼", r"点眼|眼灌流|眼科用"),
    ("点鼻", r"点鼻"),
    ("点耳", r"点耳"),
    ("吸入", r"吸入|エアゾール|ネブライザ|ジェヌエア|ディスカス|タービュヘイラー|レスピマット|エリプタ"),
    ("うがい", r"うがい|含嗽"),
    ("浣腸", r"浣腸"),
    ("スプレー", r"スプレー|フォーム|ミスト|噴霧|エロゾル|エアゾル|ゾル\d|ゾル$"),
    ("腟錠・腟剤", r"腟|膣"),
    ("口腔・歯科用", r"歯科用|口腔用|口腔内|トローチ|舌下"),
    ("パスタ", r"パスタ|ペースト"),
    ("消毒・外用液", r"消毒|消エタ|消アル|外用液|液$|液\d|液（|チンキ|エタノール|アルコール|イソプロパノール|イソプロ|アンモニア水|ポビドンヨード|オキシドール"),
    ("原末・その他", r"原末|末$|カンフル|ゴム末"),
    ("油・その他基剤", r"油$|油\d|油（|油「|石ケン|ワセリン"),
    ("ワイプ・清拭", r"ワイプ|清拭"),
  ],
  "注射薬": [
    ("キット・シリンジ", r"キット|シリンジ|オートインジェクター"),
    ("点滴静注", r"点滴|輸液|バッグ"),
    ("静注", r"静注|静脈内"),
    ("筋注", r"筋注|筋肉内"),
    ("皮下注", r"皮下注"),
    ("注射剤", r".*"),
  ],
}
FORM_SYM = {
  "内用薬": {"F":"錠","G":"錠","M":"カプセル","N":"カプセル","C":"細粒","D":"顆粒",
            "A":"散","B":"散","R":"ドライシロップ","Q":"シロップ","S":"内用液","X":"散"},
  "外用薬": {"M":"軟膏","N":"クリーム","S":"テープ","J":"坐剤","K":"浣腸","R":"点鼻",
            "Q":"点眼","G":"吸入","Y":"浣腸","F":"うがい","X":"消毒・外用液"},
  "注射薬": {"A":"注射剤","G":"キット・シリンジ","D":"注射剤","F":"点滴静注",
            "H":"静注","P":"キット・シリンジ","S":"キット・シリンジ","X":"注射剤"},
}
_FORM_CACHE = {}


def detect_form(name, kind, yj):
    """細かい剤形を返す。判定できなければ '' を返す。"""
    nm = unicodedata.normalize("NFKC", str(name or "")).replace("「", "").replace("」", "")
    for label, pat in FORM_RULES.get(kind, []):
        key = (kind, label)
        rx = _FORM_CACHE.get(key)
        if rx is None:
            rx = _FORM_CACHE[key] = re.compile(pat, re.I)
        if rx.search(nm):
            return label
    sym = (str(yj or "")[7:8] or "").upper()
    return FORM_SYM.get(kind, {}).get(sym, "")


# GitHub Actions は UTC で動くため、表示・記録はすべて日本時間に揃える
JST = datetime.timezone(datetime.timedelta(hours=9), "JST")


def now_jst():
    return datetime.datetime.now(JST)


def dict_rev(dicts, key, idx):
    """辞書インデックスから元の文字列を引く。
    辞書は生成中に増えるため、都度作り直す（件数が少ないので負荷にならない）。"""
    for k, v in dicts[key].items():
        if v == idx:
            return k
    return ""


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


# DSJPの画面に並ぶボタン等の文字列。包装名の末尾に紛れ込むため取り除く。
# 読点は全角/半角どちらもありうるので [、,] で吸収する
_PKG_NOISE_WORDS = [
    "セールス",
    "薬物データベース",
    "医薬品データベース",
    "薬価情報提供",
    "医薬品供給情報",
    "医薬品[、,]?\\s*バイオテクノロジー",
    "製薬会社情報",
    "医薬品[、,]?\\s*薬剤",
    "製薬業界動向",
    "医薬品包装価格",
    "添付文書",
    "インタビューフォーム",
    "お知らせ",
    "案内文書?",
    "詳細",
    "リンク",
    "コピー",
]
PKG_NOISE = re.compile(r"\s*(?:" + "|".join(_PKG_NOISE_WORDS) + r")\s*$")


def _is_pkg_line(s):
    return any(p.search(s) for p in PKG_PATTERNS)


def _clean_pkg(s):
    """包装名から、貼り付け時に混ざるボタン文字列を取り除く。
    末尾に複数付くことがあるので、変化しなくなるまで繰り返す。"""
    t = (s or "").strip()
    for _ in range(4):
        t2 = PKG_NOISE.sub("", t).strip()
        if t2 == t:
            break
        t = t2
    return t


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
    rec = {"n": "", "d": "", "name": "", "m": "", "p": ""}
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
                rec = {"n": "", "d": "", "name": "", "m": "", "p": ""}
                d1 = _to_iso(parts[1]) if len(parts) > 1 else ""
                recs.append({
                    "name": parts[0],
                    "n": "",
                    "p": "",
                    "d": d1,
                    "m": (parts[2] if len(parts) > 2 else
                          (parts[1] if len(parts) > 1 and not d1 else "")),
                })
                want = None
                continue

        if ln.startswith("告知日"):
            push(rec)
            rec = {"n": "", "d": "", "name": "", "m": "", "p": ""}
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

        if _is_pkg_line(ln):
            # 販売中止は包装単位で起こる。薬剤名の直後の包装表記を控えておき、
            # 「どの包装が中止か」を画面に出せるようにする。
            if rec["name"] and not rec["p"]:
                rec["p"] = _clean_pkg(ln)
            continue
        if ln in DISC_NOISE or ln.startswith(DISC_LABEL):
            continue

        if not rec["name"]:
            rec["name"] = ln
        else:
            # 名前が埋まっている状態で別の名前行が来た＝次の品目
            # （包装や会社名は上で除外済みなので、ここに来るのは薬剤名）
            push(rec)
            rec = {"n": "", "d": "", "name": ln, "m": "", "p": ""}

    push(rec)
    return recs


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

    def add(yj, rec):
        """販売中止は包装単位で起こるため、1薬剤に複数の包装を持てるようにする。
        （PTP50錠だけ中止でPTP10錠は継続、という状態を表せるようにする）"""
        cur = out.setdefault(yj, {"n": "", "d": "", "m": "", "pk": []})
        entry = {"n": rec["n"], "d": rec["d"], "p": rec.get("p", "")}
        for i, e in enumerate(cur["pk"]):
            if e.get("p") == entry["p"]:
                # 同じ包装が複数回出てくることがある（販売会社違いなど）。
                # 実施日が異なる場合は、早いほうを残す。
                # 先に入手できなくなる時期を示すほうが実務上安全なため。
                old_d, new_d = e.get("d", ""), entry.get("d", "")
                if old_d and new_d:
                    entry["d"] = min(old_d, new_d)
                elif old_d and not new_d:
                    entry["d"] = old_d
                cur["pk"][i] = entry
                break
        else:
            cur["pk"].append(entry)
        ds = [e["d"] for e in cur["pk"] if e["d"]]
        cur["d"] = min(ds) if ds else ""      # 代表は最も早い実施日
        ns = [e["n"] for e in cur["pk"] if e["n"]]
        cur["n"] = min(ns) if ns else ""
        if rec.get("m"):
            cur["m"] = rec["m"]

    for rec in parse_discontinued_text(text):
        key = rec["name"]
        # YJコード直接指定
        if re.fullmatch(r"[0-9A-Za-z]{6,12}", key) and re.search(r"\d", key) \
                and not re.search(r"[ぁ-んァ-ヶ一-龥]", key):
            add(key.upper(), rec)
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
            add(hits[0], rec)

    for k in notfound:
        print(f"  警告: 販売中止「{k}」は該当する薬剤が見つかりません（表記をご確認ください）")
    for k, hits in ambiguous:
        print(f"  警告: 販売中止「{k}」は同名が {len(hits)} 件あります。"
              f"YJコードで指定してください → {', '.join(hits)}")
    return out


def load_sentei(path):
    """選定療養の対象医薬品（薬価基準収載医薬品コード → 金額）を読む。"""
    if not path or not os.path.exists(path):
        return None
    try:
        d = json.load(open(path, encoding="utf-8"))
        if d.get("items"):
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


def _scope_css(css, scope):
    """スタイルの全セレクタを scope の下に閉じ込める。

    「データの成り立ち.html」は単体で開く前提の資料なので、
    body・p・table・.card といった広いセレクタを使っている。
    そのままページへ入れるとアプリ側の見た目を壊すため、
    すべて "#lgdoc ..." に書き換えてから入れる。
    """
    out, i, n = [], 0, len(css)
    while i < n:
        # 規則の切れ目には改行や空白が入るので、先に読み飛ばす。
        # ここを飛ばさないと @media が普通のセレクタとして扱われてしまう。
        while i < n and css[i] in " \t\r\n":
            i += 1
        if i >= n:
            break
        # @media などのブロックは、中身を再帰的に処理する
        if css[i] == "@":
            j = css.find("{", i)
            if j < 0:
                out.append(css[i:])
                break
            at = css[i:j].strip()
            depth, k = 1, j + 1
            while k < n and depth:
                if css[k] == "{":
                    depth += 1
                elif css[k] == "}":
                    depth -= 1
                k += 1
            inner = css[j + 1:k - 1]
            if at.lower().startswith(("@media", "@supports")):
                out.append(f"{at}{{{_scope_css(inner, scope)}}}")
            else:
                out.append(css[i:k])       # @font-face などはそのまま
            i = k
            continue
        j = css.find("{", i)
        if j < 0:
            break
        sel = css[i:j].strip()
        k = css.find("}", j)
        if k < 0:
            break
        body = css[j + 1:k]
        if not sel:
            i = k + 1
            continue
        parts = []
        for one in sel.split(","):
            one = one.strip()
            if not one:
                continue
            if one in (":root", "html", "body"):
                parts.append(scope)          # 資料の土台は入れ物そのもの
            elif one == "*":
                parts.append(f"{scope} *")
            else:
                parts.append(f"{scope} {one}")
        out.append(f"{','.join(parts)}{{{body}}}")
        i = k + 1
    return "".join(out)


def load_datadoc(path, scope="#lgdoc"):
    """「データの成り立ち.html」を凡例へ埋め込める形にして返す。

    資料とページで説明が食い違わないよう、資料のファイルを
    そのまま取り込む（差し替えれば両方に反映される）。
    見つからない場合は None を返し、タブごと出さない。
    """
    if not path or not os.path.exists(path):
        return None
    try:
        src = open(path, encoding="utf-8").read()
    except Exception:
        return None
    m = re.search(r"<style[^>]*>(.*?)</style>", src, re.S)
    css = _scope_css(m.group(1), scope) if m else ""
    m = re.search(r"<body[^>]*>(.*?)</body>", src, re.S)
    if not m:
        return None
    html = m.group(1)
    html = re.sub(r"<script.*?</script>", "", html, flags=re.S)  # 動作は本体側で持つ
    # 表題は凡例のタブ名と重複するので外す。ただし header の中にある
    # 4つの切り替えボタン（区分の関係／出典の対応／…）は残す必要がある。
    def _header(mo):
        tabs = re.search(r'<div class="tabs".*?</div>', mo.group(0), re.S)
        return tabs.group(0) if tabs else ""
    html = re.sub(r"<header.*?</header>", _header, html, flags=re.S)
    # 資料側の id（venn/map/join/calc）はページ側の id と衝突する
    # （calc は「計算」タブと同名）。同じ id が2つあると、
    # ページ側の動作が壊れるので接頭辞を付けて避ける。
    html = re.sub(r'(<section[^>]*\bid=")(\w+)"', r'\1doc-\2"', html)
    html = re.sub(r'(\bdata-t=")(\w+)"', r'\1doc-\2"', html)
    if "<section" not in html:
        return None
    return {"css": css, "html": html.strip()}


def load_ippanmei(path):
    """ippanmei.json（一般名処方マスタ）を読む。無ければ一般名なしで続行する。"""
    if not path or not os.path.exists(path):
        return None
    try:
        d = json.load(open(path, encoding="utf-8"))
        if d.get("items"):
            d.setdefault("map9", {})
            d.setdefault("mapx", {})
            return d
    except Exception:
        pass
    return None


def lookup_gen(ip, yj):
    """YJコードから一般名の通し番号を引く。該当が無ければ None。

    ① 例外コード品目対照表（12桁の完全一致）を先に見る
    ② 通常は「上9桁＋ZZZ」なので、上9桁で引く
    ①を先にするのは、例外コードが「上9桁では区分できない品目」のために
    作られているものだからで、順序を逆にすると
    持続性・非持続性の取り違えなどが起こりうる。
    """
    if not ip or not yj:
        return None
    n = ip["mapx"].get(yj)
    if n is not None:
        return n
    return ip["map9"].get(yj[:9])


# 和暦の経過措置期限（例: R9.3.31まで）を西暦へ
WAREKI = re.compile(
    r"(令和|平成|[RH令平])\s*(\d{1,2})\s*[\.\-/年]\s*(\d{1,2})\s*[\.\-/月]\s*(\d{1,2})")


def wareki_to_iso(t):
    """'R9.3.31まで' → '2027-03-31'。読めなければ '' を返す。"""
    m = WAREKI.search(str(t or ""))
    if not m:
        return ""
    era, y, mo, d = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))
    base = (2018 if era in ("R", "令", "令和")
            else 1988 if era in ("H", "平", "平成") else None)
    if base is None:
        return ""
    try:
        return f"{base + y:04d}-{mo:02d}-{d:02d}"
    except Exception:
        return ""


def lookup_expiry(pr, yj):
    """経過措置による使用期限。統一名収載品には設定されないため完全一致のみ。"""
    if not pr or not yj:
        return ""
    return (pr.get("expiry") or {}).get(yj, "")


def is_jpharm(pr, yj):
    """日本薬局方収載品か。薬価リストの「局」の印による。
    薬価と同様、統一名収載品は先頭9桁でも判定する
    （例: カロナール原末は統一名『アセトアミノフェン』の局方指定を受け継ぐ）。"""
    if not pr or not yj:
        return False
    js = pr.get("_jpset") or set()
    return yj in js or yj[:9] in js


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
          prev_snapshot=None, snapshot_out=None, snapshot_path=None,
          prices_path=None, keep_chg=None,
          kiso_path=None, disc_path=None, sentei_path=None,
          ippanmei_path=None, datadoc_path=None):
    """prev_snapshot: {YJコード: sc} from the previous edition, for 悪化/改善 detection.
    snapshot_out: path to write this edition's snapshot for the next run."""
    hdr = find_header_row(xlsx_path)
    df = pd.read_excel(xlsx_path, header=hdr)
    c = list(df.columns)
    if len(c) < 21:
        raise ValueError(f"列数が想定と異なります（{len(c)}列）。様式変更の可能性があります。")

    dicts = {k: {} for k in ['st','vol','rsn','out','cls','m','k','note','pc','i','fm']}
    def idx(key, val):
        d = dicts[key]
        if val not in d: d[val] = len(d)
        return d[val]

    pr = load_prices(prices_path)
    kiso = load_kiso(kiso_path)
    sentei = load_sentei(sentei_path)
    ippan = load_ippanmei(ippanmei_path)
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
    chg_map = {}
    n_price = 0
    n_kiso = 0
    n_kchg = 0
    n_sen = 0
    n_disc = 0
    n_gen = 0
    for _, r in df.iterrows():
        name = clean(r[c[5]])
        if not name: continue
        st = strip_prefix(r[c[11]])
        ing, mk, sp, yj = clean(r[c[2]]), clean(r[c[6]]), clean(r[c[3]]), clean(r[c[4]])
        sc = SC.get(st, 3)
        # 悪化/改善: compare severity against the previous edition (0<1<2)
        chg = 0
        if keep_chg is not None:
            # 再生成時は前回の判定結果をそのまま引き継ぐ（比較し直さない）
            chg = keep_chg.get(yj, 0)
        elif prev_snapshot is not None and yj in prev_snapshot:
            old_sc = prev_snapshot[yj]
            if sc in (0,1,2) and old_sc in (0,1,2):
                if   sc > old_sc: chg = 1   # 悪化
                elif sc < old_sc: chg = 2   # 改善
        elif prev_snapshot:
            chg = 3                          # 新規掲載
        snap[yj] = sc
        chg_map[yj] = chg
        price = lookup_price(pr, yj)
        exp_raw = lookup_expiry(pr, yj)
        jp = 1 if is_jpharm(pr, yj) else 0
        # 選定療養の対象。薬価基準収載医薬品コード（＝YJコード）で突合する。
        sv = 0
        if sentei:
            it = (sentei.get("items") or {}).get(yj)
            if it:
                n_sen += 1
                sv = [it.get("h"), it.get("p"), it.get("g")]   # 差額分 / 選定療養時 / 後発品最高価格
        exp_iso = wareki_to_iso(exp_raw)
        if price is not None:
            n_price += 1
        # ⑨基礎的医薬品（1：対象）
        kind_txt = strip_prefix(r[c[0]])
        form = detect_form(name, kind_txt, yj)
        kb = 1 if clean(r[c[8]]).startswith("1") else 0
        n_kiso += kb
        # 変更調剤が認められる基礎的医薬品（品名で突合）
        kc = 1 if (kiso is not None and name in kiso) else 0
        n_kchg += kc
        # 販売中止（手動登録）。供給状況とは別軸の情報として持つ
        dc = disc.get(yj)
        if dc:
            n_disc += 1
            # [告知日, 実施日, メモ, 包装ごとの明細]
            dcv = [dc.get("n", ""), dc.get("d", ""), dc.get("m", ""),
                   dc.get("pk", [])]
        else:
            dcv = 0
        # 一般名処方の標準的な記載（一般名コードで突合）
        gi = lookup_gen(ippan, yj)
        if gi is None:
            gi = -1
        else:
            n_gen += 1
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
            idx('fm', form),                    # [23] 細かい剤形
            exp_iso or exp_raw,                 # [24] 経過措置期限
            jp,                                 # [25] 1=日本薬局方収載品
            sv,                                 # [26] 選定療養 [差額分, 選定療養時薬価] / 0
            0,                                  # [27] 併売品の行番号リスト / 0
            gi,                                 # [28] 一般名の通し番号 / -1
        ])
    if not rows:
        raise ValueError("有効なデータ行が0件です。")

    data = {
        "date": as_of or datetime.date.today().isoformat(),
        "n": len(rows),
        "src": source_label, "srcurl": source_url,
        "gen": now_jst().strftime("%Y/%m/%d %H:%M"),
        "d": {k: [s for s,_ in sorted(v.items(), key=lambda x: x[1])] for k,v in dicts.items()},
        "r": rows,
    }
    if snapshot_out:
        # sc … 今回の出荷状況（次回の比較に使う）
        # chg … 今回検出した悪化/改善（再生成しても消えないよう結果を保存する）
        # hist … 状態が変わった日だけを記録する。
        #        毎回の全件を残すと膨らむため、変化点のみを積む。
        prev_hist = {}
        if os.path.exists(snapshot_out):
            try:
                prev_hist = json.load(
                    open(snapshot_out, encoding="utf-8")).get("hist", {})
            except Exception:
                prev_hist = {}

        today = data["date"]
        prev_dates = []
        if os.path.exists(snapshot_out):
            try:
                prev_dates = json.load(
                    open(snapshot_out, encoding="utf-8")).get("dates", [])
            except Exception:
                prev_dates = []
        dates = sorted(set(prev_dates) | {today})[-60:]   # 直近60版まで
        hist = dict(prev_hist)
        for yj, sc in snap.items():
            h = hist.get(yj)
            if not h:
                hist[yj] = [[today, sc]]        # 初回
            elif h[-1][1] != sc:
                h.append([today, sc])           # 変化した日だけ足す
                if len(h) > 40:                 # 古い分は間引く
                    hist[yj] = h[-40:]

        with open(snapshot_out, "w", encoding="utf-8") as f:
            json.dump({
                "date": today,
                "dates": dates,
                "sc": snap,
                "chg": {yj: v for yj, v in chg_map.items() if v},
                "hist": hist,
            }, f, separators=(',', ':'))

    # ---- 成分ごとの逼迫度（お知らせ版・グラフ用） ----
    # 同一成分の中で、限定出荷以下（限定・停止）が占める割合を出す。
    # 供給停止だけの割合も別に持ち、グラフで積み上げられるようにする。
    # 剤形ごとに分けて数える。薬局では内用薬・外用薬が主で、
    # 注射薬まで混ぜると逼迫率が実態とずれるため。
    KIDX = {"内用薬": 0, "外用薬": 1, "注射薬": 2}
    # 行の生成が終わった時点で辞書は確定しているので、逆引きを一度だけ作る
    K_REV = {v: k for k, v in dicts["k"].items()}
    ing_stat = {}
    for row in rows:
        key = row[1]                          # 成分名の辞書インデックス
        d = ing_stat.setdefault(key, {
            "n": 0, "lim": 0, "stop": 0, "worse": 0,
            "k": [[0, 0, 0], [0, 0, 0], [0, 0, 0]],   # 剤形別 [総数, 限定, 停止]
        })
        d["n"] += 1
        ki = KIDX.get(K_REV.get(row[5]), None)
        if ki is not None:
            d["k"][ki][0] += 1
        sc_v = row[8]
        if sc_v == 1:
            d["lim"] += 1
            if ki is not None:
                d["k"][ki][1] += 1
        elif sc_v == 2:
            d["stop"] += 1
            if ki is not None:
                d["k"][ki][2] += 1
        if row[18] == 1:                      # 前回版から悪化
            d["worse"] += 1

    # 2品目以上ある成分はすべて収録する。
    # 逼迫率で足切りすると、お知らせ掲示板（上昇幅5ポイント以上など、
    # より緩い条件で載る）からタップした成分が下段に存在せず、
    # 「該当する成分はありません」になってしまうため。
    # 実際の絞り込みは画面側のスライダーとボタンに任せる。
    alerts = []
    for key, d in ing_stat.items():
        if d["n"] < 2:                        # 1品目だけの成分は比較にならない
            continue
        bad = d["lim"] + d["stop"]
        ratio = bad / d["n"]
        alerts.append({
            "i": key, "n": d["n"], "lim": d["lim"], "stop": d["stop"],
            "r": round(ratio, 4), "w": d["worse"], "k": d["k"],
        })
    alerts.sort(key=lambda a: (-a["w"], -a["r"], -a["n"]))
    data["alerts"] = alerts

    # ---- 併売品（同じ中身で商品名が違う先発品どうし） ----
    # YJコードの先頭9桁は「薬効分類+成分+剤形+規格」を表すので、
    # ここが一致する＝中身が同じ。ただし後発品も同じ9桁を共有するため、
    # 先発品・準先発品・長期収載品に限ったうえで、
    # さらに商品名の語幹が異なるものだけを併売品とみなす。
    # 剤形を表す語を「まとまり」として区切る。
    # 文字クラス [錠カプセル…] にすると「カ」「プ」「セ」…と1文字ずつ
    # 区切ってしまい、「ラクティオン」が「ラ」で切れるなど語幹が壊れる。
    # 長い語を先に並べて、部分一致で誤って切れないようにする。
    _FORMS = ["ドライシロップ", "カプセル", "シロップ", "ローション", "クリーム",
              "細粒", "顆粒", "散剤", "軟膏", "ゲル", "テープ", "パップ",
              "坐剤", "点眼", "点鼻", "吸入", "配合",
              "錠", "散", "液", "注"]
    BRAND_SPLIT = re.compile("(?:" + "|".join(_FORMS) + r"|[０-９0-9])")

    def brand_of(name):
        return BRAND_SPLIT.split(name)[0].strip()

    ORIG = ("先発品", "準先発品", "長期収載品")
    by9 = {}
    for i, row in enumerate(rows):
        if dict_rev(dicts, "pc", row[17]) not in ORIG:
            continue
        yj = row[4]
        if len(yj) < 9:
            continue
        by9.setdefault(yj[:9], []).append(i)

    comarket = {}          # 行番号 → 併売相手の行番号のリスト
    for key, idxs in by9.items():
        if len(idxs) < 2:
            continue
        brands = {i: brand_of(rows[i][0]) for i in idxs}
        if len(set(brands.values())) < 2:
            continue                      # 同じブランドの規格違いにすぎない
        for i in idxs:
            others = [j for j in idxs if brands[j] != brands[i]]
            if others:
                comarket[i] = others

    for i, others in comarket.items():
        rows[i][27] = others
    data["comarket"] = len(comarket)

    # ---- 出荷状況の推移（折れ線グラフ用） ----
    # 取得のたびに記録した「変化点」から、各日付での件数を組み立てる。
    # 全成分ぶんを持つと重いので、お知らせに出る成分と全体の合計だけにする。
    series = {}
    src = snapshot_out or snapshot_path
    if src and os.path.exists(src):
        try:
            sj = json.load(open(src, encoding="utf-8"))
            sdates = sj.get("dates") or ([sj["date"]] if sj.get("date") else [])
            shist = sj.get("hist") or {}
        except Exception:
            sdates, shist = [], {}

        if sdates and shist:
            want = {a["i"] for a in alerts}
            # YJコード → (成分名インデックス, 剤形インデックス)
            yj2 = {r[4]: (r[1], KIDX.get(K_REV.get(r[5]))) for r in rows}

            # 各品目の「日付ごとの状態」を、変化点から前方に埋めて復元する。
            # 剤形ごとに分けて数え、画面側で内用・外用・注射を選べるようにする。
            def blank():
                return [[0, 0], [0, 0], [0, 0]]     # 剤形別 [限定, 停止]
            per = {d: {} for d in sdates}
            total = {d: blank() for d in sdates}
            for yj, pts in shist.items():
                v = yj2.get(yj)
                if not v or v[1] is None:
                    continue
                ing, ki = v
                track = ing in want
                pi, cur = 0, None
                for d in sdates:
                    while pi < len(pts) and pts[pi][0] <= d:
                        cur = pts[pi][1]
                        pi += 1
                    if cur == 1 or cur == 2:
                        j = 0 if cur == 1 else 1
                        total[d][ki][j] += 1
                        if track:
                            per[d].setdefault(ing, blank())[ki][j] += 1

            def pack(src):
                # [日付, 内用[限定,停止], 外用[...], 注射[...]]
                return [[d, src[d][0], src[d][1], src[d][2]] for d in sdates]

            series["_all"] = pack(total)
            for ing in want:
                series[str(ing)] = [
                    [d] + (per[d].get(ing) or blank()) for d in sdates
                ]
    data["series"] = series
    data["sdates"] = len(series.get("_all") or [])

    # ---- お知らせ掲示板（事実の提示） ----
    # 履歴から読み取れる「起きたこと」を、ジャンル別に文章化する。
    # 予測ではなく事実だけを出す（供給停止は外部要因が支配的で、
    # 過去の遷移から統計的に予測できる性質のものではないため）。
    news = []
    if src and os.path.exists(src):
        try:
            sj2 = json.load(open(src, encoding="utf-8"))
            nd = sj2.get("dates") or []
            nh = sj2.get("hist") or {}
        except Exception:
            nd, nh = [], {}

        if len(nd) >= 2:
            ing_of = {r[4]: r[1] for r in rows}
            kind_of = {r[4]: KIDX.get(K_REV.get(r[5])) for r in rows}

            # 成分ごと・剤形ごとに、日付別の [限定, 停止, 総数] を作る。
            # 画面で剤形を選び直せるよう、剤形別のまま持っておく。
            def blank3():
                return [[0, 0, 0], [0, 0, 0], [0, 0, 0]]
            agg = {}
            for yj, pts in nh.items():
                ing = ing_of.get(yj)
                ki = kind_of.get(yj)
                if ing is None or ki is None:
                    continue
                a = agg.setdefault(ing, {d: blank3() for d in nd})
                pi, cur = 0, None
                for d in nd:
                    while pi < len(pts) and pts[pi][0] <= d:
                        cur = pts[pi][1]
                        pi += 1
                    a[d][ki][2] += 1
                    if cur == 1:
                        a[d][ki][0] += 1
                    elif cur == 2:
                        a[d][ki][1] += 1

            first, last = nd[0], nd[-1]
            for ing, a in agg.items():
                tot = lambda d, j: sum(a[d][k][j] for k in range(3))
                n0, n1 = tot(first, 2), tot(last, 2)
                if n1 < 2:
                    continue
                r0 = (tot(first, 0) + tot(first, 1)) / n0 if n0 else 0
                r1 = (tot(last, 0) + tot(last, 1)) / n1
                # 剤形別の内訳を添え、画面側で選び直せるようにする
                f0 = [list(a[first][k]) for k in range(3)]
                f1 = [list(a[last][k]) for k in range(3)]
                # ① 逼迫率が上がった
                if r1 - r0 >= 0.10:
                    news.append({
                        "t": "rise", "i": ing,
                        "a": round(r0 * 100), "b": round(r1 * 100),
                        "d0": first, "d1": last, "f0": f0, "f1": f1,
                    })
                # ② 限定出荷以下が一定割合を超えた
                if r1 >= 0.30 and r0 < 0.30:
                    news.append({"t": "cross", "i": ing,
                                 "b": round(r1 * 100), "d1": last,
                                 "f0": f0, "f1": f1})

            # ③④ 品目ごとの推移から、悪化の連続と往復を拾う
            for yj, pts in nh.items():
                ing = ing_of.get(yj)
                if ing is None or len(pts) < 2:
                    continue
                seq = [p[1] for p in pts]
                # 3段階以上の悪化（0→1→2 のように単調に悪くなった）
                run = 1
                for i in range(1, len(seq)):
                    if seq[i] > seq[i - 1]:
                        run += 1
                        if run >= 3:
                            news.append({"t": "down", "i": ing, "yj": yj,
                                         "d0": pts[0][0], "d1": pts[-1][0],
                                         "n": run, "kf": kind_of.get(yj)})
                            break
                    else:
                        run = 1
                # 往復（上がったり下がったりを繰り返す）
                turns = sum(1 for i in range(1, len(seq) - 1)
                            if (seq[i] - seq[i - 1]) * (seq[i + 1] - seq[i]) < 0)
                if turns >= 3:
                    news.append({"t": "swing", "i": ing, "yj": yj,
                                 "n": turns + 1, "kf": kind_of.get(yj)})

    # 同じ成分の同じ種類は1件にまとめる。
    # 種類ごとに枠を分けないと、件数の多い種類だけで埋まってしまう。
    def rank(z):
        if z["t"] == "rise":
            return -(z.get("b", 0) - z.get("a", 0))
        if z["t"] == "cross":
            return -z.get("b", 0)
        return -z.get("n", 0)

    seen, buckets = set(), {"rise": [], "cross": [], "down": [], "swing": []}
    for x in sorted(news, key=rank):
        k = (x["t"], x["i"])
        if k in seen:
            continue
        seen.add(k)
        buckets[x["t"]].append(x)
    data["news"] = ([*buckets["rise"][:150], *buckets["cross"][:150],
                     *buckets["down"][:150], *buckets["swing"][:150]])

    # ---- 成分ごとの推移（折れ線グラフ用） ----
    # snapshot の変化点履歴から「各日にちで何品目が限定出荷／供給停止だったか」を
    # 復元する。お知らせに出る成分だけに絞るので、容量は小さく収まる。
    trend = {}
    if snapshot_out and os.path.exists(snapshot_out):
        try:
            hist = json.load(open(snapshot_out, encoding="utf-8")).get("hist", {})
        except Exception:
            hist = {}
        if hist:
            yj_to_ing = {r[4]: r[1] for r in rows}
            target = {a["i"] for a in alerts}
            # 記録に出てくる日付をすべて集める
            all_days = sorted({d for h in hist.values() for d, _ in h})
            for ing in target:
                yjs = [y for y, i in yj_to_ing.items() if i == ing and y in hist]
                if not yjs:
                    continue
                ds, ls, ss = [], [], []
                for day in all_days:
                    lim = stop = seen = 0
                    for y in yjs:
                        # その日時点で有効な状態（それ以前の最後の記録）
                        cur = None
                        for d, sc in hist[y]:
                            if d <= day:
                                cur = sc
                            else:
                                break
                        if cur is None:
                            continue
                        seen += 1
                        if cur == 1:
                            lim += 1
                        elif cur == 2:
                            stop += 1
                    if seen:
                        ds.append(day)
                        ls.append(lim)
                        ss.append(stop)
                if ds:
                    trend[ing] = {"d": ds, "l": ls, "s": ss}
    data["trend"] = trend

    data["price"] = {
        "available": pr is not None,
        "as_of": (pr or {}).get("as_of", ""),
        "matched": n_price,
        "files": (pr or {}).get("files", {}),
        "fetched": (pr or {}).get("updated_at", ""),
    }

    # データ元ごとの日付（画面の「データの鮮度」で表示する）
    kj = {}
    if kiso_path and os.path.exists(kiso_path):
        try:
            kj = json.load(open(kiso_path, encoding="utf-8"))
        except Exception:
            kj = {}
    data["sources"] = {
        "supply": {"label": "医療用医薬品供給状況",
                   "date": data["date"],
                   "file": (source_url or "").rsplit("/", 1)[-1]},
        "sentei": {"label": "選定療養の対象医薬品リスト",
                   "date": (sentei or {}).get("as_of", ""),
                   "file": (sentei or {}).get("file", ""),
                   "fetched": (sentei or {}).get("updated_at", "")},
        "kiso":   {"label": "変更調剤が認められる基礎的医薬品等",
                   "date": kj.get("as_of", ""),
                   "file": kj.get("file", ""),
                   "fetched": kj.get("updated_at", "")},
        "ippan":  {"label": "一般名処方マスタ",
                   "date": (ippan or {}).get("as_of", ""),
                   "files": (ippan or {}).get("files", {}),
                   "fetched": (ippan or {}).get("updated_at", "")},
    }

    data["disc"] = {"count": n_disc, "registered": len(disc)}
    data["sentei"] = {
        "available": sentei is not None,
        "count": n_sen,
        "as_of": (sentei or {}).get("as_of", ""),
        "file": (sentei or {}).get("file", ""),
    }

    data["kiso"] = {
        "basic": n_kiso,
        "swap_available": kiso is not None,
        "swap": n_kchg,
    }

    # ---- 一般名処方の標準的な記載 ----
    # 画面側は行データ[28]の通し番号で items を引く。
    # どの一般名にどの品目がぶら下がるかは、閲覧時にブラウザ側で組み立てる
    # （生成時に持たせると同じ情報を二重に抱えることになるため）。
    if ippan:
        # cur=1 … 現行版に載っている / cur=0 … 過去の版から引き継いだ記載。
        # 引き継ぎ分は「一般名処方加算の対象ではなくなったため削除された」
        # ものなので、画面では加算区分を出さず「旧版」と表示する。
        gitems, n_old = [], 0
        for it in ippan["items"]:
            old = 0 if it.get("cur") else 1
            n_old += old
            gitems.append([it["t"], it["k"], it["i"], it["s"],
                           it["a"], it["p"], it["b"], it["bs"], it["x"],
                           old, it.get("v", "")])
        with_items = len({r[28] for r in rows if r[28] >= 0})
        data["ippan"] = {
            "available": True,
            "as_of": ippan.get("as_of", ""),
            "files": ippan.get("files", {}),
            "fetched": ippan.get("updated_at", ""),
            "matched": n_gen,
            "with_items": with_items,
            "old": n_old,
            "items": gitems,
        }
    else:
        data["ippan"] = {"available": False, "items": []}

    data["chg"] = {
        "worse":    sum(1 for r in rows if r[18] == 1),
        "better":   sum(1 for r in rows if r[18] == 2),
        "new":      sum(1 for r in rows if r[18] == 3),
        "compared": prev_snapshot is not None and len(prev_snapshot) > 0,
    }

    head = open(TEMPLATE_HEAD, encoding="utf-8").read()
    tail = open(TEMPLATE_TAIL, encoding="utf-8").read()

    # 「データの成り立ち」を凡例の2つ目のタブとして埋め込む。
    # 資料の側を直せばページにも反映されるので、説明が二重管理にならない。
    doc = load_datadoc(datadoc_path)
    head = head.replace("/*__DOCCSS__*/", doc["css"] if doc else "")
    head = head.replace("<!--__DOCHTML__-->", doc["html"] if doc else "")
    data["doc"] = bool(doc)
    if doc:
        print(f"  データの成り立ち: 取り込みました（{len(doc['html']):,}字）")
    else:
        print("  データの成り立ち: 見つからないため、凡例のタブは出しません")

    # JSONを <script> 内に埋め込むため、HTMLとして解釈されうる文字列を無害化する。
    # 特に "</script>" が本文（薬剤名・包装名など）に含まれると
    # scriptタグが途中で閉じ、ページ全体が動かなくなる。
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    payload = (payload.replace("</", "<\\/")
                      .replace("\u2028", "\\u2028")
                      .replace("\u2029", "\\u2029"))

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(head)
        f.write(payload)
        f.write(tail)
    return len(rows)

if __name__ == "__main__":
    import sys
    n = build(sys.argv[1], sys.argv[2] if len(sys.argv)>2 else "out.html")
    print(f"{n} items")
