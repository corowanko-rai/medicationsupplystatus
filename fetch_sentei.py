#!/usr/bin/env python3
"""
長期収載品の選定療養の対象医薬品リストを取得する。

厚労省「後発医薬品のある先発医薬品（長期収載品）の選定療養について」
  https://www.mhlw.go.jp/stf/newpage_39830.html
の「対象医薬品リストについて」の表から、Excelを取得して sentei.json を作る。

ファイル名が 001684477.xlsx のような通し番号で、日付から推測できない。
そのため表のセルに書かれた「※令和８年６月１日から」という適用開始日を読み、
今日すでに始まっているもののうち最新のものを選ぶ。
（将来分が先に公表されることがあるため、開始前のものは採らない）

  python3 fetch_sentei.py           # 変更があれば更新
  python3 fetch_sentei.py --check   # 確認のみ

終了コード: 0=更新した / 10=変更なし / 1=エラー

失敗しても供給状況ページの生成は続行できる設計。
"""
import sys, os, re, json, hashlib, datetime, urllib.error

import fetch_prices as fp   # HTTP取得とデコードの共通処理を再利用

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "sentei.json")
PAGE = "https://www.mhlw.go.jp/stf/newpage_39830.html"

JST = datetime.timezone(datetime.timedelta(hours=9), "JST")


def now_jst():
    return datetime.datetime.now(JST)


def log(m):
    print(f"[{now_jst():%Y-%m-%d %H:%M:%S}] {m}", flush=True)


# ---- 和暦の読み取り ------------------------------------------------
WAREKI = re.compile(
    r"令和\s*([０-９0-9一二三四五六七八九十元]+)\s*年"
    r"\s*([０-９0-9]+)\s*月\s*([０-９0-9]+)\s*日")
_KAN = {"元": 1, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
        "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def _num(s):
    s = s.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    if s.isdigit():
        return int(s)
    if s in _KAN:
        return _KAN[s]
    if s.startswith("十"):
        return 10 + (_KAN.get(s[1:], 0) if len(s) > 1 else 0)
    return 0


START_RE = re.compile(r"※\s*(令和[^、。\n]*?)から")


def start_date(text):
    """「※令和8年6月1日から」の適用開始日を返す。読めなければ None。

    1つのウィンドウに複数の「※〜から」が含まれることがある
    （例: 旧版のブロックが直前に残っている場合）ため、
    直後のExcelリンクに対応するのは常に最後に出てくるものとみなす。
    また「※」を必須にして、「令和8年3月31日 事務連絡」のような
    無関係な日付を誤って拾わないようにする。
    """
    ms = list(START_RE.finditer(text))
    if not ms:
        return None
    m = ms[-1]
    d = WAREKI.search(m.group(1))
    if not d:
        return None
    try:
        return datetime.date(2018 + _num(d.group(1)),
                             _num(d.group(2)), _num(d.group(3)))
    except ValueError:
        return None


def find_excel(html):
    """(適用開始日, URL) のうち、今日すでに始まっている最新のものを返す。

    HTMLParserでtd/thのネストを逐一追うのはページの実際のマークアップ
    （colspan・入れ子span・コメント等）に弱いため、ここでは
    「.xlsxへのリンク」を基準に、その直前のテキスト（数千字）から
    直近の適用開始日を逆読みする方式にする。壊れたHTMLでも
    リンク自体さえ見つかれば日付とセットで拾える。
    """
    today = now_jst().date()

    # <a ... href="....xlsx" ...>ラベル</a> をすべて拾う（属性順序を問わない）
    link_re = re.compile(
        r'<a\b[^>]*?\bhref\s*=\s*(["\'])(?P<href>[^"\']+?\.xlsx)\1[^>]*>',
        re.IGNORECASE)

    cands = []
    for m in link_re.finditer(html):
        url = fp.absolutize(m.group("href"))
        if not url:
            continue
        # リンクの手前 1500 文字程度をテキスト化して、直近の
        # 「※令和８年６月１日から」を探す（タグは大まかに除去する）
        window = html[max(0, m.start() - 1500):m.start()]
        window = re.sub(r"<[^>]+>", "", window)
        window = re.sub(r"&nbsp;?", " ", window)
        d = start_date(window)
        cands.append((d, url, m.start()))

    if not cands:
        return None

    started = [(d, u) for d, u, _ in cands if d and d <= today]
    if started:
        return max(started, key=lambda x: x[0])

    log("警告: 適用開始日を読み取れませんでした。ページ内で最後に"
        "現れたExcelリンクを使います。")
    cands.sort(key=lambda x: x[2])
    return (None, cands[-1][1])


def parse_excel(path):
    """薬価基準収載医薬品コード → {差額分, 選定療養時の薬価} を作る。"""
    import pandas as pd

    df = pd.read_excel(path, sheet_name=0, header=None, dtype=str)
    hdr = None
    for i in range(min(10, len(df))):
        joined = "".join(str(x) for x in df.iloc[i].tolist())
        if "薬価基準収載医薬品コード" in joined:
            hdr = i
            break
    if hdr is None:
        raise ValueError("見出し行が見つかりません。様式変更の可能性があります。")

    df = pd.read_excel(path, sheet_name=0, header=hdr, dtype=str)
    cols = {str(c).strip(): c for c in df.columns}

    def pick(*names):
        for n in names:
            if n in cols:
                return cols[n]
        for n in names:
            for k, v in cols.items():
                if n in k:
                    return v
        return None

    c_code = pick("薬価基準収載医薬品コード")
    c_half = pick("価格差の２分の１", "価格差の2分の1", "２分の１")
    c_pay = pick("保険外併用療養費の算出に用いる価格", "保険外併用療養費")
    c_gmax = pick("後発医薬品最高価格", "後発品最高価格")
    if c_code is None:
        raise ValueError(f"コード列が見つかりません: {list(cols)[:8]}")

    def f(v):
        try:
            return round(float(str(v).replace(",", "").strip()), 2)
        except (TypeError, ValueError):
            return None

    out = {}
    for _, r in df.iterrows():
        code = str(r[c_code] or "").strip().upper()
        if not re.fullmatch(r"[0-9A-Z]{12}", code):
            continue
        out[code] = {
            "h": f(r[c_half]) if c_half is not None else None,   # 差額分薬価
            "p": f(r[c_pay]) if c_pay is not None else None,     # 選定療養時の薬価
            "g": f(r[c_gmax]) if c_gmax is not None else None,   # 後発品最高価格
        }
    return out


def main():
    check = "--check" in sys.argv
    try:
        log("選定療養の対象医薬品リストのページを確認中…")
        raw = fp.http_get(PAGE)
        log(f"ページ取得: {len(raw):,} bytes")
        html = fp.decode_html(raw)
        n_xlsx = len(re.findall(r'\.xlsx["\']', html, re.IGNORECASE))
        n_mark = html.count("※")
        log(f"  .xlsx形式のリンク候補: {n_xlsx} 件 / ※の出現: {n_mark} 件")
        got = find_excel(html)
        if not got:
            log("ERROR: 対象医薬品リストのExcelが見つかりません。")
            log("  ページの取得自体はできていますが、期待する構造が"
                "見つかりませんでした。ページの様式が変わった可能性があります。")
            # 診断用に、最初に見つかった.xlsxリンク周辺を少しログへ出す
            m = re.search(r'.{80}\.xlsx["\'].{20}', html, re.IGNORECASE | re.S)
            if m:
                log(f"  参考（最初の.xlsx付近）: ...{m.group(0)!r}...")
            return 1
        sdate, url = got
        log(f"対象リスト: {url.rsplit('/', 1)[-1]}"
            + (f"（{sdate} から適用）" if sdate else ""))

        prev = {}
        if os.path.exists(OUT):
            try:
                prev = json.load(open(OUT, encoding="utf-8"))
            except Exception:
                prev = {}

        if check:
            log("--check のため、ダウンロードは行いません。")
            log(f"  前回: {prev.get('as_of', '(未取得)')} / 今回: {sdate}")
            return 0

        blob = fp.http_get(url)
        h = hashlib.sha256(blob).hexdigest()
        log(f"取得完了 {len(blob):,} bytes / sha256={h[:16]}…")

        REQUIRED = ("items", "as_of", "file")
        missing = [k for k in REQUIRED if k not in prev]
        if prev.get("sha256") == h and not missing:
            log("変更なし（前回と同一）。処理を終了します。")
            return 10
        if prev.get("sha256") == h and missing:
            log(f"内容は同じですが記録に不足があります（{', '.join(missing)}）。作り直します。")

        tmp = os.path.join(HERE, "_sentei_tmp.xlsx")
        with open(tmp, "wb") as f:
            f.write(blob)
        items = parse_excel(tmp)
        os.remove(tmp)

        with open(OUT, "w", encoding="utf-8") as f:
            json.dump({
                "as_of": sdate.isoformat() if sdate else "",
                "file": url.rsplit("/", 1)[-1],
                "source": url,
                "sha256": h,
                "updated_at": now_jst().isoformat(timespec="seconds"),
                "items": items,
            }, f, ensure_ascii=False, separators=(",", ":"))
        log(f"sentei.json を更新（{len(items):,}品目）")
        return 0

    except urllib.error.URLError as e:
        log(f"ERROR: 通信に失敗しました: {e}")
        return 1
    except Exception as e:
        import traceback
        traceback.print_exc()
        log(f"ERROR: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
