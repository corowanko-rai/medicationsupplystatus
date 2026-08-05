#!/usr/bin/env python3
"""
薬価基準収載品目リスト 取得スクリプト

厚労省の「医療保険が適用される医薬品について」（固定URL）から
最新の薬価基準収載品目リストのページを辿り、
内用薬・注射薬・外用薬のExcelを取得して prices.json を作る。

  python3 fetch_prices.py           # 変更があれば更新
  python3 fetch_prices.py --check   # 確認のみ（DL・生成なし）

終了コード: 0=更新した / 10=変更なし / 1=エラー

このスクリプトが失敗しても供給状況ページの生成は続行できるよう、
呼び出し側（ワークフロー）では失敗を許容する設計にしている。
"""
import sys, os, re, json, hashlib, datetime, urllib.request, urllib.error
from html.parser import HTMLParser

# GitHub Actions は UTC で動くため、記録・ログは日本時間に揃える
JST = datetime.timezone(datetime.timedelta(hours=9), "JST")


def now_jst():
    return datetime.datetime.now(JST)

INDEX = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000078916.html"
BASE  = "https://www.mhlw.go.jp"
UA    = "Mozilla/5.0 (compatible; drug-price-updater/1.0)"
HERE  = os.path.dirname(os.path.abspath(__file__))
OUT   = os.path.join(HERE, "prices.json")

# ファイル末尾の連番 → 区分
KIND = {"01": "内用薬", "02": "注射薬", "03": "外用薬"}


def log(m):
    print(f"[{now_jst():%Y-%m-%d %H:%M:%S}] {m}", flush=True)


def http_get(url, timeout=120):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def decode_html(raw):
    """厚労省のページはUTF-8とShift-JISが混在する。URLはASCIIなので
    どちらで読めても抽出できるが、判定は順に試す。"""
    for enc in ("utf-8", "cp932", "euc-jp"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


class Links(HTMLParser):
    def __init__(self):
        super().__init__()
        self.hrefs = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            h = dict(attrs).get("href")
            if h:
                self.hrefs.append(h)


def absolutize(h):
    if h.startswith("http"):
        return h
    if h.startswith("/"):
        return BASE + h
    return None


def find_list_page(html):
    """薬価基準収載品目リストの最新ページURLを返す。
    /topics/YYYY/MM/tpYYYYMMDD-01.html 形式のうち日付が最大のもの。"""
    p = Links(); p.feed(html)
    best = None
    for h in p.hrefs:
        full = absolutize(h)
        if not full:
            continue
        m = re.search(r"/topics/\d{4}/\d{2}/tp(\d{8})-01\.html$", full)
        if m:
            d = m.group(1)
            if best is None or d > best[0]:
                best = (d, full)
    return best  # (YYYYMMDD, url) or None


def find_excels(html):
    """一覧ページから 内用薬/注射薬/外用薬 のExcel URLを取る。
    連番ごとに日付が異なりうるため、各区分で最新日付を採用する。"""
    p = Links(); p.feed(html)
    best = {}
    for h in p.hrefs:
        full = absolutize(h)
        if not full:
            continue
        m = re.search(r"/xls/tp(\d{8})-01_(\d{2})\.xlsx$", full, re.I)
        if not m:
            continue
        date, seq = m.group(1), m.group(2)
        if seq not in KIND:
            continue
        if seq not in best or date > best[seq][0]:
            best[seq] = (date, full)
    return best  # {"01": (date,url), ...}


def parse_price_excel(path):
    """1ファイルを読み、(完全一致辞書, 統一名辞書) を返す。

    薬価基準収載医薬品コードとYJコードは、銘柄別収載品では一致するが、
    統一名収載品では下3桁が異なる。統一名収載の行は
    「メーカー名」が空欄という規則性があるため、
    先頭9桁をキーにした辞書を別に作って取りこぼしを防ぐ。
    """
    import pandas as pd

    df = pd.read_excel(path, header=0, dtype=str)
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

    c_code  = pick("薬価基準収載医薬品コード")
    c_price = pick("薬価")
    c_maker = pick("メーカー名")
    c_exp   = pick("経過措置による使用期限", "経過措置")
    # 日本薬局方収載品は「規格」列の右隣（見出しの無い列）に「局」と入る
    # 注意: 変数 cols は pick() が参照する辞書なので上書きしないこと
    c_jp = None
    _order = list(df.columns)
    _spec = pick("規格")
    if _spec is not None and _spec in _order:
        _i = _order.index(_spec)
        if _i + 1 < len(_order):
            c_jp = _order[_i + 1]
    if c_code is None or c_price is None:
        raise ValueError(f"必要な列が見つかりません: {list(cols)[:12]}")

    exact, uni, expiry, jpharm = {}, {}, {}, set()
    for _, r in df.iterrows():
        code = str(r[c_code]).strip() if r[c_code] is not None else ""
        if not code or code == "nan" or len(code) < 9:
            continue
        try:
            price = float(str(r[c_price]).replace(",", "").strip())
        except (ValueError, TypeError):
            continue
        exact[code] = price
        maker = "" if c_maker is None else str(r[c_maker] or "").strip()
        if maker in ("", "nan"):
            uni.setdefault(code[:9], price)
        # 経過措置による使用期限（例: R9.3.31まで）
        if c_exp is not None:
            e = str(r[c_exp] or "").strip()
            if e and e != "nan":
                expiry[code] = e
        # 日本薬局方収載品（「局」の印）。
        # 統一名収載の行にも付くため、薬価と同じく先頭9桁でも引けるようにする。
        if c_jp is not None:
            v = str(r[c_jp] or "").strip()
            if v == "局":
                jpharm.add(code)
                if maker in ("", "nan"):
                    jpharm.add(code[:9])
    return exact, uni, expiry, jpharm


def main():
    check = "--check" in sys.argv
    try:
        log("薬価基準収載品目リストのページを確認中…")
        idx = decode_html(http_get(INDEX))
        found = find_list_page(idx)
        if not found:
            log("ERROR: 最新の収載品目リストページが見つかりません。")
            return 1
        page_date, page_url = found
        log(f"最新ページ: {page_url}")

        lst = decode_html(http_get(page_url))
        excels = find_excels(lst)
        if not excels:
            log("ERROR: Excelリンクが見つかりません。")
            return 1
        for seq in sorted(excels):
            log(f"  {KIND[seq]}: {excels[seq][1].rsplit('/',1)[-1]}")

        missing = [KIND[s] for s in KIND if s not in excels]
        if missing:
            log(f"警告: {', '.join(missing)} のExcelが見つかりません。取得できた分のみ反映します。")

        # 3ファイルのハッシュをまとめて比較する
        sig_src = "|".join(f"{s}:{excels[s][1]}" for s in sorted(excels))
        prev = {}
        if os.path.exists(OUT):
            try:
                prev = json.load(open(OUT, encoding="utf-8"))
            except Exception:
                prev = {}

        if check:
            log("--check のため、ダウンロードは行いません。")
            log(f"  前回の版: {prev.get('as_of','(未取得)')}")
            log(f"  今回の版: {max(excels[s][0] for s in excels)}")
            return 0

        blobs, digest = {}, hashlib.sha256()
        for seq in sorted(excels):
            b = http_get(excels[seq][1])
            blobs[seq] = b
            digest.update(hashlib.sha256(b).digest())
        h = digest.hexdigest()
        log(f"取得完了。統合ハッシュ={h[:16]}…")

        excel_date = max(excels[s][0] for s in excels)
        files_meta = {KIND[s]: {"date": excels[s][0],
                                "name": excels[s][1].rsplit("/", 1)[-1]}
                      for s in sorted(excels)}

        if prev.get("sha256") == h:
            # 中身は同じでも、記録している日付やファイル名が古い形式・古い値なら
            # そこだけ更新する（再解析はしないので処理は一瞬で終わる）
            if prev.get("as_of") != excel_date or prev.get("files") != files_meta:
                prev["as_of"] = excel_date
                prev["files"] = files_meta
                prev["updated_at"] = now_jst().isoformat(timespec="seconds")
                with open(OUT, "w", encoding="utf-8") as f:
                    json.dump(prev, f, ensure_ascii=False, separators=(",", ":"))
                log(f"変更なし。日付情報のみ更新しました（{excel_date}）")
                return 0
            log("変更なし（前回と同一）。処理を終了します。")
            return 10

        exact, uni, expiry, jpharm = {}, {}, {}, set()
        tmp = os.path.join(HERE, "_price_tmp.xlsx")
        for seq in sorted(blobs):
            with open(tmp, "wb") as f:
                f.write(blobs[seq])
            e, u, x, jp = parse_price_excel(tmp)
            exact.update(e)
            expiry.update(x)
            jpharm |= jp
            for k, v in u.items():
                uni.setdefault(k, v)
            log(f"  {KIND[seq]}: {len(e):,}件"
                f"（統一名 {len(u):,} / 経過措置 {len(x):,} / 局方 {len(jp):,}）")
        if os.path.exists(tmp):
            os.remove(tmp)

        data = {
            "as_of": excel_date,
            "files": files_meta,
            "sha256": h,
            "source": page_url,
            "updated_at": now_jst().isoformat(timespec="seconds"),
            "exact": exact,
            "uni": uni,
            "expiry": expiry,
            "jpharm": sorted(jpharm),
        }
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
        log(f"prices.json を更新（完全一致 {len(exact):,}件 / 統一名 {len(uni):,}件）")
        return 0

    except urllib.error.URLError as e:
        log(f"ERROR: 通信に失敗しました: {e}")
        return 1
    except Exception as e:
        import traceback; traceback.print_exc()
        log(f"ERROR: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
