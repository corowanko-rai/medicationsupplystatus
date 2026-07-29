#!/usr/bin/env python3
"""
基礎的医薬品のうち「後発医薬品と同様に変更調剤が認められるもの」の一覧を取得する。

厚労省「薬価基準収載品目リスト」ページ内の
  基礎的リスト（Excel版）= /xls/tpYYYY_kiso.xlsx
を取得し、品名の集合を kiso.json に保存する。

このExcelにはYJコード・薬価基準収載医薬品コードが無く、
品名・成分・規格単位・メーカー名・薬価のみが載っている。
そのため供給状況Excelの⑥品名と突き合わせる。

  python3 fetch_kiso.py           # 変更があれば更新
  python3 fetch_kiso.py --check   # 確認のみ

終了コード: 0=更新した / 10=変更なし / 1=エラー

失敗しても供給状況ページの生成は続行できる設計。
"""
import sys, os, re, json, hashlib, datetime, urllib.request, urllib.error

import fetch_prices as fp   # ページ探索とHTTPの共通処理を再利用

HERE = os.path.dirname(os.path.abspath(__file__))
OUT  = os.path.join(HERE, "kiso.json")


def log(m):
    print(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {m}", flush=True)


def find_kiso_xlsx(html):
    """基礎的リスト（Excel版）のURLを返す。/xls/tpYYYY_kiso.xlsx 形式。"""
    p = fp.Links(); p.feed(html)
    best = None
    for h in p.hrefs:
        full = fp.absolutize(h)
        if not full:
            continue
        m = re.search(r"/xls/tp(\d{4})_kiso\.xlsx$", full, re.I)
        if m:
            year = m.group(1)
            if best is None or year > best[0]:
                best = (year, full)
    return best


def parse_kiso_excel(path):
    """品名の集合と、参考情報（区分ごとの件数）を返す。"""
    import pandas as pd

    df = pd.read_excel(path, header=None, dtype=str)

    # 見出し行を探す（「品名」を含む行）
    hdr = None
    for i in range(min(12, len(df))):
        joined = "".join(str(x) for x in df.iloc[i].tolist())
        if "品名" in joined:
            hdr = i
            break
    if hdr is None:
        raise ValueError("見出し行（品名）が見つかりません。様式変更の可能性があります。")

    df = pd.read_excel(path, header=hdr, dtype=str)
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

    c_name = pick("品名")
    c_kind = pick("内注外歯区分", "区分")
    if c_name is None:
        raise ValueError(f"「品名」列が見つかりません: {list(cols)[:10]}")

    names, kinds = set(), {}
    for _, r in df.iterrows():
        nm = str(r[c_name] or "").strip()
        if not nm or nm == "nan" or nm == "品名":
            continue
        names.add(nm)
        if c_kind is not None:
            k = str(r[c_kind] or "").strip()
            if k and k != "nan":
                kinds[k] = kinds.get(k, 0) + 1
    return names, kinds


def main():
    check = "--check" in sys.argv
    try:
        log("薬価基準収載品目リストのページを確認中…")
        idx = fp.decode_html(fp.http_get(fp.INDEX))
        found = fp.find_list_page(idx)
        if not found:
            log("ERROR: 収載品目リストのページが見つかりません。")
            return 1
        _, page_url = found

        lst = fp.decode_html(fp.http_get(page_url))
        got = find_kiso_xlsx(lst)
        if not got:
            log("ERROR: 基礎的リスト（Excel版）のリンクが見つかりません。")
            return 1
        year, url = got
        log(f"基礎的リスト: {url.rsplit('/',1)[-1]}")

        prev = {}
        if os.path.exists(OUT):
            try:
                prev = json.load(open(OUT, encoding="utf-8"))
            except Exception:
                prev = {}

        if check:
            log("--check のため、ダウンロードは行いません。")
            log(f"  前回: {prev.get('as_of','(未取得)')} / 今回: {year}")
            return 0

        blob = fp.http_get(url)
        h = hashlib.sha256(blob).hexdigest()
        log(f"取得完了 {len(blob):,} bytes / sha256={h[:16]}…")

        fname = url.rsplit("/", 1)[-1]
        if prev.get("sha256") == h:
            # 中身が同じでも記録が古い形式なら、ファイル名などだけ更新する
            if prev.get("file") != fname or prev.get("as_of") != year:
                prev["file"] = fname
                prev["as_of"] = year
                prev["updated_at"] = datetime.datetime.now().isoformat(timespec="seconds")
                with open(OUT, "w", encoding="utf-8") as f:
                    json.dump(prev, f, ensure_ascii=False, separators=(",", ":"))
                log(f"変更なし。ファイル情報のみ更新しました（{fname}）")
                return 0
            log("変更なし（前回と同一）。処理を終了します。")
            return 10

        tmp = os.path.join(HERE, "_kiso_tmp.xlsx")
        with open(tmp, "wb") as f:
            f.write(blob)
        names, kinds = parse_kiso_excel(tmp)
        os.remove(tmp)

        data = {
            "as_of": year,
            "file": fname,
            "sha256": h,
            "source": url,
            "updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "names": sorted(names),
        }
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
        log(f"kiso.json を更新（{len(names):,}品名） 内訳={kinds}")
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
