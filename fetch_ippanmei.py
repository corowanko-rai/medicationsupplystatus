#!/usr/bin/env python3
"""
処方箋に記載する一般名処方の標準的な記載（一般名処方マスタ）を取得する。

厚労省「医療保険が適用される医薬品について」（固定URL、fetch_prices.py と同じ）から
  /iryouhoken/shohosen_YYYYMM.html
を辿り、そのページにある
  /dl/ippanmeishohoumaster_YYMMDD.xlsx         … 通常の一般名処方マスタ
  /dl/ippannmeishohoumaster_bs_YYMMDD(…).xlsx  … バイオ後続品の一般名処方マスタ
を取得して ippanmei.json を作る。

  python3 fetch_ippanmei.py           # 変更があれば更新
  python3 fetch_ippanmei.py --check   # 確認のみ

終了コード: 0=更新した / 10=変更なし / 1=エラー

■ 突合の考え方（重要）
一般名コードは「薬価基準収載医薬品コードの上9桁＋ZZZ」で作られる。
したがって供給状況データのYJコードは、上9桁で一般名に紐付けられる。

ただし上9桁では適切に区分できない成分・規格があり、それらは
9桁目をアルファベットに置き換えた「例外コード」になっている
（例：【般】チモロール点眼液０．２５％（持続性）＝1319702QAZZZ）。
例外コードは上9桁では絶対に当たらないため、マスタに同梱されている
「例外コード品目対照表」から12桁の完全一致表を別に作る。
順序としては例外（12桁）を先に引き、無ければ9桁で引く。

■ 過去分の引き継ぎ（重要）
マスタは版が変わると品目が減ることがある。
（例：R7.12.5版 1,381件 → R8.6.12版 1,230件。
  アムロジピン錠２．５ｍｇ・５ｍｇ などが現行版から消えている）

厚労省の説明では、削除されるのは「一般名処方加算の対象ではなくなったため」
であって、その記載での処方や調剤ができなくなるわけではない。
古い記載の処方箋も現に来るため、消さずに残す。

そこで、ページに並んでいる過去分も含めて古い順に読み、
  ・同じ一般名コードがあれば新しい版で上書きする
  ・新しい版に無いものは、古い版のものを残す
という積み上げ方式にしている。前回作った ippanmei.json の内容も
同じ規則で引き継ぐので、厚労省が過去分のリンクを消しても失われない。

引き継いだものは cur=0 とし、画面では「旧版」と表示して
加算区分は出さない。一般名処方加算の対象と誤解させないため。

失敗しても供給状況ページの生成は続行できる設計。
"""
import sys, os, re, json, hashlib, datetime, urllib.error

import fetch_prices as fp   # ページ探索とHTTPの共通処理を再利用

JST = datetime.timezone(datetime.timedelta(hours=9), "JST")


def now_jst():
    return datetime.datetime.now(JST)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "ippanmei.json")

KIDX = {"内用薬": 0, "外用薬": 1, "注射薬": 2}
FIELDS = ("c", "t", "k", "i", "s", "a", "p", "b", "bs", "x", "v", "cur")


def log(m):
    print(f"[{now_jst():%Y-%m-%d %H:%M:%S}] {m}", flush=True)


def find_page(html):
    """一般名処方マスタのページURLを返す。/iryouhoken/shohosen_NNNNNN.html 形式のうち
    数字が最大のもの（過去分へのリンクも同じ書式で並んでいるため）。"""
    p = fp.Links(); p.feed(html)
    best = None
    for h in p.hrefs:
        full = fp.absolutize(h)
        if not full:
            continue
        m = re.search(r"/iryouhoken/shohosen_(\d{6})\.html$", full)
        if m and (best is None or m.group(1) > best[0]):
            best = (m.group(1), full)
    return best


def find_masters(html):
    """{"通常": [(日付, URL), …], "バイオ": […]} を日付の昇順で返す。

    ファイル名は
      ippanmeishohoumaster_260612.xlsx
      ippannmeishohoumaster_bs_260601(260520).xlsx
    のように「n」の数と括弧書きが揺れるため、まとめて拾って区別する。
    「ippannmeishohou_sakujo_…」（削除リスト）は master を含まないので当たらない。
    過去分も同じ書式で並ぶので、ここでは全部を集めて古い順に積み上げる。
    """
    p = fp.Links(); p.feed(html)
    got = {"通常": {}, "バイオ": {}}
    for h in p.hrefs:
        full = fp.absolutize(h)
        if not full:
            continue
        m = re.search(r"/ippann?meishohoumaster_(bs_)?(\d{6})[^/]*\.xlsx$",
                      full, re.I)
        if not m:
            continue
        key = "バイオ" if m.group(1) else "通常"
        got[key][m.group(2)] = full          # 同じ日付が2つあれば後勝ち
    return {k: [(d, v[d]) for d in sorted(v)] for k, v in got.items() if v}


def _pick(cols, *names):
    for n in names:
        if n in cols:
            return cols[n]
    for n in names:
        for k, v in cols.items():
            if n in k:
                return v
    return None


def _find_header(df, must):
    for i in range(min(10, len(df))):
        joined = "".join(str(x) for x in df.iloc[i].tolist())
        if must in joined:
            return i
    return None


def parse_master(path, is_bs, date):
    """(一般名コード→項目 の辞書, 例外の {YJコード: 一般名コード}) を返す。"""
    import pandas as pd

    probe = pd.read_excel(path, sheet_name=0, header=None, dtype=str)
    hdr = _find_header(probe, "一般名コード")
    if hdr is None:
        raise ValueError("見出し行（一般名コード）が見つかりません。様式変更の可能性があります。")

    df = pd.read_excel(path, sheet_name=0, header=hdr, dtype=str)
    cols = {str(c).strip(): c for c in df.columns}
    c_code = _pick(cols, "一般名コード")
    c_text = _pick(cols, "一般名処方の標準的な記載", "標準的な記載")
    c_kind = _pick(cols, "区分")
    c_ing = _pick(cols, "成分名")
    c_spec = _pick(cols, "規格")
    c_add = _pick(cols, "一般名処方加算対象", "加算対象")
    c_exc = _pick(cols, "例外コード")
    c_low = _pick(cols, "同一剤形・規格内の最低薬価", "最低薬価")
    c_note = _pick(cols, "備考")
    if c_code is None or c_text is None:
        raise ValueError(f"必要な列が見つかりません: {list(cols)[:9]}")

    def s(r, c):
        if c is None:
            return ""
        v = str(r[c] or "").strip()
        return "" if v == "nan" else v

    items = {}
    for _, r in df.iterrows():
        code = s(r, c_code).upper()
        if not re.fullmatch(r"[0-9A-Z]{9}ZZZ", code):
            continue
        text = s(r, c_text)
        if not text:
            continue
        add = s(r, c_add)
        try:
            low = round(float(s(r, c_low).replace(",", "")), 2)
        except ValueError:
            low = None
        items[code] = {
            "c": code,
            "t": text,
            "k": KIDX.get(s(r, c_kind), 0),
            "i": s(r, c_ing),
            "s": s(r, c_spec),
            "a": 12 if "2" in add else (1 if "1" in add else 0),
            "p": low,
            "b": s(r, c_note),
            "bs": 1 if is_bs else 0,
            "x": 1 if s(r, c_exc) else 0,
            "v": date,      # この版に載っていた
            "cur": 0,       # 現行版かどうかは呼び出し側で決める
        }

    # ---- 例外コード品目対照表 ----
    mapx = {}
    try:
        ex = pd.read_excel(path, sheet_name="例外コード品目対照表",
                           header=None, dtype=str)
    except Exception:
        ex = None
    if ex is not None:
        h2 = _find_header(ex, "薬価基準収載医薬品コード")
        if h2 is None:
            raise ValueError("例外コード品目対照表の見出し行が見つかりません。")
        ex = pd.read_excel(path, sheet_name="例外コード品目対照表",
                           header=h2, dtype=str)
        ec = {str(c).strip(): c for c in ex.columns}
        e_code = _pick(ec, "一般名コード")
        e_yj = _pick(ec, "薬価基準収載医薬品コード")
        if e_code is None or e_yj is None:
            raise ValueError(f"対照表の列が見つかりません: {list(ec)[:8]}")
        cur = None
        for _, r in ex.iterrows():
            # 一般名コードのセルは縦結合されており、続きの行は空欄になる
            c = str(r[e_code] or "").strip().upper()
            if re.fullmatch(r"[0-9A-Z]{9}ZZZ", c):
                cur = c
            yj = str(r[e_yj] or "").strip().upper()
            if cur is None or not re.fullmatch(r"[0-9A-Z]{12}", yj):
                continue
            if cur in items:
                mapx[yj] = cur

    missing = sorted({c for c, it in items.items() if it["x"]}
                     - set(mapx.values()))
    if missing:
        log(f"  警告: 対照表に載っていない例外コードが {len(missing)} 件あります"
            f"（{', '.join(missing[:3])}…）。該当する一般名は品目が出ません。")
    return items, mapx


def prev_as_codes(prev):
    """前回の ippanmei.json を「コード基準」に戻す。
    項目の並び順は版ごとに変わるので、必ずコードで持ち直す。"""
    items, mapx = {}, {}
    try:
        arr = prev.get("items") or []
        for it in arr:
            c = it.get("c")
            if not c:
                continue
            d = {f: it.get(f) for f in FIELDS}
            d["c"] = c
            if not d.get("v"):
                d["v"] = prev.get("as_of", "")
            d["cur"] = 0          # 引き継ぎ分はいったん旧版に戻す
            items[c] = d
        for yj, n in (prev.get("mapx") or {}).items():
            if isinstance(n, int) and 0 <= n < len(arr):
                c = arr[n].get("c")
                if c:
                    mapx[yj] = c
    except Exception:
        return {}, {}
    return items, mapx


def main():
    check = "--check" in sys.argv
    try:
        log("一般名処方マスタのページを確認中…")
        idx = fp.decode_html(fp.http_get(fp.INDEX))
        found = find_page(idx)
        if not found:
            log("ERROR: 一般名処方マスタのページが見つかりません。")
            return 1
        _, page_url = found
        log(f"最新ページ: {page_url}")

        lst = fp.decode_html(fp.http_get(page_url))
        masters = find_masters(lst)
        if "通常" not in masters:
            log("ERROR: 一般名処方マスタのExcelが見つかりません。")
            return 1
        for k in sorted(masters):
            names = ", ".join(u.rsplit("/", 1)[-1] for _, u in masters[k])
            log(f"  {k}（{len(masters[k])}版）: {names}")
        if "バイオ" not in masters:
            log("警告: バイオ後続品の一般名処方マスタが見つかりません。"
                "通常分のみ反映します。")

        prev = {}
        if os.path.exists(OUT):
            try:
                prev = json.load(open(OUT, encoding="utf-8"))
            except Exception:
                prev = {}

        newest = {k: v[-1][0] for k, v in masters.items()}
        as_of = max(newest.values())

        if check:
            log("--check のため、ダウンロードは行いません。")
            log(f"  前回: {prev.get('as_of', '(未取得)')} / 今回: {as_of}")
            return 0

        blobs, digest = {}, hashlib.sha256()
        for k in sorted(masters):
            for date, url in masters[k]:
                b = fp.http_get(url)
                blobs[(k, date)] = b
                digest.update(hashlib.sha256(b).digest())
        h = digest.hexdigest()
        log(f"取得完了（{len(blobs)}ファイル）。統合ハッシュ={h[:16]}…")

        files_meta = {k: {"date": masters[k][-1][0],
                          "name": masters[k][-1][1].rsplit("/", 1)[-1],
                          "past": [d for d, _ in masters[k][:-1]]}
                      for k in sorted(masters)}

        REQUIRED = ("items", "map9", "mapx", "files", "as_of")
        missing = [k for k in REQUIRED if k not in prev]
        if prev.get("sha256") == h and not missing:
            if prev.get("as_of") != as_of or prev.get("files") != files_meta:
                prev["as_of"] = as_of
                prev["files"] = files_meta
                prev["updated_at"] = now_jst().isoformat(timespec="seconds")
                with open(OUT, "w", encoding="utf-8") as f:
                    json.dump(prev, f, ensure_ascii=False, separators=(",", ":"))
                log(f"変更なし。日付情報のみ更新しました（{as_of}）")
                return 0
            log("変更なし（前回と同一）。処理を終了します。")
            return 10
        if prev.get("sha256") == h and missing:
            log(f"内容は同じですが記録に不足があります（{', '.join(missing)}）。作り直します。")

        # ---- 古い順に積み上げる ----
        # 起点は前回の ippanmei.json（厚労省が過去分のリンクを消しても
        # 一度取り込んだ記載が失われないようにするため）
        items, mapx_c = prev_as_codes(prev)
        if items:
            log(f"前回の記録から {len(items):,}件を引き継ぎます。")

        tmp = os.path.join(HERE, "_ippanmei_tmp.xlsx")
        for k in sorted(masters):
            for date, _ in masters[k]:
                with open(tmp, "wb") as f:
                    f.write(blobs[(k, date)])
                it, mx = parse_master(tmp, is_bs=(k == "バイオ"), date=date)
                is_newest = (date == newest[k])
                added = sum(1 for c in it if c not in items)
                for c, v in it.items():
                    v["cur"] = 1 if is_newest else 0
                    items[c] = v          # 同じコードは新しい版で上書き
                mapx_c.update(mx)
                log(f"  {k} {date}版: {len(it):,}件"
                    f"（新規 {added:,} / 例外 {len(mx):,}）"
                    + ("  ← 現行版" if is_newest else ""))
        if os.path.exists(tmp):
            os.remove(tmp)

        old = [c for c, v in items.items() if not v["cur"]]
        if old:
            log(f"現行版に無い記載を {len(old):,}件、旧版として残します。")

        # ---- 索引を作り直す ----
        order = sorted(items, key=lambda c: (items[c]["bs"], c))
        arr = [{f: items[c].get(f) for f in FIELDS} for c in order]
        pos = {c: n for n, c in enumerate(order)}

        # 9桁の索引は「現行版を優先」。旧版のコードは、
        # 同じ9桁を現行版が使っていない場合だけ採る。
        map9 = {}
        for c in order:
            if items[c]["x"]:
                continue
            k9 = c[:9]
            if k9 not in map9 or (items[c]["cur"] and not arr[map9[k9]]["cur"]):
                map9[k9] = pos[c]
        mapx = {yj: pos[c] for yj, c in mapx_c.items() if c in pos}

        clash = [yj for yj in mapx if yj[:9] in map9]
        if clash:
            log(f"  注意: 例外品目 {len(clash)} 件が9桁でも当たります。例外を優先します。")

        with open(OUT, "w", encoding="utf-8") as f:
            json.dump({
                "as_of": as_of,
                "files": files_meta,
                "sha256": h,
                "source": page_url,
                "updated_at": now_jst().isoformat(timespec="seconds"),
                "items": arr,
                "map9": map9,
                "mapx": mapx,
            }, f, ensure_ascii=False, separators=(",", ":"))
        cur_n = sum(1 for v in arr if v["cur"])
        log(f"ippanmei.json を更新（一般名 {len(arr):,}件"
            f" / 現行版 {cur_n:,} ・旧版 {len(arr) - cur_n:,}）")
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
