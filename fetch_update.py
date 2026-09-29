#!/usr/bin/env python3
"""
医療用医薬品供給状況 自動更新スクリプト

厚労省ページを巡回し、掲載中Excelが前回と異なる場合のみ
ダウンロードして検索用HTMLを再生成する。

  python3 fetch_update.py            # 通常実行（変更時のみ更新）
  python3 fetch_update.py --local    # 厚労省へアクセスせず、保存済みExcelから再生成
  python3 fetch_update.py --force    # 変更がなくても強制再生成
  python3 fetch_update.py --check    # 確認のみ（DLも生成もしない）

終了コード: 0=更新した / 10=変更なし / 1=エラー
"""
import sys, os, re, json, hashlib, datetime, urllib.request, urllib.error
from html.parser import HTMLParser

# GitHub Actions は UTC で動くため、記録・ログは日本時間に揃える
JST = datetime.timezone(datetime.timedelta(hours=9), "JST")


def now_jst():
    return datetime.datetime.now(JST)

PAGE = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
BASE = "https://www.mhlw.go.jp"
UA   = "Mozilla/5.0 (compatible; supply-status-updater/1.0)"
HERE = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(HERE, "state.json")
XLSX  = os.path.join(HERE, "latest.xlsx")
SNAP  = os.path.join(HERE, "snapshot.json")   # 前版の出荷状況（悪化/改善判定用）
DISC  = os.path.join(HERE, "discontinued.txt") # 販売中止の手動登録
SENTEI= os.path.join(HERE, "sentei.json")    # 選定療養（fetch_sentei.py が作成）
KISO  = os.path.join(HERE, "kiso.json")       # 変更調剤可の基礎的医薬品（fetch_kiso.py が作成）
PRICES= os.path.join(HERE, "prices.json")     # 薬価（fetch_prices.py が作成。無ければ薬価なしで動く）
IPPAN = os.path.join(HERE, "ippanmei.json")   # 一般名処方マスタ（fetch_ippanmei.py が作成）
REGUL = os.path.join(HERE, "regulation.json") # 規制区分バッジ（医薬品コードマスタの公開版。手動で差し替え）
DOC   = os.path.join(HERE, "データの成り立ち.html")  # 凡例の2つ目のタブに埋め込む資料
OUT   = os.path.join(HERE, "医薬品供給状況_検索.html")

def log(m): print(f"[{now_jst():%Y-%m-%d %H:%M:%S}] {m}", flush=True)

NOCACHE = {"Cache-Control": "no-cache, max-age=0", "Pragma": "no-cache"}


def http_get(url, timeout=90, nocache=False):
    h = {"User-Agent": UA}
    if nocache:
        h.update(NOCACHE)
    req = urllib.request.Request(url, headers=h)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def http_exists(url, timeout=30):
    """そのURLにファイルがあるかを本体を取らずに確かめる。
    厚労省の配信は途中に中継が入り、掲載ページの方が古いまま
    見えることがあるため、ファイルの有無を直接確かめる用途で使う。
    戻り値: (ある/ない, バイト数または None)"""
    req = urllib.request.Request(url, headers={"User-Agent": UA, **NOCACHE},
                                 method="HEAD")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            if r.status != 200:
                return False, None
            try:
                return True, int(r.headers.get("Content-Length") or 0) or None
            except Exception:
                return True, None
    except urllib.error.HTTPError:
        return False, None
    except Exception as e:
        log(f"  （確認できませんでした: {url.rsplit('/', 1)[-1]} … {e}）")
        return False, None


# 厚労省のファイル名は「YYMMDDiyakuhinkyoukyu.xlsx」で規則的なため、
# 掲載ページに載る前でも日付から直接たどれる。
FNAME_FMT = "{ymd}iyakuhinkyoukyu.xlsx"
PROBE_DAYS = 14        # 何日分さかのぼって探すか
_FW = str.maketrans("0123456789", "０１２３４５６７８９")


def date_to_ymd(d):
    return f"{d.year % 100:02d}{d.month:02d}{d.day:02d}"


def make_label(iso):
    """2026-09-28 → 医療用医薬品供給状況（令和８年９月28日現在）
    掲載ページの書き方（年・月は全角、日は半角）に合わせる。"""
    try:
        y, m, d = (int(x) for x in iso.split("-"))
    except Exception:
        return "医療用医薬品供給状況"
    r = str(y - 2018).translate(_FW)
    return f"医療用医薬品供給状況（令和{r}年{str(m).translate(_FW)}月{d}日現在）"


def probe_newer(dir_url, base_iso, today=None):
    """掲載ページより新しい版がサーバに置かれていないか、
    日付を並べて直接確かめる。見つかった一番新しいものを返す。

    掲載ページは中継の都合で古いまま見えることがあり（実例：
    令和8年9月28日版はファイルが置かれているのに、
    ページのリンクは9月25日版のままだった）、
    ページだけを頼りにすると更新を取りこぼす。

    戻り値: (url, iso) または None
    """
    today = today or now_jst().date()
    try:
        base = datetime.date.fromisoformat(base_iso) if base_iso else None
    except Exception:
        base = None
    start = today - datetime.timedelta(days=PROBE_DAYS)
    if base and base >= start:
        start = base + datetime.timedelta(days=1)
    if start > today:
        return None
    found = None
    d = today
    while d >= start:
        url = dir_url.rstrip("/") + "/" + FNAME_FMT.format(ymd=date_to_ymd(d))
        ok, size = http_exists(url)
        if ok:
            found = (url, d.isoformat(), size)
            break                      # 新しい日から見ているので最初の1件が最新
        d -= datetime.timedelta(days=1)
    if not found:
        return None
    url, iso, size = found
    log(f"掲載ページより新しい版がサーバにありました: {iso}"
        + (f"（{size:,} bytes）" if size else ""))
    log(f"  URL : {url}")
    return url, iso

class LinkFinder(HTMLParser):
    """Collect <a href=...>text</a> pairs."""
    def __init__(self):
        super().__init__(); self.links=[]; self._href=None; self._buf=[]
    def handle_starttag(self, tag, attrs):
        if tag == "a":
            d = dict(attrs); self._href = d.get("href"); self._buf = []
    def handle_data(self, data):
        if self._href is not None: self._buf.append(data)
    def handle_endtag(self, tag):
        if tag == "a" and self._href is not None:
            self.links.append((self._href, "".join(self._buf).strip()))
            self._href = None; self._buf = []

def find_xlsx(html):
    """Return (url, label, yymmdd) for the供給状況 Excel link."""
    p = LinkFinder(); p.feed(html)
    cands = []
    for href, text in p.links:
        if not href or ".xlsx" not in href.lower():
            continue
        full = href if href.startswith("http") else BASE + href
        fname = full.rsplit("/", 1)[-1]
        # Primary signal: filename contains 'iyakuhinkyoukyu'
        score = 0
        if "iyakuhinkyoukyu" in fname.lower(): score += 100
        if "供給状況" in text: score += 50
        if "医療用医薬品" in text: score += 20
        m = re.search(r"(\d{6})iyakuhinkyoukyu", fname, re.I)
        ymd = m.group(1) if m else None
        if not ymd:
            m2 = re.search(r"(\d{6})", fname)
            ymd = m2.group(1) if m2 else None
        if score > 0:
            cands.append((score, ymd or "", full, text))
    if not cands:
        return None
    # Highest score, then newest date string
    cands.sort(key=lambda c: (c[0], c[1]), reverse=True)
    _, ymd, url, text = cands[0]
    return url, text, ymd

def load_state():
    if os.path.exists(STATE):
        try: return json.load(open(STATE, encoding="utf-8"))
        except Exception: pass
    return {}

def save_state(s):
    json.dump(s, open(STATE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

def ymd_to_iso(ymd):
    """260722 -> 2026-07-22 (YY is years since 2000)."""
    try:
        return f"20{ymd[0:2]}-{ymd[2:4]}-{ymd[4:6]}"
    except Exception:
        return datetime.date.today().isoformat()

def _load_keep_chg():
    """保存済みの悪化/改善の判定結果を読む。再生成で結果が消えないようにする。"""
    if not os.path.exists(SNAP):
        return None
    try:
        sj = json.load(open(SNAP, encoding="utf-8"))
        c = sj.get("chg")
        if c is None:
            return None
        return {k: int(v) for k, v in c.items()}
    except Exception:
        return None


def _load_keep_vdir():
    """保存済みの「出荷量の変化（改善/悪化）」を読む。
    再生成では比較元が今回の値に置き換わるため、これが無いと
    バッジの矢印がすべて「不変」になってしまう。"""
    if not os.path.exists(SNAP):
        return None
    try:
        v = json.load(open(SNAP, encoding="utf-8")).get("vdir")
        return {k: int(x) for k, x in v.items()} if v else None
    except Exception:
        return None


def _load_keep_newdel():
    """保存済みの「新たに薬価削除予定になった品目」を読む。"""
    if not os.path.exists(SNAP):
        return None
    try:
        v = json.load(open(SNAP, encoding="utf-8")).get("newdel")
        return {k: int(x) for k, x in v.items()} if v is not None else None
    except Exception:
        return None


def _load_keep_osc():
    """保存済みの「変化前の出荷状況」を読む。
    再生成では比較元が今回の値に置き換わってしまうため、
    これが無いと「供給停止 → 供給停止」のような表示になる。"""
    if not os.path.exists(SNAP):
        return None
    try:
        o = json.load(open(SNAP, encoding="utf-8")).get("osc")
        return {k: int(v) for k, v in o.items()} if o else None
    except Exception:
        return None


STALE_DAYS = 12   # 厚労省の公表間隔は週1回程度。これを超えたら知らせる


def warn_if_stale(as_of):
    """取り込み済みの版が古くなりすぎていないか。
    取得の仕組みが黙って止まっても気づけるようにする。"""
    if not as_of:
        return
    try:
        age = (now_jst().date() - datetime.date.fromisoformat(as_of)).days
    except Exception:
        return
    if age > STALE_DAYS:
        log(f"警告: 取り込み済みの版が {age} 日前（{as_of}）のままです。")
        print(f"::warning::供給状況が {age} 日間更新されていません（{as_of} 現在の版）。"
              "厚労省ページの掲載状況と取得の仕組みをご確認ください。")
        # ワークフローがこれを見てIssueを作る（ログだけでは気づけないため）
        try:
            with open(os.path.join(HERE, "_stale.md"), "w", encoding="utf-8") as f:
                f.write(
                    f"取り込み済みの供給状況が **{age} 日間** 更新されていません。\n\n"
                    f"- 掲載中として記録している版: **{as_of} 現在**\n"
                    f"- しきい値: {STALE_DAYS} 日\n\n"
                    "### 確認すること\n\n"
                    "1. 厚労省ページに新しい版が載っているか\n"
                    "   <https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/"
                    "kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html>\n"
                    "2. 載っているのに取り込めていない場合は、"
                    "ファイル名の付け方が変わった可能性があります"
                    "（`fetch_update.py` の `FNAME_FMT`）。\n"
                    "3. 厚労省側の公表が止まっているだけなら、対応は不要です。\n\n"
                    "※ 新しい版が取り込まれると、この通知は出なくなります。\n")
        except Exception:
            pass


def rebuild_only():
    """厚労省へアクセスせず、保存済みのExcelから作り直すだけ。
    表示の調整や販売中止の登録だけを反映したいときに使う。"""
    if not os.path.exists(XLSX):
        log("ERROR: 保存済みのExcelがありません。先に通常実行してください。")
        return 1
    st = load_state()
    as_of = st.get("as_of") or datetime.date.today().isoformat()
    log("ローカル再生成モード（厚労省へのアクセスはしません）")
    log(f"  使用するExcel: {XLSX}")
    log(f"  基準日: {as_of}")

    import build_html
    prev = None
    if os.path.exists(SNAP):
        try:
            sj = json.load(open(SNAP, encoding="utf-8"))
            prev = {k: int(v) for k, v in sj.get("sc", {}).items()}
        except Exception:
            prev = None

    # 再生成では悪化/改善の基準を動かさない（スナップショットは書き換えない）
    n = build_html.build(XLSX, OUT, as_of=as_of,
                         source_label=st.get("label", ""), source_url=st.get("url", ""),
                         prev_snapshot=prev, snapshot_out=None, snapshot_path=SNAP,
                         keep_chg=_load_keep_chg(), keep_osc=_load_keep_osc(),
                         keep_vdir=_load_keep_vdir(),
                         keep_newdel=_load_keep_newdel(),
                         prices_path=PRICES, kiso_path=KISO, disc_path=DISC,
                         sentei_path=SENTEI, ippanmei_path=IPPAN,
                         datadoc_path=DOC, regulation_path=REGUL)
    log(f"生成完了: {OUT}（{n:,}品目 / {as_of} 現在）")
    return 0


def main():
    force = "--force" in sys.argv
    check = "--check" in sys.argv
    if "--local" in sys.argv:
        return rebuild_only()
    try:
        log("厚労省ページを確認中…")
        html = http_get(PAGE, nocache=True).decode("utf-8", "replace")
        found = find_xlsx(html)
        if not found:
            log("ERROR: Excelリンクが見つかりません。ページ構造が変わった可能性があります。")
            return 1
        url, label, ymd = found
        log(f"掲載中: {label}")
        log(f"  URL : {url}")

        st = load_state()
        prev_url  = st.get("url")
        prev_hash = st.get("sha256")
        prev_as_of = st.get("as_of")

        # 掲載ページのリンクが古いまま見えることがあるため、
        # 日付から直接、新しい版が置かれていないかも確かめる。
        page_as_of = ymd_to_iso(ymd) if ymd else None
        base_iso = max([x for x in (page_as_of, prev_as_of) if x], default=None)
        probed = probe_newer(url.rsplit("/", 1)[0], base_iso)
        if probed:
            url, new_iso = probed
            ymd = date_to_ymd(datetime.date.fromisoformat(new_iso))
            label = make_label(new_iso)
            print("::notice::掲載ページに載る前の版を直接取得しました"
                  f"（ページ {page_as_of or '?'} → 実際 {new_iso}）。")

        # 掲載ページのリンクだけが古いまま見えている場合。
        # 直接取得で先に新しい版を入れていると、翌日以降ページが
        # 追いつくまでこの状態になる。取り込み済みのファイルが
        # サーバにまだあるなら、取り下げではなくページが遅れているだけ。
        this_as_of = ymd_to_iso(ymd) if ymd else None
        stale_page = False
        if (not probed and prev_as_of and this_as_of and this_as_of < prev_as_of
                and prev_url):
            stale_page, _ = http_exists(prev_url)
        if stale_page:
            log(f"掲載ページのリンクは {this_as_of} のままですが、"
                f"取り込み済みの {prev_as_of} 版はサーバに残っています。")
            log("  ページの反映が遅れているだけと判断し、そのまま維持します。")
            if not check:
                st["last_checked"] = now_jst().isoformat(timespec="seconds")
                st["last_seen_label"] = label
                save_state(st)
                warn_if_stale(prev_as_of)
                return 10

        # 掲載中の版が前回より古くなっていないか。
        # 厚労省は、公表した版を取り下げて前の版に戻すことがある
        # （実例：令和8年9月17日版が一時掲載され、のち9月16日版に戻った）。
        # 黙って古い版で上書きすると、誰も差し替えに気づけないため知らせる。
        rollback = False
        if not stale_page and prev_as_of and this_as_of and this_as_of < prev_as_of:
            rollback = True
            log(f"警告: 掲載中の版が前回より古くなっています"
                f"（前回 {prev_as_of} → 今回 {this_as_of}）。")
            log("  厚労省側で差し替え・取り下げがあった可能性があります。")
            log("  掲載中の版をそのまま取り込みます（厚労省の現在の公表内容に合わせます）。")
            print("::warning::供給状況の掲載版が前回より古くなりました"
                  f"（{prev_as_of} → {this_as_of}）。厚労省側の差し替えの可能性があります。")

        if check:
            log("--check のため、ダウンロードは行いません。")
            if prev_url:
                log(f"  前回取得: {st.get('label','?')}（{st.get('updated_at','?')}）")
                if prev_url == url:
                    log("  同一URLです。")
                elif stale_page:
                    log("  ★掲載ページの反映が遅れています（取り込み済みの版を維持）")
                elif probed:
                    log("  ★掲載ページに載る前の新しい版があります")
                elif rollback:
                    log("  ★前回より古い版に差し替わっています")
                else:
                    log("  ★URLが変わっています（更新の可能性）")
            else:
                log("  未実行の状態です（state.json なし）。")
            log("リンク取得は正常です。")
            return 0

        log("ファイルを取得中…")
        blob = http_get(url, nocache=True)
        h = hashlib.sha256(blob).hexdigest()
        log(f"  size={len(blob):,} bytes  sha256={h[:16]}…")

        # 中継がエラーページを返すことがあるため、Excelであることを確かめる。
        # 壊れたものを保存すると、次回まで古い内容が残り続ける。
        if len(blob) < 100_000 or blob[:2] != b"PK":
            log("ERROR: 取得したファイルがExcelとして読めません"
                f"（先頭 {blob[:8]!r} / {len(blob):,} bytes）。今回は更新しません。")
            return 1

        # 供給Excelが前回と同一なら、悪化/改善の比較基準を動かしてはいけない。
        # （--force は薬価更新などによる再生成のためのもので、
        #   同じExcelでスナップショットを取り直すと変化が消えてしまう）
        same_excel = (h == prev_hash)

        if not force and same_excel:
            # 変更が無い日も「最終確認時刻」を残す（定期実行の自動停止対策）
            st["last_checked"] = now_jst().isoformat(timespec="seconds")
            st["last_seen_label"] = label
            save_state(st)
            log("変更なし（前回と同一ファイル）。処理を終了します。")
            log(f"  掲載中の版: {st.get('as_of','?')}")
            warn_if_stale(st.get("as_of"))
            return 10

        if prev_hash:
            log(f"変更を検出しました（前回 {prev_hash[:16]}… → 今回 {h[:16]}…）")
        else:
            log("初回実行です。")

        with open(XLSX, "wb") as f:
            f.write(blob)
        log(f"保存: {XLSX}")

        # Regenerate the HTML
        import build_html
        as_of = ymd_to_iso(ymd)

        # 前版のスナップショットを読み、状況の悪化/改善を判定する
        prev = None
        if os.path.exists(SNAP):
            try:
                sj = json.load(open(SNAP, encoding="utf-8"))
                prev = {k: int(v) for k, v in sj.get("sc", {}).items()}
                log(f"前版スナップショットを読込: {sj.get('date','?')}（{len(prev):,}品目）")
            except Exception as e:
                log(f"警告: スナップショットを読めませんでした（{e}）。今回は変化判定を行いません。")
                prev = None
        else:
            log("スナップショット未作成。今回は変化判定を行わず、次回以降有効になります。")

        # Excelが変わっていない再生成では、スナップショットを書き換えない
        snap_out = None if same_excel else SNAP
        keep = None
        keep_o = None          # 先に定義しておく（同一Excelでない経路で未定義になるため）
        keep_v = None
        keep_nd = None
        if same_excel:
            # Excelが同じなら比較し直さず、前回の判定結果をそのまま使う
            keep = _load_keep_chg()
            keep_o = _load_keep_osc()
            keep_v = _load_keep_vdir()
            keep_nd = _load_keep_newdel()
            log("  供給Excelは前回と同一のため、前回の変化判定を引き継ぎます")

        n = build_html.build(XLSX, OUT, as_of=as_of, source_label=label, source_url=url,
                             prev_snapshot=prev, snapshot_out=snap_out, snapshot_path=SNAP,
                             keep_chg=keep, keep_osc=keep_o, keep_vdir=keep_v,
                             keep_newdel=keep_nd,
                             prices_path=PRICES, kiso_path=KISO,
                             disc_path=DISC, sentei_path=SENTEI,
                             ippanmei_path=IPPAN, datadoc_path=DOC,
                             regulation_path=REGUL)
        log(f"生成完了: {OUT}（{n:,}品目 / {as_of} 現在）")

        now = now_jst().isoformat(timespec="seconds")
        new_state = {"url": url, "sha256": h, "label": label,
                     "as_of": as_of, "items": n,
                     "updated_at": now, "last_checked": now}
        if rollback:
            # いつ・どの版からどの版へ戻されたかを残す。
            # あとから経緯をたどれるようにするため、履歴として積む。
            hist = st.get("rollbacks") or []
            hist.append({"at": now, "from": prev_as_of, "to": as_of})
            new_state["rollbacks"] = hist[-20:]
        elif st.get("rollbacks"):
            new_state["rollbacks"] = st["rollbacks"]
        save_state(new_state)
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
