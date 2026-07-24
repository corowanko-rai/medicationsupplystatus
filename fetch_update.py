#!/usr/bin/env python3
"""
医療用医薬品供給状況 自動更新スクリプト

厚労省ページを巡回し、掲載中Excelが前回と異なる場合のみ
ダウンロードして検索用HTMLを再生成する。

  python3 fetch_update.py            # 通常実行（変更時のみ更新）
  python3 fetch_update.py --force    # 変更がなくても強制再生成
  python3 fetch_update.py --check    # 確認のみ（DLも生成もしない）

終了コード: 0=更新した / 10=変更なし / 1=エラー
"""
import sys, os, re, json, hashlib, datetime, urllib.request, urllib.error
from html.parser import HTMLParser

PAGE = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
BASE = "https://www.mhlw.go.jp"
UA   = "Mozilla/5.0 (compatible; supply-status-updater/1.0)"
HERE = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(HERE, "state.json")
XLSX  = os.path.join(HERE, "latest.xlsx")
OUT   = os.path.join(HERE, "医薬品供給状況_検索.html")

def log(m): print(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {m}", flush=True)

def http_get(url, timeout=90):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()

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

def main():
    force = "--force" in sys.argv
    check = "--check" in sys.argv
    try:
        log("厚労省ページを確認中…")
        html = http_get(PAGE).decode("utf-8", "replace")
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

        if check:
            log("--check のため、ダウンロードは行いません。")
            if prev_url:
                log(f"  前回取得: {st.get('label','?')}（{st.get('updated_at','?')}）")
                log("  同一URLです。" if prev_url == url else "  ★URLが変わっています（更新の可能性）")
            else:
                log("  未実行の状態です（state.json なし）。")
            log("リンク取得は正常です。")
            return 0

        log("ファイルを取得中…")
        blob = http_get(url)
        h = hashlib.sha256(blob).hexdigest()
        log(f"  size={len(blob):,} bytes  sha256={h[:16]}…")

        if not force and h == prev_hash:
            log("変更なし（前回と同一ファイル）。処理を終了します。")
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
        n = build_html.build(XLSX, OUT, as_of=as_of, source_label=label, source_url=url)
        log(f"生成完了: {OUT}（{n:,}品目 / {as_of} 現在）")

        save_state({"url": url, "sha256": h, "label": label,
                    "as_of": as_of, "items": n,
                    "updated_at": datetime.datetime.now().isoformat(timespec="seconds")})
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
