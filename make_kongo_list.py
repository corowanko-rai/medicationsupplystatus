#!/usr/bin/env python3
"""
軟膏混合チェック（別のアプリ）が毎日読みに行く「外用薬の製品一覧」を作る。

  入力  gaiyo_latest.xlsx … fetch_prices.py が保存した「3．外用薬」のExcel
        latest.xlsx       … 医療用医薬品供給状況のExcel
  出力  kongo-list.json          … 製品の一覧（薬価は含めない）
        kongo-list-version.json  … 版の情報（アプリはまずこちらを見る）

厚労省へはアクセスしない。すでに取得済みのファイルから作るだけ。
一覧の内容が同じなら何も書き換えない（無駄なコミットを増やさない）。

終了コード  0 = 更新した  10 = 変更なし  1 = 異常（前の一覧をそのまま残す）

名前の整形と剤形の分類は、アプリの public/importers.js と同じ規則にしてある。
片方だけ変えると、同じ製品が別の行に分かれる原因になるので注意。
"""
import hashlib
import json
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta, timezone

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
LIST_XLSX = os.path.join(HERE, "gaiyo_latest.xlsx")
SUPPLY_XLSX = os.path.join(HERE, "latest.xlsx")
OUT = os.path.join(HERE, "kongo-list.json")
OUT_VER = os.path.join(HERE, "kongo-list-version.json")
PRICES = os.path.join(HERE, "prices.json")
STATE = os.path.join(HERE, "state.json")

JST = timezone(timedelta(hours=9), "JST")
CODE_RE = re.compile(r"^[0-9A-Z]{12}$")
EXCLUDE_NAME = re.compile(
    r"眼|点鼻|鼻腔|口腔|耳科|吸入|坐剤|坐薬|腟|膣|注腸|貼付|テープ|パップ|シャンプー|ウォッシュ|外用液|スプレー|エアゾール|エアゾル|フォーム|うがい|消毒|ゼリー|歯科|用液|点耳"
)
EXCLUDE_YJ = re.compile(r"^(131|132|2399|2559|247|261|226|225|229)")
SUPPLY_TYPE = {"1": "先発", "4": "先発", "3": "準先発", "2": "後発", "5": "その他"}
# 件数がこの範囲を外れたら、様式が変わった可能性があるので一覧を更新しない
MIN_LIST, MAX_LIST = 1500, 3000
MIN_SUPPLY = 1000
MAX_CHANGE = 0.20


def log(m):
    print(m, flush=True)


def text(v):
    return "" if v is None else str(v)


def classify(name, code):
    n = unicodedata.normalize("NFKC", str(name))
    if EXCLUDE_NAME.search(n) or EXCLUDE_YJ.search(code or ""):
        return "その他"
    if re.search(r"プロペト|ワセリン「|^白色ワセリン|吸水クリーム|親水クリーム|親水軟膏|単軟膏$", n) and not re.search(r"サリチル酸|亜鉛華", n):
        return "基剤"
    if re.search(r"軟膏|パスタ|ペースト|^エキザルベ", n) and "ローション" not in n:
        return "軟膏"
    if "クリーム" in n:
        return "クリーム"
    if "ゲル" in n:
        return "ゲル"
    if re.search(r"ローション|ゾル|リニメント", n):
        return "ローション"
    return "その他"


def display_name(name):
    s = unicodedata.normalize("NFKC", str(name))
    s = re.sub(r"\s+", " ", s).strip()
    return re.sub(r" [0-9,]+(mg|単位)$", "", s)


def key_of(name):
    return re.sub(r"[\s「」・]", "", display_name(name).lower())


def list_type(sen, gen):
    sen, gen = text(sen), text(gen)
    if "準先発" in sen:
        return "準先発"
    if "先発" in sen:
        return "先発"
    return "後発" if gen else ""


def read_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()


def read_list():
    rows = read_rows(LIST_XLSX)
    hi = head = None
    for i, r in enumerate(rows[:30]):
        h = [re.sub(r"\s", "", text(x)) for x in r]
        if "薬価基準収載医薬品コード" in h and "品名" in h:
            hi, head = i, h
            break
    if hi is None:
        raise ValueError("外用薬のExcelに必要な列が見つかりません")

    def col(n):
        return head.index(n) if n in head else -1

    c_code, c_name, c_ing, c_spec = col("薬価基準収載医薬品コード"), col("品名"), col("成分名"), col("規格")
    c_maker, c_kind = col("メーカー名"), col("区分")
    c_sen, c_gen = col("先発医薬品"), col("診療報酬において加算等の算定対象となる後発医薬品")
    c_exp = -1
    for i, h in enumerate(head):
        if "経過措置" in h:
            c_exp = i

    def cell(r, c):
        return r[c] if 0 <= c < len(r) else None

    items, seen = [], set()
    for r in rows[hi + 1 :]:
        code = text(cell(r, c_code)).strip()
        name = re.sub(r"^\s*※\s*", "", text(cell(r, c_name))).strip()
        if not code and not name:
            continue
        kind = text(cell(r, c_kind))
        if c_kind >= 0 and kind and "外用" not in kind:
            continue
        if not CODE_RE.match(code) or not name or code in seen:
            continue
        seen.add(code)
        items.append(
            {
                "code": code,
                "name": display_name(name),
                "key": key_of(name),
                "ingredient": text(cell(r, c_ing)).strip()[:120],
                "spec": text(cell(r, c_spec)).strip()[:40],
                "maker": text(cell(r, c_maker)).strip()[:60],
                "expiry": text(cell(r, c_exp)).strip()[:20],
                "form": classify(name, code),
                "type": list_type(cell(r, c_sen), cell(r, c_gen)),
            }
        )
    return items


def read_supply():
    rows = read_rows(SUPPLY_XLSX)
    hi = None
    for i, r in enumerate(rows[:20]):
        if any(text(x).startswith("⑤YJ") for x in r):
            hi = i
            break
    if hi is None:
        raise ValueError("供給状況のExcelに必要な列が見つかりません")
    head = [text(x) for x in rows[hi]]

    def col(mark):
        for i, h in enumerate(head):
            if h.startswith(mark):
                return i
        raise ValueError(f"供給状況のExcelに {mark} の列がありません")

    c_kind, c_ing, c_spec, c_yj, c_name, c_maker, c_type = (col(m) for m in "①③④⑤⑥⑦⑧")
    out = []
    for r in rows[hi + 1 :]:
        if len(r) <= c_type:
            continue
        yj = text(r[c_yj]).strip()
        if "外用" not in text(r[c_kind]) or not CODE_RE.match(yj):
            continue
        out.append(
            {
                "yj": yj, "ingredient": text(r[c_ing]), "spec": text(r[c_spec]), "name": text(r[c_name]),
                "maker": text(r[c_maker]), "kind": text(r[c_type])[:1],
            }
        )
    return out


def build():
    items = read_list()
    supply = read_supply()
    if not MIN_LIST <= len(items) <= MAX_LIST:
        raise ValueError(f"外用薬リストの件数が想定外です: {len(items)}")
    if len(supply) < MIN_SUPPLY:
        raise ValueError(f"供給状況の外用薬の件数が想定外です: {len(supply)}")
    by_yj = {s["yj"]: s for s in supply}
    for it in items:
        s = by_yj.get(it["code"])
        if s:
            t = SUPPLY_TYPE.get(s["kind"])
            if t:
                it["type"] = t
            if not it["maker"] and s["maker"]:
                it["maker"] = s["maker"]
    have = {it["code"] for it in items}
    for s in supply:
        if s["yj"] in have or s["kind"] == "6":  # 6 = 未収載は入れない
            continue
        items.append(
            {
                "code": s["yj"], "name": display_name(s["name"]), "key": key_of(s["name"]),
                "ingredient": s["ingredient"][:120], "spec": s["spec"][:40], "maker": s["maker"][:60],
                "expiry": "", "form": classify(s["name"], s["yj"]), "type": SUPPLY_TYPE.get(s["kind"], ""), "source": "supply",
            }
        )
    return items


def load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def main():
    try:
        items = build()
    except Exception as e:  # 様式変更などでは前の一覧を残す
        log(f"::warning::軟膏混合チェック用の一覧を作れませんでした（前の一覧を残します）: {e}")
        return 1

    prev = load_json(OUT_VER)
    if prev.get("count") and abs(len(items) - prev["count"]) > prev["count"] * MAX_CHANGE:
        log(f"::warning::製品数が大きく変わったため一覧を更新しません（{prev['count']} → {len(items)}）")
        return 1

    body = json.dumps(items, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    sha = hashlib.sha256(body).hexdigest()
    if prev.get("sha256") == sha and os.path.exists(OUT):
        log("軟膏混合チェック用の一覧に変更はありません。")
        return 10

    files = load_json(PRICES).get("files", {})
    list_file = files.get("外用薬", {}).get("name", "")
    label = load_json(STATE).get("label", "")
    with open(OUT, "wb") as f:
        f.write(body)
    with open(OUT_VER, "w", encoding="utf-8") as f:
        json.dump(
            {
                "sha256": sha, "count": len(items), "list_file": list_file, "supply_label": label,
                "generated_at": datetime.now(JST).isoformat(timespec="seconds"),
            },
            f, ensure_ascii=False, indent=1,
        )
    log(f"軟膏混合チェック用の一覧を更新しました（{len(items):,}品目）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
