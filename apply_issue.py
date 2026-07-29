#!/usr/bin/env python3
"""
GitHub Issue の本文を discontinued.txt に取り込む。

Issue に DrugShortage.JP からコピーした文章を貼って作成すると、
ワークフローがこのスクリプトを呼び、貼り付け欄に追記する。

  python3 apply_issue.py <本文が入ったファイル> [Issue番号]

終了コード: 0=追記した / 10=追記する内容が無かった / 1=エラー
"""
import sys, os, re, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
DISC = os.path.join(HERE, "discontinued.txt")
MARK_BEGIN = "# ▼▼▼ ここに貼り付け ▼▼▼"
MARK_END = "# ▲▲▲ ここまで ▲▲▲"


def log(m):
    print(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {m}", flush=True)


def clean_body(text):
    """Issue本文から余計なものを取り除く。

    - Markdownの見出し・引用・コードフェンス記号
    - テンプレートの説明文（<!-- --> のコメント）
    - チェックボックス行
    """
    text = re.sub(r"<!--.*?-->", "", text, flags=re.S)
    # Issueテンプレートが自動で付ける見出し・案内文は解析対象から外す
    DROP = ("貼り付け欄", "販売中止の登録", "_No response_", "### ")
    out = []
    for ln in text.splitlines():
        if ln.strip() in ("貼り付け欄", "_No response_"):
            continue
        t = ln.rstrip()
        s = t.strip()
        if s in ("```", "~~~") or s.startswith("```"):
            continue
        if re.match(r"^\s*[-*]\s*\[[ xX]\]", s):     # チェックボックス
            continue
        s2 = re.sub(r"^\s*(#{1,6}\s*|>\s*)", "", t)  # 見出し・引用
        if s2.strip() in ("貼り付け欄", "販売中止の登録"):
            continue
        out.append(s2)
    # 前後の空行を落とす
    while out and not out[0].strip():
        out.pop(0)
    while out and not out[-1].strip():
        out.pop()
    return "\n".join(out)


def _load_name_index():
    """供給状況Excelから 品名→YJコード の索引を作る。
    latest.xlsx が無い場合は None を返し、照合を省く。"""
    xlsx = os.path.join(HERE, "latest.xlsx")
    if not os.path.exists(xlsx):
        return None
    try:
        import pandas as pd, build_html
        hdr = build_html.find_header_row(xlsx)
        df = pd.read_excel(xlsx, header=hdr, dtype=str)
        c = list(df.columns)
        idx = {}
        for _, r in df.iterrows():
            nm = build_html.clean(r[c[5]])
            yj = build_html.clean(r[c[4]])
            if nm and yj:
                idx.setdefault(build_html.norm_name(nm), []).append(yj)
        return idx
    except Exception:
        return None


def main():
    if len(sys.argv) < 2:
        log("ERROR: 本文ファイルを指定してください")
        return 1
    src = sys.argv[1]
    issue_no = sys.argv[2] if len(sys.argv) > 2 else ""

    try:
        body = clean_body(open(src, encoding="utf-8").read())
    except Exception as e:
        log(f"ERROR: 本文を読めませんでした: {e}")
        return 1

    if not body.strip():
        log("追記する内容がありません。")
        return 10

    # 本番と同じ解析器で読み取り、実在する薬剤かどうかまで確かめる。
    # 実在しない名前を追記すると、以後ずっと警告が出続けるため。
    try:
        sys.path.insert(0, HERE)
        import build_html
        recs = build_html.parse_discontinued_text(body)
        names = _load_name_index()
    except Exception as e:
        log(f"警告: 事前チェックに失敗しました（{e}）。そのまま追記します。")
        recs, names = None, None

    if recs is not None:
        if not recs:
            log("ERROR: 薬剤名を読み取れませんでした。貼り付け範囲をご確認ください。")
            return 1
        ok, ng = [], []
        for r in recs:
            nm = r["name"]
            if names is None:
                ok.append(r)
                continue
            key = build_html.norm_name(nm)
            hits = names.get(key)
            if not hits:
                ng.append((nm, "該当する薬剤がありません"))
            elif len(hits) > 1:
                ng.append((nm, f"同名が{len(hits)}件あります → {', '.join(hits)}"))
            else:
                ok.append(r)
        for nm, why in ng:
            log(f"  × {nm}：{why}")
        if not ok:
            log("ERROR: 登録できる薬剤がありませんでした。貼り付け内容をご確認ください。")
            return 1
        log(f"{len(ok)}件を読み取りました:")
        for r in ok:
            log(f"  ・{r['name']}"
                + (f"／{r['p']}" if r.get("p") else "")
                + (f"（実施 {r['d']}）" if r.get("d") else ""))
        if ng:
            log(f"（{len(ng)}件は登録できないため、そのまま追記します。"
                f"不要なら該当行を削除してください）")

    if not os.path.exists(DISC):
        log(f"ERROR: {DISC} がありません")
        return 1

    lines = open(DISC, encoding="utf-8").read().splitlines()
    try:
        idx = next(i for i, l in enumerate(lines) if l.strip() == MARK_BEGIN)
    except StopIteration:
        # 目印が無ければ末尾に足す
        idx = len(lines) - 1
        log("警告: 貼り付け位置の目印が見つかりません。末尾に追記します。")

    stamp = f"# --- Issue #{issue_no} より取込（{datetime.date.today()}） ---" \
            if issue_no else f"# --- 取込（{datetime.date.today()}） ---"
    add = ["", stamp] + body.splitlines()

    lines[idx + 1:idx + 1] = add
    with open(DISC, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    log(f"discontinued.txt に {len(body.splitlines())} 行を追記しました。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
