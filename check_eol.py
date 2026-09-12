#!/usr/bin/env python3
"""
保守期限の見張り番。

「保守期限.txt」に書いた期限を毎朝確認し、
残り90日を切ったものと、期限を過ぎたものを知らせる。

  python3 check_eol.py                # 確認して結果を表示する
  python3 check_eol.py --days 120     # 何日前から知らせるかを変える
  python3 check_eol.py --today 2027-09-01   # 動作確認用に今日の日付を偽装する

終了コード:
   0 = 期限が近いものは無い
  20 = 期限が近い／過ぎたものがある（ワークフローがIssueを作る合図）
   1 = エラー

■ なぜ「失敗」にしないのか
このスクリプトが失敗扱いになると、毎朝の更新が赤くなり続けてしまい、
本当の異常に気づけなくなる。そこで通知は Issue で行い、
ワークフロー自体は成功のまま通す設計にしている。
"""
import sys, os, datetime

JST = datetime.timezone(datetime.timedelta(hours=9), "JST")
HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DAYS = 90


def resolve(path):
    """日本語ファイル名を、濁点の表し方（NFC/NFD）の違いを越えて見つける。
    macOSからアップロードするとNFDで登録されるため。"""
    import unicodedata
    if os.path.exists(path):
        return path
    folder = os.path.dirname(path) or "."
    want = unicodedata.normalize("NFC", os.path.basename(path))
    try:
        for n in os.listdir(folder):
            if unicodedata.normalize("NFC", n) == want:
                return os.path.join(folder, n)
    except OSError:
        pass
    return None


# ---- 期限を自分で調べる仕組み ------------------------------------
# 手で書き写した日付は、版を上げたときに直し忘れて必ず古くなる。
# 規則で決まっているものは、設定ファイルから読んで計算する。

WORKFLOW = os.path.join(HERE, ".github", "workflows", "update.yml")


def _read_python_version():
    """update.yml の python-version を読む。見つからなければ None。"""
    import re
    f = resolve(WORKFLOW)
    if not f:
        return None
    try:
        t = open(f, encoding="utf-8").read()
    except OSError:
        return None
    m = re.search(r"""python-version:\s*['"]?(\d+)\.(\d+)""", t)
    return f"{m.group(1)}.{m.group(2)}" if m else None


def auto_python():
    """使っているPythonの保守終了日を求める。

    Pythonは毎年1版が出て、5年間保守される（PEP 602）。
    保守終了は「リリース年＋5」の10月末で、
    3.9〜3.14 の実日付6件と一致することを確認済み。
        3.x の終了年 = 2016 + x
    版を上げれば、この日付も自動で追随する。
    """
    v = _read_python_version()
    if not v:
        return None, "Python（版が読めません）", "update.yml の python-version を確認する"
    minor = int(v.split(".")[1])
    return (datetime.date(2016 + minor, 10, 31),
            f"Python {v} の保守終了",
            "次の版へ上げる。pandas・openpyxl の対応も確認する")


AUTO = {"python": auto_python}


def parse(path):
    """登録ファイルを読む。戻り値は (期限, 名前, 場所, 対処, URL) のリスト。
    期限が読めない行は 期限=None（＝日付を調べてもらう対象）。"""
    items, bad = [], []
    with open(path, encoding="utf-8") as f:
        for n, raw in enumerate(f, 1):
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            cols = [c.strip() for c in line.split("|")]
            while len(cols) < 5:
                cols.append("")
            due = None
            if cols[0].startswith("自動:"):
                # 期限を自分で調べる行。名前と対処も上書きする
                key = cols[0].split(":", 1)[1].strip()
                fn = AUTO.get(key)
                if not fn:
                    bad.append((n, line))
                    continue
                due, name, what = fn()
                items.append((due, name, cols[2], what, cols[4]))
                continue
            if cols[0] not in ("未確認", ""):
                try:
                    due = datetime.date.fromisoformat(cols[0])
                except ValueError:
                    bad.append((n, line))
                    continue
            items.append((due, cols[1], cols[2], cols[3], cols[4]))
    return items, bad


def main():
    args = sys.argv[1:]
    days = DEFAULT_DAYS
    if "--days" in args:
        days = int(args[args.index("--days") + 1])
    today = datetime.datetime.now(JST).date()
    if "--today" in args:
        today = datetime.date.fromisoformat(args[args.index("--today") + 1])

    path = resolve(os.path.join(HERE, "保守期限.txt"))
    if not path:
        print("保守期限.txt がありません。確認するものはありません。")
        return 0

    items, bad = parse(path)
    for n, line in bad:
        print(f"::warning::保守期限.txt の {n} 行目が読めません（期限は YYYY-MM-DD）: {line}")

    over, soon, unknown, later = [], [], [], []
    for due, name, where, what, url in items:
        if due is None:
            unknown.append((name, where, what, url))
        elif due < today:
            over.append(((due - today).days, due, name, where, what, url))
        elif (due - today).days <= days:
            soon.append(((due - today).days, due, name, where, what, url))
        else:
            later.append(((due - today).days, due, name, where, what, url))

    print(f"今日 {today} 時点の保守期限（{days}日前から通知）")
    print(f"  期限切れ {len(over)} / 期限が近い {len(soon)} / "
          f"日付が未確認 {len(unknown)} / まだ先 {len(later)}")

    lines = []
    for d, due, name, where, what, url in sorted(over):
        msg = f"【期限切れ {-d}日経過】{due} {name}／{where}／{what}"
        print("  " + msg)
        print(f"::error::{msg}")
        lines.append("- **" + msg + "**" + (f"　{url}" if url else ""))
    for d, due, name, where, what, url in sorted(soon):
        msg = f"【あと{d}日】{due} {name}／{where}／{what}"
        print("  " + msg)
        print(f"::warning::{msg}")
        lines.append("- " + msg + (f"　{url}" if url else ""))
    for name, where, what, url in unknown:
        msg = f"【日付が未確認】{name}／{where}／{what}"
        print("  " + msg)
        lines.append("- " + msg + (f"　{url}" if url else ""))

    if later:
        print("  --- まだ先のもの ---")
        for d, due, name, where, what, _ in sorted(later):
            print(f"  （あと{d}日）{due} {name}")

    # ワークフローがIssueの本文に使う
    out = os.environ.get("GITHUB_OUTPUT")
    body_path = os.path.join(HERE, "_eol_body.md")
    if over or soon or unknown:
        with open(body_path, "w", encoding="utf-8") as f:
            f.write(f"{today} 時点で、対応が必要な保守期限があります。\n\n")
            f.write("\n".join(lines))
            f.write("\n\n---\n片付いたら `保守期限.txt` の該当行を消すか "
                    "`#` を付けて、このIssueを閉じてください。\n"
                    "行が残っている限り、毎朝ここに載り続けます。\n")
        if out:
            with open(out, "a", encoding="utf-8") as f:
                f.write("hit=1\n")
        return 20
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write("hit=0\n")
    print("  対応が必要なものはありません。")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except BrokenPipeError:
        # head などで出力を途中で打ち切られた場合。異常ではない
        try:
            sys.stdout.close()
        except Exception:
            pass
        sys.exit(0)
    except Exception as e:
        import traceback; traceback.print_exc()
        print(f"ERROR: {e}")
        sys.exit(1)
