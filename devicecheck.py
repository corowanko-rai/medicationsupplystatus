#!/usr/bin/env python3
"""
医薬品供給状況ページ 端末横断チェック

Android（Pixel / Galaxy）と iOS（iPhone）の実機定義で、
表示崩れ・機能・タップ領域を検証する。

  python3 devicecheck.py <検索HTMLのパス>
"""
import sys, os, re
from playwright.sync_api import sync_playwright

# 実機定義名（Playwright組み込み）。狭い順に並べる。
DEVICES = [
    "Galaxy S9+",      # 320  Android 最狭クラス
    "Pixel 4",         # 353  Android
    "Galaxy S24",      # 360  Android 現行
    "Pixel 5",         # 393  Android
    "Pixel 3",         # 393  Android  dsf=2.75
    "Pixel 2",         # 411  Android  dsf=2.625
    "Pixel 4a (5G)",   # 412  Android
    "Pixel 7",         # 412  Android 現行
    "Pixel 2 XL",      # 411  Android  dsf=3.5（最も高精細）
    "Galaxy A55",      # 480  Android 大画面
    "iPhone 12 Mini",  # 375  iOS
    "iPhone 15",        # 393  iOS
    "iPhone 15 Pro Max",# 430  iOS
]

# (検索語, 期待件数)
SEARCHES = [
    ("カルボシステインＤＳ５０％", 3),   # 全角記号
    ("カルボシステインDS50%",     3),   # 半角
    ("ニゾラールクリーム２％",     1),
    ("ニゾラール",               2),
    ("カロナール",              11),
    ("ろきそ",                 110),   # ひらがな
    ("ロキソ",                 110),
]

# 「データの成り立ち」に必ず説明があるべき語。
# 機能を足したら、ここにも1語足すこと。資料の更新漏れを検知するための一覧。
DOC_TERMS = [
    "一般名処方",     # 【般】一般名モード
    "例外コード",     # 一般名コードの例外
    "併売品",         # 併売バッジと一覧
    "選定療養",       # 計算タブ
    "変更調剤",       # 基／変更可
    "経過措置",       # 使用期限
    "販売中止",       # 手動登録
]

MIN_TAP = 40   # タップ領域の最低px（Androidの推奨48dp、iOSの44ptを踏まえた実務下限）


def num(t):
    t = t.strip()
    return 0 if "該当なし" in t else int(t.split("件")[0].strip().replace(",", ""))


def check_device(browser, p, name, url):
    dev = p.devices[name]
    ctx = browser.new_context(**dev)
    pg = ctx.new_page()
    errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    fails = []

    pg.goto(url)
    pg.wait_for_timeout(3000)

    def ovf():
        return pg.evaluate(
            "document.documentElement.scrollWidth-document.documentElement.clientWidth")

    if ovf() != 0:
        fails.append(f"初期表示で横溢れ {ovf()}px")

    # 検索
    for q, exp in SEARCHES:
        pg.fill("#q", q)
        pg.wait_for_timeout(380)
        got = num(pg.inner_text("#cnt"))
        if got != exp:
            fails.append(f'"{q}" {got}件≠{exp}件')
    pg.fill("#q", "")
    pg.wait_for_timeout(250)

    # 成分名モード
    pg.click('.sm[data-m="i"]')
    pg.wait_for_timeout(250)
    pg.fill("#q", "アセトアミノフェン")
    pg.wait_for_timeout(420)
    if num(pg.inner_text("#cnt")) != 73:
        fails.append("成分名検索が73件でない")
    pg.fill("#q", "")
    pg.click('.sm[data-m="n"]')
    pg.wait_for_timeout(250)

    # 状況フィルタ
    pg.click('.chip[data-f="normal"]')
    pg.wait_for_timeout(380)
    want = pg.evaluate("() => R.filter(r => r[8] === 0).length")
    if num(pg.inner_text("#cnt")) != want:
        fails.append("通常出荷の件数がデータと合わない（表示 %s / 期待 %d）"
                     % (pg.inner_text("#cnt"), want))
    if ovf() != 0:
        fails.append(f"チップ選択で横溢れ {ovf()}px")
    pg.click("#clrall")
    pg.wait_for_timeout(250)

    # 製品区分
    pg.click("#advbtn")
    pg.wait_for_timeout(200)
    # 併売品：バッジ・カード内一覧・一覧画面
    pg.fill("#q", "コンスタン０．４")
    pg.wait_for_timeout(420)
    if pg.locator(".card").count():
        if pg.locator(".card .cmb").count() == 0:
            fails.append("併売バッジが出ない")
        pg.locator(".card").first.click()
        pg.wait_for_timeout(350)
        if pg.locator(".card.open .cmitem").count() == 0:
            fails.append("カード内に併売品が出ない")
        if ovf() != 0:
            fails.append(f"併売品の表示で横溢れ {ovf()}px")
    pg.fill("#q", "")
    pg.wait_for_timeout(250)

    cb = pg.locator(".chip.cm")
    if cb.count() == 0:
        fails.append("併売品ボタンが無い")
    else:
        cb.click()
        pg.wait_for_timeout(600)
        if not pg.locator("#cmv").is_visible():
            fails.append("併売品一覧が開かない")
        if pg.locator(".cmvg").count() == 0:
            fails.append("併売品一覧が空")
        if ovf() != 0:
            fails.append(f"併売品一覧で横溢れ {ovf()}px")
        # 剤形の切り替え
        for k in ["注射薬", "注射薬"]:
            kb = pg.locator(f'#cmvk .kb2[data-k="{k}"]')
            if kb.count() == 0:
                fails.append("併売品一覧に剤形ボタンが無い")
                break
            kb.click()
            pg.wait_for_timeout(300)
            if ovf() != 0:
                fails.append(f"併売品の剤形切替で横溢れ {ovf()}px")
                break
        for f in ["split", "all"]:
            fb = pg.locator(f'#cmvf .sortb[data-f="{f}"]')
            if fb.count():
                fb.click()
                pg.wait_for_timeout(300)
                if ovf() != 0:
                    fails.append(f"併売品の絞り込みで横溢れ {ovf()}px")
                    break
        pg.locator(".cmvg").first.locator("summary").click()
        pg.wait_for_timeout(300)
        if ovf() != 0:
            fails.append(f"併売品を開いた時に横溢れ {ovf()}px")
        pg.click("#cmvx")
        pg.wait_for_timeout(300)
        if pg.locator("#cmv").is_visible():
            fails.append("併売品一覧が閉じない")

    # 並び替えは状況チップと同じ行の右端にあること。
    # CSSが壊れると行が分かれ、一覧の表示領域が40px削られる。
    lay = pg.evaluate("""() => {
      const q=(s)=>document.querySelector(s).getBoundingClientRect();
      const rw=q('.chiprow'), c=q('#chips'), b=q('#cobtn');
      return {row: Math.round(rw.height),
              same: Math.abs(c.top - b.top) < 2,
              disp: getComputedStyle(document.querySelector('.chiprow')).display};
    }""")
    if lay["disp"] != "flex":
        fails.append("chiprow の display が flex でない（%s／CSSの記述ミスの可能性）" % lay["disp"])
    if not lay["same"]:
        fails.append("並び替えがチップと別の行にある")
    if lay["row"] > 60:
        fails.append("チップ行が高すぎる %dpx" % lay["row"])

    # データの鮮度：一般名処方マスタの表記
    pg.click("#lgbtn")
    pg.wait_for_timeout(400)
    lg = pg.inner_text("#lgbody")
    if "一般名処方マスタ（過去分）" in lg:
        fails.append("現行版のマスタに（過去分）が付いている")
    if "一般名処方マスタ" in lg and not re.search(r"20\d\d/\d\d/\d\d", lg):
        fails.append("マスタの日付が YYYY/MM/DD になっていない")

    # 凡例の2枚目（データの成り立ち）
    dtab = pg.locator('.lgtb[data-p="doc"]')
    if dtab.count() and dtab.is_visible():
        dtab.click()
        pg.wait_for_timeout(400)
        if not pg.locator("#lgdoc").is_visible():
            fails.append("データの成り立ちが表示されない")
        if pg.locator("#lgbody").is_visible():
            fails.append("2枚目を開いてもバッジの説明が残る")
        secs = ["doc-venn", "doc-map", "doc-join", "doc-calc"]
        if pg.locator("#lgdoc .tab").count() != len(secs):
            fails.append("データの成り立ちの切り替えが4つでない")
        # 4つ目まで画面内に収まっていること（隠れると押せない）
        over = pg.evaluate("""() => {
          const t = document.querySelector('#lgdoc .tabs');
          return t.scrollWidth - t.clientWidth;
        }""")
        if over > 0:
            fails.append(f"データの成り立ちの切り替えが画面に収まらない（{over}px）")
        for k in secs:
            pg.click(f'#lgdoc .tab[data-t="{k}"]')
            pg.wait_for_timeout(260)
            vis = pg.evaluate(
                "() => [...document.querySelectorAll('#lgdoc section')]"
                ".filter(s => s.offsetParent).map(s => s.id)")
            if vis != [k]:
                fails.append(f"データの成り立ち「{k}」の切り替えが効かない（{vis}）")
            if ovf() != 0:
                fails.append(f"データの成り立ち「{k}」で横溢れ {ovf()}px")
        # 資料とページの数字がずれていないか。
        # 「データの成り立ち」は実測値を載せているので、機能を足したのに
        # 資料を直し忘れると、ここで食い違いが出る。
        # 非表示の章も含めて読む（inner_text は表示中の章しか返さない）
        doc = pg.evaluate(
            "() => document.getElementById('lgdoc').textContent")
        # 目印が置換されずに残っていないか（自動更新の失敗を検知）
        if "{{" in doc:
            import re as _re
            left = sorted(set(_re.findall(r"\{\{\w+\}\}", doc)))[:3]
            fails.append("資料に未置換の目印が残っている：%s" % "／".join(left))

        want = pg.evaluate("() => R.length.toLocaleString()")
        # 「全16,393品目（…版）での実測値です」の宣言文そのものを見る。
        # 文書のどこかに同じ数字があれば通る、という緩い判定にすると
        # 冒頭だけ古いまま残っていても気づけない。
        m = re.search(r"全\s*([\d,]+)\s*品目", doc)
        if not m:
            fails.append("資料に「全◯◯品目」の記載が無い")
        elif m.group(1) != want:
            fails.append("資料の総品目数がページと違う（資料 %s / ページ %s）"
                         % (m.group(1), want))
        # 2つのコードを同一視していないか。
        # YJコード（個別医薬品コード）と薬価基準収載医薬品コードは、
        # 銘柄別収載品では一致するが統一名収載品では下3桁が異なる。
        # 「＝」で結ぶ書き方は誤りなので、資料と凡例の両方で禁止する。
        badge = pg.evaluate("() => document.getElementById('lgbody').textContent")
        for where, text in (("資料", doc), ("凡例", badge)):
            for ng in ("薬価基準収載医薬品コード（＝YJコード）",
                       "薬価基準収載医薬品コード＝YJコード",
                       "YJコード（＝薬価基準収載医薬品コード）",
                       "YJコード＝薬価基準収載医薬品コード"):
                if ng in text:
                    fails.append("%sで2つのコードを同一視している：%s" % (where, ng))
        if DOC_TERMS:
            missing = [t for t in DOC_TERMS if t not in doc]
            if missing:
                fails.append("資料に説明が無い機能：%s" % "／".join(missing))
        pg.click('.lgtb[data-p="badge"]')
        pg.wait_for_timeout(300)
        if not pg.locator("#lgbody").is_visible():
            fails.append("1枚目に戻れない")
    else:
        fails.append("データの成り立ちのタブが無い")

    pg.click("#lgx")
    pg.wait_for_timeout(300)

    # 一般名処方（【般】モード）
    gb = pg.locator('.sm[data-m="g"]')
    if gb.count() and gb.is_visible():
        gb.click()
        pg.wait_for_timeout(600)
        if not pg.locator("#genbar").is_visible():
            fails.append("一般名モードで専用の絞り込みが出ない")
        if pg.locator("#chiprow").is_visible():
            fails.append("一般名モードで状況チップが残る")
        if ovf() != 0:
            fails.append(f"一般名モードで横溢れ {ovf()}px")
        want = pg.evaluate("() => GEN.filter(g => g[1] === 0 || g[1] === 1).length")
        if num(pg.inner_text("#cnt")) != want:
            fails.append("一般名（内用+外用）の件数がデータと合わない（期待 %d）" % want)
        # 現行マスタから外れた記載の扱い
        if pg.locator("#gob").count() == 0:
            fails.append("旧版の切り替えが無い")
        else:
            pg.click("#gob")
            pg.wait_for_timeout(500)
            want = pg.evaluate(
                "() => GEN.filter(g => (g[1]===0||g[1]===1) && !g[9]).length")
            if num(pg.inner_text("#cnt")) != want:
                fails.append("旧版を外したときの件数が合わない（期待 %d）" % want)
            pg.click("#gob")
            pg.wait_for_timeout(500)
        pg.fill("#q", "アムロジピン錠")
        pg.wait_for_timeout(450)
        if num(pg.inner_text("#cnt")) != 3:
            fails.append("アムロジピン錠の一般名が3件でない")
        if pg.locator(".gcard .gold").count() != 2:
            fails.append("旧版バッジが2件でない")
        # 旧版は一般名処方加算の対象外なので、加算バッジを出してはいけない
        for i in range(pg.locator(".gcard").count()):
            c = pg.locator(".gcard").nth(i)
            if c.locator(".gold").count() and c.locator(".gadd").count():
                fails.append("旧版に加算バッジが出ている")
                break
        pg.fill("#q", "")
        pg.wait_for_timeout(350)
        # 例外コードの品目が正しく出るか（持続性製剤の取り違え防止）
        pg.fill("#q", "チモロール点眼液０．２５％（持続性）")
        pg.wait_for_timeout(450)
        if pg.locator(".gcard").count() != 1:
            fails.append("チモロール点眼液０．２５％（持続性）が1件でない")
        else:
            pg.locator(".gcard").first.click()
            pg.wait_for_timeout(350)
            items = pg.locator(".gcard.open .cmitem").count()
            if items == 0:
                fails.append("持続性チモロールに品目が1つも出ない")
            # 例外コードの要は「持続性でないものを混ぜないこと」。
            # 品目数は供給状況データで増減するので、中身で判定する。
            names = pg.locator(".gcard.open .cmnm").all_inner_texts()
            bad = [t for t in names if "チモロール" in t or "チモプトール" in t
                   or "リズモン" in t]
            wrong = [t for t in bad if "ＸＥ" not in t and "ＴＧ" not in t]
            if wrong:
                fails.append("持続性でない製剤が混ざっている：%s" % "／".join(wrong))
            if ovf() != 0:
                fails.append(f"一般名の展開で横溢れ {ovf()}px")
        # 「口腔内崩壊錠」を OD でも引けること
        pg.fill("#q", "アムロジピンOD錠")
        pg.wait_for_timeout(450)
        if num(pg.inner_text("#cnt")) != 3:
            fails.append("OD錠での言い換え検索が効かない")
        pg.fill("#q", "")
        pg.wait_for_timeout(350)
        # 通常出荷が無いものへの絞り込み
        pg.click('.gfb[data-g="risk"]')
        pg.wait_for_timeout(500)
        want = pg.evaluate("""() => {
          let n = 0;
          GEN.forEach((g, i) => {
            if (g[1] !== 0 && g[1] !== 1) return;
            const st = gStat(i);
            if (st[3] > 0 && st[0] === 0) n++;
          });
          return n;
        }""")
        if num(pg.inner_text("#cnt")) != want:
            fails.append("「通常出荷が無い」の件数が合わない（期待 %d）" % want)
        if ovf() != 0:
            fails.append(f"一般名の絞り込みで横溢れ {ovf()}px")
        pg.click('.gfb[data-g="all"]')
        pg.wait_for_timeout(400)
        pg.click('.sm[data-m="n"]')
        pg.wait_for_timeout(400)
        if not pg.locator("#chiprow").is_visible():
            fails.append("品名モードに戻すと状況チップが復活しない")
        if pg.locator("#genbar").is_visible():
            fails.append("品名モードで一般名の絞り込みが残る")
    else:
        fails.append("一般名モードのボタンが無い")

    # 選定療養の表示
    pg.fill("#q", "ムコダインシロップ")
    pg.wait_for_timeout(420)
    if pg.locator(".card").count():
        if pg.locator(".card .sen").count() == 0:
            fails.append("選バッジが出ない")
        pg.locator(".card").first.click()
        pg.wait_for_timeout(350)
        if pg.locator(".card.open .senbox").count() == 0:
            fails.append("選定療養の金額が出ない")
        if ovf() != 0:
            fails.append(f"選定療養の表示で横溢れ {ovf()}px")
    pg.fill("#q", "")
    pg.wait_for_timeout(250)

    # 選定療養：後発品との差額（自己負担割合の切替）
    pg.fill("#q", "ヒルドイドソフト軟膏")
    pg.wait_for_timeout(420)
    if pg.locator(".card").count():
        pg.locator(".card").first.click()
        pg.wait_for_timeout(350)
        diff = pg.locator(".card.open .sendiff")
        if diff.count() == 0:
            fails.append("後発品との差額ブロックが出ない")
        else:
            before = diff.locator(".sdresult").inner_text()
            diff.locator('.sdb[data-r="0.1"]').click()
            pg.wait_for_timeout(300)
            if pg.locator(".card.open").count() == 0:
                fails.append("割合ボタンを押すとカードが閉じる")
            elif diff.locator(".sdresult").inner_text() == before:
                fails.append("割合を変えても差額が変わらない")
            if ovf() != 0:
                fails.append(f"差額の割合切替で横溢れ {ovf()}px")
            # 処方量を入れると合計が出る
            q = pg.locator(".card.open .sdq")
            if q.count() == 0:
                fails.append("処方量の入力欄が無い")
            else:
                q.fill("50")
                pg.wait_for_timeout(400)
                # 内服薬なら日数も入れないと金額が出ない
                dd = pg.locator(".card.open .sdday")
                if dd.count():
                    dd.fill("14")
                    pg.wait_for_timeout(400)
                txt = pg.locator(".card.open .sdtotal").inner_text()
                if not txt.strip():
                    fails.append("処方量を入れても合計が出ない")
                elif "円" not in txt:
                    fails.append("差額の金額が出ていない")
                if ovf() != 0:
                    fails.append(f"処方量の入力で横溢れ {ovf()}px")
    pg.fill("#q", "")
    pg.wait_for_timeout(250)

    pg.select_option("#ffm", "錠")
    pg.wait_for_timeout(380)
    if ovf() != 0:
        fails.append(f"剤形（詳細）で横溢れ {ovf()}px")
    pg.select_option("#ffm", "")
    pg.wait_for_timeout(250)
    pg.select_option("#fpc", "後発品")
    pg.wait_for_timeout(380)
    want = pg.evaluate("() => R.filter(r => D.pc[r[17]] === '後発品').length")
    if num(pg.inner_text("#cnt")) != want:
        fails.append("後発品の件数がデータと合わない（表示 %s / 期待 %d）"
                     % (pg.inner_text("#cnt"), want))
    pg.select_option("#fpc", "")
    pg.wait_for_timeout(250)
    pg.click("#advbtn")
    pg.wait_for_timeout(200)

    # 並び順の切替
    for _ in range(3):
        btn = pg.locator("#ordbtn")
        if btn.count() == 0:
            fails.append("並び順ボタンが無い")
            break
        btn.click()
        pg.wait_for_timeout(380)
        if ovf() != 0:
            fails.append(f"並び順切替で横溢れ {ovf()}px")
            break

    # 「この成分で検索」ボタン
    pg.fill("#q", "ノルバスク")
    pg.wait_for_timeout(420)
    if pg.locator(".card").count():
        pg.locator(".card").first.click()
        pg.wait_for_timeout(300)
        ib = pg.locator(".card.open .ingbtn")
        if ib.count() == 0:
            fails.append("「この成分で検索」ボタンが無い")
        else:
            ib.click()
            pg.wait_for_timeout(500)
            if pg.get_attribute('.sm[data-m="i"]', "aria-pressed") != "true":
                fails.append("成分名モードに切り替わらない")
            if not pg.input_value("#q"):
                fails.append("成分名が入力されない")
            if ovf() != 0:
                fails.append(f"成分切替後に横溢れ {ovf()}px")
    pg.fill("#q", "")
    pg.click('.sm[data-m="n"]')
    pg.wait_for_timeout(300)

    # 同一成分薬比較（剤形順／規格順）
    pg.fill("#q", "カルボシステインＤＳ")
    pg.wait_for_timeout(400)
    if pg.locator(".card").count() == 0:
        fails.append("比較用の検索結果が0件")
    else:
        pg.locator(".card").first.click()
        pg.wait_for_timeout(280)
        btn = pg.locator(".card.open .cmpbtn")
        if btn.count() == 0:
            fails.append("比較ボタンが出ない")
        else:
            btn.click()
            pg.wait_for_timeout(600)
            if not pg.locator("#ov").is_visible():
                fails.append("比較画面が開かない")
            if pg.locator(".grp").count() == 0:
                fails.append("剤形順のグループ見出しが無い")

            if ovf() != 0:
                fails.append(f"比較画面で横溢れ {ovf()}px")
            # 内部スクロール
            pg.locator(".ovb").evaluate("e=>e.scrollTop=9999")
            pg.wait_for_timeout(250)
            # 3つの並び替えを順に確認
            before = pg.locator(".grp").first.inner_text()
            for mode, label in [("a", "入手できる順"), ("s", "規格順")]:
                btn = pg.locator(f'#cmpsort .sortb[data-s="{mode}"]')
                if btn.count() == 0 or not btn.is_visible():
                    fails.append(f"{label}のボタンが無い")
                    continue
                btn.click()
                pg.wait_for_timeout(420)
                if pg.locator(".grp").count() == 0:
                    fails.append(f"{label}のグループ見出しが無い")
                if ovf() != 0:
                    fails.append(f"{label}で横溢れ {ovf()}px")
                if label == "入手できる順" and pg.locator(".grp").first.inner_text() == before:
                    fails.append("並び替えを押しても表示が変わらない")
            pg.click("#ovx")
            pg.wait_for_timeout(250)
            if pg.locator("#ov").is_visible():
                fails.append("比較画面が閉じない")
            if pg.evaluate("getComputedStyle(document.body).overflow==='hidden'"):
                fails.append("閉じた後もbodyがスクロール不可")

    # 経過措置フィルタ
    for f in ["has", "near", "over"]:
        btn = pg.locator(f'.chip[data-f="{f}"]')
        if btn.count() == 0:
            fails.append(f"経過措置チップ({f})が無い")
            break
        btn.click()
        pg.wait_for_timeout(300)
        if ovf() != 0:
            fails.append(f"経過措置フィルタで横溢れ {ovf()}px")
            break
        btn.click()
        pg.wait_for_timeout(200)

    # 計算ページ
    ct = pg.locator('.ptab[data-p="calc"]')
    if ct.count() == 0:
        fails.append("計算タブが無い")
    else:
        ct.click()
        pg.wait_for_timeout(600)
        if not pg.locator("#calc").is_visible():
            fails.append("計算ページが開かない")
        pg.click("#caadd")
        pg.wait_for_timeout(350)
        if pg.locator(".carp").count() == 0:
            fails.append("剤を追加できない")
        else:
            pg.locator(".caadddrug").first.click()
            pg.wait_for_timeout(500)
            if not pg.locator("#capick").is_visible():
                fails.append("薬の選択画面が開かない")
            pg.fill("#capinput", "ユーロジン２")
            pg.wait_for_timeout(500)
            if pg.locator(".capitem").count() == 0:
                fails.append("薬の候補が出ない")
            else:
                pg.locator(".capitem").first.click()
                pg.wait_for_timeout(400)
                pg.fill(".cadqin", "2")
                pg.wait_for_timeout(300)
                pg.fill(".cadaysin", "30")
                pg.wait_for_timeout(500)
                if not pg.locator("#cares").inner_text().strip():
                    fails.append("計算結果が出ない")
            if ovf() != 0:
                fails.append(f"計算ページで横溢れ {ovf()}px")
        pg.locator('.ptab[data-p="search"]').click()
        pg.wait_for_timeout(500)

    # 画面タブ（検索／お知らせ）
    tb = pg.locator('.ptab[data-p="board"]')
    if tb.count() == 0:
        fails.append("お知らせタブが無い")
    else:
        tb.click()
        pg.wait_for_timeout(700)
        if not pg.locator("#board").is_visible():
            fails.append("お知らせページが開かない")
        if pg.locator(".bdrow").count() == 0:
            fails.append("掲示板の中身が空")
        # お知らせの並び替え
        ns = pg.locator("#nwsort")
        if ns.count() == 0:
            fails.append("お知らせの並び替えボタンが無い")
        else:
            before = ns.inner_text()
            ns.click()
            pg.wait_for_timeout(350)
            if ns.inner_text() == before:
                fails.append("並び替えが切り替わらない")
            if ovf() != 0:
                fails.append(f"並び替えで横溢れ {ovf()}px")
            ns.click()
            pg.wait_for_timeout(250)

        # 設定は折りたたまれた状態で始まること（開く操作より前に見る）
        if pg.eval_on_selector("#cfg", "e=>e.open"):
            fails.append("設定が最初から開いている")
        if ovf() != 0:
            fails.append(f"お知らせページで横溢れ {ovf()}px")
        if pg.evaluate("getComputedStyle(document.getElementById('more')).display") != "none":
            fails.append("検索用の要素が残っている")
        # 設定は折りたたみなので、開いてから操作する
        pg.eval_on_selector("#cfg", "e=>e.open=true")
        pg.wait_for_timeout(250)
        # しきい値（帯の切り替え）
        for lo, hi in [("25", "50"), ("75", "100"), ("50", "100")]:
            b2 = pg.locator(f'#bdth .thb[data-lo="{lo}"][data-hi="{hi}"]')
            if b2.count() == 0:
                fails.append("しきい値ボタンが無い")
                break
            b2.click()
            pg.wait_for_timeout(280)
            if ovf() != 0:
                fails.append(f"しきい値切替で横溢れ {ovf()}px")
                break
        # 任意%（スライダー）
        pg.eval_on_selector("#rng", "e=>{e.value=30;e.dispatchEvent(new Event('input'))}")
        pg.wait_for_timeout(350)
        if pg.inner_text("#rnglb") != "30":
            fails.append("スライダーの値が反映されない")
        # 剤形の切り替え
        pg.locator('.kb2[data-k="2"]').click()
        pg.wait_for_timeout(350)
        if ovf() != 0:
            fails.append(f"剤形切替で横溢れ {ovf()}px")
        pg.locator('.kb2[data-k="2"]').click()
        pg.wait_for_timeout(250)
        # お知らせ掲示板
        if pg.locator("#nwlist").count() == 0:
            fails.append("お知らせ掲示板が無い")
        for t in ["cross", "down", "swing", ""]:
            nb = pg.locator(f'.nwt[data-t="{t}"]')
            if nb.count() == 0:
                fails.append("ジャンルタブが無い")
                break
            nb.click()
            pg.wait_for_timeout(280)
            if ovf() != 0:
                fails.append(f"ジャンル切替で横溢れ {ovf()}px")
                break
        # グラフ
        if pg.locator("#chart svg").count() == 0 and pg.locator(".bdempty").count() == 0:
            fails.append("グラフも案内も出ていない")
        pg.locator('.ptab[data-p="search"]').click()
        pg.wait_for_timeout(500)
        if not pg.locator("#list").is_visible():
            fails.append("検索へ戻れない")


    # 販売中止フィルタ（登録が無い環境では0件が正しい）
    pg.click('.chip[data-f="disc"]')
    pg.wait_for_timeout(420)
    if ovf() != 0:
        fails.append(f"販売中止フィルタで横溢れ {ovf()}px")
    pg.click("#clrall")
    pg.wait_for_timeout(250)


    # 後発品が高い薬
    pg.fill("#q", "")
    pg.wait_for_timeout(250)
    pg.click("#rvbtn")
    pg.wait_for_timeout(700)
    if not pg.locator("#rv").is_visible():
        fails.append("「後発品が高い」が開かない")
    if pg.locator(".rvgrp").count() == 0 and pg.locator(".rvempty").count() == 0:
        fails.append("逆転一覧の中身が空")
    if ovf() != 0:
        fails.append(f"逆転一覧で横溢れ {ovf()}px")
    pg.locator("#rvbody").evaluate("e=>e.scrollTop=9999")
    pg.wait_for_timeout(250)
    # 剤形フィルタ
    kbtns = [b for b in pg.locator("#rvk .sortb").all() if b.is_visible()]
    if len(kbtns) < 1:
        fails.append("逆転一覧の剤形ボタンが無い")
    for b in kbtns[:2]:
        b.click()
        pg.wait_for_timeout(320)
        if ovf() != 0:
            fails.append(f"逆転一覧の剤形フィルタで横溢れ {ovf()}px")
            break
    # 高い/同額の絞り込み
    for g in ["over", "eq", ""]:
        btn = pg.locator(f'#rvgap .sortb[data-g="{g}"]')
        if btn.count() == 0:
            fails.append("高い/同額のボタンが無い")
            break
        btn.click()
        pg.wait_for_timeout(300)
        if ovf() != 0:
            fails.append(f"高い/同額の絞り込みで横溢れ {ovf()}px")
            break
    pg.fill("#rvq", "錠")
    pg.wait_for_timeout(420)
    if ovf() != 0:
        fails.append(f"逆転一覧の検索で横溢れ {ovf()}px")
    pg.click("#rvord")
    pg.wait_for_timeout(380)
    pg.fill("#rvq", "")
    pg.wait_for_timeout(300)
    pg.click("#rvx")
    pg.wait_for_timeout(250)
    if pg.locator("#rv").is_visible():
        fails.append("逆転一覧が閉じない")

    # 凡例
    pg.fill("#q", "")
    pg.wait_for_timeout(250)
    pg.click("#lgbtn")
    pg.wait_for_timeout(500)
    if not pg.locator("#lg").is_visible():
        fails.append("凡例が開かない")
    if pg.locator(".lgrow").count() < 15:
        fails.append(f"凡例の項目が少ない({pg.locator('.lgrow').count()})")
    if ovf() != 0:
        fails.append(f"凡例で横溢れ {ovf()}px")
    pg.locator("#lgbody").evaluate("e=>e.scrollTop=9999")
    pg.wait_for_timeout(250)
    pg.click("#lgx")
    pg.wait_for_timeout(250)
    if pg.locator("#lg").is_visible():
        fails.append("凡例が閉じない")

    # タップ領域（Androidは48dp推奨。主要な操作要素を確認）
    small = pg.evaluate(f"""() => {{
      const sels=['#q','.sm','.chip','#advbtn','.cmpbtn','#logic','.sortb','#ovx','#clrall','#lgbtn','#lgx','#rvbtn','#rvx','#rvk .sortb','#rvq','#rvord','#rvgap .sortb','#ordbtn','.ingbtn','.ptab','.nwt','#nwsort','.chip.cm','.cmitem','.sdq','.sdday','#cmvk .kb2','.gfb','#genk .kb2','#gob','#caadd','.caadddrug','.carb'];
      const out=[];
      sels.forEach(s=>document.querySelectorAll(s).forEach(el=>{{
        const r=el.getBoundingClientRect();
        if(r.width>0 && r.height>0 && r.height < {MIN_TAP})
          out.push(s+' h='+Math.round(r.height));
      }}));
      return [...new Set(out)];
    }}""")
    if small:
        fails.append("タップ領域が小さい: " + ", ".join(small[:4]))

    if errs:
        fails.append(f"JSエラー {len(errs)}件: {errs[0][:60]}")

    vp = dev["viewport"]
    tag = "Android" if "Pixel" in name or "Galaxy" in name else "iOS  "
    status = "OK" if not fails else "NG"
    print(f'  [{tag}] {name:19s} {vp["width"]:>3}x{vp["height"]:<4} '
          f'dsf={dev["device_scale_factor"]:<5} {status}')
    for f in fails:
        print(f"        └ {f}")
    ctx.close()
    return not fails


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "医薬品供給状況_検索.html"
    if not os.path.exists(path):
        print(f"ファイルが見つかりません: {path}")
        return 1
    url = "file://" + os.path.abspath(path)
    print(f"検証対象: {path}")
    print(f"端末数: {len(DEVICES)}（Android {sum(1 for d in DEVICES if 'Pixel' in d or 'Galaxy' in d)} / "
          f"iOS {sum(1 for d in DEVICES if 'iPhone' in d)}）")
    print()
    with sync_playwright() as p:
        b = p.chromium.launch(args=["--no-sandbox"])
        results = [check_device(b, p, n, url) for n in DEVICES]
        b.close()
    print()
    print(f"合格 {sum(results)}/{len(results)} 端末 → "
          f"{'ALL PASS' if all(results) else 'FAIL'}")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
