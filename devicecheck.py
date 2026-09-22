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

GOOD_TYPES = {"clear", "uncross", "fall", "up"}


def isGoodType(t):
    return t in GOOD_TYPES


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
        # 削除リスト由来の記載が拾えているか。
        # 現行版にも過去版にも無い記載（例：バルプロ酸Ｎａ錠２００ｍｇ）が
        # 落ちると、対応する品目の一般名が分からなくなる。
        pg.fill("#q", "バルプロ酸Ｎａ錠２００")
        pg.wait_for_timeout(450)
        if pg.locator(".gcard").count() != 1:
            fails.append("削除リストの記載（バルプロ酸Ｎａ錠２００ｍｇ）が出ない")
        else:
            c = pg.locator(".gcard").first
            if c.locator(".gdel").count() == 0:
                fails.append("削除リストの記載に「削除」バッジが出ていない")
            if c.locator(".gadd").count():
                fails.append("削除リストの記載に加算バッジが出ている")
            c.click()
            pg.wait_for_timeout(350)
            names = c.locator(".cmnm").all_inner_texts()
            if not any("ＤＳＰ" in t for t in names):
                fails.append("削除リストの記載に「ＤＳＰ」が紐づいていない")
        pg.fill("#q", "アムロジピン錠")
        pg.wait_for_timeout(450)
        # 2.5mg と 5mg は現行版に無い（削除リストにも載っているため「削除」表示）。
        # 現行版に無い印が何も付かなくなったら、積み上げが壊れている。
        marks = (pg.locator(".gcard .gold").count()
                 + pg.locator(".gcard .gdel").count())
        if marks != 2:
            fails.append("現行版に無い記載の印が2件でない（%d件）" % marks)
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

    # 計算タブ：所定単位が剤形ごとに正しいか。
    # 頓服薬は1調剤（全量で1単位）。1回分ごとに点数化して回数を掛けると
    # 回数の分だけ金額が跳ね上がるため、内服との違いを必ず確かめる。
    amt = pg.evaluate(r"""() => {
      const r = R.find(x => get(x,'n').includes('ロキソニン錠６０'));
      if (!r) return null;
      const yj = get(r,'yj'), keep = caRps.slice();
      const run = (kind, qty, times) => {
        caRps = [{kind, days: times, drugs: [{yj, name: get(r,'n'), qty}]}];
        caRatio = 0.3; drawCalc();
        const t = document.getElementById('cares').innerText;
        const m = t.match(/特別の料金（消費税込み）\s*([\d,]+)/);
        return m ? Number(m[1].replace(/,/g,'')) : null;
      };
      const out = {oral: run('内服',1,10), ton: run('頓服',1,10),
                   ext: run('外用',10,1)};
      caRps = keep; drawCalc();
      return out;
    }""")
    if amt is None:
        fails.append("計算タブの検証薬（ロキソニン錠６０ｍｇ）が見つからない")
    else:
        if amt["ton"] != amt["ext"]:
            fails.append("頓服薬が1調剤で計算されていない（頓服%s円／外用%s円）"
                         % (amt["ton"], amt["ext"]))
        if amt["ton"] != 11:
            fails.append("頓服 1回1錠10回分が11円でない（%s円）" % amt["ton"])
        if amt["oral"] != 110:
            fails.append("内服 1日1錠10日分が110円でない（%s円）" % amt["oral"])

    # カード内の差額も、計算タブと同じ所定単位で求めているか。
    # 画面が実際に呼ぶ senTotalText() の出力を読み、
    # 期待値は同じ関数を使わず、規則から独立に計算して突き合わせる
    # （同じ関数で検算すると、分岐の誤りをそのまま素通りさせてしまう）。
    card = pg.evaluate(r"""() => {
      const r = R.find(x => get(x,'n').includes('ロキソニン錠６０'));
      if (!r) return null;
      const v = senOf(r), u = unitOf(r), qu = u.replace(/^1/,'') || '個';
      const ratio = 0.3, qty = 1, times = 10;
      const pick = (html) => {
        const m = html.match(/長期収載品\s*([\d,]+)円/);
        return m ? Number(m[1].replace(/,/g,'')) : null;
      };
      // 画面の出力
      const got = {
        teiki: pick(senTotalText(v, ratio, qty, u, qu, times, true, false)),
        ton:   pick(senTotalText(v, ratio, qty, u, qu, times, true, true))
      };
      // 規則からの独立計算（15円以下は1点、15円超は10円で割って五捨五超入）
      const ten = (yen) => {
        if (yen <= 0) return 0;
        if (yen <= 15) return 1;
        const q = yen/10, i = Math.floor(q);
        return (q - i) <= 0.5 ? i : i + 1;
      };
      const want = {
        // 定期＝1日量で点数化して日数を掛ける
        teiki: ten(v[0]*qty)*times*10*1.1 + ten(v[1]*qty)*times*10*ratio,
        // 頓用＝1回量×回数の全量で1単位
        ton:   ten(v[0]*qty*times)*10*1.1 + ten(v[1]*qty*times)*10*ratio
      };
      return {got, want: {teiki: Math.round(want.teiki),
                          ton:   Math.round(want.ton)}};
    }""")
    if card is None:
        fails.append("カードの差額検証用の薬が見つからない")
    else:
        for k, ja in (("teiki", "定期"), ("ton", "頓用")):
            if card["got"][k] != card["want"][k]:
                fails.append("カードの%sが所定単位どおりでない（画面%s円／規則%s円）"
                             % (ja, card["got"][k], card["want"][k]))
        if card["got"]["ton"] == card["got"]["teiki"]:
            fails.append("カードで定期と頓用の金額が同じ（切り替えが効いていない）")

    # お知らせ：通常出荷に戻った薬。
    # 前後が同じ（例：供給停止→供給停止）の行が出ていないか。
    # 履歴から「戻った日」を拾う作りなので、from は 1 か 2 に限られる。
    rec = pg.evaluate(r"""() => {
      const a = DATA.recover || [];
      return {n: a.length,
              bad: a.filter(x => x.from !== 1 && x.from !== 2).length,
              noday: a.filter(x => !x.d).length};
    }""")
    if rec["bad"]:
        fails.append("通常出荷に戻った薬に、戻る前が通常出荷の行が %d件ある" % rec["bad"])
    if rec["noday"]:
        fails.append("通常出荷に戻った薬に、日付の無い行が %d件ある" % rec["noday"])

    # 画面に固定した「？」。スクロールしても押せること。
    pg.evaluate("window.scrollTo(0, 4000)")
    pg.wait_for_timeout(500)
    if not pg.locator("#helpfab").is_visible():
        fails.append("スクロール後に「？」が消えている")
    else:
        a = pg.locator("#helpfab").bounding_box()
        t = pg.locator("#top").bounding_box()
        if t and a and not (a["y"] + a["height"] <= t["y"]
                            or t["y"] + t["height"] <= a["y"]):
            fails.append("「？」と「先頭へ戻る」が重なっている")
        if a and (a["width"] < MIN_TAP or a["height"] < MIN_TAP):
            fails.append("「？」のタップ領域が小さい（%dx%d）"
                         % (a["width"], a["height"]))
    pg.evaluate("window.scrollTo(0, 0)")
    pg.wait_for_timeout(400)

    # 凡例の矢印の説明が、実際の表示と食い違っていないこと。
    # 記録がたまったのに「まだ表示していません」と書いてあると誤解を招く。
    pg.click("#lgbtn")
    pg.wait_for_timeout(450)
    lg = pg.inner_text("#lgbody")
    gen = pg.evaluate("DATA.vgen || 0")
    if gen >= 2:
        if "いまは表示していません" in lg:
            fails.append("矢印は出ているのに、凡例が「表示していません」のまま")
        for w in ("前回より改善した", "前回から変わっていない", "前回より悪化した"):
            if w not in lg:
                fails.append("凡例に矢印つきバッジの説明が無い（%s）" % w)
    else:
        if "いまは表示していません" not in lg:
            fails.append("矢印が出ないのに、凡例がその旨を説明していない")
    pg.click("#lgx")
    pg.wait_for_timeout(300)

    # 出荷量の矢印は、履歴が2世代未満なら出さないこと。
    arr = pg.evaluate(r"""() => {
      const gen = DATA.vgen || 0;
      const any = R.some(r => {
        const v = get(r,'vc');
        return v != null && v >= 0 && v !== 4 && volBadge(r).includes('vbd');
      });
      return {gen, any};
    }""")
    if arr["gen"] < 2 and arr["any"]:
        fails.append("履歴が%d世代しか無いのに出荷量の矢印が出ている" % arr["gen"])

    # お知らせタブのボタンが、枠からはみ出していないこと。
    # ボタンの数が増えると1行に収まらず、右へ飛び出す。
    pg.click('.ptab[data-p="board"]')
    pg.wait_for_timeout(500)
    over = pg.evaluate("""() => {
      const out = [];
      document.querySelectorAll('.bdsec .chgtabs, .bdsec .nwside').forEach(t => {
        const sec = t.closest('.bdsec').getBoundingClientRect();
        t.querySelectorAll('button').forEach(b => {
          if (!b.offsetParent) return;
          const x = b.getBoundingClientRect();
          if (x.right > sec.right - 1 || b.scrollWidth > b.clientWidth + 1)
            out.push(b.textContent.trim());
        });
      });
      return out;
    }""")
    if over:
        fails.append("お知らせのボタンが枠からはみ出している：%s" % "／".join(over))

    # 出荷量の動き（傾向と新規の薬価削除予定）。
    # 傾向は履歴全体から判定するので、型は5種のいずれかに収まること。
    vt = pg.evaluate(r"""() => {
      const ok = ['imp','rec','wor','rel','swing'];
      const a = DATA.voltrend || [];
      return {n: a.length,
              bad: a.filter(x => !ok.includes(x.p)).length,
              noseq: a.filter(x => !(x.n >= 2)).length,
              del: (DATA.newdel || []).length,
              deldup: (DATA.voltrend || []).filter(x => x.b === 4).length};
    }""")
    if vt["bad"]:
        fails.append("出荷量の傾向に未知の型が %d件ある" % vt["bad"])
    if vt["noseq"]:
        fails.append("出荷量の傾向に履歴2点未満の行が %d件ある" % vt["noseq"])
    if vt["deldup"]:
        fails.append("薬価削除予定が傾向に混じっている（%d件）" % vt["deldup"])

    # ⑰出荷量。バッジ・絞り込み・並び替えが動くこと。
    # 直前でお知らせタブに移っているので、検索画面へ戻す。
    # 【般】一般名モードのままだと詳細フィルタ自体が隠れるので、品名モードに戻す。
    pg.click('.ptab[data-p="search"]')
    pg.wait_for_timeout(400)
    pg.click('.sm[data-m="n"]')
    pg.wait_for_timeout(400)
    vol = pg.evaluate(r"""() => {
      const c = {};
      R.forEach(r => { const v = get(r,'vc'); c[v] = (c[v]||0)+1; });
      return c;
    }""")
    if not any(vol.get(str(k), 0) for k in range(5)):
        fails.append("出荷量（⑰）が1件も読み取れていない")
    # 詳細フィルタは折りたたみ。開いていなければ開く
    if not pg.locator("#fvc").is_visible():
        pg.click("#advbtn")
        pg.wait_for_timeout(350)
    pg.select_option("#fvc", "4")           # 薬価削除予定
    pg.wait_for_timeout(500)
    want = vol.get("4", 0)
    if num(pg.inner_text("#cnt")) != want:
        fails.append("出荷量での絞り込みが合わない（表示 %s / 期待 %d）"
                     % (pg.inner_text("#cnt"), want))
    pg.select_option("#fvc", "")
    pg.wait_for_timeout(400)
    for _ in range(4):
        if "出荷量" in pg.inner_text("#ordbtn"):
            break
        pg.click("#ordbtn")
        pg.wait_for_timeout(350)
    if "出荷量" not in pg.inner_text("#ordbtn"):
        fails.append("出荷量での並び替えが選べない")
    else:
        head = pg.evaluate("() => results.slice(0,5).map(r => get(r,'vc'))")
        if head != sorted(head):
            fails.append("出荷量の並び替えが良い順になっていない（%s）" % head)
        pg.click("#ordbtn")
        pg.wait_for_timeout(300)
    # 開いたままにすると、狭い端末で #fmk などが他の要素を覆い、
    # 後続のクリックを遮ってしまう。必ず閉じてから次へ進む。
    if pg.locator("#fvc").is_visible():
        pg.click("#advbtn")
        pg.wait_for_timeout(300)

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

    # 詳細フィルタは折りたたみ。閉じていれば開いてから操作する
    if not pg.locator("#ffm").is_visible():
        pg.click("#advbtn")
        pg.wait_for_timeout(350)
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
    if pg.locator("#ffm").is_visible():
        pg.click("#advbtn")
        pg.wait_for_timeout(250)

    # 並び順の切替。選択肢は4つ（状況／成分／名前／出荷量）なので、
    # 4回押して元に戻すところまで確かめる。
    for _ in range(4):
        btn = pg.locator("#ordbtn")
        if btn.count() == 0:
            fails.append("並び順ボタンが無い")
            break
        btn.click()
        pg.wait_for_timeout(380)
        if ovf() != 0:
            fails.append(f"並び順切替で横溢れ {ovf()}px")
            break
    if "状況" not in pg.inner_text("#ordbtn"):
        fails.append("並び順が4回で一周しない（%s）" % pg.inner_text("#ordbtn"))

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

    # お知らせ掲示板：良い知らせ／悪い知らせの切り替え。
    # 種類ボタンは選んだ向きに合うものだけが出ること。
    pg.click('.ptab[data-p="board"]')
    pg.wait_for_timeout(500)
    for v, want_good in (("good", True), ("bad", False)):
        pg.click(f'#nwside .nws[data-v="{v}"]')
        pg.wait_for_timeout(350)
        vis = pg.evaluate("""() => [...document.querySelectorAll('#nwtabs .nwt')]
            .filter(b => b.offsetParent).map(b => b.dataset.t).filter(Boolean)""")
        bad = [t for t in vis if isGoodType(t) != want_good]
        if bad:
            fails.append("掲示板の種類ボタンが向きと合わない（%s に %s）"
                         % (v, "／".join(bad)))
    pg.click('#nwside .nws[data-v=""]')
    pg.wait_for_timeout(300)

    # 剤形の設定が、お知らせタブの各段すべてに効くこと
    before = pg.evaluate("""() => ({
      rec: (document.getElementById('chgcnt')||{}).textContent || '',
      news: (document.getElementById('nwcnt')||{}).textContent || '',
      bd: (document.getElementById('bdcnt')||{}).textContent || ''})""")
    pg.evaluate("document.getElementById('cfg').open = true")
    pg.click('#bdk .kb2[data-k="2"]')          # 注射薬を足す
    pg.wait_for_timeout(500)
    after = pg.evaluate("""() => ({
      rec: (document.getElementById('chgcnt')||{}).textContent || '',
      news: (document.getElementById('nwcnt')||{}).textContent || '',
      bd: (document.getElementById('bdcnt')||{}).textContent || ''})""")
    if before == after:
        fails.append("剤形を変えてもお知らせタブの表示が変わらない")
    pg.click('#bdk .kb2[data-k="2"]')          # 元に戻す
    pg.wait_for_timeout(400)


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
