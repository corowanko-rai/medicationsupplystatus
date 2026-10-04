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
    "規制区分",       # 毒・劇・麻などのバッジ（regulation.json）
    "生薬・漢方",     # 出荷量の動きの折りたたみ・除外
    "切替余地",       # カードの切替余地／残りわずかの成分
    "回復までの日数", # 分析タブ
    "マイ薬局",       # 採用薬（端末内保存）
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

    # 出荷量の動き：並び・まとめ・生薬漢方の折りたたみ・設定。
    # 期待値は画面の関数を使わず、DATA と行データから独立に作る
    # （画面と同じ関数で期待値を作ると、分岐の誤りを見逃すため）。
    VT_EXPECT = r"""(T) => {
      const KAN={'漢方製剤':1,'生薬':2,'その他の生薬及び漢方処方に基づく医薬品':2};
      const KI={'内用薬':0,'外用薬':1,'注射薬':2};
      const on=new Set([0,1]);
      const kan=(r)=>KAN[D.cls[r[6]]]||0;
      const all=[...(DATA.voltrend||[]), ...(DATA.newdel||[])]
        .filter(x=>{const r=R[x.i]; return r && on.has(KI[D.k[r[5]]]);});
      const sec=[[],[],[]]; all.forEach(x=>sec[kan(R[x.i])].push(x));
      const grp=(items)=>{const m={}; items.forEach(x=>{
          const r=R[x.i]; const mv = x.from!==undefined ? 'd'+x.from : x.p+x.a+x.b;
          const k=r[2]+'_'+r[6]+'_'+mv; m[k]=(m[k]||0)+1;});
        const big=Object.values(m).filter(n=>n>=T);
        return {groups:big.length, grouped:big.reduce((a,b)=>a+b,0)};};
      return {n:all.length, s0:sec[0].length, s1:sec[1].length, s2:sec[2].length,
              g0:grp(sec[0]), g2:grp(sec[2]),
              s0sc: sec[0].map(x=>R[x.i][8])};
    }"""
    VT_SHOWN = r"""() => {
      const box=document.getElementById('vtlist');
      const top=[...box.children];
      const iFold=top.findIndex(e=>e.classList.contains('vtfold'));
      const head= iFold<0 ? top : top.slice(0,iFold);
      const KAN={'漢方製剤':1,'生薬':2,'その他の生薬及び漢方処方に基づく医薬品':2};
      const byName={}; R.forEach(r=>{byName[r[0]]=r;});
      const rows=head.filter(e=>e.classList.contains('chgrow'))
        .map(e=>byName[e.dataset.nm]).filter(Boolean);
      return {
        folds: top.filter(e=>e.classList.contains('vtfold')).map(e=>+e.dataset.f),
        open: box.querySelectorAll('.vtfbody').length,
        headGrp: head.filter(e=>e.classList.contains('vtgrp')).length,
        headKan: rows.filter(r=>KAN[D.cls[r[6]]]).length,
        headSc: rows.map(r=>r[8]),
        anyGrp: box.querySelectorAll('.vtgrp').length,
      };
    }"""
    pg.click('.ptab[data-p="board"]')
    pg.wait_for_timeout(400)
    exp = pg.evaluate(VT_EXPECT, 5)
    got = pg.evaluate(VT_SHOWN)
    if got["headKan"]:
        fails.append("出荷量の動き：生薬・漢方が折りたたみの外に %d件出ている" % got["headKan"])
    if got["headSc"] != sorted(got["headSc"], reverse=True):
        fails.append("出荷量の動き：出荷対応の悪い順に並んでいない（%s）" % got["headSc"][:12])
    want_f = [k for k, n in ((1, exp["s1"]), (2, exp["s2"])) if n]
    if got["folds"] != want_f:
        fails.append("出荷量の動き：折りたたみが %s（期待 %s：漢方製剤→生薬）" % (got["folds"], want_f))
    if got["open"]:
        fails.append("出荷量の動き：生薬・漢方の折りたたみが最初から開いている")
    if got["headGrp"] != exp["g0"]["groups"]:
        fails.append("出荷量の動き：生薬・漢方以外のまとめが %d行（期待 %d）"
                     % (got["headGrp"], exp["g0"]["groups"]))
    # 生薬を開き、まとめの行数と品目数を確かめる
    if exp["s2"]:
        pg.click('.vtfold[data-f="2"]')
        pg.wait_for_timeout(300)
        g2 = pg.evaluate("""() => {
          const b=document.querySelector('#vtlist .vtfbody');
          if(!b) return null;
          const gs=[...b.querySelectorAll(':scope > .vtgrp')];
          return {groups: gs.length,
                  grouped: gs.reduce((n,e)=>n+parseInt(e.querySelector('.vtgn').textContent),0),
                  singles: b.querySelectorAll(':scope > .chgrow').length};
        }""")
        if not g2:
            fails.append("出荷量の動き：生薬の折りたたみが開かない")
        else:
            if g2["groups"] != exp["g2"]["groups"] or g2["grouped"] != exp["g2"]["grouped"]:
                fails.append("出荷量の動き：生薬のまとめが %d行・%d品目（期待 %d行・%d品目）"
                             % (g2["groups"], g2["grouped"],
                                exp["g2"]["groups"], exp["g2"]["grouped"]))
            if g2["grouped"] + g2["singles"] != exp["s2"]:
                fails.append("出荷量の動き：生薬の品目数が合わない（%d＋%d≠%d）"
                             % (g2["grouped"], g2["singles"], exp["s2"]))
            # まとめを開くと中身が出ること
            if g2["groups"]:
                pg.locator('#vtlist .vtfbody > .vtgrp').first.click()
                pg.wait_for_timeout(250)
                if not pg.locator('#vtlist .vtgbody .chgrow').count():
                    fails.append("出荷量の動き：まとめを開いても中身が出ない")
        if ovf() != 0:
            fails.append(f"出荷量の動きを開いて横溢れ {ovf()}px")
    # まとめる件数の設定：99なら1つもまとめない／2なら2件以上をまとめる
    pg.evaluate("document.getElementById('cfg').open=true")
    pg.wait_for_timeout(200)
    for T in (99, 2):
        pg.fill("#vtgm", str(T))
        pg.dispatch_event("#vtgm", "change")
        pg.wait_for_timeout(300)
        e2 = pg.evaluate(VT_EXPECT, T)
        shown = pg.evaluate(VT_SHOWN)
        if T == 99 and shown["anyGrp"]:
            fails.append("まとめる件数を99にしてもまとめ行が残る")
        if T == 2 and shown["headGrp"] != e2["g0"]["groups"]:
            fails.append("まとめる件数2で、まとめ行が %d（期待 %d）"
                         % (shown["headGrp"], e2["g0"]["groups"]))
    saved = pg.evaluate("() => { try { return localStorage.getItem('vtGroupMin'); } catch(e) { return 'x'; } }")
    if saved not in ("2", "x"):
        fails.append("まとめる件数が端末に保存されていない（%s）" % saved)
    pg.click("#vtgmup")                        # 2 → 3
    pg.wait_for_timeout(250)
    if pg.input_value("#vtgm") != "3":
        fails.append("まとめる件数の＋ボタンが効かない（%s）" % pg.input_value("#vtgm"))
    pg.fill("#vtgm", "1")                      # 下限未満は2に寄せる
    pg.dispatch_event("#vtgm", "change")
    pg.wait_for_timeout(250)
    if pg.input_value("#vtgm") != "2":
        fails.append("まとめる件数が下限2未満を受け付けた（%s）" % pg.input_value("#vtgm"))
    pg.fill("#vtgm", "5")                      # 既定に戻す
    pg.dispatch_event("#vtgm", "change")
    pg.wait_for_timeout(250)
    # 設定欄の操作要素が押しやすく、枠からはみ出していないこと
    cfg = pg.evaluate(f"""() => {{
      const c=document.querySelector('#cfg .cfgin').getBoundingClientRect();
      const out=[], small=[];
      document.querySelectorAll('#cfg .cfgin *').forEach(el=>{{
        const r=el.getBoundingClientRect();
        if(r.width>0 && r.right>c.right+1) out.push(el.id||el.className||el.tagName);
      }});
      document.querySelectorAll('.gmb,#vtgm,#bdkan .kb2,.vtfold,.vtgrp').forEach(el=>{{
        const r=el.getBoundingClientRect();
        if(r.width>0 && r.height<{MIN_TAP}) small.push((el.id||el.className)+' h='+Math.round(r.height));
      }});
      return {{out:[...new Set(out)], small:[...new Set(small)]}};
    }}""")
    if cfg["out"]:
        fails.append("表示の設定で枠からはみ出す要素：%s" % "／".join(cfg["out"][:4]))
    if cfg["small"]:
        fails.append("出荷量の動き・設定のタップ領域が小さい：%s" % "／".join(cfg["small"][:4]))
    # 生薬・漢方を一覧から除く
    before_news = pg.inner_text("#nwcnt")
    pg.click('#bdkan .kb2[data-v="1"]')
    pg.wait_for_timeout(450)
    ex = pg.evaluate(r"""() => {
      const KAN={'漢方製剤':1,'生薬':2,'その他の生薬及び漢方処方に基づく医薬品':2};
      const byName={}; R.forEach(r=>{byName[r[0]]=r;});
      const kanRows=(sel)=>[...document.querySelectorAll(sel)]
        .map(e=>byName[e.dataset.nm]).filter(r=>r && KAN[D.cls[r[6]]]).length;
      /* 成分単位：その成分の品目がすべて生薬・漢方なら除かれているはず */
      const all=new Map();
      R.forEach(r=>{const k=KAN[D.cls[r[6]]]?1:0, p=all.get(r[1]);
        all.set(r[1], p===undefined?k:(p&k));});
      const kanIng=new Set([...all].filter(([i,v])=>v).map(([i])=>i));
      return {
        vtFold: document.querySelectorAll('#vtlist .vtfold').length,
        vtKan: kanRows('#vtlist .chgrow'),
        chgKan: kanRows('#chglist .chgrow'),
        bdKan: [...document.querySelectorAll('#bdlist .bdrow')]
          .filter(e=>kanIng.has(+e.dataset.ing)).length,
        saved: (()=>{ try { return localStorage.getItem('bdNoKan'); } catch(e) { return 'x'; } })(),
      };
    }""")
    if ex["vtFold"] or ex["vtKan"] or ex["chgKan"] or ex["bdKan"]:
        fails.append("生薬・漢方を除いても残っている（動き%d・折りたたみ%d・通常出荷に戻った%d・逼迫%d）"
                     % (ex["vtKan"], ex["vtFold"], ex["chgKan"], ex["bdKan"]))
    if ex["saved"] not in ("1", "x"):
        fails.append("生薬・漢方を除く設定が端末に保存されていない")
    pg.click('#bdkan .kb2[data-v="0"]')        # 元に戻す
    pg.wait_for_timeout(450)
    if pg.inner_text("#nwcnt") != before_news:
        fails.append("生薬・漢方を含めに戻しても、お知らせの件数が戻らない")
    if pg.locator('#vtlist .vtfold').count() != len(want_f):
        fails.append("生薬・漢方を含めに戻しても、折りたたみが戻らない")
    pg.evaluate("document.getElementById('cfg').open=false")
    pg.wait_for_timeout(150)

    # ===== v45：残りわずかの成分／切替余地／マイ薬局／分析タブ =====
    # 期待値は画面の関数を使わず、行データから独立に数える。
    GRP_JS = r"""
      const POWD=new Set(['散','細粒','顆粒','ドライシロップ','シロップ','内用液']);
      const SOLID=new Set(['錠','OD錠','カプセル','チュアブル']);
      const grpOf=(r)=>{ const k=D.k[r[5]], f=D.fm[r[23]]||'';
        if(k==='内用薬') return POWD.has(f)?'粉・液':SOLID.has(f)?'錠・カプセル':'その他';
        if(k==='外用薬') return f||'その他'; return '注射'; };
    """
    # タブが4つになっても1行に収まり、文字が切れていないこと
    tabs = pg.evaluate("""() => [...document.querySelectorAll('.ptab')].map(b => ({
        t: b.dataset.p, h: Math.round(b.getBoundingClientRect().height),
        cut: b.scrollWidth > b.clientWidth + 1}))""")
    if len(tabs) != 4:
        fails.append("画面のタブが %d個（期待 4）" % len(tabs))
    if any(t["cut"] for t in tabs) or len({t["h"] for t in tabs}) != 1 or tabs[0]["h"] > 52:
        fails.append("タブが1行に収まっていない（%s）" % tabs)

    # 残りわずかの成分：件数が独立に数えた値と合うこと
    pg.click('.ptab[data-p="board"]')
    pg.wait_for_timeout(450)
    low = pg.evaluate("() => {" + GRP_JS + r"""
      const KAN={'漢方製剤':1,'生薬':2,'その他の生薬及び漢方処方に基づく医薬品':2};
      const g=new Map();
      R.forEach(r=>{ if(!r[4]||r[4].length!==12) return;
        const k=D.k[r[5]]; if(k!=='内用薬'&&k!=='外用薬') return;
        const key=r[1]+'|'+k+'|'+grpOf(r);
        let v=g.get(key); if(!v){ v={n:0,ok:0}; g.set(key,v); }
        v.n++; if(r[8]===0) v.ok++; });
      let few=0, zero=0;
      g.forEach(v=>{ if(v.n>=3&&v.ok>=1&&v.ok<=2&&(v.n-v.ok)/v.n>=0.6) few++;
                     else if(v.n>=2&&v.ok===0) zero++; });
      const lab=[...document.querySelectorAll('#lowsub .chgs')].map(b=>b.textContent);
      return {few, zero, lab, rows: document.querySelectorAll('#lowlist .lowrow').length};
    }""")
    want_lab = ["残り1〜2品目（%d）" % low["few"], "通常出荷なし（%d）" % low["zero"]]
    if low["lab"] != want_lab:
        fails.append("残りわずかの成分の件数が合わない（表示 %s / 期待 %s）" % (low["lab"], want_lab))
    if low["rows"] != min(low["few"], 200):
        fails.append("残りわずかの成分の行数が %d（期待 %d）" % (low["rows"], min(low["few"], 200)))
    if low["zero"]:
        pg.click('#lowsub .chgs[data-v="zero"]')
        pg.wait_for_timeout(250)
        if pg.locator('#lowlist .lowrow.zero').count() != min(low["zero"], 200):
            fails.append("「通常出荷なし」の行数が合わない")
        pg.click('#lowsub .chgs[data-v="few"]')
        pg.wait_for_timeout(200)
    if ovf() != 0:
        fails.append(f"残りわずかの成分で横溢れ {ovf()}px")

    # 切替余地：限定出荷・供給停止の品目を1つ選び、①②の件数が独立計算と合うこと
    pg.click('.ptab[data-p="search"]')
    pg.wait_for_timeout(350)
    pg.click('.sm[data-m="n"]')
    pg.wait_for_timeout(250)
    sw = pg.evaluate("() => {" + GRP_JS + r"""
      const pick = R.find(r=>r[4]==='6132002R1141'&&r[8]>0) ||
                   R.find(r=>r[8]>0&&r[4]&&r[4].length===12&&D.k[r[5]]==='内用薬');
      if(!pick) return null;
      const k=D.k[pick[5]], g=grpOf(pick), y9=pick[4].slice(0,9);
      const same=R.filter(r=>r[4]&&r[4].length===12&&r[1]===pick[1]&&D.k[r[5]]===k);
      const t1=R.filter(r=>r!==pick&&r[4]&&r[4].length===12&&r[4].slice(0,9)===y9);
      const t2=same.filter(r=>r[4].slice(0,9)!==y9&&grpOf(r)===g);
      const t2b=same.filter(r=>grpOf(r)!==g);
      const f=(a)=>a.length?`通常出荷 ${a.filter(r=>r[8]===0).length}／全${a.length}`:null;
      return {name: pick[0], want:[f(t1),f(t2),f(t2b)].filter(Boolean),
              ok1: t1.filter(r=>r[8]===0).map(r=>r[0]).sort()};
    }""")
    if sw:
        pg.fill("#q", sw["name"])
        pg.wait_for_timeout(450)
        card = pg.locator(".card").first
        card.click()
        pg.wait_for_timeout(200)
        if card.locator(".swbtn").count() != 1:
            fails.append("カードに切替余地のボタンが無い")
        else:
            card.locator(".swbtn").click()
            pg.wait_for_timeout(250)
            got = card.locator(".swt .swcnt").all_inner_texts()
            if got[:len(sw["want"])] != sw["want"]:
                fails.append("切替余地の件数が合わない（表示 %s / 期待 %s）" % (got, sw["want"]))
            if "open" not in (card.get_attribute("class") or ""):
                fails.append("切替余地を押すとカードが閉じてしまう")
            shown1 = sorted(card.locator(".swt").first.locator(".swi .cmnm").all_inner_texts()) if sw["ok1"] else []
            if sw["ok1"] and shown1 != sw["ok1"][:6] and len(sw["ok1"]) <= 6:
                fails.append("切替余地①の品名が合わない（%s / 期待 %s）" % (shown1, sw["ok1"]))
            if ovf() != 0:
                fails.append(f"切替余地を開いて横溢れ {ovf()}px")
            cut = pg.evaluate("""() => [...document.querySelectorAll('.swi .cmnm')]
                .filter(e => e.getBoundingClientRect().width < 90).length""")
            if cut:
                fails.append("切替余地の品名が細く折れている（%d件）" % cut)
        pg.fill("#q", "")
        pg.wait_for_timeout(250)

    # マイ薬局：店舗の作成・一覧の読み込み・絞り込み・切り替え・書き出し/読み込み
    pg.evaluate("() => { try { localStorage.removeItem('myph'); localStorage.removeItem('myOnly'); } catch(e) {} }")
    pg.reload()
    pg.wait_for_timeout(2600)
    if pg.locator('#mybar.on').count() != 0:
        fails.append("店舗が無いのに店舗の切り替えが出ている")
    pg.click('.chip.myset')
    pg.wait_for_timeout(300)
    pg.click('#myadd'); pg.wait_for_timeout(150)
    pg.click('#myadd'); pg.wait_for_timeout(150)
    for i, nm in enumerate(("本店", "駅前店")):
        inp = pg.locator('.mystore').nth(i).locator('.myname')
        inp.fill(nm); inp.dispatch_event('change'); pg.wait_for_timeout(150)
    yjs = pg.evaluate("""() => {
      const a=R.filter(r=>r[4]&&r[4].length===12&&r[8]===0).slice(0,2).map(r=>r[4]);
      const b=R.filter(r=>r[4]&&r[4].length===12&&r[8]>0).slice(0,2).map(r=>r[4]);
      const nm=R.find(r=>r[8]===0&&R.filter(x=>x[0]===r[0]).length===1&&!a.includes(r[4]));
      return {a, b, nm:[nm[0], nm[4]]};
    }""")
    pg.locator('.mystore').nth(0).locator('[data-act="imp"]').click(); pg.wait_for_timeout(150)
    pg.fill('.mytxt', "品名,コード\n,%s\n%s\n%s\n存在しない薬ＸＹＺ\n%s" % (
        yjs["a"][0], yjs["b"][0], yjs["nm"][0], yjs["a"][1].lower()))
    pg.locator('.mystore').nth(0).locator('[data-act="rep"]').click(); pg.wait_for_timeout(250)
    # レセコンのCSV（シフトJIS・見出しより上に出力条件の行・在庫0の行つき）を読めること
    csv_path = "/tmp/_dc_list.csv"
    with open(csv_path, "w", encoding="cp932", newline="") as f:
        f.write("店舗名：,検査用薬局,,\r\n日付：,2026/10/ 4,,\r\n")
        f.write("管理区分,YJコード,薬品名,論理在庫\r\n")
        f.write("調剤,%s,何かの薬,12\r\n" % yjs["a"][0])
        f.write("調剤,%s,何かの薬,0\r\n" % yjs["b"][0])
        f.write('調剤,,"%s",3\r\n' % yjs["nm"][0])
        f.write("調剤,9999999X9999,一覧に無い薬,5\r\n調剤,,ペンニードル,1\r\n")
    pg.click('#myadd'); pg.wait_for_timeout(150)
    pg.locator('.mystore').nth(2).locator('[data-act="imp"]').click(); pg.wait_for_timeout(150)
    with pg.expect_file_chooser() as fc:
        pg.locator('.mystore').nth(2).locator('[data-act="file"]').click()
    fc.value.set_files(csv_path); pg.wait_for_timeout(500)
    pv = pg.inner_text('#mypv')
    if "3品目" not in pv or "検査用薬局" not in pv or "2件" not in pv:
        fails.append("レセコンCSVの読み取りが合わない（%s）" % pv[:70])
    pg.locator('.mystore').nth(2).locator('[data-act="zero"]').click(); pg.wait_for_timeout(200)
    pg.locator('.mystore').nth(2).locator('[data-act="rep"]').click(); pg.wait_for_timeout(300)
    s3 = pg.evaluate("() => [MY.stores[2].name, MY.stores[2].src, MY.stores[2].items.slice().sort()]")
    if s3 != ["店舗3", "検査用薬局", sorted([yjs["a"][0], yjs["nm"][1]])]:
        fails.append("レセコンCSV（在庫0を除く）の登録が合わない（%s）" % s3)
    # 表示名を付け直しても、CSVの店舗名から同じ店舗だと自動で見分けること
    inp = pg.locator('.mystore').nth(2).locator('.myname')
    inp.fill("古淵店"); inp.dispatch_event('change'); pg.wait_for_timeout(200)
    pg.locator('.mystore').nth(2).locator('[data-act="zero"]').click() if pg.locator('.mystore').nth(2).locator('[data-act="zero"]').count() else None
    with pg.expect_file_chooser() as fc:
        pg.click('#mycsv')
    fc.value.set_files(csv_path); pg.wait_for_timeout(500)
    au = pg.inner_text('#myauto') if pg.locator('#myauto').count() else ""
    if "検査用薬局" not in au or "古淵店" not in au:
        fails.append("CSVの店舗名から店舗を自動で見分けていない（%s）" % au[:60])
    else:
        if pg.locator('#myauto [data-auto="zero"]').inner_text().startswith("✓"):
            pg.click('#myauto [data-auto="zero"]'); pg.wait_for_timeout(250)
        pg.click('#myauto [data-auto="rep"]'); pg.wait_for_timeout(350)
        s4 = pg.evaluate("() => [MY.stores.length, MY.stores[2].name, MY.stores[2].items.length, MYLAB[2]]")
        if s4 != [3, "古淵店", 3, "古"]:   # 印は頭文字の「古」
            fails.append("自動判別での読み込み結果が合わない（%s）" % s4)
    if ovf() != 0:
        fails.append(f"CSVの自動判別で横溢れ {ovf()}px")
    if "何かの薬" in pg.inner_text('#mymsg') or "文字化け" in pg.inner_text('#mymsg'):
        fails.append("レセコンCSVの結果表示がおかしい")
    pg.locator('.mystore').nth(2).locator('[data-act="imp"]').click(); pg.wait_for_timeout(150)
    pg.locator('.mystore').nth(2).locator('[data-act="zero"]').click(); pg.wait_for_timeout(150)  # 既定に戻す
    for _ in range(2):   # 検査用の店舗を消す（2回押しで削除）
        pg.locator('.mystore').nth(2).locator('[data-act="del"]').click(); pg.wait_for_timeout(200)
    if pg.evaluate("() => MY.stores.length") != 2:
        fails.append("店舗の削除（2回押し）が効かない")
    pg.locator('.mystore').nth(1).locator('[data-act="imp"]').click(); pg.wait_for_timeout(150)
    pg.fill('.mytxt', "%s\n%s" % (yjs["b"][0], yjs["b"][1]))
    pg.locator('.mystore').nth(1).locator('[data-act="add"]').click(); pg.wait_for_timeout(250)
    # 店舗名を変えると、保存されたことが表示されること
    inp = pg.locator('.mystore').nth(0).locator('.myname')
    inp.fill("本店"); inp.dispatch_event('change'); pg.wait_for_timeout(200)
    if "保存済み" not in (pg.inner_text('#mymsg') if pg.locator('#mymsg').count() else ""):
        fails.append("店舗名を変えても「保存済み」と出ない")
    st = pg.evaluate("() => MY.stores.map(s => [s.name, s.items.slice().sort()])")
    want0 = sorted([yjs["a"][0], yjs["b"][0], yjs["nm"][1], yjs["a"][1]])
    if len(st) != 2 or st[0][1] != want0 or st[1][1] != sorted(yjs["b"]):
        fails.append("マイ薬局の一覧読み込みが合わない（%s）" % [[x[0], len(x[1])] for x in st])
    small = pg.evaluate(f"""() => [...document.querySelectorAll('#my .myb,#my .mycb,#my .myname,#myx')]
        .filter(e => e.getBoundingClientRect().height>0 && e.getBoundingClientRect().height < {MIN_TAP})
        .map(e => e.className||e.id)""")
    if small:
        fails.append("マイ薬局の設定でタップ領域が小さい：%s" % "／".join(sorted(set(small))[:3]))
    if ovf() != 0:
        fails.append(f"マイ薬局の設定で横溢れ {ovf()}px")
    exp_all = exp_one = None
    try:
        with pg.expect_download(timeout=8000) as dl:
            pg.click('#myall')
        exp_all = "/tmp/_dc_all.json"; dl.value.save_as(exp_all)
        with pg.expect_download(timeout=8000) as dl:
            pg.locator('.mystore').nth(1).locator('[data-act="exp"]').click()
        exp_one = "/tmp/_dc_one.json"; dl.value.save_as(exp_one)
    except Exception as e:
        fails.append("設定ファイルの書き出しに失敗：%s" % str(e)[:50])
    if exp_one:
        import json as _json
        one = _json.load(open(exp_one, encoding="utf-8"))
        if [x["name"] for x in one.get("stores", [])] != ["駅前店"]:
            fails.append("店舗別の書き出しに、ほかの店舗が混じっている")
        if len(_json.load(open(exp_all, encoding="utf-8")).get("stores", [])) != 2:
            fails.append("全店舗の書き出しが2店舗になっていない")
    pg.click('#myx'); pg.wait_for_timeout(400)
    # 切り替えは「すべての薬・全店舗・本店・駅前店」。手で作った直後は「すべての薬」のまま
    if pg.locator('#mybar.on .myv').all_inner_texts()[:2] != ["すべての薬", "全店舗"] or \
            pg.locator('#mybar.on .myv').count() != 4:
        fails.append("店舗の切り替え（すべての薬・全店舗・各店舗）が出ていない")
    total_all = pg.evaluate("() => R.length")
    if num(pg.inner_text("#cnt")) != total_all:
        fails.append("「すべての薬」なのに件数が絞られている（%s）" % pg.inner_text("#cnt"))
    pg.locator('#mybar .myv').nth(1).click(); pg.wait_for_timeout(450)       # 全店舗
    if num(pg.inner_text("#cnt")) != 5:
        fails.append("全店舗を選んでも採用薬5件にならない（%s）" % pg.inner_text("#cnt"))
    both = pg.evaluate("""(y) => { const c=[...document.querySelectorAll('.card')]
        .find(e => ((e.querySelector('.nm')||{}).textContent||'').includes(R.find(r=>r[4]===y)[0]));
        return c ? c.querySelectorAll('.nmmy .mym').length : -1; }""", yjs["b"][0])
    if both != 2:
        fails.append("2店舗で採用している薬の印が %d個（期待 2）" % both)
    pg.locator('#mybar .myv').nth(3).click(); pg.wait_for_timeout(450)       # 駅前店
    if num(pg.inner_text("#cnt")) != 2:
        fails.append("駅前店に切り替えても2件にならない（%s）" % pg.inner_text("#cnt"))
    if ovf() != 0:
        fails.append(f"マイ薬局の表示で横溢れ {ovf()}px")
    card = pg.locator(".card").first
    card.click(); pg.wait_for_timeout(200)
    before = pg.evaluate("() => MY.stores[0].items.length")
    card.locator('.myt').nth(0).click(); pg.wait_for_timeout(250)
    after = pg.evaluate("() => MY.stores[0].items.length")
    if abs(after - before) != 1 or "open" not in (card.get_attribute("class") or ""):
        fails.append("カードの「採用にする」が効かない、またはカードが閉じる")
    card.locator('.myt').nth(0).click(); pg.wait_for_timeout(250)   # 元に戻す
    # 店舗を選んでいるとき：打った検索語は採用薬だけ。採用薬に無ければ採用薬以外を出して知らせる
    un = pg.evaluate("""() => { const mine=new Set(MY.stores.flatMap(s=>s.items));
        const r=R.find(r=>r[4]&&r[4].length===12&&!mine.has(r[4])&&R.filter(x=>x[0]===r[0]).length===1);
        return r[0]; }""")
    pg.fill("#q", un); pg.wait_for_timeout(500)
    if num(pg.inner_text("#cnt")) < 1 or "採用薬には見つからない" not in pg.inner_text("#mypeek"):
        fails.append("採用薬に無い薬を検索したとき、採用薬以外を出していない")
    pg.fill("#q", ""); pg.wait_for_timeout(400)
    if pg.locator('#mypeek.on').count():
        fails.append("検索語を消しても「採用薬以外も表示」の帯が残る")
    # お知らせ：採用薬に関係するものだけ。切替候補の段が出て、数が独立計算と合う
    pg.locator('#mybar .myv').nth(1).click(); pg.wait_for_timeout(350)       # 全店舗
    pg.click('.ptab[data-p="board"]'); pg.wait_for_timeout(500)
    ms = pg.evaluate("""() => {
      const mine=new Set(MY.stores.flatMap(s=>s.items));
      const byName={}; R.forEach(r=>{byName[r[0]]=r;});
      const rows=[...document.querySelectorAll('#vtlist .chgrow,#chglist .chgrow')]
        .map(e=>byName[e.dataset.nm]).filter(Boolean);
      const want=(sc)=>R.filter(r=>mine.has(r[4])&&r[8]===sc&&r[4].length===12&&
        (D.k[r[5]]==='内用薬'||D.k[r[5]]==='外用薬')).length;
      return {bad: rows.filter(r=>!mine.has(r[4])).length,
              hidden: document.getElementById('mswsec').hidden,
              lab: [...document.querySelectorAll('#mswsub .chgs')].map(b=>b.textContent),
              want: [`供給停止（${want(2)}）`, `限定出荷（${want(1)}）`]}; }""")
    if ms["bad"]:
        fails.append("店舗を選んでも、採用していない薬が %d件残る" % ms["bad"])
    # 切替候補・残りわずかの候補の薬に、出荷状況のバッジが付いていること
    nob = pg.evaluate("""() => {
      let n=0;
      for(const v of ['2','1']){
        document.querySelector(`#mswsub .chgs[data-v="${v}"]`).click();
        n+=[...document.querySelectorAll('#mswlist .mswrow .lowi:not(.mswh)')]
            .filter(e=>!e.querySelector('.cmst')).length;
      }
      n+=[...document.querySelectorAll('#lowlist .lowi[data-nm]')].filter(e=>!e.querySelector('.cmst')).length;
      return n; }""")
    if nob:
        fails.append("候補の薬に出荷状況のバッジが無いものが %d件ある" % nob)
    pg.wait_for_timeout(200)
    if ms["hidden"] or ms["lab"] != ms["want"]:
        fails.append("採用薬の切替候補が合わない（表示 %s / 期待 %s）" % (ms["lab"], ms["want"]))
    # 切替候補の「未採用」の薬をタップすると、採用薬以外も含めた検索に移ること
    for v in ("2", "1"):
        pg.click(f'#mswsub .chgs[data-v="{v}"]'); pg.wait_for_timeout(250)
        if pg.locator('#mswlist .lowi:has(.unad)').count():
            break
    if pg.locator('#mswlist .lowi:has(.unad)').count():
        nm = pg.locator('#mswlist .lowi:has(.unad)').first.get_attribute("data-nm")
        pg.locator('#mswlist .lowi:has(.unad)').first.click(); pg.wait_for_timeout(550)
        first = pg.locator(".card .nm").first.inner_text() if pg.locator(".card").count() else ""
        if nm not in first or "採用薬以外も" not in pg.inner_text("#mypeek"):
            fails.append("未採用の候補へ移っても表示されない（%s）" % first[:20])
        pg.click("#mypeekx"); pg.wait_for_timeout(400)
        pg.click("#clr"); pg.wait_for_timeout(350)
    if ovf() != 0:
        fails.append(f"採用薬の切替候補で横溢れ {ovf()}px")
    small = pg.evaluate(f"""() => [...document.querySelectorAll('#mybar .myv,#mypeek button,#mswsub .chgs')]
        .filter(e => e.getBoundingClientRect().height>0 && e.getBoundingClientRect().height < {MIN_TAP})
        .map(e => e.className||e.id)""")
    if small:
        fails.append("マイ薬局の切り替えでタップ領域が小さい：%s" % "／".join(sorted(set(small))[:3]))
    pg.locator('#mybar .myv').nth(0).click(); pg.wait_for_timeout(350)       # すべての薬
    pg.click('.ptab[data-p="board"]'); pg.wait_for_timeout(400)
    if not pg.evaluate("document.getElementById('mswsec').hidden"):
        fails.append("「すべての薬」なのに採用薬の切替候補が出ている")
    # 再読み込みで残り、消してから設定ファイルで戻せること
    pg.reload(); pg.wait_for_timeout(2600)
    if pg.evaluate("() => MY.stores.length") != 2:
        fails.append("再読み込みでマイ薬局の設定が消える")
    pg.evaluate("() => { try { localStorage.removeItem('myph'); } catch(e) {} }")
    pg.reload(); pg.wait_for_timeout(2600)
    if exp_one:
        pg.click('.chip.myset'); pg.wait_for_timeout(300)
        with pg.expect_file_chooser() as fc:
            pg.click('#myload')
        fc.value.set_files(exp_one); pg.wait_for_timeout(450)
        got = pg.evaluate("() => MY.stores.map(s => [s.name, s.items.length])")
        if got != [["駅前店", 2]]:
            fails.append("設定ファイルの読み込みが合わない（%s）" % got)
        pg.click('#myx'); pg.wait_for_timeout(450)
        # 何も無い端末に店舗別のファイルを入れたら、その店舗の表示から始まる
        if pg.locator('#mybar.on .myv').count() != 2 or num(pg.inner_text("#cnt")) != 2:
            fails.append("設定ファイルを入れた直後に、その店舗の表示になっていない（%s）"
                         % pg.inner_text("#cnt"))
    pg.evaluate("() => { try { localStorage.removeItem('myph'); } catch(e) {} }")
    pg.reload(); pg.wait_for_timeout(2600)

    # 分析タブ：3つの段が出て、数が独立計算と合うこと
    pg.click('.ptab[data-p="ana"]'); pg.wait_for_timeout(600)
    an = pg.evaluate(r"""() => {
      const KI={'内用薬':0,'外用薬':1,'注射薬':2}, on=new Set([0,1]);
      const ok=(r)=>r&&on.has(KI[D.k[r[5]]]);
      const rows=R.filter(ok);
      const iso=(s)=>{const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s); return Date.UTC(+m[1],+m[2]-1,+m[3]);};
      const s=new Date(iso(DATA.date)-6*86400000).toISOString().slice(0,10), e=DATA.date;
      const net=new Map();
      (DATA.ev||[]).forEach(x=>{ if(x[1]<s||x[1]>e||!ok(R[x[0]])) return;
        const o=net.get(x[0]); if(o) o.to=x[3]; else net.set(x[0],{from:x[2],to:x[3]}); });
      const a=[...net.values()].filter(o=>o.from!==o.to);
      return {
        total: rows.length, bad: rows.filter(r=>r[8]>0).length,
        stop: a.filter(o=>o.to===2).length,
        lim: a.filter(o=>o.to===1&&o.from===0).length,
        back: a.filter(o=>o.to===0).length,
        rc: (DATA.rcv||[]).filter(x=>ok(R[x[0]])).length,
        kpi: [...document.querySelectorAll('#wkbody .kpi b')].map(b=>parseInt(b.textContent)),
        st: document.getElementById('stcnt').textContent,
        rccnt: document.getElementById('rccnt').textContent,
        bars: document.querySelectorAll('#stbody .sbr').length,
        hist: document.querySelectorAll('#rcbody .hb').length,
      };
    }""")
    if an["kpi"][:3] != [an["stop"], an["lim"], an["back"]]:
        fails.append("週次レポートの件数が合わない（表示 %s / 期待 %s）"
                     % (an["kpi"][:3], [an["stop"], an["lim"], an["back"]]))
    if an["st"] != "{:,}品目".format(an["total"]):
        fails.append("供給不足の構造の品目数が合わない（%s / 期待 %d）" % (an["st"], an["total"]))
    if an["rc"] and an["rccnt"] != "回復 {:,}件".format(an["rc"]):
        fails.append("回復までの日数の件数が合わない（%s / 期待 %d）" % (an["rccnt"], an["rc"]))
    if not an["bars"] or (an["rc"] and not an["hist"]):
        fails.append("分析タブの図が描かれていない")
    for d in ("pc", "fm", "kb", "mk", "cls", "price"):
        pg.click(f'#stdim .chgs[data-d="{d}"]'); pg.wait_for_timeout(200)
        if not pg.locator('#stbody .sbr').count():
            fails.append(f"供給不足の構造（{d}）が空")
    # 薬価の帯の合計が全体と一致すること（帯の取りこぼしが無い）
    ssum = pg.evaluate("""() => [...document.querySelectorAll('#stbody .sbn')]
        .reduce((n,e)=>n+parseInt(e.textContent.replace(/,/g,'')),0)""")
    if ssum != an["total"]:
        fails.append("薬価の帯の合計が全体と合わない（%d / %d）" % (ssum, an["total"]))
    pg.click('#sttbl'); pg.wait_for_timeout(200)
    if not pg.locator('#stbody table.anat').count():
        fails.append("供給不足の構造が表に切り替わらない")
    pg.click('#sttbl'); pg.wait_for_timeout(200)
    pg.click('#wkper .chgs[data-p="30"]'); pg.wait_for_timeout(250)
    pg.click('#wkper .chgs[data-p="7"]'); pg.wait_for_timeout(250)
    if ovf() != 0:
        fails.append(f"分析タブで横溢れ {ovf()}px")
    anx = pg.evaluate(f"""() => {{
      const sec=[...document.querySelectorAll('#ana .bdsec')];
      const out=[], small=[];
      sec.forEach(s=>{{ const b=s.getBoundingClientRect();
        s.querySelectorAll('*').forEach(el=>{{ const r=el.getBoundingClientRect();
          if(r.width>0 && r.right>b.right+1) out.push(el.className||el.tagName); }}); }});
      document.querySelectorAll('#ana .chgs,#ana .kb2,#ana .myb,#ana .tblbtn').forEach(el=>{{
        const r=el.getBoundingClientRect();
        if(r.width>0 && r.height<{MIN_TAP}) small.push((el.className||el.id)+' h='+Math.round(r.height)); }});
      return {{out:[...new Set(out)], small:[...new Set(small)]}};
    }}""")
    if anx["out"]:
        fails.append("分析タブで枠からはみ出す要素：%s" % "／".join(anx["out"][:4]))
    if anx["small"]:
        fails.append("分析タブのタップ領域が小さい：%s" % "／".join(anx["small"][:4]))
    # 凡例に新機能の説明があること
    pg.click('.ptab[data-p="search"]'); pg.wait_for_timeout(350)
    pg.click("#lgbtn"); pg.wait_for_timeout(400)
    lg = pg.inner_text("#lgbody")
    for t in ("切替余地", "残りわずかの成分", "マイ薬局", "週次レポート", "回復までの日数", "供給不足の構造"):
        if t not in lg:
            fails.append(f"凡例に「{t}」の説明が無い")
    pg.click("#lgx"); pg.wait_for_timeout(300)
    pg.click('.ptab[data-p="board"]'); pg.wait_for_timeout(400)

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

    # 規制区分バッジ。regulation.json が置かれていれば、代表品目に正しく付くこと。
    # 凡例と表示の食い違い（凡例に無い略号が出る等）もここで見る。
    reg = pg.evaluate("() => (DATA.reg||{}).available ? DATA.reg.codes : null")
    if reg:
        for q, want in (("オキノーム散２．５", ["劇", "麻", "処"]),
                        ("ニンラーロカプセル２．３", ["毒", "処"]),
                        ("コンサータ錠１８", ["劇", "向1", "処"])):
            pg.fill("#q", q)
            pg.wait_for_timeout(420)
            if pg.locator(".card").count() == 0:
                fails.append(f"規制区分の検証薬（{q}）が見つからない")
                continue
            got = pg.locator(".card").first.locator(".rg").all_inner_texts()
            if got != want:
                fails.append(f"{q} の規制区分が {got}（期待 {want}）")
        pg.fill("#q", "")
        pg.wait_for_timeout(250)
        pg.click("#lgbtn")
        pg.wait_for_timeout(400)
        shown = pg.locator("#lgbody .rg").all_inner_texts()
        if shown != reg:
            fails.append("凡例の規制区分がデータと食い違う（%s）" % "／".join(shown))
        pg.click("#lgx")
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
