#!/usr/bin/env python3
"""
医薬品供給状況ページ 端末横断チェック

Android（Pixel / Galaxy）と iOS（iPhone）の実機定義で、
表示崩れ・機能・タップ領域を検証する。

  python3 devicecheck.py <検索HTMLのパス>
"""
import sys, os
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
    if num(pg.inner_text("#cnt")) != 14550:
        fails.append("通常出荷が14,550件でない")
    if ovf() != 0:
        fails.append(f"チップ選択で横溢れ {ovf()}px")
    pg.click("#clrall")
    pg.wait_for_timeout(250)

    # 製品区分
    pg.click("#advbtn")
    pg.wait_for_timeout(200)
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

    pg.select_option("#ffm", "錠")
    pg.wait_for_timeout(380)
    if ovf() != 0:
        fails.append(f"剤形（詳細）で横溢れ {ovf()}px")
    pg.select_option("#ffm", "")
    pg.wait_for_timeout(250)
    pg.select_option("#fpc", "後発品")
    pg.wait_for_timeout(380)
    if num(pg.inner_text("#cnt")) != 7225:
        fails.append("後発品が7,225件でない")
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
      const sels=['#q','.sm','.chip','#advbtn','.cmpbtn','#logic','.sortb','#ovx','#clrall','#lgbtn','#lgx','#rvbtn','#rvx','#rvk .sortb','#rvq','#rvord','#rvgap .sortb','#ordbtn','.ingbtn','.ptab','.nwt','#nwsort'];
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
