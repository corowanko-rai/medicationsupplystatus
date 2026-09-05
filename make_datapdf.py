#!/usr/bin/env python3
"""
「データの成り立ち.html」から PDF を作る。

  python3 make_datapdf.py [入力.html] [出力.pdf]

HTMLは4つのタブで切り替える作りだが、紙では切り替えられないので
4章すべてを開いた状態にし、章の頭で改ページする。

必要なもの（この処理のためだけに使う。自動更新には不要）:
  pip install playwright
  playwright install chromium

自動更新のワークフローからは呼んでいない。
表示を直したときに手元で作り直すためのもの。
"""
import sys, os, pathlib

HERE = os.path.dirname(os.path.abspath(__file__))
# 既定は「値入り」の資料。build_html.py が生成時に書き出す。
# 元の「データの成り立ち.html」は数値が {{目印}} のままなので、
# そのままPDFにすると目印が印刷されてしまう。
_FILLED = os.path.join(HERE, "データの成り立ち_値入り.html")
_RAW = os.path.join(HERE, "データの成り立ち.html")
SRC = sys.argv[1] if len(sys.argv) > 1 else (_FILLED if os.path.exists(_FILLED) else _RAW)
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "データの成り立ち.pdf")

# 章の見出し（HTML側のタブと同じ順序・同じ id）
NAMES = {"venn": "1. 区分の関係", "map": "2. 出典の対応",
         "join": "3. 結合の仕組み", "calc": "4. 計算の根拠"}

PREP = """(names) => {
  document.querySelector('.tabs').style.display='none';
  Object.keys(names).forEach((k,i)=>{
    const s=document.getElementById(k);
    if(!s) return;
    s.classList.add('on');
    // 章ごとに改ページすると、章末に大きな余白が残る（実測で最大85%）。
    // 紙では見出しで区切りが分かるので、改ページは強制しない。
    s.style.pageBreakBefore = 'auto';
    s.style.marginTop = i ? '18px' : '0';
    const b=document.createElement('div');
    b.textContent=names[k];
    b.className='chapmark';
    b.style.cssText='font-size:11px;font-weight:800;letter-spacing:.08em;'
      +'color:#6b7a88;margin:0 0 6px;padding-bottom:6px;'
      +'border-bottom:2px solid #d8e0e8';
    s.insertBefore(b, s.firstChild);
  });
  // 表は原則として1ページに収める（分割すると見出しと本体が離れる）。
  // ただし行数の多い表まで丸ごと送ると、前のページに大きな空白が残る。
  // 行数で線を引き、長い表だけ分割を許す。行数は紙面の幅に依らないので
  // ブラウザ上で数えても結果がぶれない。
  const LONG_ROWS = 7;
  document.querySelectorAll('.tw').forEach(w=>{
    const n = w.querySelectorAll('tbody tr').length
           || Math.max(0, w.querySelectorAll('tr').length - 1);
    if (n > LONG_ROWS) {
      w.style.breakInside = 'auto';
      w.style.pageBreakInside = 'auto';
      const th = w.querySelector('thead');
      if (th) th.style.display = 'table-header-group';  // 続きにも見出しを出す
    }
  });
  const st=document.createElement('style');
  st.textContent='@page{size:A4;margin:11mm 11mm}'
    +'body{background:#fff}'
    /* 囲みと図は途中で切らない。ただし表は「行」単位で切れれば十分で、
       表ごと次ページへ送ると章末に大きな余白ができる。 */
    +'.card,.flow,.dfbox,.dfw{break-inside:avoid;page-break-inside:avoid}'
    /* 桁の図・コードの対応図・コードの比較表は、途中で切れると
       意味が取れなくなるので必ず1ページに収める */
    +'.cdx,.cgrid{break-inside:avoid;page-break-inside:avoid}'
    /* 桁の図は縦に長い。丸ごと1ページに収めようとすると
       前のページに大きな空白が出るので、
       「12桁の並び＋桁番号」だけを分割禁止にし、凡例は流す。
       意味が取れなくなる切れ方はこれで防げる。 */
    +'.dgrow{break-inside:avoid;page-break-inside:avoid}'
    +'.dg .dgrow:first-child{break-after:avoid;page-break-after:avoid}'
    +'.dgl li{break-inside:avoid}'
    +'tr,li{break-inside:avoid;page-break-inside:avoid}'
    /* 見出しだけがページ末尾に取り残されないようにする。
       表が複数ページに渡るときは、各ページの先頭に見出し行を繰り返す。 */
    /* Chromium は thead の break-after:avoid を無視するため、
       表そのものを分割対象から外す。1ページに収まらない表は
       ブラウザが自動で分割するので、はみ出す心配はない。
       分割された場合に備えて見出し行は各ページで繰り返す。 */
    +'.tw{break-inside:avoid;page-break-inside:avoid}'
    /* 章の見出しと小見出しは、直後の内容と切り離さない。
       これが無いと、見出しだけがページ末尾に取り残される。 */
    +'h2,h3,.dgnum,.chapmark{break-after:avoid;page-break-after:avoid}'
    +'h2,h3{break-inside:avoid}'
    /* 画面では横スクロールさせている表を、紙では全幅で出す */
    +'.tw{overflow:visible}table{min-width:0}'
    /* 紙は画面より情報を詰められる。読みにくくならない範囲で余白を詰める */
    +'main{padding:0}section{margin:0}'
    +'h2{margin:16px 0 8px}h3{margin:14px 0 6px}'
    +'p{margin:0 0 8px}ul{margin:0 0 8px}'
    +'.card{padding:10px 12px;margin:8px 0}'
    +'.flow{padding:10px 12px;margin:8px 0}'
    +'.tw{margin:8px 0}th,td{padding:6px 9px}'
    +'hr{margin:14px 0}'
    /* ベン図などのSVGは画面幅を前提に大きめ。紙では縮めて、
       囲みごと次ページへ送られて空白が空くのを防ぐ */
    +'svg{max-height:210px;height:auto}'
    /* ---- 日本語フォントの指定（重要） ----
       画面用の指定は Hiragino / Yu Gothic / Noto Sans JP を前提にしている。
       これらが入っていない環境で作ると、Chromiumが中国語フォント
       （WenQuanYi Zen Hei など）へ落ちてしまい、漢字が中国語字形になる。
       日本語のCJKフォントを明示して、かな・カナ・漢字を正しく出す。
       日本語フォントが入っている環境では先頭の指定がそのまま使われる。 */
    +'body,body *{font-family:-apple-system,"Hiragino Sans","Hiragino Kaku Gothic ProN",'
    +'"Yu Gothic UI","Yu Gothic","Meiryo","IPAPGothic","IPAexGothic",'
    +'"Noto Sans CJK JP","Noto Sans JP",sans-serif !important}'
    /* 等幅の図はCJK対応の等幅フォントに。和欧の幅が揃い、罫線がずれない */
    +'pre,code,.flow pre{font-family:ui-monospace,Menlo,Consolas,"IPAGothic",'
    +'"Noto Sans Mono CJK JP",monospace !important}';
  document.head.appendChild(st);
}"""


# PDFに実際に埋め込まれたフォントを調べる。
# 中国語・韓国語のフォントが混ざっていたら字形が変わるので警告する。
NG_FONTS = ("WenQuanYi", "CJKSC", "CJK SC", "CJKTC", "CJK TC",
            "CJKKR", "CJK KR", "CJKHK", "CJK HK", "Droid Sans Fallback",
            "Noto Sans SC", "Noto Sans TC", "Noto Sans KR")


def embedded_fonts(path):
    from pypdf import PdfReader
    out = set()
    for pg in PdfReader(path).pages:
        res = pg.get("/Resources")
        if res is None:
            continue
        if hasattr(res, "get_object"):
            res = res.get_object()
        fonts = res.get("/Font")
        if fonts is None:
            continue
        if hasattr(fonts, "get_object"):
            fonts = fonts.get_object()
        for v in fonts.values():
            bf = v.get_object().get("/BaseFont")
            if bf:
                out.add(str(bf).split("+")[-1])
    return sorted(out)


def main():
    if not os.path.exists(SRC):
        print(f"ERROR: {SRC} がありません")
        return 1
    if "{{" in open(SRC, encoding="utf-8").read():
        print(f"ERROR: {os.path.basename(SRC)} に未置換の目印が残っています。")
        print("  先に fetch_update.py を実行して"
              "「データの成り立ち_値入り.html」を作ってください。")
        return 1
    from playwright.sync_api import sync_playwright
    url = pathlib.Path(SRC).resolve().as_uri()
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": 900, "height": 1200})
        pg.goto(url)
        pg.wait_for_timeout(1200)
        pg.evaluate(PREP, NAMES)
        pg.wait_for_timeout(300)
        pg.emulate_media(media="print")
        pg.pdf(path=OUT, format="A4", print_background=True,
               prefer_css_page_size=True)
        b.close()
    fonts = embedded_fonts(OUT)
    print(f"作成しました: {OUT}（{os.path.getsize(OUT):,} bytes）")
    print("埋め込みフォント: " + ", ".join(fonts))
    bad = [f for f in fonts if any(x.lower() in f.lower() for x in NG_FONTS)]
    if bad:
        print("警告: 日本語以外のCJKフォントが混ざっています → "
              + ", ".join(bad))
        print("  漢字が中国語・韓国語の字形になっている可能性があります。")
        print("  日本語フォント（Noto Sans CJK JP など）を入れて作り直してください。")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
