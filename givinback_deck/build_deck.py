# -*- coding: utf-8 -*-
"""
ジンジブ様_CoreUniq オンボーディング・POC 資料
Givin' Back 標準フォーマット（ネイビー×コーラル／ロゴ右上／Confidentialフッター）で生成する。

素材:
  logo_gb.png / logo_gb_white.png : Drive「Givin' Backロゴ」から切り出し
  lock.png / lockw.png            : ヘッダー用 横組みロックアップ
"""
import copy
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------- brand tokens
NAVY = RGBColor(0x14, 0x4B, 0x7F)   # 基調（見出し・帯）
NAVY_DP = RGBColor(0x0E, 0x35, 0x5C)  # 濃ネイビー（グラデ端／表紙）
INK = RGBColor(0x1E, 0x29, 0x3B)   # 本文
GRAY = RGBColor(0x6B, 0x77, 0x87)   # 補足
GRAY_L = RGBColor(0x93, 0x9E, 0xAC)   # フッター
CORAL = RGBColor(0xF9, 0x61, 0x67)   # アクセント
CORAL_D = RGBColor(0xE0, 0x45, 0x4B)
PINK_L = RGBColor(0xFD, 0xEC, 0xED)   # 淡ピンク面
BLUE_L = RGBColor(0xEE, 0xF3, 0xF8)   # 淡ブルー面
LINE = RGBColor(0xD8, 0xE0, 0xEA)   # 罫線
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

JP_FONT = "Meiryo"

SW, SH = 13.333, 7.5           # スライドサイズ（inch）
ML, MR = 0.62, 12.71           # 左右マージン
CW = MR - ML                   # コンテンツ幅

FOOTER = "© Givin' Back Inc. All Rights Reserved.  |  Confidential"

LOGO = os.path.join(HERE, "lock.png")
LOGO_W = os.path.join(HERE, "lockw.png")


# ------------------------------------------------------------------- utilities
def _set_font(run, size, bold, color, font=JP_FONT):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = font
    # 日本語グリフにも同じフォントを割り当てる
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.makeelement(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag.split(":")[1],
            {"typeface": font},
        )
        rPr.append(el)


def text(slide, x, y, w, h, lines, size=12, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.25, space_after=0):
    """lines: str か [(文字列, {size/bold/color/font}), ...] の段落リスト"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    if isinstance(lines, str):
        lines = [lines]
    for i, item in enumerate(lines):
        if isinstance(item, tuple):
            content, opt = item
        else:
            content, opt = item, {}
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = opt.get("align", align)
        p.line_spacing = opt.get("spacing", spacing)
        if space_after:
            p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = content
        _set_font(run, opt.get("size", size), opt.get("bold", bold),
                  opt.get("color", color), opt.get("font", JP_FONT))
    return box


def shape(slide, x, y, w, h, kind=MSO_SHAPE.RECTANGLE, fill=None,
          line=None, line_w=1.0, adj=None, shadow=False):
    sp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if adj is not None:
        try:
            sp.adjustments[0] = adj
        except (IndexError, ValueError):
            pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    if not shadow:
        sp.shadow.inherit = False
    sp.text_frame.word_wrap = True
    return sp


def gradient_bg(slide):
    """表紙・裏表紙用のネイビー→コーラル斜めグラデーション"""
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SW), Inches(SH))
    sp.line.fill.background()
    sp.shadow.inherit = False
    fill = sp.fill
    fill.gradient()
    fill.gradient_angle = 25.0
    stops = fill.gradient_stops
    stops[0].color.rgb = NAVY_DP
    stops[0].position = 0.0
    stops[1].color.rgb = CORAL
    stops[1].position = 1.0
    return sp


def logo(slide, x, y, h, white=False):
    path = LOGO_W if white else LOGO
    return slide.shapes.add_picture(path, Inches(x), Inches(y), height=Inches(h))


def header(slide, title, lead, page):
    """Givin' Back 標準ヘッダー＋リード＋フッター"""
    text(slide, ML, 0.40, CW - 1.5, 0.42, title, size=21, bold=True, color=NAVY)
    logo(slide, MR - 1.24, 0.40, 0.26)
    shape(slide, ML, 0.95, CW, 0.014, fill=LINE)

    if lead:
        shape(slide, ML, 1.16, 0.055, 0.30, fill=CORAL)
        text(slide, ML + 0.20, 1.14, CW - 0.20, 0.34, lead,
             size=14.5, bold=True, color=NAVY)

    shape(slide, ML, 7.02, CW, 0.012, fill=LINE)
    text(slide, ML, 7.13, CW - 0.35, 0.20, FOOTER, size=7.5, color=GRAY_L,
         align=PP_ALIGN.RIGHT)
    text(slide, MR - 0.32, 7.11, 0.32, 0.22, str(page), size=9, bold=True,
         color=NAVY, align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, fill=WHITE, line=LINE):
    return shape(slide, x, y, w, h, MSO_SHAPE.ROUNDED_RECTANGLE,
                 fill=fill, line=line, line_w=1.0, adj=0.035)


def card_title(slide, x, y, w, label, sub=None):
    """カード見出し（ネイビー帯 + 白抜き文字）"""
    shape(slide, x, y, w, 0.40, MSO_SHAPE.ROUNDED_RECTANGLE, fill=NAVY, adj=0.25)
    text(slide, x + 0.22, y + 0.075, w - 0.44, 0.26, label, size=12, bold=True,
         color=WHITE)
    if sub:
        text(slide, x, y + 0.46, w, 0.24, sub, size=9.5, color=GRAY)


def num_badge(slide, x, y, d, n, fill=CORAL):
    shape(slide, x, y, d, d, MSO_SHAPE.OVAL, fill=fill)
    text(slide, x, y + d * 0.16, d, d * 0.7, str(n), size=13, bold=True,
         color=WHITE, align=PP_ALIGN.CENTER)


def check_row(slide, x, y, w, label, size=11.5):
    text(slide, x, y, 0.22, 0.26, "✓", size=11.5, bold=True, color=CORAL)
    text(slide, x + 0.26, y, w - 0.26, 0.26, label, size=size, color=INK)


def notes(slide, body):
    slide.notes_slide.notes_text_frame.text = body


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ------------------------------------------------------------------ build deck
prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)

SRC = "[Sources]\n- Givin' Back社内資料・受講前アンケート\n\n[Talk Track]\n"

# ---------------------------------------------------------------- 1 : 表紙
s = new(prs)
gradient_bg(s)
logo(s, ML, 0.62, 0.42, white=True)
text(s, ML, 1.85, 9.6, 0.34, "ジンジブ様｜CoreUniq オンボーディング・POC",
     size=13, bold=True, color=RGBColor(0xFF, 0xD9, 0xDB))
text(s, ML, 2.42, 10.6, 2.1,
     [("若手育成を、研修だけで終わらせない。", {"size": 40, "bold": True}),
      ("若手が、自分の強みを仕事の成果につなげる組織へ。", {"size": 26, "bold": True})],
     color=WHITE, spacing=1.22)
shape(s, ML, 4.72, 1.5, 0.05, fill=WHITE)
text(s, ML, 5.02, 10.2, 0.9,
     [("キャリア開発から、現場での挑戦・役割形成までを一体で支援", {"size": 15, "bold": True}),
      ("自分を知る。仕事で試す。挑戦と貢献を、継続的に広げる。", {"size": 13})],
     color=RGBColor(0xEF, 0xF4, 0xFA), spacing=1.5)
text(s, ML, 6.62, 6.0, 0.3, "導入意図・検証方針の目線合わせ", size=11,
     color=RGBColor(0xE6, 0xEE, 0xF7))
text(s, 7.6, 6.62, 5.1, 0.3,
     "Givin' Back株式会社　|　ONBOARDING GUIDE 2026", size=10,
     color=RGBColor(0xE6, 0xEE, 0xF7), align=PP_ALIGN.RIGHT)
notes(s, SRC +
      "本日は操作説明だけではなく、CoreUniqを導入する意図と、今回一緒に検証したいことを共有します。\n"
      "CoreUniqは日報確認ツールではなく、キャリア開発を仕事上の実践につなげる継続支援です。")

# --------------------------------------------- 2 : TODAY 本日の目的とゴール
s = new(prs)
header(s, "TODAY｜本日の目的とゴール",
       "操作説明だけで終わらず、導入意図と検証方針を合わせる", 2)

card(s, ML, 1.78, 5.35, 4.05, fill=BLUE_L)
card_title(s, ML + 0.28, 2.04, 4.79, "本日のゴール")
goals = [("1", "CoreUniqが生まれた背景と", "支援したい課題を共有する"),
         ("2", "ジンジブ様社内のPOCで", "検証する内容を合意する")]
for i, (n, l1, l2) in enumerate(goals):
    y = 2.78 + i * 1.32
    num_badge(s, ML + 0.30, y, 0.44, n)
    text(s, ML + 0.94, y - 0.02, 4.10, 0.86,
         [(l1, {}), (l2, {})], size=12.5, bold=True, color=INK, spacing=1.35)

x2 = ML + 5.62
card(s, x2, 1.78, CW - 5.62, 4.05)
card_title(s, x2 + 0.30, 2.04, CW - 5.62 - 0.60, "本日の3つの論点")
points = [("導入意図", "何を解決するサービスか"),
          ("利用設計", "本人と管理職がどう使うか"),
          ("検証方針", "何を学び、次に活かすか")]
for i, (k, v) in enumerate(points):
    y = 2.82 + i * 0.94
    shape(s, x2 + 0.30, y, 1.72, 0.42, MSO_SHAPE.ROUNDED_RECTANGLE,
          fill=PINK_L, line=CORAL, line_w=0.75, adj=0.28)
    text(s, x2 + 0.30, y + 0.09, 1.72, 0.26, k, size=11.5, bold=True,
         color=CORAL_D, align=PP_ALIGN.CENTER)
    text(s, x2 + 2.14, y + 0.02, 0.34, 0.3, "→", size=13, bold=True, color=GRAY_L)
    text(s, x2 + 2.58, y + 0.08, CW - 5.62 - 2.90, 0.3, v, size=12.5, color=INK)

text(s, ML, 6.10, CW, 0.34,
     "最後に、対象者・期間・振り返り方法・次回日程まで確認します。",
     size=11.5, bold=True, color=NAVY)
notes(s, SRC +
      "今日のゴールは、導入意図・利用設計・検証方針の3点を合わせることです。\n"
      "最後に、対象者、期間、振り返り方法、次回日程まで確認します。")

# --------------------------------------------------- 3 : AGENDA タイムテーブル
s = new(prs)
header(s, "AGENDA｜本日の進め方（60分）",
       "説明は前半に集約し、後半はデモと議論に時間を使う", 3)

rows = [("00-05", "本日の目的・前提合わせ", "操作説明ではなく、導入意図と検証方針の共有であると目線を合わせる"),
        ("05-15", "CoreUniqが生まれた背景", "日報ツールではなく、キャリア開発研修＋実践伴走であることを明確に伝える"),
        ("15-25", "若手の課題とアンケートデータ", "目標から行動への翻訳が途切れる構造を提示し、課題感をすり合わせる"),
        ("25-35", "POCの目的・検証したいこと", "検証テーマを明示し、将来の展開判断に向けた材料づくりを共有する"),
        ("35-48", "メンバー・管理職画面のデモ", "機能一覧ではなく、行動設定→実践→振り返り→FBの一連の流れを見せる"),
        ("48-56", "ディスカッション", "検証項目の追加や画面UIについてフィードバックをいただく"),
        ("56-60", "次回アクションの合意", "対象者・期間・着眼点・次回日程の4点を確定させる")]

hy = 1.78
shape(s, ML, hy, CW, 0.40, fill=NAVY)
for cx, cw, label, al in [(ML + 0.20, 1.05, "時間", PP_ALIGN.LEFT),
                          (ML + 1.42, 3.30, "アジェンダ", PP_ALIGN.LEFT),
                          (ML + 4.92, CW - 5.12, "目的・進行のポイント", PP_ALIGN.LEFT)]:
    text(s, cx, hy + 0.09, cw, 0.26, label, size=11, bold=True, color=WHITE, align=al)

for i, (t, a, p) in enumerate(rows):
    y = hy + 0.40 + i * 0.60
    if i % 2 == 0:
        shape(s, ML, y, CW, 0.60, fill=BLUE_L)
    shape(s, ML, y + 0.595, CW, 0.008, fill=LINE)
    text(s, ML + 0.20, y + 0.17, 1.15, 0.28, t, size=11, bold=True, color=CORAL_D)
    text(s, ML + 1.42, y + 0.17, 3.35, 0.30, a, size=11.5, bold=True, color=NAVY)
    text(s, ML + 4.92, y + 0.18, CW - 5.12, 0.30, p, size=10.5, color=INK)

text(s, ML, 6.56, CW, 0.30,
     "デモと議論に25分を確保し、その場で検証項目を追加できる進行にします。",
     size=11.5, bold=True, color=NAVY)
notes(s, SRC +
      "前半は導入意図と課題認識の共有、後半はデモとディスカッションに時間を配分しています。\n"
      "残り4分になったら、対象者・期間・着眼点・次回日程の4点の確認に必ず入ります。")

# ------------------------------------------------ 4 : DATA 行動に移れない背景
s = new(prs)
header(s, "DATA｜若手が行動に移れない背景",
       "悩みは『社内での方向性』と『強みの活かし方』に集中している", 4)

card(s, ML, 1.78, 6.72, 4.15)
card_title(s, ML + 0.30, 2.04, 6.12, "受講前の悩み上位（複数回答）")
bars = [("やりたいことが見えない", 36.7), ("強みの活かし方が不明", 34.7),
        ("自分の軸が分からない", 28.6), ("見つけ方が分からない", 24.5),
        ("始め方が分からない", 20.4)]
bx, bw = ML + 2.92, 3.05
for i, (label, val) in enumerate(bars):
    y = 2.68 + i * 0.60
    text(s, ML + 0.30, y + 0.02, 2.55, 0.28, label, size=11, color=INK)
    shape(s, bx, y + 0.05, bw, 0.24, fill=RGBColor(0xEE, 0xF1, 0xF5))
    shape(s, bx, y + 0.05, bw * val / 40.0, 0.24,
          fill=CORAL if i < 2 else RGBColor(0x6F, 0x9A, 0xC8))
    text(s, bx + bw + 0.10, y + 0.03, 0.72, 0.26, f"{val}%", size=11, bold=True,
         color=NAVY)

x2 = ML + 6.98
card(s, x2, 1.78, CW - 6.98, 4.15, fill=PINK_L, line=RGBColor(0xF6, 0xC8, 0xCA))
card_title(s, x2 + 0.30, 2.04, CW - 6.98 - 0.60, "『協働』と『提案』のギャップ")
metrics = [("周囲との協力", "8.02", "／10"), ("率先した提案", "5.38", "／10")]
for i, (k, v, u) in enumerate(metrics):
    y = 2.66 + i * 0.92
    shape(s, x2 + 0.30, y, CW - 6.98 - 0.60, 0.76, MSO_SHAPE.ROUNDED_RECTANGLE,
          fill=WHITE, line=RGBColor(0xF2, 0xD3, 0xD5), adj=0.12)
    text(s, x2 + 0.52, y + 0.24, 2.0, 0.30, k, size=11.5, color=INK)
    text(s, x2 + 2.40, y + 0.11, 1.4, 0.46, v, size=22, bold=True,
         color=CORAL_D if i else NAVY, align=PP_ALIGN.RIGHT)
    text(s, x2 + 3.82, y + 0.28, 0.6, 0.26, u, size=10, color=GRAY)
check_row(s, x2 + 0.32, 4.62, CW - 6.98 - 0.64, "協力8点以上・提案5点以下が 38.8%")
check_row(s, x2 + 0.32, 5.06, CW - 6.98 - 0.64, "協働の土台はあるが、意思を行動に変えにくい")

shape(s, ML, 6.12, CW, 0.62, MSO_SHAPE.ROUNDED_RECTANGLE, fill=BLUE_L, adj=0.16)
text(s, ML + 0.26, 6.30, CW - 0.52, 0.30,
     "主体性がないのではなく、方向性を言語化し、仕事で試す方法が見えていない可能性があります。",
     size=12, bold=True, color=NAVY)
notes(s, SRC +
      "若手は意欲がないのではなく、強みや方向性を具体的な行動に翻訳する方法が見えていない可能性があります。\n"
      "協力は8.02に対して、率先した提案は5.38です。協働の土台はあっても、自分から動くところにギャップがあります。\n"
      "質問：ジンジブ様社内や支援先企業でも、似た傾向はありますか。")

# ------------------------------------------- 5 : WHAT CoreUniqが支援すること
s = new(prs)
header(s, "WHAT｜CoreUniqが支援すること",
       "研修で終わらず、若手の行動変容と役割形成まで伴走する", 5)

shape(s, ML, 1.82, CW, 1.72, MSO_SHAPE.ROUNDED_RECTANGLE, fill=NAVY, adj=0.10)
text(s, ML + 0.50, 2.12, 3.2, 0.6, "CoreUniq", size=30, bold=True, color=WHITE)
text(s, ML + 4.10, 2.10, CW - 4.60, 1.1,
     [("キャリア開発を起点に、現場での実践・対話・役割形成を", {}),
      ("半年〜年間で支える若手活躍支援プログラムです。", {})],
     size=14, bold=True, color=RGBColor(0xE9, 0xF0, 0xF8), spacing=1.4)

blocks = [("キャリア開発", "強み・価値観・目指す姿を\n自分の言葉で整理する", NAVY),
          ("現場実践", "強みを活かす小さな行動を\n実際の業務で試す", CORAL),
          ("1on1・育成支援", "対話を通じて経験を付与し\n次の役割を設計する", NAVY)]
bw2 = (CW - 1.30) / 3
for i, (title, body, col) in enumerate(blocks):
    x = ML + i * (bw2 + 0.65)
    card(s, x, 3.90, bw2, 2.05)
    shape(s, x, 3.90, bw2, 0.10, fill=col)
    text(s, x + 0.26, 4.22, bw2 - 0.52, 0.34, title, size=14, bold=True, color=col)
    text(s, x + 0.26, 4.74, bw2 - 0.52, 1.0,
         [(l, {}) for l in body.split("\n")], size=11.5, color=INK, spacing=1.35)
    if i < 2:
        text(s, x + bw2 + 0.10, 4.62, 0.45, 0.4, "×", size=18, bold=True,
             color=GRAY_L, align=PP_ALIGN.CENTER)

text(s, ML, 6.22, CW, 0.32,
     "若手本人を主役に、管理職との対話を通じて活躍をつくります。",
     size=12, bold=True, color=NAVY)
notes(s, SRC +
      "CoreUniqの価値は、研修、現場実践、管理職との対話を一つの支援プロセスとしてつなぐことです。\n"
      "研修部分を初期設定や付帯作業として扱わず、若手の自己理解と行動設計をつくる中核サービスとして位置づけています。")

# ---------------------------------------------- 6 : 3 DESIGNS 若手活躍の設計
s = new(prs)
header(s, "3 DESIGNS｜若手活躍を支える設計",
       "若手活躍は、自己理解・実践・役割形成の循環で育つ", 6)

steps = [("1", "自分を、知る", "強み・価値観・目指す姿を、自分の言葉で整理する"),
         ("2", "仕事で、試す", "強みを活かす小さな行動を、実際の業務で実践する"),
         ("3", "役割を、広げる", "対話を通じて、挑戦・役割・周囲への貢献を広げる")]
cw3 = (CW - 1.30) / 3
for i, (n, t, d) in enumerate(steps):
    x = ML + i * (cw3 + 0.65)
    card(s, x, 1.92, cw3, 3.55)
    shape(s, x, 1.92, cw3, 0.10, fill=CORAL if i == 1 else NAVY)
    num_badge(s, x + (cw3 - 0.62) / 2, 2.34, 0.62, n,
              fill=CORAL if i == 1 else NAVY)
    text(s, x + 0.24, 3.24, cw3 - 0.48, 0.42, t, size=17, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER)
    shape(s, x + (cw3 - 1.1) / 2, 3.82, 1.1, 0.04, fill=LINE)
    text(s, x + 0.36, 4.06, cw3 - 0.72, 1.2, d, size=11.5, color=INK,
         align=PP_ALIGN.CENTER, spacing=1.45)
    if i < 2:
        text(s, x + cw3 + 0.10, 3.42, 0.45, 0.4, "→", size=18, bold=True,
             color=GRAY_L, align=PP_ALIGN.CENTER)

shape(s, ML, 5.78, CW, 0.66, MSO_SHAPE.ROUNDED_RECTANGLE, fill=BLUE_L, adj=0.16)
text(s, ML + 0.26, 5.98, CW - 0.52, 0.30,
     "キャリア開発を入口に、本人の行動と職場での貢献が広がる状態をつくります。",
     size=12, bold=True, color=NAVY)
notes(s, SRC +
      "自己理解だけで終わらず、仕事で試し、対話を通じて役割を広げる循環をつくります。\n"
      "システムだけで変化を生むのではなく、研修と人の支援を前提にした設計です。")

# ------------------------------------------------------- 7 : POC 検証したいこと
s = new(prs)
header(s, "POC｜ジンジブ様社内で検証したいこと",
       "完成品の評価ではなく、若手支援の運用仮説を一緒に検証する", 7)

cols = [("1", "社内で試す",
         ["対象となる若手と管理職を設定", "実際の業務の中で継続利用", "利用場面と対話の流れを確認"]),
        ("2", "運用仮説を検証",
         ["本人が行動に落とし込めるか", "管理職の対話に役立つか", "無理なく続けられる頻度か"]),
        ("3", "学びを整理する",
         ["有効だった支援方法を整理", "改善すべき機能・運用を特定", "顧客企業への展開可能性を検討"])]
cw4 = (CW - 0.90) / 3
for i, (n, t, items) in enumerate(cols):
    x = ML + i * (cw4 + 0.45)
    card(s, x, 1.90, cw4, 3.70)
    shape(s, x, 1.90, cw4, 0.10, fill=CORAL if i == 1 else NAVY)
    num_badge(s, x + 0.26, 2.26, 0.46, n, fill=CORAL if i == 1 else NAVY)
    text(s, x + 0.88, 2.32, cw4 - 1.14, 0.34, t, size=14.5, bold=True, color=NAVY)
    shape(s, x + 0.26, 2.96, cw4 - 0.52, 0.03, fill=LINE)
    for j, it in enumerate(items):
        check_row(s, x + 0.26, 3.24 + j * 0.68, cw4 - 0.52, it, size=11)

shape(s, ML, 5.90, CW, 0.66, MSO_SHAPE.ROUNDED_RECTANGLE, fill=PINK_L, adj=0.16)
text(s, ML + 0.26, 6.10, CW - 0.52, 0.30,
     "社内POCの結果をもとに、再現できる若手支援の形を一緒に整理します。",
     size=12, bold=True, color=CORAL_D)
notes(s, SRC +
      "今回のPOCでは、利用率だけではなく、本人の行動、管理職との対話、継続できる運用条件を確認したいと考えています。\n"
      "成功事例を先に約束するのではなく、何が有効かを一緒に見つける検証です。")

# --------------------------------------- 8 : FOR ORGANIZATION 2者と仕組み
s = new(prs)
header(s, "FOR ORGANIZATION｜若手活躍を支える2者と仕組み",
       "若手本人と管理職の対話を、CoreUniqが支える", 8)

roles = [("PERSON 01", "若手本人", "強みを仕事に活かし、自分から役割と貢献を広げていく。",
          ["仕事上の役割が見える", "挑戦と貢献を自分から広げる"], NAVY),
         ("MANAGER 02", "管理職", "本人の強みと実践を踏まえ、経験付与や次の役割を設計する。",
          ["1on1で活躍を後押しする", "本人に合う経験を与えられる"], NAVY),
         ("COREUNIQ 03", "対話を支える記録", "本人の振り返りと管理職のフィードバックを蓄積し、次の対話につなげる。",
          ["本人・管理職のみが閲覧", "1on1の材料を残せる"], CORAL)]
cw5 = (CW - 0.90) / 3
for i, (tag, t, d, items, col) in enumerate(roles):
    x = ML + i * (cw5 + 0.45)
    card(s, x, 1.88, cw5, 3.72, fill=PINK_L if col == CORAL else WHITE,
         line=RGBColor(0xF6, 0xC8, 0xCA) if col == CORAL else LINE)
    shape(s, x, 1.88, cw5, 0.10, fill=col)
    text(s, x + 0.26, 2.16, cw5 - 0.52, 0.24, tag, size=9, bold=True, color=col)
    text(s, x + 0.26, 2.48, cw5 - 0.52, 0.38, t, size=16, bold=True, color=NAVY)
    shape(s, x + 0.26, 3.00, cw5 - 0.52, 0.03, fill=LINE)
    text(s, x + 0.26, 3.20, cw5 - 0.52, 0.90, d, size=11, color=INK, spacing=1.4)
    for j, it in enumerate(items):
        check_row(s, x + 0.26, 4.42 + j * 0.52, cw5 - 0.52, it, size=11)

shape(s, ML, 5.90, CW, 0.66, MSO_SHAPE.ROUNDED_RECTANGLE, fill=BLUE_L, adj=0.16)
text(s, ML + 0.26, 6.10, CW - 0.52, 0.30,
     "キャリア開発は入口。若手本人の役割と貢献が広がることが到達点です。"
     "（人事向けの閲覧機能は現時点の提供範囲外）",
     size=11.5, bold=True, color=NAVY)
notes(s, SRC +
      "現在、画面を閲覧できるのは若手本人と管理職の2者です。人事向け閲覧機能は現時点の提供範囲には含まれていません。\n"
      "本人の振り返りと管理職のフィードバックを蓄積し、次の1on1につなげます。")

# -------------------------------------------------------- 9 : DEMO 利用の流れ
s = new(prs)
header(s, "DEMO｜この後ご覧いただく利用の流れ",
       "機能一覧ではなく、本人と管理職の利用場面に沿ってご説明します", 9)

flow = [("START", "メンバー画面", "目標を行動に分け、実践内容を振り返る", NAVY),
        ("MONTHLY", "管理職画面", "本人の振り返りを確認し、フィードバックする", CORAL),
        ("REVIEW", "ディスカッション", "現場で使ううえでの不足や改善点を伺う", NAVY)]
cw6 = (CW - 1.30) / 3
for i, (tag, t, d, col) in enumerate(flow):
    x = ML + i * (cw6 + 0.65)
    card(s, x, 2.02, cw6, 2.70)
    shape(s, x + 0.26, 2.32, 1.28, 0.32, MSO_SHAPE.ROUNDED_RECTANGLE,
          fill=col, adj=0.35)
    text(s, x + 0.26, 2.39, 1.28, 0.24, tag, size=9, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)
    text(s, x + 0.26, 2.86, cw6 - 0.52, 0.40, t, size=16, bold=True, color=NAVY)
    shape(s, x + 0.26, 3.38, cw6 - 0.52, 0.03, fill=LINE)
    text(s, x + 0.26, 3.60, cw6 - 0.52, 0.90, d, size=11.5, color=INK, spacing=1.4)
    if i < 2:
        text(s, x + cw6 + 0.10, 3.10, 0.45, 0.4, "→", size=18, bold=True,
             color=GRAY_L, align=PP_ALIGN.CENTER)

shape(s, ML, 5.08, CW, 0.62, MSO_SHAPE.ROUNDED_RECTANGLE, fill=BLUE_L, adj=0.16)
text(s, ML + 0.26, 5.26, CW - 0.52, 0.30,
     "メンバーが振り返る　→　管理職が応答する　→　次の対話につなげる",
     size=12.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
text(s, ML, 6.06, CW, 0.30,
     "実際の画面を共有しながら、使いやすさと運用上の課題を伺います。",
     size=11.5, bold=True, color=CORAL_D)
notes(s, SRC +
      "ここから実際の画面に切り替えます。\n"
      "質問：管理職の負担を抑えながら運用するには、どの場面・頻度が現実的でしょうか。")

# ---------------------------------------------------- 10 : REVIEW 振り返り観点
s = new(prs)
header(s, "REVIEW｜POCの振り返りで確認すること",
       "利用率だけではなく、行動・対話・運用の変化を確認する", 10)

card(s, ML, 1.86, 6.95, 3.95)
card_title(s, ML + 0.30, 2.12, 6.35, "活躍の状態を、5つの観点で確認")
views = ["強み・価値観の理解", "仕事上の役割の理解", "強みを活かした実践",
         "新しい挑戦・役割の拡大", "周囲・組織への働きかけ"]
for i, v in enumerate(views):
    y = 2.78 + i * 0.58
    shape(s, ML + 0.30, y, 6.35, 0.46, MSO_SHAPE.ROUNDED_RECTANGLE,
          fill=BLUE_L, adj=0.20)
    num_badge(s, ML + 0.42, y + 0.06, 0.34, i + 1, fill=NAVY)
    text(s, ML + 0.94, y + 0.10, 5.6, 0.28, v, size=11.5, bold=True, color=INK)

x2 = ML + 7.21
card(s, x2, 1.86, CW - 7.21, 3.95, fill=PINK_L,
     line=RGBColor(0xF6, 0xC8, 0xCA))
card_title(s, x2 + 0.30, 2.12, CW - 7.21 - 0.60, "本人・管理職と確認する内容")
asks = ["本人が行動に落とし込めたか", "管理職の1on1に役立ったか",
        "継続できる頻度・運用だったか", "改善すべき機能・支援は何か"]
for i, a in enumerate(asks):
    check_row(s, x2 + 0.32, 2.84 + i * 0.68, CW - 7.21 - 0.64, a, size=11.5)

shape(s, ML, 6.02, CW, 0.62, MSO_SHAPE.ROUNDED_RECTANGLE, fill=BLUE_L, adj=0.16)
text(s, ML + 0.26, 6.20, CW - 0.52, 0.30,
     "本人と管理職の声をもとに、次の運用とサービス改善を判断します。",
     size=12, bold=True, color=NAVY)
notes(s, SRC +
      "POCでは、本人と管理職の変化、そして運用可能性を組み合わせて振り返ります。\n"
      "人事が画面上で過程を追う機能を前提にはしていません。必要な共有方法はPOCを通じて検討します。")

# ------------------------------------------------- 11 : NEXT ACTION 合意4点
s = new(prs)
header(s, "NEXT ACTION｜本日、合意したい4点",
       "POCの対象者・期間・着眼点・次回日程をこの場で確定する", 11)

agree = [("01", "POCの対象者", "どの部門の、どのメンバーと管理職で検証するか"),
         ("02", "スケジュール", "利用開始日と、いつまでを検証期間とするか"),
         ("03", "検証の着眼点", "特にご確認されたい変化・課題を追加でいただく"),
         ("04", "ネクストアクション", "中間確認、または次回打ち合わせの日程を確保する")]
cwa = (CW - 0.72) / 2
for i, (n, t, d) in enumerate(agree):
    x = ML + (i % 2) * (cwa + 0.72)
    y = 1.92 + (i // 2) * 1.72
    card(s, x, y, cwa, 1.44)
    shape(s, x, y, 0.10, 1.44, fill=CORAL if i % 2 else NAVY)
    text(s, x + 0.36, y + 0.24, 0.6, 0.3, n, size=13, bold=True,
         color=CORAL if i % 2 else NAVY)
    text(s, x + 1.06, y + 0.20, cwa - 1.4, 0.36, t, size=15, bold=True, color=NAVY)
    text(s, x + 1.06, y + 0.74, cwa - 1.4, 0.46, d, size=11.5, color=INK)

shape(s, ML, 5.62, CW, 0.94, MSO_SHAPE.ROUNDED_RECTANGLE, fill=BLUE_L, adj=0.12)
text(s, ML + 0.30, 5.82, CW - 0.60, 0.60,
     [("次回は、実際の利用状況と改善点を一緒に振り返らせてください。", {"size": 12.5, "bold": True, "color": NAVY}),
      ("顧客企業への展開は、社内POCの学びを整理したうえで改めて協議します。", {"size": 11, "color": GRAY})],
     spacing=1.4)
notes(s, SRC +
      "最後に、対象となるメンバーと管理職、開始日、検証期間、中間確認または次回振り返りの日程を確認します。\n"
      "顧客企業への展開については、社内POCの学びを整理したうえで、可能性を改めて協議します。")

# ------------------------------------------------------------- 12 : 裏表紙
s = new(prs)
gradient_bg(s)
logo(s, ML, 0.68, 0.42, white=True)
text(s, ML, 2.30, 9.0, 0.34, "CoreUniq｜若手活躍支援サービス", size=13, bold=True,
     color=RGBColor(0xFF, 0xD9, 0xDB))
text(s, ML, 2.86, 10.4, 1.5,
     [("自分を知る。仕事で試す。", {"size": 34, "bold": True}),
      ("挑戦と貢献を、継続的に広げる。", {"size": 34, "bold": True})],
     color=WHITE, spacing=1.2)
shape(s, ML, 4.88, 1.5, 0.05, fill=WHITE)
text(s, ML, 5.22, 6.2, 0.9,
     [("Givin' Back株式会社", {"size": 14, "bold": True}),
      ("info@givinback.co.jp　|　https://givinback.co.jp/", {"size": 12})],
     color=RGBColor(0xEF, 0xF4, 0xFA), spacing=1.5)
text(s, 7.6, 6.62, 5.1, 0.3, "お問い合わせ", size=10,
     color=RGBColor(0xE6, 0xEE, 0xF7), align=PP_ALIGN.RIGHT)
notes(s, SRC +
      "本日はありがとうございました。POCの進行中も、随時ご相談いただけます。")

OUT = os.path.join(HERE, "ジンジブ様_CoreUniqオンボーディング・POC資料_GivinBackフォーマット.pptx")
prs.save(OUT)
print("saved:", OUT, len(prs.slides.__iter__.__self__._sldIdLst), "slides")
