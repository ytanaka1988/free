# -*- coding: utf-8 -*-
"""
ジンジブ様_CoreUniq オンボーディング・POC 資料（全11枚）

体裁は gb_format.py（Givin' Back 標準フォーマット）に寄せてある。
このファイルは「中身」だけを持つので、別案件の資料をつくるときは
gb_format を import して同じ書き方でスライドを積めばよい。
"""
import os

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

from gb_format import (BLUE_L, CORAL, CREAM, CW, GRAY, GRAY_L, INK, LINE, ML,
                       NAVY, PAD, PINK_L, SZ_BODY, WHITE, Deck)

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")

SRC = "[Sources]\n- Givin' Back社内資料・受講前アンケート\n\n[Talk Track]\n"


def art(slide, name, x, y, h):
    """assets/ に人物イラストがあれば置く（無ければ何もしない）"""
    path = os.path.join(ASSETS, name)
    if os.path.exists(path):
        slide.picture(path, x, y, h=h)


deck = Deck()

# ------------------------------------------------------------------ 1 表紙
s = deck.cover(
    headline="若手育成を、研修だけで終わらせない。",
    subline="若手が、自分の強みを仕事の成果につなげる組織へ。",
    eyebrow="ジンジブ様｜CoreUniq オンボーディング・POC",
    supports=("キャリア開発から、現場での挑戦・役割形成までを一体で支援",
              "自分を知る。仕事で試す。挑戦と貢献を、継続的に広げる。"),
    footnote="導入意図・検証方針の目線合わせ",
    right_note="Givin' Back株式会社　|　ONBOARDING GUIDE 2026")
s.notes(SRC +
        "本日は操作説明だけではなく、CoreUniqを導入する意図と、今回一緒に検証したいことを共有します。\n"
        "CoreUniqは日報確認ツールではなく、キャリア開発を仕事上の実践につなげる継続支援です。")

# ------------------------------------------------------- 2 TODAY 目的とゴール
s = deck.slide("TODAY｜本日の目的とゴール",
               "操作説明だけで終わらず、導入意図と検証方針を合わせる")

s.pill(0.90, 2.04, 4.79, "本日のゴール", size=30, h=0.72, align=PP_ALIGN.CENTER)
goals = [("CoreUniqが生まれた背景と", "支援したい課題を共有する"),
         ("ジンジブ様社内のPOCで", "検証する内容を合意する")]
for i, (l1, l2) in enumerate(goals):
    y = 3.27 + i * 1.32
    s.shape(0.90, y, 0.44, 0.44, MSO_SHAPE.OVAL, fill=CORAL)
    s.text(1.54, y - 0.06, 4.82, 1.10,
           [(l1, {}), (l2, {})], size=23.5, bold=True, color=INK, spacing=1.28)

s.pill(6.54, 2.04, 5.87, "本日の3つの論点", size=30, h=0.72, align=PP_ALIGN.CENTER)
points = [("導入意図", "何を解決するサービスか"),
          ("利用設計", "本人と管理職がどう使うか"),
          ("検証方針", "何を学び、次に活かすか")]
for i, (k, v) in enumerate(points):
    y = 3.34 + i * 0.94
    s.text(6.52, y, 1.86, 0.36, k, size=21.5, bold=True, color=CORAL)
    s.text(8.44, y - 0.02, 0.40, 0.36, "→", size=23, bold=True, color=GRAY_L)
    s.text(8.92, y, 3.79, 0.36, v, size=22.5, color=INK)
s.notes(SRC +
        "今日のゴールは、導入意図・利用設計・検証方針の3点を合わせることです。\n"
        "最後に、対象者、期間、振り返り方法、次回日程まで確認します。")

# ------------------------------------------------------------- 3 AGENDA
s = deck.slide("AGENDA｜本日の進め方（60分）")
s.agenda(["本日の目的・前提合わせ",
          "CoreUniqが生まれた背景",
          "若手の課題とアンケートデータ",
          "POCの目的・検証したいこと",
          "メンバー・管理職画面のデモ",
          "ディスカッション",
          "次回アクションの合意"])
s.notes(SRC +
        "前半は導入意図と課題認識の共有、後半はデモとディスカッションに時間を配分しています。\n"
        "デモに13分、ディスカッションに8分を確保しています。\n"
        "残り4分になったら、対象者・期間・着眼点・次回日程の4点の確認に必ず入ります。")

# ----------------------------------------------------------- 4 DATA 背景
s = deck.slide("DATA｜若手が行動に移れない背景",
               "悩みは『社内での方向性』と『強みの活かし方』に集中している")

s.card(ML, 1.78, 6.72, 4.15)
s.pill(0.92, 2.04, 6.12, "弊社キャリア開発研修受講前の悩み上位（複数回答）", size=16)
s.bars(0.92, 2.66, 6.12,
       [("やりたいことが見えない", 36.7), ("強みの活かし方が不明", 34.7),
        ("自分の軸が分からない", 28.6), ("見つけ方が分からない", 24.5),
        ("始め方が分からない", 20.4)], vmax=40.0)

s.pill(7.90, 2.04, 4.51, "『協働』と『提案』のギャップ", size=16)
metrics = [("周囲との協力意識", "8.02", NAVY), ("率先提案の意識", "5.38", CORAL)]
for i, (k, v, col) in enumerate(metrics):
    y = 2.67 + i * 0.92
    s.card(7.90, y, 4.51, 0.76, border=LINE)
    s.text(8.12, y + 0.23, 2.6, 0.32, k, size=16.5, color=INK)
    s.text(9.90, y + 0.10, 1.30, 0.50, v, size=25.7, bold=True, color=col,
           align=PP_ALIGN.RIGHT)
    s.text(11.30, y + 0.27, 0.6, 0.28, "／10", size=12, color=GRAY)
s.checks(7.92, 4.62, 4.45, ["協働の土台はあるが、意思を行動に変えにくい"])
art(s, "person_data.png", 8.20, 4.95, 1.90)

s.closing("主体性がないのではなく、方向性を言語化し、仕事で試す方法が見えていない可能性があります。",
          y=6.24, size=18)
s.notes(SRC +
        "若手は意欲がないのではなく、強みや方向性を具体的な行動に翻訳する方法が見えていない可能性があります。\n"
        "協力意識は8.02に対して、率先提案の意識は5.38です。協働の土台はあっても、自分から動くところにギャップがあります。\n"
        "質問：ジンジブ様社内や支援先企業でも、似た傾向はありますか。")

# ------------------------------------------------------------- 5 WHAT
s = deck.slide("WHAT｜CoreUniqが支援すること",
               "研修で終わらず、若手の行動変容と役割形成まで伴走")

s.shape(ML, 1.82, CW, 1.72, fill=NAVY)
s.text(1.02, 2.30, 3.2, 0.70, "CoreUniq", size=41, bold=True, color=WHITE)
s.text(4.77, 2.28, CW - 4.50, 1.1,
       [("キャリア開発を起点に、現場での実践・対話・役割形成を", {}),
        ("半年〜年間で支える若手活躍支援プログラムです。", {})],
       size=18, bold=True, color=CREAM, spacing=1.45)

s.text(ML, 3.90, CW, 0.40,
       "若手本人を主役に、管理職との対話を通じて活躍をつくります。",
       size=24, bold=True, color=NAVY)

blocks = [("キャリア開発", "強み・価値観・目指す姿を\n自分の言葉で整理する", NAVY),
          ("現場実践", "強みを活かす小さな行動を\n実際の業務で試す", CORAL),
          ("1on1・育成支援", "対話を通じて経験を付与し\n次の役割を設計する", NAVY)]
bw = 3.60
for i, (title, body, col) in enumerate(blocks):
    x = ML + i * (bw + 0.65)
    s.card(x, 4.76, bw, 2.05, accent=col)
    s.text(x + PAD, 5.02, bw - PAD * 2, 0.36, title, size=20, bold=True,
           color=col)
    s.text(x + PAD, 5.52, bw - PAD * 2, 1.0,
           [(l, {}) for l in body.split("\n")], size=SZ_BODY, color=INK,
           spacing=1.35)
    if i < 2:
        s.text(x + bw + 0.10, 5.42, 0.45, 0.4, "×", size=18, bold=True,
               color=GRAY_L, align=PP_ALIGN.CENTER)
art(s, "person_what.png", 9.60, 3.70, 1.60)
s.notes(SRC +
        "CoreUniqの価値は、研修、現場実践、管理職との対話を一つの支援プロセスとしてつなぐことです。\n"
        "研修部分を初期設定や付帯作業として扱わず、若手の自己理解と行動設計をつくる中核サービスとして位置づけています。")

# --------------------------------------------------------- 6 3 DESIGNS
s = deck.slide("3 DESIGNS｜若手活躍を支える設計",
               "若手活躍は、自己理解・実践・役割形成の3つの循環で育つ")

steps = [("自分を、知る", "強み・価値観・目指す姿を、\n自分の言葉で整理する"),
         ("仕事で、試す", "強みを活かす小さな行動を、\n実際の業務で実践する"),
         ("役割を、広げる", "対話を通じて、\n挑戦・役割・周囲への貢献を広げる")]
cw = 3.60
for i, (t, d) in enumerate(steps):
    x = ML + i * (cw + 0.65)
    col = CORAL if i == 1 else NAVY
    s.card(x, 1.92, cw, 3.55, accent=col)
    s.num_badge(x + (cw - 0.62) / 2, 2.33, 0.62, i + 1, fill=col)
    s.text(x + PAD, 3.14, cw - PAD * 2, 0.40, t, size=17, bold=True, color=NAVY,
           align=PP_ALIGN.CENTER)
    s.text(x + PAD, 3.58, cw - PAD * 2, 1.0,
           [(l, {}) for l in d.split("\n")], size=SZ_BODY, color=INK,
           spacing=1.45, align=PP_ALIGN.CENTER)
    art(s, "person_step{0}.png".format(i + 1), x + (cw - 1.5) / 2, 4.55, 1.45)
    if i < 2:
        s.chevron(x + cw + 0.16, 2.72)

s.closing("キャリア開発を入口に、本人の行動と職場での貢献が広がる状態をつくります。",
          y=5.90, size=23)
s.notes(SRC +
        "自己理解だけで終わらず、仕事で試し、対話を通じて役割を広げる循環をつくります。\n"
        "システムだけで変化を生むのではなく、研修と人の支援を前提にした設計です。")

# -------------------------------------------------------------- 7 POC
s = deck.slide("POC｜ジンジブ様社内で検証したいこと",
               "完成品の評価ではなく、若手支援の運用仮説を一緒に検証")

cols = [("社内で試す",
         ["対象となる若手と管理職を設定", "実際の業務の中で継続利用",
          "利用場面と対話の流れを確認"]),
        ("運用仮説を検証",
         ["本人が行動に落とし込めるか", "管理職の対話に役立つか",
          "無理なく続けられる頻度か"]),
        ("学びを整理する",
         ["有効だった支援方法を整理", "改善すべき機能・運用を特定",
          "顧客企業への展開可能性を検討"])]
cw = 3.73
for i, (t, items) in enumerate(cols):
    x = ML + i * (cw + 0.45)
    col = CORAL if i == 1 else NAVY
    s.card(x, 1.90, cw, 3.70, accent=col)
    s.num_badge(x + 0.20, 2.17, 0.62, i + 1, fill=col)
    s.text(x + 1.09, 2.30, cw - 1.35, 0.34, t, size=14.5, bold=True, color=col)
    s.rule(x + PAD, 2.96, cw - PAD * 2, h=0.03)
    s.checks(x + PAD, 3.20, cw - PAD * 2, items)

s.closing("社内POCの結果をもとに、再現できる若手支援の形を一緒に整理", accent=True,
          y=6.05, size=25)
s.notes(SRC +
        "今回のPOCでは、利用率だけではなく、本人の行動、管理職との対話、継続できる運用条件を確認したいと考えています。\n"
        "成功事例を先に約束するのではなく、何が有効かを一緒に見つける検証です。")

# ------------------------------------------------- 8 FOR ORGANIZATION
s = deck.slide("FOR ORGANIZATION｜若手活躍を支える2者と仕組み",
               "若手本人と管理職の対話を、CoreUniqが支援")

roles = [("MEMBER 01", "若手本人", "強みを仕事に活かし、\n自分から役割と貢献を広げていく。",
          ["仕事上の役割が見える", "挑戦と貢献を自分から広げる"], NAVY, WHITE),
         ("MANAGER 02", "管理職", "本人の強みと実践を踏まえ、\n経験付与や次の役割を設計する。",
          ["1on1で活躍を後押しする", "本人に合う経験を与えられる"], NAVY, WHITE),
         ("COREUNIQ 03", "対話を支える記録", "本人の振り返りと管理職の\nフィードバックを蓄積し、次の対話につなげる。",
          ["本人・管理職のみが閲覧", "1on1の材料を残せる"], CORAL, PINK_L)]
cw = 3.73
for i, (tag, t, d, items, col, tint) in enumerate(roles):
    x = ML + i * (cw + 0.45)
    s.card(x, 1.88, cw, 3.72, accent=col, tint=tint)
    s.text(x + PAD, 2.05, cw - PAD * 2, 0.28, tag, size=17, bold=True, color=col)
    s.text(x + PAD, 2.44, cw - PAD * 2, 0.40, t, size=23, bold=True, color=NAVY)
    s.rule(x + PAD, 3.00, cw - PAD * 2, h=0.03)
    s.text(x + PAD, 3.20, cw - PAD * 2, 1.0,
           [(l, {}) for l in d.split("\n")], size=SZ_BODY, color=INK,
           spacing=1.4)
    s.checks(x + PAD, 4.42, cw - PAD * 2, items, pitch=0.52)

s.closing([("キャリア開発は入口。若手本人の役割と貢献が広がることが到達点",
            {"align": PP_ALIGN.CENTER}),
           ("（人事向けの閲覧機能は現時点の提供範囲外）",
            {"align": PP_ALIGN.CENTER})], y=5.85, size=20.5)
s.notes(SRC +
        "現在、画面を閲覧できるのは若手本人と管理職の2者です。人事向け閲覧機能は現時点の提供範囲には含まれていません。\n"
        "本人の振り返りと管理職のフィードバックを蓄積し、次の1on1につなげます。")

# ------------------------------------------------------------- 9 DEMO
s = deck.slide("DEMO｜この後ご覧いただく利用の流れ",
               "機能一覧ではなく、本人と管理職の利用場面に沿ってご説明します")

flow = [("メンバー画面", "目標を行動に分け、\n実践内容を振り返る"),
        ("管理職画面", "本人の振り返りを確認し、\nフィードバックする"),
        ("ディスカッション", "現場で使ううえでの不足や\n改善点を伺う")]
cw = 3.60
for i, (t, d) in enumerate(flow):
    x = ML + i * (cw + 0.65)
    s.card(x, 2.01, cw, 2.70)
    s.text(x + PAD, 2.79, cw - PAD * 2, 0.36, t, size=16, bold=True, color=NAVY)
    s.rule(x + PAD, 3.37, cw - PAD * 2, h=0.03)
    s.text(x + PAD, 3.58, cw - PAD * 2, 1.0,
           [(l, {}) for l in d.split("\n")], size=SZ_BODY, color=INK,
           spacing=1.4)
    if i < 2:
        s.text(x + cw + 0.10, 3.05, 0.45, 0.4, "→", size=18, bold=True,
               color=GRAY_L, align=PP_ALIGN.CENTER)

s.shape(ML, 5.08, CW, 0.62, fill=BLUE_L)
s.text(ML, 5.24, CW, 0.34,
       "メンバーが振り返る　→　管理職が応答する　→　次の対話につなげる",
       size=22.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
s.closing("実際の画面を共有しながら、使いやすさと運用上の課題を伺います。",
          accent=True, y=6.05, size=27.5)
s.notes(SRC +
        "ここから実際の画面に切り替えます。\n"
        "質問：管理職の負担を抑えながら運用するには、どの場面・頻度が現実的でしょうか。")

# ------------------------------------------------------------ 10 REVIEW
s = deck.slide("REVIEW｜POCの振り返りで確認すること",
               "利用率だけではなく、行動・対話・運用の変化を確認")

s.pill(0.92, 2.12, 6.35, "活躍の状態を、5つの観点で確認", size=17)
s.rows(0.92, 2.78, 6.35,
       ["強み・価値観の理解", "仕事上の役割の理解", "強みを活かした実践",
        "新しい挑戦・役割の拡大", "周囲・組織への働きかけ"])

s.pill(8.13, 2.12, 4.28, "本人・管理職と確認する内容", size=16, fill=CORAL)
s.checks(8.15, 2.78, 4.26,
         ["本人が行動に落とし込めたか", "管理職の1on1に役立ったか",
          "継続できる頻度・運用だったか", "改善すべき機能・支援は何か"],
         pitch=0.68, size=19.5, inline=True)

s.closing("本人と管理職の声をもとに、次の運用とサービス改善を判断します。",
          y=6.15, size=26)
s.notes(SRC +
        "POCでは、本人と管理職の変化、そして運用可能性を組み合わせて振り返ります。\n"
        "人事が画面上で過程を追う機能を前提にはしていません。必要な共有方法はPOCを通じて検討します。")

# ------------------------------------------------------- 11 NEXT ACTION
s = deck.gradient_slide("NEXT ACTION｜本日、合意したい4点",
                        "POCの対象者・期間・着眼点・次回日程をこの場で確定する")

agree = [("POCの対象者", "どの部門の、どのメンバーと管理職で検証するか", NAVY),
         ("スケジュール", "利用開始日と、いつまでを検証期間とするか", CORAL),
         ("検証の着眼点", "特にご確認されたい変化・課題を追加でいただく", NAVY),
         ("ネクストアクション", "中間確認、または次回打ち合わせの日程を確保する", CORAL)]
cwa = 5.685
for i, (t, d, col) in enumerate(agree):
    x = ML + (i % 2) * (cwa + 0.72)
    y = 1.92 + (i // 2) * 1.72
    s.card(x, y, cwa, 1.44, side=col)
    s.text(x + 1.06, y + 0.20, cwa - 1.4, 0.40, t, size=24.5, bold=True,
           color=col)
    s.text(x + 1.05, y + 0.74, cwa - 1.4, 0.46, d, size=13, color=INK)

s.shape(ML, 5.62, CW, 0.94, fill=BLUE_L)
s.text(ML + 0.30, 5.79, CW - 0.60, 0.70,
       [("次回は、実際の利用状況と改善点を一緒に振り返らせてください。",
         {"size": 16.5, "bold": True, "color": NAVY}),
        ("顧客企業への展開は、社内POCの学びを整理したうえで改めて協議します。",
         {"size": 15, "color": GRAY})], spacing=1.45)
s.notes(SRC +
        "最後に、対象となるメンバーと管理職、開始日、検証期間、中間確認または次回振り返りの日程を確認します。\n"
        "顧客企業への展開については、社内POCの学びを整理したうえで、可能性を改めて協議します。\n"
        "お問い合わせ先（info@givinback.co.jp）は口頭でお伝えする。")

OUT = os.path.join(HERE, "ジンジブ様_CoreUniqオンボーディング・POC資料_GivinBackフォーマット.pptx")
deck.save(OUT)
print("saved:", OUT, deck.page, "slides")
