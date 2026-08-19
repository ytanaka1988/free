# Givin' Back 資料フォーマット ＋ ジンジブ様_CoreUniq オンボーディング・POC 資料

`gb_format.py` が Givin' Back の資料フォーマット本体、`build_deck.py` がジンジブ様向け
デッキ（全11枚）の中身。別案件の資料は `gb_format` を import して同じ書き方で組める。

```
pip install python-pptx pillow
python3 build_deck.py
```

## フォーマットの由来

田中さんが手を入れた修正版PDFから寸法・級数・配色を採寸して `gb_format.py` に反映済み。
初版（Walnut Exporter 出力のオレンジ基調）からの主な変更は次のとおり。

| 要素 | 初版 | 現行フォーマット |
| --- | --- | --- |
| アクセント色 | オレンジ `#F7941D` | コーラル `#F96167` |
| リード文 | コーラルの縦バー＋14.5pt | バー無し・**24pt ネイビー太字**（ページの主役） |
| 締め文 | 淡ブルーの角丸ボックス入り12pt | **ボックス無し・24pt 太字**（ネイビー or コーラル） |
| 本文 | 11〜12pt | 14pt、一覧行は19.5pt |
| フォント | Meiryo | **Zen Kaku Gothic Antique** |
| 最終ページ | 裏表紙（お問い合わせ） | **NEXT ACTION をグラデーション面で最終ページ化** |
| AGENDA | 時間・アジェンダ・ポイントの3列表 | アジェンダ項目のみの一覧 |

## デザイントークン（`gb_format.py`）

| 種別 | 値 |
| --- | --- |
| スライド | 16:9（13.333 × 7.5 inch = 960 × 540 pt） |
| マージン | 左右 `ML=0.62` / `MR=12.71`、コンテンツ幅 `CW=12.09` |
| ネイビー | `#144B7F`（濃 `#0E355C`） |
| コーラル | `#F96167`（文字用 `#E0454B`） |
| 面 | 淡ブルー `#EEF3F8` / 淡ピンク `#FDECED` / バー下地 `#EEF1F5` |
| 文字 | 本文 `#1E293B` / 補足 `#6B7787` / フッター `#939EAC` |
| 罫線・カード枠 | `#D8E0EA` |
| ヘッダー | 見出し21pt（左 0.62in）＋右上ロゴ（1.23 × 0.26in）＋罫線 y=0.95in |
| リード | 24pt 太字、x=0.82in、y=1.17in |
| 締め文 | 24pt 太字、x=0.88in、y≒6.15in |
| フッター | 罫線 y=7.02in ＋ `© Givin' Back Inc. All Rights Reserved. \| Confidential` 7.5pt ＋ ページ番号9pt |
| 表紙・最終面 | ネイビー→コーラルのグラデーション、白ロゴ（最終面は 1.99 × 0.42in） |

**フォント注意**: Zen Kaku Gothic Antique は Google Fonts。Google スライドならそのまま出る。
PowerPoint で未インストールの場合は `gb_format.py` の `FONT` を `"Meiryo"` などに変えて再生成する。

## `gb_format` の使い方

```python
from gb_format import Deck, ML, CW, NAVY, CORAL

deck = Deck()

deck.cover("大見出し", "サブ見出し", eyebrow="◯◯様｜案件名",
           supports=("支援文1", "支援文2"),
           footnote="左下の注記", right_note="右下の注記")

s = deck.slide("SECTION｜見出し", "リード文")     # 白地ページ（ヘッダー・フッター自動）
s.card(ML, 1.86, 6.0, 3.9, accent=NAVY)          # 上辺アクセントの白カード
s.pill(0.92, 2.12, 6.35, "カード見出し")           # ネイビーの見出し帯
s.rows(0.92, 2.78, 6.35, ["項目1", "項目2"])       # 連番付きの淡ブルー一覧
s.checks(8.15, 2.78, 4.26, ["確認1"], inline=True) # コーラルのチェック項目
s.bars(0.92, 2.66, 6.12, [("ラベル", 36.7)])       # 横棒グラフ
s.closing("締めの一文", accent=True)                # 最下部の一文

deck.slide("AGENDA｜本日の進め方").agenda(["項目1", "項目2"])
deck.gradient_slide("NEXT ACTION｜見出し", "リード文")   # 締めのグラデーション面
deck.contact(["行1", "行2"], "Givin' Back株式会社", "info@givinback.co.jp")  # 裏表紙（任意）

deck.save("out.pptx")
```

ページ番号は `deck.slide()` / `deck.gradient_slide()` の呼び出し順に自動採番される。

## 構成（全11枚）

| No. | スライド | ページ種別 |
| --- | --- | --- |
| 1 | 表紙 | グラデーション |
| 2 | TODAY｜本日の目的とゴール | 白地 |
| 3 | AGENDA｜本日の進め方（60分） | 白地 |
| 4 | DATA｜若手が行動に移れない背景 | 白地 |
| 5 | WHAT｜CoreUniqが支援すること | 白地 |
| 6 | 3 DESIGNS｜若手活躍を支える設計 | 白地 |
| 7 | POC｜ジンジブ様社内で検証したいこと | 白地 |
| 8 | FOR ORGANIZATION｜若手活躍を支える2者と仕組み | 白地 |
| 9 | DEMO｜この後ご覧いただく利用の流れ | 白地 |
| 10 | REVIEW｜POCの振り返りで確認すること | 白地 |
| 11 | NEXT ACTION｜本日、合意したい4点 | グラデーション |

スピーカーノート（Sources / Talk Track）は全ページに入っている。

## 素材

| ファイル | 用途 |
| --- | --- |
| `logo_gb.png` / `logo_gb_white.png` | Drive「Givin' Backロゴ」から切り出した縦組みロゴ |
| `lock.png` / `lockw.png` | ヘッダー・表紙用の横組みロックアップ（濃色／白） |
| `assets/person_*.png` | 人物イラスト（**未収録**） |

修正版の DATA・WHAT・3 DESIGNS には人物イラストが入っているが、素材ファイルが手元に無いため
リポジトリには含めていない。`assets/` に以下の名前で置けば `build_deck.py` が自動で配置する。
置かなければ何も描かれない（レイアウトは崩れない）。

```
assets/person_data.png    … DATA ページ右下
assets/person_what.png    … WHAT ページ右上
assets/person_step1.png   … 3 DESIGNS カード1
assets/person_step2.png   … 3 DESIGNS カード2
assets/person_step3.png   … 3 DESIGNS カード3
```

## 既知の申し送り

- 最終ページを NEXT ACTION にしたため、**お問い合わせ先（info@givinback.co.jp）が資料内に出てこない**。
  口頭で伝える前提であればそのまま。載せるなら `deck.contact(...)` を1行足せば裏表紙が戻る。
- AGENDA から時間配分の列を外したので、60分の内訳はスピーカーノートに退避してある。
