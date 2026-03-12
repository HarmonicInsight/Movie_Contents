#!/usr/bin/env python3
"""
中編：タイプ別おすすめ通信制高校＋比較データシート（約10分 / 7スライド）
YouTube SEOキーワード: 通信制高校 おすすめ 比較 ヒューマンキャンパス クラーク ルネサンス
"""
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx_helpers import *

SERIES = "通信制高校選び完全ガイド【中編】"
prs = create_presentation()

# ════════════════════════════════════════════════════
# SLIDE 1: タイトル
# ════════════════════════════════════════════════════
slide = add_title_slide(
    prs,
    "タイプ別おすすめ通信制高校",
    "あなたに合う学校が見つかる比較ガイド",
    "〜 通信制高校選び完全ガイド【中編】 〜",
    "Harmonic Insight 2026年3月"
)

# フィードバック②: 問題提起から入る
add_notes(slide,
    "「通信制高校が気になるけど、数が多すぎてどこがいいか分からない」\n"
    "そう思っていませんか？\n"
    "実は通信制高校は300校以上あって、それぞれ特徴が全然違います。\n"
    "でも大丈夫。あなたのタイプに合った学校を、今日この動画で見つけましょう。\n\n"
    "こんにちは、Harmonic Insightです。\n"
    "この動画は「通信制高校選び完全ガイド」の中編です。\n"
    "前編でお伝えしたA・B・C・Dの4タイプ別に、\n"
    "おすすめの学校を具体的にご紹介します。\n\n"
    "前編をまだ見ていない方は、まず自分のタイプを確認してから\n"
    "この動画に戻ってきてくださいね。概要欄にリンクを貼っています。\n\n"
    "今日の動画は概要欄にチャプター（タイムスタンプ）を用意しています。\n"
    "自分のタイプだけ見たい方は、チャプターから該当部分に飛んでください。"
)


# ════════════════════════════════════════════════════
# SLIDE 2: Bタイプ - 専門特化校
# フィードバック⑤: チャプター案内を追加
# ════════════════════════════════════════════════════
slide = new_slide(prs)
add_title_bar(slide, "【Bタイプ】「好き」を仕事に！専門特化校",
              "専門性追求タイプ向け")
add_footer(slide, 1, SERIES)

schools = [
    ("ヒューマンキャンパス\n高等学校", "声優・eスポーツ・美容\nマンガなど40分野\nプロ講師の実践授業", "40以上の\n専門分野", C_ACCENT_GREEN),
    ("AOIKE高等学校", "パティシエ・調理師の\n夢を育む専門校\n在学中に資格取得可能", "製菓・調理\n特化", C_ACCENT_ORANGE),
    ("IT・プログラミング\n特化校", "最先端のIT技術を\n高校から学べる\n就職直結のスキル", "テック系\nスキル", C_ACCENT_BLUE),
]

card_w = Emu(2651760)
gap = Emu(182880)
for i, (name, desc, tag, color) in enumerate(schools):
    x = MARGIN + (card_w + gap) * i
    y = Emu(1097280)
    add_rounded_rect(slide, x, y, card_w, Emu(2743200), C_WHITE)
    add_rect(slide, x, y, card_w, Emu(54864), color)
    tag_shape = add_rounded_rect(slide, x + Emu(91440), y + Emu(137160),
                                 Emu(1097280), Emu(365760), color)
    add_textbox(slide, x + Emu(91440), y + Emu(137160), Emu(1097280), Emu(365760),
                tag, "Calibri", Pt(9), True, C_WHITE, PP_ALIGN.CENTER)
    add_textbox(slide, x + Emu(91440), y + Emu(594360), card_w - Emu(182880), Emu(548640),
                name, "Calibri", Pt(14), True, C_DARK, PP_ALIGN.LEFT)
    add_textbox(slide, x + Emu(91440), y + Emu(1188720), card_w - Emu(182880), Emu(1097280),
                desc, "Calibri", Pt(11), False, C_GRAY, PP_ALIGN.LEFT)

add_rounded_rect(slide, MARGIN, Emu(4023360), CONTENT_W, Emu(457200), C_BG_LIGHT)
add_textbox(slide, Emu(548640), Emu(4069080), Emu(8046720), Emu(365760),
            "在学中から実践スキルが身につく！卒業後の就職・デビューサポートも充実",
            "Calibri", Pt(12), True, C_GOLD, PP_ALIGN.LEFT)

add_notes(slide,
    "まずBタイプ、好きなことを深めたい専門性追求タイプの方へ。\n\n"
    "Aタイプの方はこのパートを飛ばして、概要欄のチャプターから\n"
    "Aタイプ向けに進んでも大丈夫です。\n\n"
    "ヒューマンキャンパス高等学校は、40以上の専門分野を学べる学校です。\n"
    "声優、eスポーツ、美容、マンガ、イラストなど、\n"
    "プロの講師から実践的なスキルを身につけられるのが最大の特徴です。\n\n"
    "AOIKE高等学校は、パティシエや調理師を目指す方に特化。\n"
    "在学中に実際に資格を取得できるのが強みです。\n\n"
    "IT・プログラミング特化校では、最先端の技術を高校生のうちから習得。\n"
    "卒業後すぐに活かせるスキルが身につきます。\n\n"
    "これらの学校の共通点は、卒業後の就職やデビューのサポートが充実していること。\n"
    "「好きなことで生きていく」を本気で実現できる環境です。"
)


# ════════════════════════════════════════════════════
# SLIDE 3: Dタイプ - グローバル校
# ════════════════════════════════════════════════════
slide = new_slide(prs)
add_title_bar(slide, "【Dタイプ】世界へ羽ばたく！海外大学への道",
              "グローバル・挑戦志向タイプ向け")
add_footer(slide, 2, SERIES)

# NIC
add_rounded_rect(slide, MARGIN, Emu(1005840), Emu(3931920), Emu(2834640), C_WHITE)
add_rect(slide, MARGIN, Emu(1005840), Emu(3931920), Emu(54864), C_ACCENT_PURPLE)
add_textbox(slide, Emu(548640), Emu(1097280), Emu(3748440), Emu(365760),
            "NIC International College", "Calibri", Pt(15), True, C_DARK, PP_ALIGN.LEFT)
add_textbox(slide, Emu(548640), Emu(1463040), Emu(3748440), Emu(274320),
            "37年の実績「転換教育」", "Calibri", Pt(12), True, C_ACCENT_PURPLE, PP_ALIGN.LEFT)
nic_points = [
    "英語力と思考力を徹底強化する独自メソッド",
    "欧米トップ大学への進学実績が豊富",
    "日本にいながら海外大学準備が完結",
    "帰国後のキャリアサポートも万全",
]
y = Emu(1828800)
for p in nic_points:
    add_textbox(slide, Emu(640080), y, Emu(3657600), Emu(274320),
                f"  {p}", "Calibri", Pt(10), False, C_DARK)
    y += Emu(274320)

# AIE
add_rounded_rect(slide, Emu(4572000), Emu(1005840), Emu(4114800), Emu(2834640), C_WHITE)
add_rect(slide, Emu(4572000), Emu(1005840), Emu(4114800), Emu(54864), C_ACCENT_BLUE)
add_textbox(slide, Emu(4663440), Emu(1097280), Emu(3931920), Emu(365760),
            "AIE国際高等学校", "Calibri", Pt(15), True, C_DARK, PP_ALIGN.LEFT)
add_textbox(slide, Emu(4663440), Emu(1463040), Emu(3931920), Emu(274320),
            "国際バカロレア（IB）教育", "Calibri", Pt(12), True, C_ACCENT_BLUE, PP_ALIGN.LEFT)
aie_points = [
    "国際バカロレア認定で世界基準の教育",
    "バイリンガル環境で英語力を自然習得",
    "海外大学への出願準備を手厚くサポート",
    "少人数制のきめ細かい指導",
]
y = Emu(1828800)
for p in aie_points:
    add_textbox(slide, Emu(4754880), y, Emu(3840480), Emu(274320),
                f"  {p}", "Calibri", Pt(10), False, C_DARK)
    y += Emu(274320)

add_rounded_rect(slide, MARGIN, Emu(4023360), CONTENT_W, Emu(457200), C_BG_LIGHT)
add_textbox(slide, Emu(548640), Emu(4069080), Emu(8046720), Emu(365760),
            "海外大学進学は「遠い夢」ではない。通信制から世界へ羽ばたく道がある",
            "Calibri", Pt(12), True, C_GOLD, PP_ALIGN.LEFT)

add_notes(slide,
    "Dタイプ、世界に飛び出したい方へ。\n"
    "Bタイプの方はここを飛ばして、概要欄のチャプターからどうぞ。\n\n"
    "NIC International Collegeは37年の実績がある学校です。\n"
    "「転換教育」という独自のメソッドで、英語力と思考力を徹底的に鍛えます。\n"
    "日本にいながら海外大学の準備が完結するのが最大の魅力。\n\n"
    "AIE国際高等学校は国際バカロレア認定校。\n"
    "バイリンガル環境で英語力を自然に身につけられます。\n"
    "少人数制なので、きめ細かい指導が受けられるのもポイントです。\n\n"
    "海外大学への進学、通信制高校からでも十分に実現できます。\n"
    "むしろ時間の自由度が高い分、準備に集中しやすいメリットもあるんです。"
)


# ════════════════════════════════════════════════════
# SLIDE 4: A＆Cタイプ
# ════════════════════════════════════════════════════
slide = new_slide(prs)
add_title_bar(slide, "【A・Cタイプ】自分のペースで ＆ 仲間と成長",
              "学習自由度重視 × コミュニティ重視")
add_footer(slide, 3, SERIES)

# Left: A type
add_rounded_rect(slide, MARGIN, Emu(1005840), Emu(4023360), Emu(3520440), C_WHITE)
add_rect(slide, MARGIN, Emu(1005840), Emu(4023360), Emu(54864), C_ACCENT_BLUE)
add_textbox(slide, Emu(548640), Emu(1097280), Emu(3840480), Emu(320040),
            "Aタイプ：学習自由度重視", "Calibri", Pt(14), True, C_ACCENT_BLUE, PP_ALIGN.LEFT)

a_schools = [
    ("ルネサンス高等学校グループ", "スマホで学べる新スタイル\nスクーリングは年4日程度\n芸能・スポーツとの両立実績多数"),
    ("広域通信制高校", "完全オンライン対応\n自宅学習中心で時間を有効活用\n費用を抑えたい方にも"),
]
y = Emu(1463040)
for name, desc in a_schools:
    add_textbox(slide, Emu(548640), y, Emu(3840480), Emu(274320),
                name, "Calibri", Pt(12), True, C_DARK, PP_ALIGN.LEFT)
    add_textbox(slide, Emu(640080), y + Emu(274320), Emu(3748440), Emu(548640),
                desc, "Calibri", Pt(10), False, C_GRAY, PP_ALIGN.LEFT)
    y += Emu(914400)

add_textbox(slide, Emu(548640), y, Emu(3840480), Emu(457200),
            "→ 部活・習い事・仕事と両立したい方に\n→ 自律的に学習を進められる方に最適",
            "Calibri", Pt(10), True, C_ACCENT_BLUE, PP_ALIGN.LEFT)

# Right: C type
add_rounded_rect(slide, Emu(4663440), Emu(1005840), Emu(4023360), Emu(3520440), C_WHITE)
add_rect(slide, Emu(4663440), Emu(1005840), Emu(4023360), Emu(54864), C_ACCENT_ORANGE)
add_textbox(slide, Emu(4754880), Emu(1097280), Emu(3840480), Emu(320040),
            "Cタイプ：コミュニティ重視", "Calibri", Pt(14), True, C_ACCENT_ORANGE, PP_ALIGN.LEFT)

c_schools = [
    ("クラーク記念国際高等学校", "全国展開・多様なコース\n週5日通学も可能"),
    ("第一学院高等学校", "一人ひとりに寄り添う教育\n豊富なキャンパスライフ"),
    ("おおぞら高等学院", "手厚いメンタルケア\n充実した進路サポート"),
    ("公立通信制高校", "学費を抑えたい方に\n通学圏内で探せる"),
]
y = Emu(1463040)
for name, desc in c_schools:
    add_textbox(slide, Emu(4754880), y, Emu(3840480), Emu(274320),
                name, "Calibri", Pt(12), True, C_DARK, PP_ALIGN.LEFT)
    add_textbox(slide, Emu(4846320), y + Emu(274320), Emu(3748440), Emu(320040),
                desc, "Calibri", Pt(9), False, C_GRAY, PP_ALIGN.LEFT)
    y += Emu(594360)

add_notes(slide,
    "AタイプとCタイプ、まとめてご紹介します。\n\n"
    "Aタイプ、自分のペースで学びたい方。\n"
    "ルネサンス高等学校グループはスマホで学べる新しいスタイル。\n"
    "スクーリングは年間たった4日程度。\n"
    "芸能活動やスポーツとの両立実績も多い学校です。\n\n"
    "Cタイプ、仲間と一緒に成長したい方。\n"
    "クラーク記念国際は全国にキャンパスがあり、週5日通学も可能。\n"
    "全日制に近い感覚で通えるのが魅力です。\n"
    "第一学院は一人ひとりに寄り添った教育が特徴。\n"
    "おおぞら高等学院はメンタルケアが手厚いことで知られています。\n"
    "学費を抑えたい方には公立通信制という選択肢もありますよ。"
)


# ════════════════════════════════════════════════════
# SLIDE 5: 比較データシート
# フィードバック⑥: 1項目ずつハイライト + PDF案内
# ════════════════════════════════════════════════════
slide = new_slide(prs)
add_title_bar(slide, "主要通信制高校 比較データシート",
              "概要欄からPDFダウンロードできます")
add_footer(slide, 4, SERIES)

add_rounded_rect(slide, MARGIN, Emu(960120), CONTENT_W, Emu(3200400), C_WHITE)

# Header
comp_headers = ["学校名", "タイプ", "学費目安（年）", "特徴", "向いている人"]
comp_widths = [Emu(1645920), Emu(1005840), Emu(1280160), Emu(2194560), Emu(2103120)]
hx = MARGIN
for h, w in zip(comp_headers, comp_widths):
    add_rect(slide, hx, Emu(960120), w, Emu(320040), C_GOLD)
    add_textbox(slide, hx, Emu(960120), w, Emu(320040),
                h, "Calibri", Pt(9), True, C_WHITE, PP_ALIGN.CENTER)
    hx += w

comp_data = [
    ("N高・S高", "オンライン/通学", "25〜120万", "先進的IT教育", "自己管理力がある人"),
    ("ゼロ高", "体験型", "約50万〜", "起業・実践中心", "行動力がある人"),
    ("ヒューマンキャンパス", "専門特化", "約40〜80万", "40以上の専門分野", "好きを極めたい人"),
    ("クラーク記念国際", "通学型", "約60〜100万", "全国キャンパス", "仲間と学びたい人"),
    ("ルネサンス", "オンライン", "約25〜40万", "スマホ学習", "自分のペースの人"),
    ("NIC International", "グローバル", "約150万〜", "海外大学進学", "海外を目指す人"),
    ("公立通信制", "通信型", "約3〜5万", "低コスト", "費用を抑えたい人"),
]

for ri, (name, typ, cost, feat, suited) in enumerate(comp_data):
    bg = C_BG_LIGHT if ri % 2 == 0 else C_WHITE
    ry = Emu(1280160) + Emu(320040) * ri
    vals = [name, typ, cost, feat, suited]
    rx = MARGIN
    for ci, (val, w) in enumerate(zip(vals, comp_widths)):
        add_rect(slide, rx, ry, w, Emu(320040), bg)
        add_textbox(slide, rx, ry, w, Emu(320040),
                    val, "Calibri", Pt(8), ci == 0, C_DARK, PP_ALIGN.CENTER)
        rx += w

# ハイライト情報
highlights = [
    ("学費が最も安い", "公立通信制：年3〜5万円", C_ACCENT_GREEN),
    ("サポート最充実", "クラーク・第一学院", C_ACCENT_BLUE),
    ("専門分野の幅", "ヒューマンキャンパス：40分野", C_ACCENT_ORANGE),
]
hx = MARGIN
hw = Emu(2743200)
for title, val, color in highlights:
    add_rounded_rect(slide, hx, Emu(3611880), hw, Emu(457200), color)
    add_textbox(slide, hx, Emu(3611880), hw, Emu(228600),
                title, "Calibri", Pt(9), True, C_WHITE, PP_ALIGN.CENTER)
    add_textbox(slide, hx, Emu(3840480), hw, Emu(228600),
                val, "Calibri", Pt(10), False, C_WHITE, PP_ALIGN.CENTER)
    hx += hw + Emu(91440)

add_textbox(slide, MARGIN, Emu(4160520), CONTENT_W, Emu(274320),
            "※ この比較表はPDFでダウンロードできます。概要欄のリンクからどうぞ",
            "Calibri", Pt(9), False, C_GRAY, PP_ALIGN.CENTER)

add_notes(slide,
    "ここで主要な通信制高校を一覧で比較してみましょう。\n"
    "この比較表は概要欄からPDFでダウンロードできます。\n"
    "ご家族で検討する際にぜひ印刷して使ってください。\n\n"
    "動画では、特に注目すべきポイントを3つハイライトします。\n\n"
    "まず、学費が最も安いのは公立通信制。年間3万から5万円です。\n"
    "就学支援金を使えば、実質ゼロ円に近くなる場合もあります。\n\n"
    "サポートが最も充実しているのは、クラーク記念国際や第一学院。\n"
    "週5日通学もでき、担任の先生が手厚くサポートしてくれます。\n\n"
    "専門分野の幅が最も広いのはヒューマンキャンパス。\n"
    "40以上の分野から選べるのは、他校にはない圧倒的な強みです。\n\n"
    "全部の学校を比較する必要はありません。\n"
    "前編で分かったあなたのタイプに合う学校を中心に、2〜3校に絞ってみてください。"
)


# ════════════════════════════════════════════════════
# SLIDE 6: よくある質問（FAQ）
# ════════════════════════════════════════════════════
slide = new_slide(prs)
add_title_bar(slide, "通信制高校 よくある質問", "視聴者の皆さんの疑問に回答")
add_footer(slide, 5, SERIES)

faqs = [
    ("Q. 卒業資格は全日制と同じですか？",
     "A. はい、全く同じです。「高校卒業」の資格に違いはありません。\n"
     "    履歴書にも「○○高等学校 卒業」と記載できます。",
     C_ACCENT_BLUE),
    ("Q. 大学受験に不利になりませんか？",
     "A. 不利にはなりません。推薦入試やAO入試では、\n"
     "    通信制ならではの経験がアピールポイントになることも。",
     C_ACCENT_GREEN),
    ("Q. 友達はできますか？",
     "A. できます！スクーリング、部活動、オンラインコミュニティなど、\n"
     "    交流の機会は多くの学校で用意されています。",
     C_ACCENT_ORANGE),
    ("Q. 途中で全日制に戻れますか？",
     "A. 制度上は可能ですが、カリキュラムの違いがあるため、\n"
     "    事前に転入先の学校に相談することをおすすめします。",
     C_ACCENT_PURPLE),
]

fy = Emu(960120)
for q, a, color in faqs:
    add_rounded_rect(slide, MARGIN, fy, CONTENT_W, Emu(822960), C_WHITE)
    add_rect(slide, MARGIN, fy, Emu(36576), Emu(822960), color)
    add_textbox(slide, Emu(594360), fy + Emu(45720), Emu(8046720), Emu(274320),
                q, "Calibri", Pt(11), True, C_DARK, PP_ALIGN.LEFT)
    add_textbox(slide, Emu(594360), fy + Emu(320040), Emu(8046720), Emu(457200),
                a, "Calibri", Pt(10), False, C_GRAY, PP_ALIGN.LEFT)
    fy += Emu(868680)

add_notes(slide,
    "ここで視聴者の皆さんからよくいただく質問にお答えします。\n\n"
    "「卒業資格は全日制と同じですか？」\n"
    "はい、全く同じです。通信制だから資格が違うということは一切ありません。\n\n"
    "「大学受験に不利にならない？」\n"
    "なりません。むしろ通信制で培った自主性や独自の経験は、\n"
    "推薦入試やAO入試でアピールポイントになることもあります。\n\n"
    "「友達はできる？」\n"
    "これが一番多い質問ですが、できます。\n"
    "スクーリングや部活動、オンラインのコミュニティなど、\n"
    "交流の場は想像以上にたくさん用意されています。\n\n"
    "他にも質問があれば、コメント欄に書いてください。\n"
    "次回以降の動画でお答えしますね。"
)


# ════════════════════════════════════════════════════
# SLIDE 7: エンドカード
# ════════════════════════════════════════════════════
slide = add_end_slide(
    prs,
    summary_items=[
        "Bタイプ → ヒューマンキャンパス等の専門特化校",
        "Dタイプ → NIC・AIE等のグローバル校",
        "Aタイプ → ルネサンス等のオンライン校",
        "Cタイプ → クラーク等の通学型サポート校",
    ],
    next_video_title="【後編】学費攻略＋今すぐ始める行動プラン",
    next_video_desc="支援制度で学費を大幅に減らす方法と、今日からできる具体的アクション"
)

add_notes(slide,
    "中編のまとめです。\n"
    "Bタイプの方にはヒューマンキャンパスなどの専門特化校。\n"
    "Dタイプの方にはNICやAIEなどのグローバル校。\n"
    "Aタイプの方にはルネサンスなどのオンライン校。\n"
    "Cタイプの方にはクラークなどの通学型サポート校がおすすめです。\n\n"
    "気になる学校は見つかりましたか？\n"
    "でも「学費が心配…」という方、大丈夫です。\n\n"
    "後編では、支援制度を使って学費を大幅に減らす方法と、\n"
    "今日からすぐ始められる具体的な行動プランをお伝えします。\n"
    "画面に表示されている後編の動画をぜひクリックしてください。"
)

# ── 保存 ──
output = "通信制高校ガイド_中編_タイプ別おすすめ校と比較データ.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
