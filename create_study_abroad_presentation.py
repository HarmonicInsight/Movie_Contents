#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
海外教育比較 × グローバルオンライン教育 YouTube勉強会用プレゼンテーション
Harmonic Insight テンプレートスタイル準拠

スピーカーノート = YouTube動画のナレーション台本
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ===== Color Palette (Harmonic Insight Style) =====
GOLD = RGBColor(0x8B, 0x75, 0x36)
GOLD_LIGHT = RGBColor(0xC9, 0xA8, 0x4C)
GOLD_ACCENT = RGBColor(0xB8, 0x86, 0x0B)
DARK_BG = RGBColor(0x1A, 0x1A, 0x1A)
CREAM = RGBColor(0xF5, 0xF0, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x33, 0x33, 0x33)
TEXT_GRAY = RGBColor(0x66, 0x66, 0x66)
BLUE_ACCENT = RGBColor(0x2C, 0x7A, 0xB0)
GREEN_ACCENT = RGBColor(0x27, 0xAE, 0x60)
ORANGE_ACCENT = RGBColor(0xE6, 0x7E, 0x22)
RED_ACCENT = RGBColor(0xC0, 0x39, 0x2B)
PURPLE_ACCENT = RGBColor(0x8E, 0x44, 0xAD)

SLIDE_WIDTH = Emu(9144000)
SLIDE_HEIGHT = Emu(5143500)

prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT

# ===== Helper Functions =====
def add_shape(slide, left, top, width, height, fill_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    return shape

def add_text_box(slide, left, top, width, height, text, font_size=14, bold=False,
                 color=TEXT_DARK, alignment=PP_ALIGN.LEFT, font_name="Yu Gothic"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = alignment
    return txBox

def add_bullet_text(slide, left, top, width, height, items, font_size=12,
                    color=TEXT_DARK, font_name="Yu Gothic", spacing=Pt(6)):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = font_name
        p.space_after = spacing
    return txBox

def add_gold_line(slide, left, top, width):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()
    return shape

def set_notes(slide, text):
    """スライドにスピーカーノート（ナレーション台本）を設定"""
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = text

def make_title_slide(title, subtitle="", notes=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=DARK_BG)
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(4), fill_color=GOLD)
    add_shape(slide, 0, SLIDE_HEIGHT - Pt(4), SLIDE_WIDTH, Pt(4), fill_color=GOLD)
    add_text_box(slide, Inches(0.8), Inches(1.5), Inches(8.4), Inches(1.5),
                 title, font_size=32, bold=True, color=GOLD_LIGHT,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    if subtitle:
        add_text_box(slide, Inches(1), Inches(3.2), Inches(8), Inches(1),
                     subtitle, font_size=16, color=CREAM,
                     alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    add_text_box(slide, Inches(0.5), Inches(4.8), Inches(9), Inches(0.4),
                 "Harmonic Insight", font_size=11, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    if notes:
        set_notes(slide, notes)
    return slide

def make_content_slide(title, bullet_items, notes=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=CREAM)
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(3), fill_color=GOLD)
    add_shape(slide, 0, Pt(3), SLIDE_WIDTH, Inches(0.7), fill_color=WHITE)
    add_text_box(slide, Inches(0.5), Inches(0.05), Inches(9), Inches(0.6),
                 title, font_size=20, bold=True, color=GOLD,
                 alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_gold_line(slide, Inches(0.5), Inches(0.72), Inches(9))
    add_bullet_text(slide, Inches(0.6), Inches(0.9), Inches(8.8), Inches(3.8),
                    bullet_items, font_size=13, color=TEXT_DARK)
    add_text_box(slide, Inches(0.5), Inches(5.1), Inches(9), Inches(0.3),
                 "Harmonic Insight", font_size=9, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    if notes:
        set_notes(slide, notes)
    return slide

def make_two_column_slide(title, left_title, left_items, right_title, right_items,
                          notes="", left_color=GOLD, right_color=BLUE_ACCENT):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=CREAM)
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(3), fill_color=GOLD)
    add_shape(slide, 0, Pt(3), SLIDE_WIDTH, Inches(0.7), fill_color=WHITE)
    add_text_box(slide, Inches(0.5), Inches(0.05), Inches(9), Inches(0.6),
                 title, font_size=20, bold=True, color=GOLD,
                 alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_gold_line(slide, Inches(0.5), Inches(0.72), Inches(9))
    # Left
    add_shape(slide, Inches(0.4), Inches(0.9), Inches(4.3), Inches(0.4), fill_color=left_color)
    add_text_box(slide, Inches(0.5), Inches(0.92), Inches(4.1), Inches(0.35),
                 left_title, font_size=13, bold=True, color=WHITE,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    add_bullet_text(slide, Inches(0.5), Inches(1.4), Inches(4.1), Inches(3.4),
                    left_items, font_size=11, color=TEXT_DARK)
    # Right
    add_shape(slide, Inches(5.2), Inches(0.9), Inches(4.3), Inches(0.4), fill_color=right_color)
    add_text_box(slide, Inches(5.3), Inches(0.92), Inches(4.1), Inches(0.35),
                 right_title, font_size=13, bold=True, color=WHITE,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    add_bullet_text(slide, Inches(5.3), Inches(1.4), Inches(4.1), Inches(3.4),
                    right_items, font_size=11, color=TEXT_DARK)
    add_text_box(slide, Inches(0.5), Inches(5.1), Inches(9), Inches(0.3),
                 "Harmonic Insight", font_size=9, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    if notes:
        set_notes(slide, notes)
    return slide

def make_three_column_slide(title, col_data, notes=""):
    """col_data: list of (col_title, items, color)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=CREAM)
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(3), fill_color=GOLD)
    add_shape(slide, 0, Pt(3), SLIDE_WIDTH, Inches(0.7), fill_color=WHITE)
    add_text_box(slide, Inches(0.5), Inches(0.05), Inches(9), Inches(0.6),
                 title, font_size=20, bold=True, color=GOLD,
                 alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_gold_line(slide, Inches(0.5), Inches(0.72), Inches(9))
    col_width = Inches(2.9)
    for i, (col_title, items, col_color) in enumerate(col_data):
        x = Inches(0.35 + i * 3.1)
        add_shape(slide, x, Inches(0.9), col_width, Inches(0.35), fill_color=col_color)
        add_text_box(slide, x + Inches(0.05), Inches(0.9), col_width - Inches(0.1), Inches(0.35),
                     col_title, font_size=12, bold=True, color=WHITE,
                     alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
        add_bullet_text(slide, x + Inches(0.1), Inches(1.35), col_width - Inches(0.2), Inches(3.5),
                        items, font_size=10, color=TEXT_DARK, spacing=Pt(4))
    add_text_box(slide, Inches(0.5), Inches(5.1), Inches(9), Inches(0.3),
                 "Harmonic Insight", font_size=9, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    if notes:
        set_notes(slide, notes)
    return slide

def make_highlight_slide(title, key_number, key_label, description_items, notes=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=CREAM)
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(3), fill_color=GOLD)
    add_shape(slide, 0, Pt(3), SLIDE_WIDTH, Inches(0.7), fill_color=WHITE)
    add_text_box(slide, Inches(0.5), Inches(0.05), Inches(9), Inches(0.6),
                 title, font_size=20, bold=True, color=GOLD,
                 alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_gold_line(slide, Inches(0.5), Inches(0.72), Inches(9))
    add_shape(slide, Inches(0.5), Inches(1.0), Inches(3.5), Inches(2.0),
              fill_color=RGBColor(0x2A, 0x24, 0x18))
    add_text_box(slide, Inches(0.6), Inches(1.1), Inches(3.3), Inches(1.2),
                 key_number, font_size=40, bold=True, color=GOLD_LIGHT,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    add_text_box(slide, Inches(0.6), Inches(2.2), Inches(3.3), Inches(0.6),
                 key_label, font_size=14, color=CREAM,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    add_bullet_text(slide, Inches(4.3), Inches(1.0), Inches(5.2), Inches(3.5),
                    description_items, font_size=12, color=TEXT_DARK)
    add_text_box(slide, Inches(0.5), Inches(5.1), Inches(9), Inches(0.3),
                 "Harmonic Insight", font_size=9, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    if notes:
        set_notes(slide, notes)
    return slide


# ============================================================
# SLIDE 1: フック＋タイトル
# ============================================================
make_title_slide(
    "日本の教育は世界から\nどう見えているのか？",
    "日米欧オンライン教育の徹底比較\n── あなたの学び方は、世界基準ですか？ ──",
    notes=(
        "皆さんこんにちは、Harmonic Insightです。\n"
        "突然ですが、日本の教育って世界と比べてどうなんでしょうか？\n"
        "アメリカの高校生は成績だけじゃなくて課外活動で大学合否が決まる。\n"
        "イギリスでは大学に行く前に1年間、働いたり海外に出たりするのが当たり前。\n"
        "韓国では深夜まで塾に通い続ける受験戦争が社会問題になっている。\n"
        "一方で、日本の通信制高校の進学率は過去最高を更新しています。\n\n"
        "今日は、世界のオンライン教育の最前線を日本・アメリカ・ヨーロッパの3地域で\n"
        "徹底比較しながら、「これからの学び方」を一緒に考えていきたいと思います。\n"
        "最後まで見ていただくと、教育の選択肢が一気に広がるはずです。\n"
        "ぜひ最後までご覧ください。"
    )
)

# ============================================================
# SLIDE 2: 目次
# ============================================================
make_content_slide(
    "今日お話しすること",
    [
        "1. いま世界で何が起きている？── オンライン教育市場の爆発的成長",
        "",
        "2. 海外の高校生・大学生は何をしている？── 米国・英国・アジアの実態",
        "",
        "3. アメリカのオンライン高校 ── 日本と何が違うのか？",
        "",
        "4. ヨーロッパの教育 ── 「学ぶ権利」という考え方",
        "",
        "5. 日本の通信制高校・オンライン大学 ── 実は進化している",
        "",
        "6. 学費・質保証・サポート体制 ── 日米欧まるごと比較",
        "",
        "7. MOOCsの活用法 ── Coursera・Udemyで何ができるか",
        "",
        "8. まとめ ── これからの「学び方」をどう選ぶか",
    ],
    notes=(
        "今日の動画では、この8つのテーマでお話ししていきます。\n"
        "まず世界のオンライン教育がどれだけ成長しているかという全体像をお見せした後、\n"
        "各国の高校生・大学生が実際にどんな学び方をしているのか、具体的に見ていきます。\n"
        "その上で、日本の通信制高校やオンライン大学の現状を、\n"
        "アメリカ・ヨーロッパと比較しながら解説します。\n"
        "最後に、CourseraやUdemyといったオンライン学習プラットフォームの活用法と、\n"
        "皆さんがこれからの学び方を選ぶときのヒントをお伝えします。"
    )
)

# ============================================================
# SLIDE 3: 市場の爆発的成長
# ============================================================
make_highlight_slide(
    "オンライン教育市場 ── いま世界で何が起きているか",
    "3,888億$",
    "2025年 世界市場規模",
    [
        "2000年から市場規模は900%に拡大",
        "2030年には5,647億ドルへ（年率7.75%成長）",
        "2029年までに約11億人がオンライン学習に参加",
        "コロナ後も米国の学習率は28%を維持",
        "　（パンデミック前はわずか10%未満）",
        "AI・VR・クラウド技術が成長をさらに加速",
    ],
    notes=(
        "まず全体像です。世界のオンライン教育市場は2025年時点で約3,888億ドル、\n"
        "日本円にするとおよそ58兆円という巨大な規模になっています。\n"
        "2000年と比べて900パーセント、つまり10倍近くに成長しました。\n\n"
        "注目すべきは、コロナが終わった後も成長が止まっていないということです。\n"
        "アメリカではパンデミック前のオンライン学習率はわずか10%未満でしたが、\n"
        "コロナが収束した2022年でも28%を維持しています。\n"
        "つまり、一度オンライン学習を経験した人は「これでいい」と実感したんですね。\n\n"
        "さらにAIやVR技術の進化が、この流れをますます加速させています。\n"
        "もはやオンライン教育は緊急時の代替手段ではなく、\n"
        "教育システムの中核になりつつあるということです。"
    )
)

# ============================================================
# SLIDE 4: オンライン学習の利点と課題
# ============================================================
make_two_column_slide(
    "オンライン学習 ── 世界共通の「光」と「影」",
    "光：普遍的な利点",
    [
        "・いつでもどこでも学べる",
        "  → 地理的・時間的な制約がなくなる",
        "・AIで一人ひとりに最適化された学習",
        "  → アダプティブラーニングの進化",
        "・録画で繰り返し学習が可能",
        "・通学費・寮費が不要でコスト削減",
        "・年齢・外見に左右されない公平な環境",
    ],
    "影：共通の課題",
    [
        "・友達と会えない → 社会性が育ちにくい",
        "・モチベーションを自分で維持する難しさ",
        "・デジタルデバイド",
        "  → ネット環境・デバイスを持てない人がいる",
        "・実習系の科目はオンラインに限界がある",
        "・教員のオンライン指導スキルが追いついていない",
    ],
    notes=(
        "オンライン学習には世界共通の「光」と「影」があります。\n\n"
        "まず光の部分。場所や時間に縛られずに学べるのは最大の強みですよね。\n"
        "しかもAIの進化で、一人ひとりのレベルに合わせた学習が可能になっています。\n"
        "授業の録画があれば何度でも見返せますし、\n"
        "通学費や寮費がかからないのでコストも大幅に下がります。\n"
        "さらに面白いのが、オンラインでは年齢や外見で判断されないので、\n"
        "発言の内容そのもので評価されるフェアな環境が生まれるということです。\n\n"
        "一方で影の部分もあります。\n"
        "友達に会えない、孤独感がある、モチベーションが続かない。\n"
        "これはオンライン学習の最大の課題として世界中で指摘されています。\n"
        "また、ネット環境やパソコンが手に入らないという「デジタルデバイド」も深刻です。\n"
        "技術が進歩しても、その恩恵を受けられない人が取り残されてしまう。\n"
        "この「光」と「影」のバランスをどう取るかが、各国の教育政策の最大のテーマになっています。"
    ),
    left_color=GREEN_ACCENT,
    right_color=RED_ACCENT,
)

# ============================================================
# SLIDE 5: 米国・英国・アジアの学生の「今」
# ============================================================
make_three_column_slide(
    "世界の高校生・大学生は何をしている？",
    [
        ("アメリカ", [
            "成績だけでは大学に入れない",
            "課外活動の「継続性」が評価される",
            "クラブ・ボランティア・インターンを",
            "　長く深く続けることが重要",
            "大学は「なぜその活動をしたか」を",
            "　面接で深掘りしてくる",
        ], BLUE_ACCENT),
        ("イギリス", [
            "「ギャップイヤー」が当たり前",
            "大学の前後に1年間の社会経験",
            "83%が就業、56%が海外滞在",
            "企業の94%がギャップイヤー経験者を",
            "　積極採用と回答",
            "ワーホリ費用：年144万〜226万円",
        ], GREEN_ACCENT),
        ("アジア（韓国・シンガポール）", [
            "韓国：「SKY大学」がエリートの象徴",
            "高校生は深夜まで塾に通う受験戦争",
            "大学入学後も成績・資格の競争が続く",
            "シンガポール：大学進学率は3〜4割",
            "職業訓練校（ポリテク・ITE）が主要進路",
            "「効率的に学んでキャリアに直結」が鍵",
        ], ORANGE_ACCENT),
    ],
    notes=(
        "ここからは国別に、高校生や大学生が実際に何をしているのか見ていきましょう。\n\n"
        "まずアメリカ。アメリカの大学入試は日本とまったく違います。\n"
        "テストの点数だけでは入れないんです。\n"
        "クラブ活動、ボランティア、インターンシップ。\n"
        "しかも大事なのは「たくさんやること」じゃなくて「1つか2つを長く深く続けること」。\n"
        "大学側は面接で「なぜその活動を選んだのか」「何を学んだのか」を徹底的に聞いてきます。\n"
        "つまり、知識量よりも「どんな人間か」が問われるんですね。\n\n"
        "次にイギリス。イギリスには「ギャップイヤー」という文化があります。\n"
        "大学に入る前、あるいは卒業後に1年間、働いたり海外に行ったりする期間です。\n"
        "驚くべきことに、企業の94パーセントが\n"
        "「ギャップイヤーを経験した人を積極的に採用したい」と回答しています。\n"
        "日本では「空白期間」はマイナスに見られがちですが、\n"
        "イギリスでは逆に「主体性がある」「問題解決能力が高い」と評価されるんです。\n\n"
        "アジアに目を向けると、韓国の受験戦争は壮絶です。\n"
        "ソウル大学・高麗大学・延世大学、通称「SKY大学」に入るために\n"
        "高校生は文字通り深夜まで塾に通い続けます。\n"
        "一方シンガポールは大学進学率を3〜4割に絞っていて、\n"
        "ポリテクニックやITEという職業訓練校が主要な進路として確立されています。\n"
        "「全員が大学に行く必要はない」という割り切りが、実はとても合理的なんですね。"
    )
)

# ============================================================
# SLIDE 6: アメリカのオンライン高校
# ============================================================
make_two_column_slide(
    "アメリカのオンライン高校 ── 大学進学を徹底サポート",
    "カリキュラムの特徴",
    [
        "・数学、科学、英語、社会科のコア科目",
        "・外国語、デジタル技術、美術史など選択科目",
        "・Honors / AP（大学レベル先取り）コース",
        "・NCAA承認 → 学生アスリートにも対応",
        "・週30時間の学習（80〜90%がPC上）",
        "・代表校：Pearson Online Academy",
        "　　　　 Connections Academy",
    ],
    "手厚いサポート体制",
    [
        "・科目別の専門教員が指導",
        "・ホームルーム教員が全科目をモニタリング",
        "・専任スクールカウンセラーが常駐",
        "・4年間の学習計画を一緒に作成",
        "・大学出願プロセスの支援",
        "・奨学金情報の提供",
        "・入試対策の無料オンラインコース",
    ],
    notes=(
        "アメリカのオンライン高校は、日本の通信制高校とはかなり違います。\n"
        "まずカリキュラム。数学や理科といったコア科目はもちろんですが、\n"
        "AP、つまりAdvanced Placementという大学レベルの授業を\n"
        "高校のうちから受けられるんです。\n"
        "大学進学を目指す生徒向けに、かなり戦略的なカリキュラムが組まれています。\n\n"
        "でも本当にすごいのはサポート体制のほうです。\n"
        "科目ごとに専門の先生がいるのはもちろん、\n"
        "それとは別にホームルームの先生が全科目の学習状況を見てくれます。\n"
        "さらに専任のスクールカウンセラーが常駐していて、\n"
        "4年間の学習計画を一緒に作ってくれたり、大学の出願プロセスを手伝ってくれたり、\n"
        "奨学金の情報を提供してくれたりします。\n\n"
        "つまりアメリカのオンライン高校は、単に「オンラインで授業が受けられる」だけじゃなくて、\n"
        "大学合格までの道筋を全部サポートする仕組みが整っているんです。\n"
        "学費は年間1,800ドルから2,800ドル、日本円で約27万円から42万円くらいです。"
    ),
    left_color=BLUE_ACCENT,
    right_color=BLUE_ACCENT,
)

# ============================================================
# SLIDE 7: ヨーロッパのオンライン教育
# ============================================================
make_two_column_slide(
    "ヨーロッパのオンライン教育 ── 「学ぶ権利」という思想",
    "高校教育",
    [
        "・理念：「高等教育は人権である」",
        "  → 国連世界人権宣言がベース",
        "・英国/米国カリキュラム、IBプログラム",
        "・8つのキーコンピテンシーの育成",
        "・「ヨーロピアンアワー」で異文化交流",
        "・CNED（仏）：欧州最大の遠隔教育機関",
        "・学費：年間€630〜€5,900",
    ],
    "大学教育",
    [
        "・Open University（英）：幅広い分野で学位",
        "・FernUniversität in Hagen（独）：",
        "  ドイツ最大の州立遠隔教育大学",
        "・ロンドン大学：キャンパスと同質の学位",
        "・EU/EEA圏の学生は授業料ほぼ無料",
        "・非EU学生でも年€5,000〜€18,000",
        "・ESGという共通の質保証基準で信頼性を担保",
    ],
    notes=(
        "次にヨーロッパです。ヨーロッパの教育で特徴的なのは、\n"
        "「教育は人権である」という考え方が根底にあることです。\n"
        "国連世界人権宣言をベースにしていて、\n"
        "だからこそEU圏の学生は大学の授業料がほぼ無料、\n"
        "あるいは非常に低額で学べるようになっています。\n\n"
        "オンライン高校では、英国カリキュラムや国際バカロレアなど\n"
        "複数のカリキュラムが選べるのが特徴です。\n"
        "面白いのは「ヨーロピアンアワー」という取り組みで、\n"
        "違う国の生徒同士がオンラインで異文化交流する時間が設けられています。\n\n"
        "大学レベルでは、イギリスのOpen Universityや\n"
        "ドイツのFernUniversitätが有名です。\n"
        "ロンドン大学のオンラインプログラムは、\n"
        "キャンパスで受ける授業と同等の質が保証されていると公式に言われています。\n\n"
        "そしてヨーロッパ全体で「ESG」という共通の質保証基準を持っているのが強みです。\n"
        "どの国のどの大学でも、一定の基準を満たしているという安心感がある。\n"
        "これは日本のオンライン教育にはまだ確立されていない仕組みです。"
    ),
    left_color=GREEN_ACCENT,
    right_color=GREEN_ACCENT,
)

# ============================================================
# SLIDE 8: 日本の通信制高校
# ============================================================
make_two_column_slide(
    "日本の通信制高校 ── 実は進化し続けている",
    "制度の仕組み",
    [
        "・毎日通わなくていい（単位制）",
        "・レポート提出＋スクーリング＋試験で単位取得",
        "・卒業要件：3年以上在籍＋74単位以上",
        "・特別活動への30時間以上の参加",
        "・公立：年間3万〜5万円",
        "・私立：年間10万〜100万円",
        "・就学支援金で実質無料になるケースも",
    ],
    "進路の実態（最新データ）",
    [
        "・大学進学率：約27%（2023年度）",
        "  → 通信制の歴史上、過去最高",
        "・専門学校等含む合計進学率：52%超",
        "・N高等学校：東大・海外大学への進学実績も",
        "・ゼロ高等学院：起業家育成に特化",
        "  留学・インターンシップが充実",
        "・柔軟な学習で早期から受験勉強が可能",
    ],
    notes=(
        "さて、ここからは日本の話です。\n"
        "日本の通信制高校って、皆さんどんなイメージですか？\n"
        "「全日制に通えなかった人の受け皿」というイメージがあるかもしれません。\n"
        "でも実は、通信制高校はいま大きく進化しています。\n\n"
        "まず制度の仕組みですが、毎日学校に通わなくていい単位制です。\n"
        "レポートを出して、スクーリングに行って、試験を受けて単位を取る。\n"
        "3年以上在籍して74単位を取れば卒業できます。\n"
        "公立なら年間3万円から5万円、就学支援金を使えば実質無料のケースもあります。\n\n"
        "そして注目すべきは進路の実態です。\n"
        "2023年度のデータでは、通信制高校の大学進学率が約27パーセント。\n"
        "これは通信制の歴史上、過去最高の数字です。\n"
        "専門学校を含めた合計進学率は52パーセントを超えました。\n\n"
        "N高等学校からは東京大学や海外の大学に進学する生徒も出ていますし、\n"
        "ゼロ高等学院は「座学よりも行動」をモットーに、\n"
        "高校生のうちから起業やインターンを経験させるユニークな教育を行っています。\n"
        "通信制＝ドロップアウトの受け皿、という時代は終わりつつあるんです。"
    ),
    left_color=ORANGE_ACCENT,
    right_color=ORANGE_ACCENT,
)

# ============================================================
# SLIDE 9: ZEN大学 ── 日本の新しい挑戦
# ============================================================
make_highlight_slide(
    "ZEN大学 ── 日本初の本格オンライン大学（2025年開学）",
    "年間38万円",
    "知能情報社会学部 定員3,500名",
    [
        "6分野を横断して学ぶオーダーメイドカリキュラム",
        "　数理／情報／文化・思想／社会／経済／デジタル産業",
        "279科目から自分だけの時間割を設計",
        "大半がオンデマンド授業（自分のペースで学習）",
        "3種のメンターが手厚くサポート",
        "　クラス・コーチ／アカデミック・アドバイザー",
        "　／キャリアアドバイザー",
        "pixiv提携科目や東浩紀氏の対話型講座も",
    ],
    notes=(
        "日本のオンライン大学の新しい動きとして、2025年4月に開学した\n"
        "ZEN大学を紹介しておきたいと思います。\n\n"
        "年間の授業料は38万円。これはアメリカのオンライン大学と比べると\n"
        "圧倒的に安いです。知能情報社会学部という1つの学部に定員3,500名、\n"
        "279科目から自分だけの時間割を組むオーダーメイドカリキュラムが特徴です。\n\n"
        "数理、情報、文化・思想、社会、経済、デジタル産業という6つの分野を横断して学べて、\n"
        "大半の授業がオンデマンド。自分のペースで進められます。\n\n"
        "サポート体制も充実していて、クラス・コーチ、アカデミック・アドバイザー、\n"
        "キャリアアドバイザーという3種類のメンターがつきます。\n"
        "アメリカのオンライン高校のサポート体制に近い手厚さですよね。\n\n"
        "pixivとの提携科目があったり、東浩紀さんの対話型講座があったり、\n"
        "従来の日本の大学にはないユニークな試みが始まっています。\n"
        "日本のオンライン大学はまだ黎明期ですが、\n"
        "ZEN大学がその成功モデルになれるかどうか、注目されています。"
    )
)

# ============================================================
# SLIDE 10: 日米欧 費用比較
# ============================================================
make_three_column_slide(
    "学費の比較 ── 日米欧でこれだけ違う",
    [
        ("アメリカ", [
            "【オンライン高校】",
            "年間$1,800〜$2,800",
            "（約27万〜42万円）",
            "高額校は$20,000超もあり",
            "",
            "【オンライン大学】",
            "学士号：$40,536〜$63,185",
            "（約600万〜950万円）",
            "修士号：$19,000〜$27,000",
            "ただし通学費・寮費は不要",
        ], BLUE_ACCENT),
        ("ヨーロッパ", [
            "【オンライン高校】",
            "年間€630〜€5,900",
            "（約10万〜90万円）",
            "",
            "【オンライン大学】",
            "EU/EEA圏：無料〜低額",
            "非EU圏：年€5,000〜€18,000",
            "ドイツの州立大学は非常に安価",
            "「教育は人権」の理念が根底に",
        ], GREEN_ACCENT),
        ("日本", [
            "【通信制高校】",
            "公立：年間1.6万〜5万円",
            "私立：年間10万〜100万円",
            "就学支援金で実質無料も可能",
            "",
            "【オンライン大学】",
            "ZEN大学：年間38万円",
            "",
            "全日制高校の平均：約51万円/年",
            "通信制は大幅にコスト削減可能",
        ], ORANGE_ACCENT),
    ],
    notes=(
        "ここで学費を一気に比較してみましょう。3地域でかなり違いがあります。\n\n"
        "アメリカのオンライン高校は年間1,800ドルから2,800ドル、\n"
        "日本円で約27万円から42万円です。ただし高額なところは2万ドル、\n"
        "つまり300万円を超える学校もあります。\n"
        "大学レベルになると学士号の取得に600万円から950万円。\n"
        "アメリカの教育費はやはり高いですが、\n"
        "通学費や寮費がかからないのはオンラインの大きなメリットです。\n\n"
        "ヨーロッパは「教育は人権」という理念があるので、\n"
        "EU/EEA圏の学生は大学の授業料がほぼ無料というケースが多いです。\n"
        "非EU圏の学生でも年間5,000ユーロからで、アメリカと比べると格段に安い。\n\n"
        "そして日本。公立の通信制高校なら年間1万6千円から5万円。\n"
        "就学支援金を使えば実質無料になることもあります。\n"
        "全日制の高校は平均51万円ですから、通信制のコストメリットは明白です。\n"
        "ZEN大学は年間38万円。国際的に見ても非常に安い設定と言えます。\n"
        "ただし「安かろう悪かろう」ではなく、\n"
        "中身の質をどう保証するかが次のポイントになります。"
    )
)

# ============================================================
# SLIDE 11: 質保証の比較
# ============================================================
make_content_slide(
    "教育の質をどう保証している？── 認定制度の比較",
    [
        "【アメリカ】「地域認定」が信頼性の最高基準",
        "  ・Cogniaなどの認定機関が厳格に審査",
        "  ・認定なし = 単位互換できない、就職で不利になる",
        "  ・「ディプロマミル」（偽の学位販売業者）を排除する仕組み",
        "",
        "【ヨーロッパ】ESG = 欧州全体の共通品質フレームワーク",
        "  ・どの国・どの大学でも同じ基準で質を担保",
        "  ・国境を越えた単位互換・学位認定のベースになっている",
        "",
        "【日本】文部科学省の学習指導要領に基づく運営",
        "  ・通信制高校は文科省の規定で卒業要件が定められている",
        "  ・ただし、オンライン教育に特化した包括的な質保証の",
        "    フレームワークはまだ確立途上",
        "  ・国際的な認定の調和が今後の課題",
    ],
    notes=(
        "学費が安いのは良いことですが、問題は教育の質ですよね。\n"
        "各地域でどうやって質を保証しているのか見てみましょう。\n\n"
        "アメリカでは「地域認定」という仕組みが最も信頼されています。\n"
        "Cogniaなどの認定機関が学校を厳しく審査して、基準を満たしていれば認定を出す。\n"
        "認定を受けていない学校の単位は他の大学に持っていけないし、\n"
        "就職でも不利になります。\n"
        "実はアメリカには「ディプロマミル」と呼ばれる\n"
        "偽の学位を売る業者も存在するので、\n"
        "認定は消費者保護の役割も果たしているんです。\n\n"
        "ヨーロッパにはESGという共通の質保証基準があります。\n"
        "フランスの大学で取った単位がドイツの大学でも認められる。\n"
        "国をまたいで学んだり働いたりすることを前提にした仕組みで、\n"
        "これはヨーロッパの大きな強みです。\n\n"
        "日本はどうかというと、通信制高校は文科省の規定でしっかり管理されていますが、\n"
        "オンライン教育に特化した包括的な質保証フレームワークは、\n"
        "まだ確立されていないのが現状です。\n"
        "今後ZEN大学のようなオンライン専門大学が増えていくなかで、\n"
        "この仕組み作りが急務になってくると思います。"
    )
)

# ============================================================
# SLIDE 12: サポート体制の比較
# ============================================================
make_content_slide(
    "学習サポートの比較 ── 「放っておかれる」か「伴走してもらえる」か",
    [
        "【アメリカ】キャリアまで見据えた「伴走型サポート」",
        "  ・科目別教員＋ホームルーム教員＋専任カウンセラーの3層構造",
        "  ・大学出願、奨学金、入試対策まで一貫してサポート",
        "  ・課外活動やクラブで社会性も補完",
        "",
        "【ヨーロッパ】テクノロジーを活かした学習支援",
        "  ・FernUniversitätの演習システム：即時フィードバック＋難易度自動調整",
        "  ・COIL（国際共同オンライン学習）で海外の学生と協働",
        "",
        "【日本】学校によって差が大きい",
        "  ・N高：アバターチャットで交流促進＋多彩な課外授業",
        "  ・ゼロ高：専属コーチが目標設定から伴走",
        "  ・サポート校で学習面・精神面を補完するケースも",
        "  ・ただし公立通信制では自学自習が基本 → 孤立リスク",
    ],
    notes=(
        "オンライン学習で一番大事なのは、実はサポート体制です。\n"
        "放っておかれるのか、伴走してもらえるのかで、結果はまったく違ってきます。\n\n"
        "アメリカのオンライン高校は、ここが本当に手厚い。\n"
        "科目の先生、ホームルームの先生、そしてカウンセラーという3層のサポートがあって、\n"
        "大学の出願から奨学金まで一貫して面倒を見てくれます。\n"
        "社会性の面でも、オンラインのクラブ活動やNational Honor Societyのような\n"
        "課外活動で補完する仕組みが用意されています。\n\n"
        "ヨーロッパでは、テクノロジーを活かしたサポートが特徴的です。\n"
        "ドイツのFernUniversitätには演習システムがあって、\n"
        "解答すると即座にフィードバックが返ってきて、\n"
        "しかも難易度が自動で調整されるんです。\n"
        "またCOILという仕組みで、別の国の学生とオンラインで\n"
        "一緒にプロジェクトをやる機会も作られています。\n\n"
        "日本はどうかというと、正直なところ学校によって差が大きいです。\n"
        "N高やゼロ高のように手厚いサポートを提供している学校がある一方で、\n"
        "公立の通信制高校では基本的に自学自習で、\n"
        "孤立して学習が止まってしまうリスクも指摘されています。\n"
        "この格差をどう埋めていくかが、日本の通信制教育の大きな課題です。"
    )
)

# ============================================================
# SLIDE 13: MOOCsの活用
# ============================================================
make_content_slide(
    "MOOCsの世界 ── Coursera・Udemyで何ができるか",
    [
        "【Coursera】 月額$59〜 / 年額$399〜",
        "  ・Google、MIT、ハーバード等の講座が9,000コース以上",
        "  ・修了証は履歴書に書ける → 転職・社内評価に繋がった事例も",
        "  ・企業向けサービスも充実（Coursera for Teams）",
        "",
        "【Udemy】 コースごとの買い切り（セールで1,200円台も）",
        "  ・20万コース以上、最も幅広い分野をカバー",
        "",
        "【edX】 非営利の教育プラットフォーム",
        "  ・東大がMIT・ハーバードと連携して講座を公開",
        "",
        "【オンライン英会話】 月額2,000〜13,000円",
        "  ・スキマ時間で学習、費用対効果が高い",
        "",
        "ポイント：講義は無料でも、修了証を取るには有料プランが必要なケースが多い",
    ],
    notes=(
        "ここまで高校や大学の制度を見てきましたが、\n"
        "もう1つ大きな選択肢があります。MOOCsです。\n"
        "Massive Open Online Courses、大規模公開オンライン講座のことです。\n\n"
        "代表格はCoursera。Googleの認定証やMITの講座が受けられて、\n"
        "修了証は実際に履歴書に書けます。\n"
        "「Courseraの修了証がきっかけで転職できた」\n"
        "「社内の評価が上がった」という事例も報告されています。\n"
        "月額59ドルから、年間プランだと399ドルからです。\n\n"
        "Udemyはもっと手軽で、セールのときは1,200円台でコースが買えることもあります。\n"
        "20万コース以上あるので、ほぼどんな分野でも見つかるでしょう。\n\n"
        "edXは非営利性が特徴で、東京大学がMITやハーバードと\n"
        "連携した講座を公開していたりもします。\n\n"
        "1つ注意していただきたいのは、講義自体は無料で見られるケースが多いんですが、\n"
        "修了証を取得するには有料プランへの加入が必要になることがほとんどです。\n"
        "でもGoogle認定証のような就職に直結する資格が\n"
        "自宅から取れるというのは、大きな武器になりますよね。"
    )
)

# ============================================================
# SLIDE 14: 日本への提言
# ============================================================
make_content_slide(
    "日本のオンライン教育に必要な5つのこと",
    [
        "1. 質保証の仕組みを整える",
        "   → 欧州のESGのように、オンライン教育専用の品質基準を作る",
        "",
        "2. デジタルデバイドを解消する",
        "   → ネット環境やデバイスが手に入らない人を取り残さない",
        "",
        "3. オンラインと対面を組み合わせた「ハイブリッド型」を広げる",
        "   → 欧米のCOIL（国際共同学習）のように、交流の機会を増やす",
        "",
        "4. 教員のオンライン指導スキルを底上げする",
        "   → ファシリテーター、コースデザイナーとしての新しい役割",
        "",
        "5. キャリア支援を充実させる",
        "   → 米国のように出願・就職まで一貫してサポートする体制へ",
    ],
    notes=(
        "ここまでの比較を踏まえて、日本のオンライン教育に何が必要なのか、\n"
        "5つのポイントに整理してみました。\n\n"
        "1つ目は質保証の仕組みです。\n"
        "ヨーロッパにはESGという共通基準がありますが、日本にはまだありません。\n"
        "オンライン教育が増えていくなかで、\n"
        "「この学校は本当に信頼できるのか」を判断する仕組みが必要です。\n\n"
        "2つ目はデジタルデバイドの解消。\n"
        "ネット環境やパソコンが手に入らない人を取り残してはいけません。\n\n"
        "3つ目はハイブリッド型の学習です。\n"
        "オンラインだけだと社会性が育ちにくいという課題がありましたよね。\n"
        "ヨーロッパのCOILのように、国を超えた学生同士の協働をオンラインで実現しながら、\n"
        "対面の機会も組み合わせていく。そのバランスが大事です。\n\n"
        "4つ目は教員の育成です。\n"
        "オンライン教育では教員に「ファシリテーター」や「コースデザイナー」\n"
        "としての新しいスキルが求められますが、\n"
        "多くの教員がまだ対応できていないのが現状です。\n\n"
        "5つ目はキャリア支援の充実です。\n"
        "アメリカのオンライン高校のように、\n"
        "大学出願から就職まで一貫してサポートする体制を日本でも広げていくべきです。"
    )
)

# ============================================================
# SLIDE 15: まとめ
# ============================================================
make_content_slide(
    "まとめ ── これからの「学び方」をどう選ぶか",
    [
        "【世界の潮流】",
        "  ・オンライン教育は一時的なブームではない",
        "  ・市場は年率7.75%で成長、2030年には約85兆円規模に",
        "",
        "【各国の強み】",
        "  ・アメリカ：大学進学を徹底サポートする「伴走型」",
        "  ・ヨーロッパ：「教育は人権」、質保証の共通基準ESG",
        "  ・日本：通信制の柔軟性とコストの安さ、進学率は過去最高",
        "",
        "【あなたにとっての「最適解」は？】",
        "  ・学びの選択肢は、通信制、MOOCs、ギャップイヤー、留学…",
        "    もはや「全日制の学校に通う」だけが正解ではない時代",
        "  ・大事なのは「何を学ぶか」ではなく「どう学ぶか」を自分で選ぶこと",
    ],
    notes=(
        "最後に、今日の内容をまとめます。\n\n"
        "まず世界の潮流として、オンライン教育はもう一時的なブームではありません。\n"
        "コロナが終わっても成長は止まっておらず、\n"
        "2030年には85兆円規模になると予測されています。\n\n"
        "アメリカは大学進学を徹底的にサポートする伴走型。\n"
        "ヨーロッパは「教育は人権」という思想と共通の質保証基準。\n"
        "日本は通信制の柔軟性とコストの安さが強みで、進学率は過去最高を更新中。\n"
        "それぞれの地域に、それぞれの良さがあります。\n\n"
        "大事なのは、皆さん自身がどう学ぶかを自分で選ぶということです。\n"
        "通信制高校、MOOCs、ギャップイヤー、海外留学、ワーキングホリデー。\n"
        "もはや「全日制の学校に毎日通う」だけが唯一の正解ではない時代です。\n"
        "今日の動画が、皆さんの学び方を考えるきっかけになれば嬉しいです。"
    )
)

# ============================================================
# SLIDE 16: CTA ＋ Thank You
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=DARK_BG)
add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(4), fill_color=GOLD)
add_shape(slide, 0, SLIDE_HEIGHT - Pt(4), SLIDE_WIDTH, Pt(4), fill_color=GOLD)
add_text_box(slide, Inches(1), Inches(1.0), Inches(8), Inches(1.0),
             "ご視聴ありがとうございました", font_size=30, bold=True, color=GOLD_LIGHT,
             alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
add_text_box(slide, Inches(1), Inches(2.2), Inches(8), Inches(1.5),
             "チャンネル登録・高評価で応援していただけると嬉しいです\n\n"
             "コメント欄で感想や質問もお待ちしています\n"
             "「自分はこう学んでいる」というシェアも大歓迎です",
             font_size=15, color=CREAM,
             alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
add_text_box(slide, Inches(1), Inches(4.0), Inches(8), Inches(0.5),
             "次回予告：「海外の通信制高校に日本から入学する方法」",
             font_size=13, bold=True, color=GOLD_ACCENT,
             alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
add_text_box(slide, Inches(0.5), Inches(4.8), Inches(9), Inches(0.4),
             "Harmonic Insight", font_size=11, color=GOLD,
             alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
set_notes(slide, (
    "ご視聴いただきありがとうございました。\n\n"
    "今日は日本・アメリカ・ヨーロッパのオンライン教育を比較してきましたが、\n"
    "いかがだったでしょうか。\n\n"
    "この動画が参考になったと思っていただけたら、\n"
    "ぜひチャンネル登録と高評価をお願いします。\n"
    "コメント欄で感想や質問もお待ちしています。\n"
    "「自分はこういう学び方をしている」「こんな選択肢もあるよ」\n"
    "という皆さんのシェアも大歓迎です。\n\n"
    "次回は「海外の通信制高校に日本から入学する方法」をテーマに、\n"
    "具体的なステップを解説していく予定です。\n"
    "それでは、また次の動画でお会いしましょう。"
))

# ===== Save =====
output_path = "海外学習_留学_グローバルオンライン教育_プレゼンテーション.pptx"
prs.save(output_path)
print(f"プレゼンテーションを保存しました: {output_path}")
print(f"合計スライド数: {len(prs.slides)}")
