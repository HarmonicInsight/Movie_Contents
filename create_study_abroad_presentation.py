#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
海外学習・留学 × グローバルオンライン教育 プレゼンテーション生成スクリプト
Harmonic Insight テンプレートスタイルに準拠

2つのPDFレポートを統合:
1. グローバル教育トレンドとオンライン学習を連携させたウェブサイト構築戦略
2. グローバルオンライン教育の現状と日本の立ち位置：高校・大学における比較分析と展望
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import copy

# ===== Color Palette (Harmonic Insight Style) =====
GOLD = RGBColor(0x8B, 0x75, 0x36)
GOLD_LIGHT = RGBColor(0xC9, 0xA8, 0x4C)
GOLD_ACCENT = RGBColor(0xB8, 0x86, 0x0B)
DARK_BG = RGBColor(0x1A, 0x1A, 0x1A)
CREAM = RGBColor(0xF5, 0xF0, 0xE8)
CREAM2 = RGBColor(0xF5, 0xF0, 0xE6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x33, 0x33, 0x33)
TEXT_GRAY = RGBColor(0x66, 0x66, 0x66)
TEXT_BROWN = RGBColor(0x3C, 0x2F, 0x1E)
BROWN_LIGHT = RGBColor(0x8B, 0x73, 0x55)
RED_ACCENT = RGBColor(0xC0, 0x39, 0x2B)
GREEN_ACCENT = RGBColor(0x27, 0xAE, 0x60)
BLUE_ACCENT = RGBColor(0x2C, 0x7A, 0xB0)
ORANGE_ACCENT = RGBColor(0xE6, 0x7E, 0x22)
PURPLE_ACCENT = RGBColor(0x8E, 0x44, 0xAD)
TABLE_HEADER_BG = RGBColor(0x8B, 0x75, 0x36)
TABLE_ROW_LIGHT = RGBColor(0xF9, 0xF6, 0xF0)
TABLE_ROW_DARK = RGBColor(0xF0, 0xE8, 0xD5)

SLIDE_WIDTH = Emu(9144000)   # 10 inches
SLIDE_HEIGHT = Emu(5143500)  # 5.625 inches

prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT

# ===== Helper Functions =====
def add_shape(slide, left, top, width, height, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width or 1)
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
        p.level = 0
    return txBox

def add_gold_line(slide, left, top, width):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()
    return shape

def make_title_slide(title, subtitle=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    # Dark background
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=DARK_BG)
    # Gold accent line top
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(4), fill_color=GOLD)
    # Gold accent line bottom
    add_shape(slide, 0, SLIDE_HEIGHT - Pt(4), SLIDE_WIDTH, Pt(4), fill_color=GOLD)
    # Title
    add_text_box(slide, Inches(0.8), Inches(1.5), Inches(8.4), Inches(1.5),
                 title, font_size=32, bold=True, color=GOLD_LIGHT,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    if subtitle:
        add_text_box(slide, Inches(1), Inches(3.2), Inches(8), Inches(1),
                     subtitle, font_size=16, color=CREAM,
                     alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    # Harmonic Insight brand
    add_text_box(slide, Inches(0.5), Inches(4.8), Inches(9), Inches(0.4),
                 "Harmonic Insight", font_size=11, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    return slide

def make_section_slide(section_title, section_number=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=RGBColor(0x2A, 0x24, 0x18))
    # Gold left bar
    add_shape(slide, Inches(0.3), Inches(1.5), Pt(4), Inches(2.5), fill_color=GOLD)
    if section_number:
        add_text_box(slide, Inches(0.6), Inches(1.3), Inches(8), Inches(0.6),
                     section_number, font_size=14, color=GOLD_LIGHT,
                     alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_text_box(slide, Inches(0.6), Inches(1.9), Inches(8.5), Inches(1.5),
                 section_title, font_size=28, bold=True, color=CREAM,
                 alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_text_box(slide, Inches(0.5), Inches(4.8), Inches(9), Inches(0.4),
                 "Harmonic Insight", font_size=10, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    return slide

def make_content_slide(title, bullet_items, note_text=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=CREAM)
    # Top gold bar
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(3), fill_color=GOLD)
    # Title area
    add_shape(slide, 0, Pt(3), SLIDE_WIDTH, Inches(0.7), fill_color=WHITE)
    add_text_box(slide, Inches(0.5), Inches(0.05), Inches(9), Inches(0.6),
                 title, font_size=20, bold=True, color=GOLD,
                 alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_gold_line(slide, Inches(0.5), Inches(0.72), Inches(9))
    # Bullet content
    add_bullet_text(slide, Inches(0.6), Inches(0.9), Inches(8.8), Inches(3.8),
                    bullet_items, font_size=13, color=TEXT_DARK)
    if note_text:
        add_text_box(slide, Inches(0.6), Inches(4.6), Inches(8.8), Inches(0.5),
                     note_text, font_size=10, color=TEXT_GRAY,
                     alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    # Footer
    add_text_box(slide, Inches(0.5), Inches(5.1), Inches(9), Inches(0.3),
                 "Harmonic Insight", font_size=9, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    return slide

def make_two_column_slide(title, left_title, left_items, right_title, right_items):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=CREAM)
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(3), fill_color=GOLD)
    add_shape(slide, 0, Pt(3), SLIDE_WIDTH, Inches(0.7), fill_color=WHITE)
    add_text_box(slide, Inches(0.5), Inches(0.05), Inches(9), Inches(0.6),
                 title, font_size=20, bold=True, color=GOLD,
                 alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_gold_line(slide, Inches(0.5), Inches(0.72), Inches(9))
    # Left column
    add_shape(slide, Inches(0.4), Inches(0.9), Inches(4.3), Inches(0.4), fill_color=GOLD)
    add_text_box(slide, Inches(0.5), Inches(0.92), Inches(4.1), Inches(0.35),
                 left_title, font_size=13, bold=True, color=WHITE,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    add_bullet_text(slide, Inches(0.5), Inches(1.4), Inches(4.1), Inches(3.2),
                    left_items, font_size=11, color=TEXT_DARK)
    # Right column
    add_shape(slide, Inches(5.2), Inches(0.9), Inches(4.3), Inches(0.4), fill_color=BLUE_ACCENT)
    add_text_box(slide, Inches(5.3), Inches(0.92), Inches(4.1), Inches(0.35),
                 right_title, font_size=13, bold=True, color=WHITE,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    add_bullet_text(slide, Inches(5.3), Inches(1.4), Inches(4.1), Inches(3.2),
                    right_items, font_size=11, color=TEXT_DARK)
    add_text_box(slide, Inches(0.5), Inches(5.1), Inches(9), Inches(0.3),
                 "Harmonic Insight", font_size=9, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    return slide

def make_three_column_slide(title, col_data):
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
    return slide

def make_highlight_slide(title, key_number, key_label, description_items):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=CREAM)
    add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(3), fill_color=GOLD)
    add_shape(slide, 0, Pt(3), SLIDE_WIDTH, Inches(0.7), fill_color=WHITE)
    add_text_box(slide, Inches(0.5), Inches(0.05), Inches(9), Inches(0.6),
                 title, font_size=20, bold=True, color=GOLD,
                 alignment=PP_ALIGN.LEFT, font_name="Yu Gothic")
    add_gold_line(slide, Inches(0.5), Inches(0.72), Inches(9))
    # Key number highlight
    add_shape(slide, Inches(0.5), Inches(1.0), Inches(3.5), Inches(2.0),
              fill_color=RGBColor(0x2A, 0x24, 0x18))
    add_text_box(slide, Inches(0.6), Inches(1.1), Inches(3.3), Inches(1.2),
                 key_number, font_size=40, bold=True, color=GOLD_LIGHT,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    add_text_box(slide, Inches(0.6), Inches(2.2), Inches(3.3), Inches(0.6),
                 key_label, font_size=14, color=CREAM,
                 alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
    # Description
    add_bullet_text(slide, Inches(4.3), Inches(1.0), Inches(5.2), Inches(3.5),
                    description_items, font_size=12, color=TEXT_DARK)
    add_text_box(slide, Inches(0.5), Inches(5.1), Inches(9), Inches(0.3),
                 "Harmonic Insight", font_size=9, color=GOLD,
                 alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")
    return slide

# ============================================================
# SLIDE 1: Title Slide
# ============================================================
make_title_slide(
    "グローバル教育トレンドと\nオンライン学習の未来",
    "海外学習・留学から見る日米欧の教育比較と\nウェブサイト構築戦略"
)

# ============================================================
# SLIDE 2: Agenda / 目次
# ============================================================
make_content_slide(
    "本日のアジェンダ",
    [
        "1. オンライン教育のグローバル市場動向",
        "2. 海外の学習・キャリア観 ── 米国・英国・アジア",
        "3. 日米欧オンライン高校・大学の比較分析",
        "4. 日本の通信制教育の現状と進路状況",
        "5. MOOCs市場と学習プラットフォーム比較",
        "6. コンテンツ連携戦略とサイト構造の設計",
        "7. 収益化モデルとロードマップ",
        "8. 成功事例と提言",
    ]
)

# ============================================================
# SECTION 1: グローバル市場動向
# ============================================================
make_section_slide("オンライン教育の\nグローバル市場動向", "SECTION 01")

make_highlight_slide(
    "世界のオンライン教育市場 ── 爆発的成長",
    "3,888億$",
    "2025年 世界市場予測",
    [
        "2000年から市場規模は900%拡大",
        "2030年には5,647億ドル規模へ（CAGR 7.75%）",
        "2029年までに約11.21億人がオンライン学習にアクセス",
        "COVID-19後も米国のオンライン学習率は28%を維持",
        "   （パンデミック前は10%未満）",
        "AI・VR/AR・クラウド技術の統合が成長を加速",
    ]
)

make_content_slide(
    "オンライン学習の普遍的な利点と課題",
    [
        "【利点】",
        "  ・地理的・時間的制約からの解放 ── いつでもどこでも学べる",
        "  ・個別最適化された学習（アダプティブラーニング）",
        "  ・教師の働き方の多様化と効率化",
        "  ・対話の質向上と平等な学習環境の実現",
        "",
        "【課題】",
        "  ・社会的交流の欠如とモチベーション維持の難しさ",
        "  ・デジタルデバイド ── 低所得層・地方の格差",
        "  ・教育の質のばらつきと学業不正への懸念",
        "  ・実践的・実習的科目のオンライン化の限界",
        "  ・教員のオンライン指導スキル不足",
    ]
)

# ============================================================
# SECTION 2: 海外学生の学習・キャリア観
# ============================================================
make_section_slide("海外の高校生・大学生の「今」\n米国・英国・アジアの学習・キャリア観", "SECTION 02")

make_three_column_slide(
    "各国の教育文化と学生の「今」── 比較分析",
    [
        ("アメリカ", [
            "学業＋人間性・コミットメントを重視",
            "課外活動の「継続性」が評価される",
            "クラブ・ボランティア・インターンシップ",
            "特定分野のインターン経験が強力なアピール",
            "大学入試は「なぜその活動をしたか」を深掘り",
        ], BLUE_ACCENT),
        ("イギリス", [
            "「ギャップイヤー」文化が定着",
            "若者の83%が国内就業、56%が海外滞在",
            "企業の94%がギャップイヤー経験者の採用に前向き",
            "主体性・異文化理解・問題解決能力を評価",
            "ワーホリ：年間144万〜226万円",
        ], GREEN_ACCENT),
        ("アジア（韓国・シンガポール）", [
            "韓国：「SKY大学」がエリートの象徴",
            "高校生は深夜まで塾通いの受験競争",
            "シンガポール：大学進学率3〜4割",
            "ポリテクニック・ITEの職業訓練が主流",
            "効率的な学習とキャリア直結を重視",
        ], ORANGE_ACCENT),
    ]
)

make_content_slide(
    "日本の若者層・社会人の学習ニーズ",
    [
        "【通信制教育の利用者が抱える課題】",
        "  ・メリット：自分のペースで学習、費用が安価",
        "  ・デメリット：大学受験の難しさ、友人ができにくい",
        "  ・学習の継続性、社会的評価、孤立感の克服が必要",
        "",
        "【オンライン学習者に求められるもの】",
        "  ・知識だけでなく「社会的側面」「モチベーション維持」",
        "  ・コミュニティとしての学習プラットフォーム",
        "  ・海外学生のリアルな生活 → 「憧れ」「目標」としての機能",
        "  ・N高等学校はアバターチャットで生徒間交流を促進",
    ]
)

# ============================================================
# SECTION 3: 日米欧 オンライン教育の比較
# ============================================================
make_section_slide("日米欧オンライン高校・大学の\n比較分析", "SECTION 03")

make_two_column_slide(
    "アメリカのオンライン高校 ── 大学進学準備の徹底",
    "カリキュラム・学習モデル",
    [
        "・コア科目＋多様な選択科目",
        "・Honors / AP / NCAA承認コース",
        "・週30時間の学習推奨（80-90%がPC上）",
        "・大学進学に特化した包括的プログラム",
        "・代表校：Pearson Online Academy",
        "        Connections Academy",
    ],
    "サポート体制",
    [
        "・科目別教員＋ホームルーム教員",
        "・専任スクールカウンセラー常駐",
        "・4年間の学習計画作成支援",
        "・大学出願・奨学金・入試対策サポート",
        "・National Honor Society等の課外活動",
        "・学費：年間$1,800〜$2,800（FT）",
    ],
)

make_two_column_slide(
    "ヨーロッパのオンライン教育 ── 機会均等と質保証",
    "高校教育の特徴",
    [
        "・「高等教育は人権」という理念に基づく",
        "・英国/米国カリキュラム、IBプログラム",
        "・8つのキーコンピテンシーの育成",
        "・欧州の価値観の醸成を重視",
        "・CNED（仏）：欧州最大の遠隔教育機関",
        "・学費：年間€630〜€5,900（学校による）",
    ],
    "大学教育と質保証",
    [
        "・Open University（英）：多分野の学位提供",
        "・FernUniversität in Hagen（独）",
        "・ロンドン大学：オンライン学位プログラム",
        "・ESG（欧州質保証基準）が共通フレームワーク",
        "・EU/EEA圏の学生は授業料無料〜低額",
        "・非EU学生：年間€5,000〜€18,000",
    ],
)

make_two_column_slide(
    "日本の通信制教育 ── 柔軟性と多様なニーズへの対応",
    "制度の特徴",
    [
        "・毎日の登校不要・単位制で自分のペース",
        "・レポート提出＋スクーリング＋単位認定試験",
        "・卒業要件：3年以上在籍、74単位以上",
        "・特別活動30単位時間以上",
        "・公立：年間3〜5万円程度",
        "・私立：年間10〜100万円程度",
    ],
    "主要校と進路状況",
    [
        "・N高等学校：ネットコースで効率学習",
        "  入学金1万＋1単位7,200円＋施設費5万",
        "・ゼロ高等学院：「座学より行動」を重視",
        "  起業家育成・留学・インターンシップ",
        "・大学進学率：約27%（2023年度）",
        "・専門学校等含む合計進学率：52%超",
        "  ※通信制史上最高水準",
    ],
)

make_highlight_slide(
    "日米欧 費用比較のポイント",
    "38万円/年",
    "ZEN大学 ── 日本初の本格オンライン大学",
    [
        "【米国】オンライン高校：$1,800〜$20,000+/年",
        "　　　 オンライン大学：学士$40,536〜$63,185",
        "",
        "【欧州】EU圏：無料〜低額 / 非EU：€5,000〜€18,000/年",
        "　　　 Open University等で費用対効果が高い",
        "",
        "【日本】公立通信制：年間1.6万〜5万円（就学支援金適用）",
        "　　　 私立通信制：年間44万円程度",
        "　　　 ZEN大学：年間38万円（2025年4月開学）",
        "　　　 　知能情報社会学部・定員3,500名・279科目",
    ]
)

# ============================================================
# SECTION 4: 質保証と認定制度
# ============================================================
make_section_slide("質保証と認定制度の\n国際比較", "SECTION 04")

make_three_column_slide(
    "質保証・認定制度の日米欧比較",
    [
        ("アメリカ", [
            "地域認定が最も権威が高い",
            "CHEA・米国教育省OPEで確認推奨",
            "「ディプロマミル」対策として重要",
            "Cognia認定：世界90カ国以上で信頼",
            "DEAC：オンライン高校認定",
            "認定＝単位互換・雇用主評価に直結",
        ], BLUE_ACCENT),
        ("ヨーロッパ", [
            "ESG（欧州高等教育質保証基準）",
            "提供場所・方法に関わらず適用",
            "EUCDL：遠隔教育の品質保証",
            "英国DfE：OEAS認定スキーム",
            "QAHE：K-12国際基準認定",
            "国境を越えた学習・単位認定を促進",
        ], GREEN_ACCENT),
        ("日本", [
            "文部科学省の規定に基づく運営",
            "学習指導要領・卒業要件の規定",
            "文科省認定の社会通信教育",
            "オンライン特化の包括的フレームワークは未確立",
            "国際的調和の推進が課題",
            "ZEN大学等の新規認定が注目",
        ], ORANGE_ACCENT),
    ]
)

# ============================================================
# SECTION 5: MOOCs市場
# ============================================================
make_section_slide("オンライン学習プラットフォームの\n市場分析と連携モデル", "SECTION 05")

make_content_slide(
    "主要MOOCsプラットフォーム比較",
    [
        "【Coursera】月額$59〜 / 年額$399〜",
        "  9,000+コース、修了証・学位取得可、Google認定証が人気",
        "  有名大学・企業との連携が強く、社会的信用度が高い",
        "",
        "【Udemy】コースごとの買い切り（セール頻度高）",
        "  20万+コース、幅広い分野で学習可能",
        "",
        "【edX】非営利性が特徴、東大もMIT・ハーバードと連携",
        "",
        "【N高等学校】年額63,000円〜（ネットコース）",
        "  ICT活用の効率学習＋ネット上の交流機能",
        "",
        "【オンライン英会話】月額2,000〜13,000円、スキマ時間学習",
    ]
)

make_two_column_slide(
    "連携モデルの検討 ── サイトの独自価値を創出",
    "アフィリエイト型",
    [
        "・記事やレビューからMOOCs・留学エージェントへ送客",
        "・導入コストが低い",
        "・コンテンツ制作に集中できる",
        "・成果報酬型の収益モデル",
        "・Coursera、Udemy、留学ジャーナル等",
    ],
    "キュレーション型（LSCM）",
    [
        "・独自の「学習ロードマップ」を設計",
        "・特定MOOCs・英会話を推奨",
        "・サイトの専門性をアピール",
        "・ユーザーの信頼獲得に直結",
        "・「興味→キャリア」の学習プロセス全体を最適化",
        "  ＝ラーニング・サプライチェーン・マネジメント",
    ],
)

# ============================================================
# SECTION 6: コンテンツ連携戦略
# ============================================================
make_section_slide("コンテンツ連携戦略と\nサイト構造の設計", "SECTION 06")

make_content_slide(
    "コンセプト：「興味・関心」から「行動・学習」への導線設計",
    [
        "【ブランドコンセプト例】",
        "  「世界の学びを、あなたのキャリアに。」",
        "",
        "【コンテンツ戦略の2本柱】",
        "  1. 海外学生の「ライフスタイル」紹介コンテンツ",
        "     ・動画コンテンツ（AI音声生成ツールでコスト削減）",
        "     ・深掘りインタビュー記事",
        "     ・「なぜその大学？」「課外活動の成果は？」",
        "",
        "  2. オンライン学習プラットフォーム活用法",
        "     ・Google認定証取得の体験談",
        "     ・Coursera / Udemy / edX 比較記事",
        "     ・学習コスト・期間・履歴書記載方法の情報提供",
        "",
        "【導線の仕組み】",
        "  記事末尾に「このスキルを今すぐ身につけるには？」→ コースへ誘導",
    ]
)

make_content_slide(
    "ウェブサイト構造 ── UXを最適化するナビゲーション",
    [
        "【建設業DXの知見を応用】",
        "  ・BIM/CIM → サイトマップ・ワイヤーフレームとして捉える",
        "  ・「基本設計」→「実施設計」で手戻りを防止",
        "",
        "【運用体制のポイント】",
        "  ・CMS導入で非エンジニアでも更新可能",
        "  ・製造業のリーン生産方式を適用",
        "  ・「ムダ・ムラ・ムリ」の排除で運用効率化",
        "",
        "【動線分析】",
        "  ・ユーザーの行動を定期的に確認",
        "  ・設計した導線と実際の乖離を修正",
        "  ・ユーザー目線に立った継続改善",
    ]
)

# ============================================================
# SECTION 7: 収益化モデルとロードマップ
# ============================================================
make_section_slide("事業の持続可能性と\n収益化戦略", "SECTION 07")

make_three_column_slide(
    "3フェーズ収益化ロードマップ",
    [
        ("Phase 1：スモールスタート\n（3〜6ヶ月）", [
            "特定地域（例：米国）に絞りコンテンツ制作",
            "深掘りインタビュー＋動画",
            "CMS導入・AIナレーション活用",
            "Google AdSense＋アフィリエイト",
            "KPI：アクセス数・滞在時間・読了率",
        ], BLUE_ACCENT),
        ("Phase 2：コンテンツ拡充\n（6〜18ヶ月）", [
            "英国・アジアへ地域展開",
            "学習ロードマップ＋体験談拡充",
            "UX最適化・マッチング機能",
            "アフィリエイト本格化",
            "サブスクリプション検討",
            "KPI：CVR・CTR・修了率",
        ], GREEN_ACCENT),
        ("Phase 3：事業多角化\n（18ヶ月以降）", [
            "B2B DX人材育成サービス",
            "グローバル人材育成コンサル",
            "LMS・教育PF連携強化",
            "複数収益源で安定基盤構築",
            "KPI：顧客満足度・スキルアップ度",
        ], ORANGE_ACCENT),
    ]
)

make_content_slide(
    "収益化モデル一覧 ── ハイブリッド型が鍵",
    [
        "【広告型】Google AdSense等 → 低コスト / アクセス数依存",
        "",
        "【アフィリエイト型】MOOCs・留学エージェントへ送客 → 低コスト / 成約数依存",
        "",
        "【サブスクリプション型】有料会員限定コンテンツ → 中〜高コスト / 安定収益",
        "",
        "【マッチングサービス型】留学エージェント・企業研修紹介 → 中〜高コスト / 安定収益",
        "",
        "【コンテンツ販売型】独自学習ガイド・教材販売 → 中コスト / 販売数依存",
        "",
        "※ 単一モデルへの依存はリスク。複数モデルの組み合わせで",
        "  市場変動・アルゴリズム更新に対する耐性を確保",
    ]
)

# ============================================================
# SECTION 8: 成功事例と提言
# ============================================================
make_section_slide("成功事例から学ぶ統合メディアと\n日本への提言", "SECTION 08")

make_content_slide(
    "コンテンツマーケティング成功事例",
    [
        "【北欧、暮らしの道具店】",
        "  ライフスタイルWebマガジンで「世界観」を共有 → ブランド体験を提供",
        "",
        "【有隣堂しか知らない世界】",
        "  キャラクター「R.B.ブッコロー」でエンタメ性を掛け合わせ",
        "  → 本業に興味がなかった層にもアプローチ",
        "",
        "【となりのカインズさん】",
        "  現場の一次情報を発信 → 月間400万PV達成",
        "  専門性と信頼性がサイト価値の源泉",
        "",
        "【教訓】ユーザーが求めるのは「共感できる世界観」と",
        "  「信頼できるキャラクター」。海外学生の価値観・悩みを",
        "  ストーリーとして描くことがエンゲージメントの鍵",
    ]
)

make_content_slide(
    "日本のオンライン教育への5つの提言",
    [
        "1. 質保証と認定の国際的調和の推進",
        "   ESGを参考にオンライン教育特化の質保証基準を策定",
        "",
        "2. デジタルデバイド解消への包括的投資",
        "   高速インターネット整備・低所得世帯へのデバイス提供",
        "",
        "3. ハイブリッド学習モデルの推進と社会的交流の促進",
        "   COIL（国際共同オンライン学習）の導入・異文化理解",
        "",
        "4. 教員育成と専門性向上の体系化",
        "   オンライン指導特化の研修プログラム義務化",
        "",
        "5. キャリア・進路支援の強化と実社会との連携深化",
        "   総合型選抜対応、PBL、企業インターンシップの拡充",
    ]
)

# ============================================================
# FINAL SLIDE: Thank You
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_shape(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill_color=DARK_BG)
add_shape(slide, 0, 0, SLIDE_WIDTH, Pt(4), fill_color=GOLD)
add_shape(slide, 0, SLIDE_HEIGHT - Pt(4), SLIDE_WIDTH, Pt(4), fill_color=GOLD)
add_text_box(slide, Inches(1), Inches(1.5), Inches(8), Inches(1.2),
             "Thank You", font_size=36, bold=True, color=GOLD_LIGHT,
             alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
add_text_box(slide, Inches(1), Inches(2.8), Inches(8), Inches(1),
             "グローバル教育トレンドとオンライン学習の未来\n"
             "── 海外学習・留学から見る日米欧の教育比較と戦略提言 ──",
             font_size=14, color=CREAM,
             alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
add_text_box(slide, Inches(1), Inches(4.0), Inches(8), Inches(0.5),
             "出典：「グローバル教育トレンドとオンライン学習を連携させたウェブサイト構築戦略」\n"
             "「グローバルオンライン教育の現状と日本の立ち位置：高校・大学における比較分析と展望」",
             font_size=9, color=TEXT_GRAY,
             alignment=PP_ALIGN.CENTER, font_name="Yu Gothic")
add_text_box(slide, Inches(0.5), Inches(4.8), Inches(9), Inches(0.4),
             "Harmonic Insight", font_size=11, color=GOLD,
             alignment=PP_ALIGN.RIGHT, font_name="Yu Gothic")

# ===== Save =====
output_path = "海外学習_留学_グローバルオンライン教育_プレゼンテーション.pptx"
prs.save(output_path)
print(f"プレゼンテーションを保存しました: {output_path}")
print(f"合計スライド数: {len(prs.slides)}")
