"""pptx資料生成モジュール - 3C分析レポート"""

import io
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from config import PptxStyle
from models import ResearchResult


# --- フォント設定 ---

def _get_jp_font():
    """利用可能な日本語フォントを探す"""
    jp_fonts = [
        "Noto Sans JP", "Noto Sans CJK JP", "Hiragino Sans",
        "Hiragino Kaku Gothic ProN", "Yu Gothic", "Meiryo",
        "IPAGothic", "IPAPGothic",
    ]
    available = {f.name for f in fm.fontManager.ttflist}
    for font in jp_fonts:
        if font in available:
            return font
    return "sans-serif"


JP_FONT = _get_jp_font()
plt.rcParams["font.family"] = JP_FONT
plt.rcParams["axes.unicode_minus"] = False


# --- ヘルパー関数 ---

def _hex_to_rgb(hex_str: str) -> RGBColor:
    return RGBColor(int(hex_str[:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


def _add_text(tf, text, size=12, bold=False, color=None, alignment=None):
    """テキストフレームに段落を追加"""
    p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
    if tf.paragraphs[0].text:
        p = tf.add_paragraph()
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = _hex_to_rgb(color)
    try:
        run.font.name = PptxStyle.FONT_TITLE
    except Exception:
        run.font.name = PptxStyle.FONT_FALLBACK
    if alignment:
        p.alignment = alignment
    return p


def _set_shape_bg(shape, hex_color):
    """図形の背景色を設定"""
    shape.fill.solid()
    shape.fill.fore_color.rgb = _hex_to_rgb(hex_color)


# --- チャート生成（matplotlib → 画像バイト） ---

def _create_positioning_map(result: ResearchResult) -> bytes:
    """ポジショニングマップの画像を生成"""
    fig, ax = plt.subplots(figsize=(7, 5))

    comp = result.competitor
    ax_label_x = comp.positioning_axis_x or "軸1"
    ax_label_y = comp.positioning_axis_y or "軸2"

    # 対象企業のプロット
    target_pos = getattr(comp, "_target_position", (5, 5))
    ax.scatter([target_pos[0]], [target_pos[1]], s=200, c="#E94560", zorder=5, marker="*")
    ax.annotate(result.client_name, (target_pos[0], target_pos[1]),
                fontsize=10, fontweight="bold", color="#E94560",
                xytext=(8, 8), textcoords="offset points")

    # 直接競合
    for c in comp.direct_competitors:
        ax.scatter([c.position_x], [c.position_y], s=120, c="#0F3460", zorder=4)
        ax.annotate(c.name, (c.position_x, c.position_y),
                    fontsize=9, color="#333333",
                    xytext=(6, 6), textcoords="offset points")

    # 間接競合
    for c in comp.indirect_competitors:
        ax.scatter([c.position_x], [c.position_y], s=80, c="#999999", zorder=3, marker="D")
        ax.annotate(c.name, (c.position_x, c.position_y),
                    fontsize=8, color="#666666",
                    xytext=(6, 6), textcoords="offset points")

    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.set_xlabel(ax_label_x, fontsize=11)
    ax.set_ylabel(ax_label_y, fontsize=11)
    ax.set_title("ポジショニングマップ", fontsize=13, fontweight="bold")
    ax.grid(True, alpha=0.3)
    ax.axhline(y=5, color="#ccc", linestyle="--", linewidth=0.8)
    ax.axvline(x=5, color="#ccc", linestyle="--", linewidth=0.8)

    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _create_timeline(result: ResearchResult) -> bytes:
    """タイムライン画像を生成（最大8件、交互配置で重なり防止）"""
    events = result.company.timeline
    if not events:
        return b""

    fig, ax = plt.subplots(figsize=(10, 3.5))

    years = []
    labels = []
    for e in events:
        try:
            y = int(e.year[:4])
        except (ValueError, IndexError):
            continue
        years.append(y)
        # 説明文を15文字で切り詰め
        desc = e.description[:15] + "…" if len(e.description) > 15 else e.description
        labels.append(f"{e.year}\n{desc}")

    if not years:
        plt.close(fig)
        return b""

    # 最大8件に制限
    if len(years) > 8:
        years = years[:8]
        labels = labels[:8]

    y_pos = [0] * len(years)
    ax.scatter(years, y_pos, s=60, c="#0F3460", zorder=5)

    # 交互に上下に大きくずらして重なりを防止
    for i, (year, label) in enumerate(zip(years, labels)):
        if i % 2 == 0:
            offset_y = 30
            va = "bottom"
        else:
            offset_y = -30
            va = "top"
        ax.annotate(label, (year, 0), fontsize=6, ha="center", va=va,
                    xytext=(0, offset_y), textcoords="offset points",
                    arrowprops=dict(arrowstyle="-", color="#ccc", lw=0.5))

    ax.axhline(y=0, color="#0F3460", linewidth=2, alpha=0.5)
    ax.set_ylim(-1.5, 1.5)
    ax.set_yticks([])
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.set_title("企業沿革", fontsize=12, fontweight="bold")

    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# --- スライド生成 ---

def _slide_cover(prs: Presentation, result: ResearchResult):
    """表紙スライド"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank

    # 背景色
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = _hex_to_rgb(PptxStyle.PRIMARY)

    # タイトル
    left, top = Inches(1), Inches(2)
    width, height = Inches(10), Inches(1.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    _add_text(tf, f"{result.client_name}", size=36, bold=True, color=PptxStyle.TEXT_LIGHT, alignment=PP_ALIGN.LEFT)

    # サブタイトル
    txBox2 = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(10), Inches(1))
    tf2 = txBox2.text_frame
    _add_text(tf2, "3C分析レポート", size=20, color="CCCCCC", alignment=PP_ALIGN.LEFT)

    # 日付
    txBox3 = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(10), Inches(0.5))
    tf3 = txBox3.text_frame
    _add_text(tf3, f"分析実施日: {result.created_at}", size=12, color="999999", alignment=PP_ALIGN.LEFT)

    # 業界
    txBox4 = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(10), Inches(0.5))
    tf4 = txBox4.text_frame
    _add_text(tf4, f"業界: {result.industry}", size=12, color="999999", alignment=PP_ALIGN.LEFT)


def _slide_executive_summary(prs: Presentation, result: ResearchResult):
    """エグゼクティブサマリー"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # タイトル
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
    tf = txBox.text_frame
    _add_text(tf, "Executive Summary", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)

    # 区切り線
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
    _set_shape_bg(line, PptxStyle.HIGHLIGHT)
    line.line.fill.background()

    # 企業概要
    txBox2 = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(11), Inches(0.8))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    overview = result.company.business_overview or f"{result.client_name}は{result.industry}業界の企業です。"
    _add_text(tf2, overview, size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)

    # キーファインディング
    txBox3 = slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(11), Inches(0.5))
    tf3 = txBox3.text_frame
    _add_text(tf3, "Key Findings", size=PptxStyle.SIZE_HEADING, bold=True, color=PptxStyle.SECONDARY)

    y_offset = 2.8
    for i, finding in enumerate(result.key_findings[:5], 1):
        # 番号付きボックス
        num_shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(0.7), Inches(y_offset), Inches(0.35), Inches(0.35)
        )
        _set_shape_bg(num_shape, PptxStyle.HIGHLIGHT)
        num_shape.line.fill.background()
        num_tf = num_shape.text_frame
        num_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = num_tf.paragraphs[0].add_run()
        run.text = str(i)
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = _hex_to_rgb(PptxStyle.TEXT_LIGHT)

        # ファインディングテキスト
        txBox_f = slide.shapes.add_textbox(Inches(1.2), Inches(y_offset), Inches(10), Inches(0.4))
        tf_f = txBox_f.text_frame
        tf_f.word_wrap = True
        _add_text(tf_f, finding, size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)

        y_offset += 0.55


def _slide_company(prs: Presentation, result: ResearchResult):
    """Company分析スライド（2〜3ページ）"""
    company = result.company

    # --- ページ1: 企業概要 ---
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    # タイトル
    txBox = slide1.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
    tf = txBox.text_frame
    _add_text(tf, "Company - 企業概要", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)

    line = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
    _set_shape_bg(line, PptxStyle.ACCENT)
    line.line.fill.background()

    # 企業情報ボックス
    info_items = [
        ("企業名", company.name),
        ("公式HP", company.official_url),
        ("理念・ビジョン", company.mission_vision),
        ("事業概要", company.business_overview),
        ("主要商品・サービス", company.products_services),
    ]

    y = 1.2
    for label, value in info_items:
        if not value:
            continue
        txBox_l = slide1.shapes.add_textbox(Inches(0.5), Inches(y), Inches(2.5), Inches(0.3))
        tf_l = txBox_l.text_frame
        _add_text(tf_l, label, size=PptxStyle.SIZE_BODY, bold=True, color=PptxStyle.ACCENT)

        txBox_v = slide1.shapes.add_textbox(Inches(3.2), Inches(y), Inches(8.5), Inches(0.5))
        tf_v = txBox_v.text_frame
        tf_v.word_wrap = True
        # 長いテキストは短縮
        display_val = value[:200] + "..." if len(value) > 200 else value
        _add_text(tf_v, display_val, size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)

        y += 0.7 if len(value) <= 80 else 1.0

    # --- ページ2: 沿革 + 最新動向 ---
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    txBox = slide2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
    tf = txBox.text_frame
    _add_text(tf, "Company - 沿革・最新動向", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)

    line = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
    _set_shape_bg(line, PptxStyle.ACCENT)
    line.line.fill.background()

    # タイムライン画像
    timeline_img = _create_timeline(result)
    if timeline_img:
        img_stream = io.BytesIO(timeline_img)
        slide2.shapes.add_picture(img_stream, Inches(0.5), Inches(1.1), Inches(11), Inches(2))
        news_y = 3.3
    else:
        news_y = 1.2

    # 最新ニュース
    txBox_n = slide2.shapes.add_textbox(Inches(0.5), Inches(news_y), Inches(11), Inches(0.4))
    tf_n = txBox_n.text_frame
    _add_text(tf_n, "最新ニュース", size=PptxStyle.SIZE_HEADING, bold=True, color=PptxStyle.SECONDARY)

    # スライド下端までの残りスペースに応じてニュース件数を制限
    max_news_y = 6.8
    available = max_news_y - (news_y + 0.5)
    news_spacing = 0.85  # 各ニュース間のスペース
    max_news_count = max(1, int(available / news_spacing))
    display_news = company.recent_news[:min(4, max_news_count)]

    y = news_y + 0.5
    for news in display_news:
        if y + news_spacing > max_news_y:
            break
        txBox_item = slide2.shapes.add_textbox(Inches(0.7), Inches(y), Inches(10.5), Inches(0.7))
        tf_item = txBox_item.text_frame
        tf_item.word_wrap = True
        date_str = f"[{news.date}] " if news.date else ""
        title_text = news.title[:60] + "…" if len(news.title) > 60 else news.title
        _add_text(tf_item, f"{date_str}{title_text}", size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)
        if news.summary:
            summary_text = news.summary[:80] + "…" if len(news.summary) > 80 else news.summary
            _add_text(tf_item, f"  {summary_text}", size=PptxStyle.SIZE_SMALL, color="666666")
        y += news_spacing

    # --- ページ3: SNS・ブランド評価 ---
    if company.sns_analysis or company.brand_momentum:
        slide3 = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide3.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
        tf = txBox.text_frame
        _add_text(tf, "Company - ブランド評価・SNS分析", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)

        line = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
        _set_shape_bg(line, PptxStyle.ACCENT)
        line.line.fill.background()

        # ブランドの勢い
        if company.brand_momentum:
            txBox_m = slide3.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(11), Inches(0.4))
            tf_m = txBox_m.text_frame
            _add_text(tf_m, "ブランドの勢い・熱量", size=PptxStyle.SIZE_HEADING, bold=True, color=PptxStyle.SECONDARY)

            txBox_mv = slide3.shapes.add_textbox(Inches(0.7), Inches(1.7), Inches(10.5), Inches(0.8))
            tf_mv = txBox_mv.text_frame
            tf_mv.word_wrap = True
            _add_text(tf_mv, company.brand_momentum, size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)

        # SNS分析
        y = 2.8
        for sns in company.sns_analysis:
            txBox_s = slide3.shapes.add_textbox(Inches(0.5), Inches(y), Inches(11), Inches(0.35))
            tf_s = txBox_s.text_frame
            tone_color = {"ポジティブ": "27AE60", "ネガティブ": "E74C3C"}.get(sns.tone, PptxStyle.TEXT_DARK)
            _add_text(tf_s, f"{sns.platform}  [トーン: {sns.tone}]", size=PptxStyle.SIZE_BODY, bold=True, color=tone_color)

            txBox_sd = slide3.shapes.add_textbox(Inches(0.7), Inches(y + 0.35), Inches(10.5), Inches(0.6))
            tf_sd = txBox_sd.text_frame
            tf_sd.word_wrap = True
            _add_text(tf_sd, sns.summary, size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)
            if sns.key_topics:
                _add_text(tf_sd, f"主な話題: {', '.join(sns.key_topics)}", size=PptxStyle.SIZE_SMALL, color="666666")
            y += 1.0


def _slide_competitor(prs: Presentation, result: ResearchResult):
    """Competitor分析スライド（1〜2ページ）"""
    comp = result.competitor

    # --- ページ1: ポジショニングマップ ---
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    txBox = slide1.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
    tf = txBox.text_frame
    _add_text(tf, "Competitor - ポジショニングマップ", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)

    line = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
    _set_shape_bg(line, PptxStyle.HIGHLIGHT)
    line.line.fill.background()

    # ポジショニングマップ画像
    map_img = _create_positioning_map(result)
    img_stream = io.BytesIO(map_img)
    slide1.shapes.add_picture(img_stream, Inches(1.5), Inches(1.2), Inches(7), Inches(5))

    # 凡例
    txBox_legend = slide1.shapes.add_textbox(Inches(9), Inches(1.5), Inches(2.5), Inches(1.5))
    tf_legend = txBox_legend.text_frame
    _add_text(tf_legend, "凡例", size=PptxStyle.SIZE_SMALL, bold=True, color=PptxStyle.SECONDARY)
    _add_text(tf_legend, f"★ {result.client_name}", size=PptxStyle.SIZE_SMALL, color="E94560")
    _add_text(tf_legend, "● 直接競合", size=PptxStyle.SIZE_SMALL, color="0F3460")
    _add_text(tf_legend, "◆ 間接競合", size=PptxStyle.SIZE_SMALL, color="999999")

    # 業界ポジション
    if comp.industry_position:
        txBox_pos = slide1.shapes.add_textbox(Inches(9), Inches(3.5), Inches(2.8), Inches(2))
        tf_pos = txBox_pos.text_frame
        tf_pos.word_wrap = True
        _add_text(tf_pos, "業界ポジション", size=PptxStyle.SIZE_SMALL, bold=True, color=PptxStyle.SECONDARY)
        _add_text(tf_pos, comp.industry_position, size=PptxStyle.SIZE_SMALL, color=PptxStyle.TEXT_DARK)

    # --- ページ2: 競合比較表 ---
    all_competitors = comp.direct_competitors + comp.indirect_competitors
    if all_competitors:
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
        tf = txBox.text_frame
        _add_text(tf, "Competitor - 競合比較表", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)

        line = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
        _set_shape_bg(line, PptxStyle.HIGHLIGHT)
        line.line.fill.background()

        # テーブル
        cols = 4  # 企業名, 概要, 強み, 差別化ポイント
        rows = len(all_competitors) + 1  # ヘッダー + データ行
        table_shape = slide2.shapes.add_table(rows, cols, Inches(0.3), Inches(1.2), Inches(11.4), Inches(5))
        table = table_shape.table

        # ヘッダー
        headers = ["企業名", "概要", "強み", "差別化ポイント"]
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = _hex_to_rgb(PptxStyle.PRIMARY)
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10)
                p.font.bold = True
                p.font.color.rgb = _hex_to_rgb(PptxStyle.TEXT_LIGHT)

        # データ行
        for row_idx, c in enumerate(all_competitors, 1):
            values = [c.name, c.description[:80], c.strengths[:80], c.differentiation[:80]]
            for col_idx, val in enumerate(values):
                cell = table.cell(row_idx, col_idx)
                cell.text = val
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(9)
                    p.font.color.rgb = _hex_to_rgb(PptxStyle.TEXT_DARK)
                if row_idx % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = _hex_to_rgb(PptxStyle.BG_LIGHT)


def _slide_customer(prs: Presentation, result: ResearchResult):
    """Customer分析スライド（内容量に応じて1〜2ページ）"""
    customer = result.customer
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
    tf = txBox.text_frame
    _add_text(tf, "Customer - 市場・顧客分析", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
    _set_shape_bg(line, "27AE60")
    line.line.fill.background()

    MAX_Y = 6.5  # スライド下端の安全マージン

    # 市場規模
    y = 1.2
    if customer.market_size:
        txBox_ms = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(11), Inches(0.3))
        tf_ms = txBox_ms.text_frame
        _add_text(tf_ms, "市場規模", size=PptxStyle.SIZE_HEADING, bold=True, color=PptxStyle.SECONDARY)

        market_text = customer.market_size[:200] + "…" if len(customer.market_size) > 200 else customer.market_size
        txBox_msv = slide.shapes.add_textbox(Inches(0.7), Inches(y + 0.35), Inches(10.5), Inches(0.5))
        tf_msv = txBox_msv.text_frame
        tf_msv.word_wrap = True
        _add_text(tf_msv, market_text, size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)
        y += 1.1

    # 市場トレンド
    if customer.market_trend:
        txBox_mt = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(11), Inches(0.3))
        tf_mt = txBox_mt.text_frame
        _add_text(tf_mt, "市場トレンド", size=PptxStyle.SIZE_HEADING, bold=True, color=PptxStyle.SECONDARY)

        trend_text = customer.market_trend[:200] + "…" if len(customer.market_trend) > 200 else customer.market_trend
        txBox_mtv = slide.shapes.add_textbox(Inches(0.7), Inches(y + 0.35), Inches(10.5), Inches(0.7))
        tf_mtv = txBox_mtv.text_frame
        tf_mtv.word_wrap = True
        _add_text(tf_mtv, trend_text, size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)
        y += 1.3

    # ターゲット顧客層
    if customer.target_segments or customer.target_description:
        # 残りスペースが足りなければ次のスライドへ
        if y + 1.0 > MAX_Y:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            txBox_t2 = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
            tf_t2 = txBox_t2.text_frame
            _add_text(tf_t2, "Customer - 顧客分析（続き）", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)
            line2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
            _set_shape_bg(line2, "27AE60")
            line2.line.fill.background()
            y = 1.2

        txBox_tg = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(11), Inches(0.3))
        tf_tg = txBox_tg.text_frame
        _add_text(tf_tg, "ターゲット顧客層", size=PptxStyle.SIZE_HEADING, bold=True, color=PptxStyle.SECONDARY)

        txBox_tgv = slide.shapes.add_textbox(Inches(0.7), Inches(y + 0.35), Inches(10.5), Inches(0.7))
        tf_tgv = txBox_tgv.text_frame
        tf_tgv.word_wrap = True
        if customer.target_segments:
            segments_text = "・" + "\n・".join(s[:40] for s in customer.target_segments[:5])
            _add_text(tf_tgv, segments_text, size=PptxStyle.SIZE_BODY, color=PptxStyle.TEXT_DARK)
        if customer.target_description:
            desc_text = customer.target_description[:200] + "…" if len(customer.target_description) > 200 else customer.target_description
            _add_text(tf_tgv, desc_text, size=PptxStyle.SIZE_SMALL, color=PptxStyle.TEXT_DARK)
        y += 1.3

    # 類似事例
    if customer.similar_cases:
        # 残りスペースが足りなければ次のスライドへ
        if y + 1.0 > MAX_Y:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            txBox_t3 = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
            tf_t3 = txBox_t3.text_frame
            _add_text(tf_t3, "Customer - 類似事例", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)
            line3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
            _set_shape_bg(line3, "27AE60")
            line3.line.fill.background()
            y = 1.2

        txBox_sc = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(11), Inches(0.3))
        tf_sc = txBox_sc.text_frame
        _add_text(tf_sc, "類似事例・参考企業", size=PptxStyle.SIZE_HEADING, bold=True, color=PptxStyle.SECONDARY)
        y += 0.4

        for case in customer.similar_cases[:3]:
            if y + 0.8 > MAX_Y:
                break
            txBox_case = slide.shapes.add_textbox(Inches(0.7), Inches(y), Inches(10.5), Inches(0.7))
            tf_case = txBox_case.text_frame
            tf_case.word_wrap = True
            _add_text(tf_case, f"{case.company}（{case.industry}）", size=PptxStyle.SIZE_BODY, bold=True, color=PptxStyle.ACCENT)
            case_desc = case.description[:100] + "…" if len(case.description) > 100 else case.description
            _add_text(tf_case, case_desc, size=PptxStyle.SIZE_SMALL, color=PptxStyle.TEXT_DARK)
            y += 0.85


def _slide_appendix(prs: Presentation, result: ResearchResult):
    """付録 - 情報ソース一覧"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
    tf = txBox.text_frame
    _add_text(tf, "付録 - 情報ソース一覧", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(11), Pt(3))
    _set_shape_bg(line, "999999")
    line.line.fill.background()

    txBox_note = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(11), Inches(0.3))
    tf_note = txBox_note.text_frame
    _add_text(tf_note, f"分析実施日: {result.created_at}　|　情報ソース数: {len(result.sources)}件", size=PptxStyle.SIZE_SMALL, color="666666")

    y = 1.5
    for i, source in enumerate(result.sources[:20], 1):
        if y > 6.5:
            # 次のスライドへ
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(11), Inches(0.6))
            tf = txBox.text_frame
            _add_text(tf, "付録 - 情報ソース一覧（続き）", size=PptxStyle.SIZE_TITLE, bold=True, color=PptxStyle.PRIMARY)
            y = 1.0

        txBox_src = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(11), Inches(0.3))
        tf_src = txBox_src.text_frame
        title = source.title or source.url
        _add_text(tf_src, f"{i}. {title[:80]}", size=PptxStyle.SIZE_SMALL, color=PptxStyle.TEXT_DARK)
        _add_text(tf_src, f"   {source.url[:100]}", size=PptxStyle.SIZE_CAPTION, color="888888")
        y += 0.4


# --- メイン関数 ---

def build_pptx(result: ResearchResult, output_dir: str = "output") -> str:
    """3C分析結果からpptxファイルを生成

    Args:
        result: リサーチ結果
        output_dir: 出力ディレクトリ

    Returns:
        生成されたpptxファイルのパス
    """
    prs = Presentation()
    prs.slide_width = Emu(PptxStyle.SLIDE_WIDTH_EMU)
    prs.slide_height = Emu(PptxStyle.SLIDE_HEIGHT_EMU)

    # スライド生成
    _slide_cover(prs, result)
    _slide_executive_summary(prs, result)
    _slide_company(prs, result)
    _slide_competitor(prs, result)
    _slide_customer(prs, result)
    _slide_appendix(prs, result)

    # ファイル保存
    os.makedirs(output_dir, exist_ok=True)
    safe_name = result.client_name.replace("/", "_").replace("\\", "_")
    filename = f"3C分析_{safe_name}_{result.created_at.replace(':', '-').replace(' ', '_')}.pptx"
    filepath = os.path.join(output_dir, filename)
    prs.save(filepath)

    return filepath


def build_pptx_bytes(result: ResearchResult) -> bytes:
    """3C分析結果からpptxのバイトデータを生成（Streamlitダウンロード用）"""
    prs = Presentation()
    prs.slide_width = Emu(PptxStyle.SLIDE_WIDTH_EMU)
    prs.slide_height = Emu(PptxStyle.SLIDE_HEIGHT_EMU)

    _slide_cover(prs, result)
    _slide_executive_summary(prs, result)
    _slide_company(prs, result)
    _slide_competitor(prs, result)
    _slide_customer(prs, result)
    _slide_appendix(prs, result)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()
