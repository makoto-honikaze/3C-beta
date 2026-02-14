"""設定管理 - Streamlit Cloud / ローカル開発の両対応"""

import os

# --- APIキー管理 ---

def get_api_key() -> str:
    """ANTHROPIC_API_KEYを取得（Streamlit Cloud → .env の優先順）"""
    # 1. Streamlit Cloud の Secrets
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "ANTHROPIC_API_KEY" in st.secrets:
            return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass

    # 2. 環境変数 / .env ファイル
    from dotenv import load_dotenv
    load_dotenv()
    key = os.getenv("ANTHROPIC_API_KEY", "")
    if key:
        return key

    raise ValueError(
        "ANTHROPIC_API_KEY が設定されていません。\n"
        "ローカル: .env ファイルに設定してください。\n"
        "Streamlit Cloud: Secrets に設定してください。"
    )


# --- pptxスタイル定数 ---

class PptxStyle:
    # カラーパレット
    PRIMARY = "1A1A2E"       # ダークネイビー（タイトル）
    SECONDARY = "16213E"     # ネイビー（サブヘッダー）
    ACCENT = "0F3460"        # ブルー（アクセント）
    HIGHLIGHT = "E94560"     # レッド（ハイライト）
    BG_LIGHT = "F5F5F5"      # ライトグレー（背景）
    TEXT_DARK = "333333"      # テキスト（本文）
    TEXT_LIGHT = "FFFFFF"     # テキスト（白）

    # フォント
    FONT_TITLE = "Noto Sans JP"
    FONT_BODY = "Noto Sans JP"
    FONT_FALLBACK = "Arial"

    # フォントサイズ（Pt）
    SIZE_TITLE = 28
    SIZE_SUBTITLE = 18
    SIZE_HEADING = 20
    SIZE_BODY = 12
    SIZE_SMALL = 10
    SIZE_CAPTION = 8

    # スライドサイズ（16:9）
    SLIDE_WIDTH_EMU = 12192000
    SLIDE_HEIGHT_EMU = 6858000


# --- Claude API設定 ---

CLAUDE_MODEL = "claude-haiku-4-5-20251001"
CLAUDE_MAX_TOKENS = 8192
