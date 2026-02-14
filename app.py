"""3C分析リサーチ自動化ツール - Streamlit UI"""

import json
import os
from datetime import datetime

import streamlit as st

from researcher import run_full_research
from pptx_builder import build_pptx_bytes, build_pptx
from models import ResearchResult

# --- ページ設定 ---
st.set_page_config(
    page_title="3C分析リサーチツール",
    page_icon="📊",
    layout="wide",
)

# --- カスタムCSS ---
st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: bold; color: #1A1A2E; margin-bottom: 0.5rem; }
    .sub-title { font-size: 1rem; color: #666; margin-bottom: 2rem; }
    .section-header { font-size: 1.3rem; font-weight: bold; color: #0F3460; border-bottom: 2px solid #E94560; padding-bottom: 0.3rem; margin-top: 1.5rem; }
    .info-box { background-color: #f8f9fa; border-radius: 8px; padding: 1rem; margin: 0.5rem 0; border-left: 4px solid #0F3460; }
    .highlight-box { background-color: #fff3f5; border-radius: 8px; padding: 1rem; margin: 0.5rem 0; border-left: 4px solid #E94560; }
</style>
""", unsafe_allow_html=True)

# --- 定数 ---
HISTORY_DIR = "output"
os.makedirs(HISTORY_DIR, exist_ok=True)


# --- ユーティリティ ---

def save_result_json(result: ResearchResult) -> str:
    """分析結果をJSONファイルに保存"""
    safe_name = result.client_name.replace("/", "_").replace("\\", "_")
    filename = f"3C分析_{safe_name}_{result.created_at.replace(':', '-').replace(' ', '_')}.json"
    filepath = os.path.join(HISTORY_DIR, filename)

    # dataclassをdictに変換（簡易シリアライズ）
    def _to_dict(obj):
        if hasattr(obj, "__dataclass_fields__"):
            d = {}
            for field_name in obj.__dataclass_fields__:
                val = getattr(obj, field_name)
                d[field_name] = _to_dict(val)
            return d
        elif isinstance(obj, list):
            return [_to_dict(item) for item in obj]
        else:
            return obj

    data = _to_dict(result)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return filepath


def load_history() -> list[str]:
    """過去の分析結果ファイル一覧を取得"""
    if not os.path.exists(HISTORY_DIR):
        return []
    files = [f for f in os.listdir(HISTORY_DIR) if f.endswith(".json")]
    files.sort(reverse=True)
    return files


def load_result_from_json(filepath: str) -> dict:
    """JSONファイルから分析結果を読み込み"""
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


# --- サイドバー ---

with st.sidebar:
    st.markdown("### 📊 3C分析リサーチツール")
    st.markdown("---")

    # API キー状態チェック
    try:
        from config import get_api_key
        get_api_key()
        st.success("API Key: 設定済み")
    except ValueError:
        st.error("API Key が未設定です")
        st.markdown("""
        **設定方法:**
        - ローカル: `.env` ファイルに `ANTHROPIC_API_KEY=sk-ant-...` を記載
        - Streamlit Cloud: Secrets に `ANTHROPIC_API_KEY` を設定
        """)

    st.markdown("---")

    # 履歴
    st.markdown("### 📁 分析履歴")
    history_files = load_history()
    if history_files:
        selected_history = st.selectbox(
            "過去の分析結果",
            ["-- 選択してください --"] + history_files,
            key="history_select",
        )
    else:
        st.info("分析履歴はまだありません")
        selected_history = None


# --- メインコンテンツ ---

st.markdown('<div class="main-title">3C分析 リサーチ自動化ツール</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Claude AI による企業・競合・市場の自動分析とレポート生成</div>', unsafe_allow_html=True)

# --- 入力フォーム ---
tab_new, tab_history = st.tabs(["🔍 新規分析", "📁 履歴閲覧"])

with tab_new:
    col1, col2 = st.columns(2)

    with col1:
        client_name = st.text_input(
            "クライアント名 / ブランド名 *",
            placeholder="例: トヨタ自動車",
            help="分析対象の企業名またはブランド名を入力してください",
        )

    with col2:
        industry = st.text_input(
            "業種・業界 *",
            placeholder="例: 自動車業界",
            help="分析対象が属する業種・業界を入力してください",
        )

    orientation = st.text_area(
        "オリエンシート情報（任意）",
        placeholder="オリエンの要点やクライアントから共有された情報があれば入力してください。\n例: 若年層向けのブランディング強化を検討中。SNSでの認知拡大が課題。",
        height=120,
    )

    st.markdown("---")

    # 実行ボタン
    can_run = bool(client_name and industry)

    if st.button("🚀 分析を開始", type="primary", disabled=not can_run, use_container_width=True):
        st.markdown("---")

        # 進捗表示
        progress_container = st.container()

        with progress_container:
            status = st.status("3C分析を実行中...", expanded=True)

            phase_labels = {
                "company": "📋 Company分析: 企業・ブランド情報を収集中...",
                "competitor": "🏢 Competitor分析: 競合情報を分析中...",
                "customer": "👥 Customer分析: 市場・顧客情報を分析中...",
                "summary": "📝 エグゼクティブサマリーを生成中...",
                "done": "✅ 分析完了！",
            }

            current_phase_text = st.empty()

            def on_progress(phase, detail=""):
                label = phase_labels.get(phase, detail)
                status.update(label=label)
                current_phase_text.markdown(f"**{label}**")

            try:
                result = run_full_research(
                    company_name=client_name,
                    industry=industry,
                    orientation=orientation,
                    on_progress=on_progress,
                )

                status.update(label="✅ 分析完了！", state="complete")

                # 結果をセッションに保存
                st.session_state["last_result"] = result

                # JSONに保存
                json_path = save_result_json(result)

                # pptxを生成
                pptx_bytes = build_pptx_bytes(result)
                st.session_state["last_pptx"] = pptx_bytes

                st.success(f"分析が完了しました！ データを保存: {json_path}")

            except Exception as e:
                status.update(label="❌ エラーが発生しました", state="error")
                st.error(f"分析中にエラーが発生しました: {str(e)}")
                st.stop()

        # --- 結果表示 ---
        if "last_result" in st.session_state:
            result = st.session_state["last_result"]

            st.markdown("---")
            st.markdown('<div class="section-header">分析結果プレビュー</div>', unsafe_allow_html=True)

            # ダウンロードボタン
            if "last_pptx" in st.session_state:
                safe_name = result.client_name.replace("/", "_")
                st.download_button(
                    label="📥 pptxレポートをダウンロード",
                    data=st.session_state["last_pptx"],
                    file_name=f"3C分析_{safe_name}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True,
                )

            # タブで結果表示
            r_tab1, r_tab2, r_tab3, r_tab4 = st.tabs([
                "📝 サマリー", "📋 Company", "🏢 Competitor", "👥 Customer"
            ])

            with r_tab1:
                st.markdown("#### Key Findings")
                for i, finding in enumerate(result.key_findings, 1):
                    st.markdown(f"**{i}.** {finding}")

            with r_tab2:
                company = result.company
                st.markdown(f"**企業名:** {company.name}")
                st.markdown(f"**公式HP:** {company.official_url}")

                if company.mission_vision:
                    st.markdown("**理念・ビジョン:**")
                    st.info(company.mission_vision)

                if company.business_overview:
                    st.markdown("**事業概要:**")
                    st.markdown(company.business_overview)

                if company.timeline:
                    st.markdown("**沿革:**")
                    for event in company.timeline:
                        st.markdown(f"- **{event.year}**: {event.description}")

                if company.recent_news:
                    st.markdown("**最新ニュース:**")
                    for news in company.recent_news:
                        date_str = f"[{news.date}] " if news.date else ""
                        st.markdown(f"- {date_str}**{news.title}** - {news.summary}")

                if company.brand_momentum:
                    st.markdown("**ブランドの勢い:**")
                    st.success(company.brand_momentum)

            with r_tab3:
                comp = result.competitor

                if comp.industry_position:
                    st.markdown("**業界ポジション:**")
                    st.info(comp.industry_position)

                if comp.direct_competitors:
                    st.markdown("**直接競合:**")
                    for c in comp.direct_competitors:
                        with st.expander(f"🏢 {c.name}"):
                            st.markdown(f"**概要:** {c.description}")
                            st.markdown(f"**強み:** {c.strengths}")
                            if c.weaknesses:
                                st.markdown(f"**弱み:** {c.weaknesses}")
                            st.markdown(f"**差別化:** {c.differentiation}")

                if comp.indirect_competitors:
                    st.markdown("**間接競合:**")
                    for c in comp.indirect_competitors:
                        with st.expander(f"🔄 {c.name}"):
                            st.markdown(f"**概要:** {c.description}")
                            st.markdown(f"**強み:** {c.strengths}")

            with r_tab4:
                customer = result.customer

                if customer.market_size:
                    st.markdown("**市場規模:**")
                    st.info(customer.market_size)

                if customer.market_trend:
                    st.markdown("**市場トレンド:**")
                    st.markdown(customer.market_trend)

                if customer.target_segments:
                    st.markdown("**ターゲットセグメント:**")
                    for seg in customer.target_segments:
                        st.markdown(f"- {seg}")

                if customer.similar_cases:
                    st.markdown("**類似事例:**")
                    for case in customer.similar_cases:
                        with st.expander(f"📌 {case.company}（{case.industry}）"):
                            st.markdown(case.description)
                            if case.relevance:
                                st.markdown(f"**参考ポイント:** {case.relevance}")

    elif not can_run:
        st.info("クライアント名と業種を入力して「分析を開始」ボタンを押してください。")


with tab_history:
    if selected_history and selected_history != "-- 選択してください --":
        filepath = os.path.join(HISTORY_DIR, selected_history)
        try:
            data = load_result_from_json(filepath)
            st.markdown(f"**クライアント:** {data.get('client_name', 'N/A')}")
            st.markdown(f"**業界:** {data.get('industry', 'N/A')}")
            st.markdown(f"**分析日:** {data.get('created_at', 'N/A')}")

            st.markdown("---")
            st.json(data, expanded=False)

        except Exception as e:
            st.error(f"ファイルの読み込みに失敗しました: {e}")
    else:
        st.info("サイドバーから過去の分析結果を選択してください。")
