"""Claude API web_search を使った3C分析リサーチモジュール"""

import json
import time
import traceback
import anthropic
from config import get_api_key, CLAUDE_MODEL, CLAUDE_MAX_TOKENS
from models import (
    ResearchResult, CompanyInfo, CompetitorInfo, CustomerInfo,
    TimelineEvent, NewsItem, SNSInfo, Competitor, SimilarCase, SourceInfo,
)


# --- リトライ処理 ---

def _call_api_with_retry(func, max_retries=2):
    """429エラー（レートリミット）のリトライ処理

    Args:
        func: 実行するAPI呼び出し関数（引数なしのcallable）
        max_retries: 最大リトライ回数（デフォルト2回、合計3回まで）

    Returns:
        API呼び出しの戻り値

    Raises:
        Exception: 3回失敗した場合、ユーザー向けエラーメッセージで即座に中止
    """
    for attempt in range(max_retries + 1):
        try:
            return func()
        except anthropic.RateLimitError:
            if attempt < max_retries:
                wait_sec = 30
                print(f"[レートリミット] {wait_sec}秒待機してリトライします... (試行 {attempt + 1}/{max_retries + 1})")
                time.sleep(wait_sec)
            else:
                raise Exception(
                    "レートリミットに達しました。1〜2分待ってから再実行してください。"
                )


def _create_client() -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=get_api_key())


# --- 安全なアクセスヘルパー ---

def _safe_str(val) -> str:
    """どんな値でも安全にstrに変換"""
    if val is None:
        return ""
    if isinstance(val, str):
        return val
    return str(val)


def _safe_get(obj, key, default=None):
    """オブジェクトまたはdictから安全に値を取得"""
    if obj is None:
        return default
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _safe_get_str(obj, key, default="") -> str:
    """オブジェクトまたはdictから安全に文字列値を取得"""
    val = _safe_get(obj, key, default)
    return _safe_str(val)


def _safe_get_type(block) -> str:
    """ブロックのtypeを安全に取得"""
    return _safe_get_str(block, "type", "")


def _safe_to_dict(obj) -> dict:
    """オブジェクトをdictに安全に変換（dataclass展開用）"""
    if isinstance(obj, dict):
        return obj
    if hasattr(obj, "__dict__"):
        return {k: v for k, v in obj.__dict__.items() if not k.startswith("_")}
    return {}


# --- レスポンス解析 ---

def _extract_text_and_sources(response) -> tuple[str, list[dict]]:
    """APIレスポンスからテキストとソース一覧を抽出（dict/object両対応）"""
    text_parts = []
    sources = []

    # response.content がlistかどうかチェック
    content_list = _safe_get(response, "content", [])
    if not isinstance(content_list, list):
        print(f"[DEBUG] response.content is not list: {type(content_list)}")
        return "", []

    for block in content_list:
        block_type = _safe_get_type(block)
        print(f"[DEBUG] block_type={block_type}, python_type={type(block).__name__}")

        if block_type == "text":
            text_val = _safe_get(block, "text", "")
            text_str = _safe_str(text_val)
            if text_str:
                text_parts.append(text_str)

            # citationsからソースを抽出
            citations = _safe_get(block, "citations", None)
            if citations and isinstance(citations, list):
                for cite in citations:
                    url = _safe_get_str(cite, "url", "")
                    if url:
                        sources.append({
                            "url": url,
                            "title": _safe_get_str(cite, "title", ""),
                        })

        elif block_type == "web_search_tool_result":
            block_content = _safe_get(block, "content", None)

            # contentがlistの場合（正常な検索結果）
            if isinstance(block_content, list):
                for result_item in block_content:
                    result_type = _safe_get_type(result_item)
                    # エラーブロックはスキップ
                    if result_type == "web_search_tool_result_error":
                        print(f"[DEBUG] web_search error: {result_item}")
                        continue
                    url = _safe_get_str(result_item, "url", "")
                    if url:
                        sources.append({
                            "url": url,
                            "title": _safe_get_str(result_item, "title", ""),
                        })
            # contentがdictの場合（エラーレスポンス等）
            elif isinstance(block_content, dict):
                print(f"[DEBUG] web_search_tool_result content is dict: {block_content}")
            # contentがオブジェクトの場合
            elif block_content is not None:
                print(f"[DEBUG] web_search_tool_result content is {type(block_content).__name__}")

        elif block_type == "server_tool_use":
            # 検索クエリのブロック（処理不要）
            pass

        else:
            print(f"[DEBUG] unknown block_type={block_type}, type={type(block).__name__}")

    # ソースの重複排除
    seen_urls = set()
    unique_sources = []
    for s in sources:
        if s["url"] not in seen_urls:
            seen_urls.add(s["url"])
            unique_sources.append(s)

    combined_text = "\n".join(text_parts)
    return combined_text, unique_sources


def _search_and_analyze(client: anthropic.Anthropic, prompt: str) -> tuple[str, list[dict]]:
    """Claude web_search で検索・分析し、テキスト結果とソース一覧を返す"""

    def _api_call():
        return client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=CLAUDE_MAX_TOKENS,
            tools=[{
                "type": "web_search_20250305",
                "name": "web_search",
                "max_uses": 3,
            }],
            messages=[{"role": "user", "content": prompt}],
        )

    response = _call_api_with_retry(_api_call)

    print(f"[DEBUG] response type: {type(response).__name__}")
    print(f"[DEBUG] response.content type: {type(_safe_get(response, 'content', [])).__name__}")

    return _extract_text_and_sources(response)


def _parse_json_from_text(text) -> dict:
    """テキストからJSONブロックを抽出してパース（入力がstrでなくても安全）"""
    # 入力がstr以外の場合の防御
    if not isinstance(text, str):
        print(f"[DEBUG] _parse_json_from_text received non-str: {type(text).__name__}")
        if isinstance(text, dict):
            return text  # すでにdictならそのまま返す
        text = _safe_str(text)
        if not text:
            return {}

    # ```json ... ``` ブロックを探す
    if "```json" in text:
        try:
            start = text.index("```json") + 7
            end = text.index("```", start)
            json_str = text[start:end].strip()
        except ValueError:
            json_str = ""
    elif "```" in text:
        try:
            start = text.index("```") + 3
            end = text.index("```", start)
            json_str = text[start:end].strip()
        except ValueError:
            json_str = ""
    else:
        # { から } までを抽出
        brace_start = text.find("{")
        brace_end = text.rfind("}")
        if brace_start >= 0 and brace_end > brace_start:
            json_str = text[brace_start:brace_end + 1]
        else:
            return {}

    if not json_str:
        return {}

    try:
        return json.loads(json_str)
    except (json.JSONDecodeError, TypeError) as e:
        print(f"[DEBUG] JSON parse error: {e}")
        return {}


# --- 安全なデータモデル変換 ---

def _make_timeline_event(e) -> TimelineEvent:
    """dictまたはオブジェクトからTimelineEventを安全に生成"""
    d = _safe_to_dict(e) if not isinstance(e, dict) else e
    return TimelineEvent(
        year=_safe_str(d.get("year", "")),
        description=_safe_str(d.get("description", "")),
    )


def _make_news_item(n) -> NewsItem:
    """dictまたはオブジェクトからNewsItemを安全に生成"""
    d = _safe_to_dict(n) if not isinstance(n, dict) else n
    return NewsItem(
        title=_safe_str(d.get("title", "")),
        date=_safe_str(d.get("date", "")),
        summary=_safe_str(d.get("summary", "")),
        url=_safe_str(d.get("url", "")),
    )


def _make_sns_info(s) -> SNSInfo:
    """dictまたはオブジェクトからSNSInfoを安全に生成"""
    d = _safe_to_dict(s) if not isinstance(s, dict) else s
    key_topics = d.get("key_topics", [])
    if not isinstance(key_topics, list):
        key_topics = [_safe_str(key_topics)]
    return SNSInfo(
        platform=_safe_str(d.get("platform", "")),
        summary=_safe_str(d.get("summary", "")),
        tone=_safe_str(d.get("tone", "")),
        key_topics=[_safe_str(t) for t in key_topics],
    )


def _make_competitor(c) -> Competitor:
    """dictまたはオブジェクトからCompetitorを安全に生成"""
    d = _safe_to_dict(c) if not isinstance(c, dict) else c
    try:
        pos_x = float(d.get("position_x", 5))
    except (ValueError, TypeError):
        pos_x = 5.0
    try:
        pos_y = float(d.get("position_y", 5))
    except (ValueError, TypeError):
        pos_y = 5.0
    return Competitor(
        name=_safe_str(d.get("name", "")),
        description=_safe_str(d.get("description", "")),
        strengths=_safe_str(d.get("strengths", "")),
        weaknesses=_safe_str(d.get("weaknesses", "")),
        differentiation=_safe_str(d.get("differentiation", "")),
        position_x=pos_x,
        position_y=pos_y,
    )


def _make_similar_case(c) -> SimilarCase:
    """dictまたはオブジェクトからSimilarCaseを安全に生成"""
    d = _safe_to_dict(c) if not isinstance(c, dict) else c
    return SimilarCase(
        company=_safe_str(d.get("company", "")),
        industry=_safe_str(d.get("industry", "")),
        description=_safe_str(d.get("description", "")),
        relevance=_safe_str(d.get("relevance", "")),
    )


# --- 各分析フェーズ ---

def research_company(client: anthropic.Anthropic, company_name: str, industry: str, orientation: str = "") -> tuple[CompanyInfo, list[dict]]:
    """Company分析を実行"""
    context = f"\nオリエン情報: {orientation}" if orientation else ""

    prompt = f"""企業リサーチの専門家として、以下の企業をWeb検索で分析し、JSON形式で出力してください。

企業名: {company_name}　業界: {industry}{context}

```json
{{
  "name": "企業名", "official_url": "公式HP URL",
  "mission_vision": "理念・ビジョン",
  "business_overview": "事業概要（200字）",
  "products_services": "主要商品・サービス（200字）",
  "timeline": [{{"year": "年", "description": "出来事"}}],
  "recent_news": [{{"title": "タイトル", "date": "YYYY-MM", "summary": "概要", "url": "URL"}}],
  "ir_summary": "IR・業績要約（200字）",
  "sns_analysis": [{{"platform": "X", "summary": "傾向", "tone": "ポジティブ/ネガティブ/ニュートラル", "key_topics": ["話題"]}}],
  "brand_momentum": "ブランドの勢い評価（100字）"
}}
```
沿革は創業〜現在の主要節目、ニュースは直近1年分、SNSは「{company_name} 評判」で検索してください。"""

    text, sources = _search_and_analyze(client, prompt)
    data = _parse_json_from_text(text)

    if not data:
        return CompanyInfo(name=company_name), sources

    # timeline, recent_news, sns_analysis がlistであることを保証
    timeline_raw = data.get("timeline", [])
    if not isinstance(timeline_raw, list):
        timeline_raw = []
    news_raw = data.get("recent_news", [])
    if not isinstance(news_raw, list):
        news_raw = []
    sns_raw = data.get("sns_analysis", [])
    if not isinstance(sns_raw, list):
        sns_raw = []

    info = CompanyInfo(
        name=_safe_str(data.get("name", company_name)),
        official_url=_safe_str(data.get("official_url", "")),
        mission_vision=_safe_str(data.get("mission_vision", "")),
        business_overview=_safe_str(data.get("business_overview", "")),
        products_services=_safe_str(data.get("products_services", "")),
        timeline=[_make_timeline_event(e) for e in timeline_raw],
        recent_news=[_make_news_item(n) for n in news_raw],
        ir_summary=_safe_str(data.get("ir_summary", "")),
        sns_analysis=[_make_sns_info(s) for s in sns_raw],
        brand_momentum=_safe_str(data.get("brand_momentum", "")),
    )
    return info, sources


def research_competitor(client: anthropic.Anthropic, company_name: str, industry: str) -> tuple[CompetitorInfo, list[dict]]:
    """Competitor分析を実行"""
    prompt = f"""競合分析の専門家として、以下の企業の競合をWeb検索で分析し、JSON形式で出力してください。

企業名: {company_name}　業界: {industry}

```json
{{
  "direct_competitors": [
    {{"name": "競合名", "description": "概要", "strengths": "強み", "weaknesses": "弱み",
      "differentiation": "{company_name}との違い", "position_x": 7.0, "position_y": 8.0}}
  ],
  "indirect_competitors": [
    {{"name": "間接競合名", "description": "説明", "strengths": "強み",
      "differentiation": "{company_name}との違い", "position_x": 3.0, "position_y": 5.0}}
  ],
  "industry_position": "{company_name}の業界ポジション（100字）",
  "positioning_axis_x": "X軸ラベル（例: 価格帯 低←→高）",
  "positioning_axis_y": "Y軸ラベル（例: 品質 低←→高）",
  "target_company_position_x": 6.0,
  "target_company_position_y": 7.0
}}
```
直接競合3〜5社、間接競合1〜3社。座標は0〜10の範囲で設定してください。"""

    text, sources = _search_and_analyze(client, prompt)
    data = _parse_json_from_text(text)

    if not data:
        return CompetitorInfo(), sources

    direct_raw = data.get("direct_competitors", [])
    if not isinstance(direct_raw, list):
        direct_raw = []
    indirect_raw = data.get("indirect_competitors", [])
    if not isinstance(indirect_raw, list):
        indirect_raw = []

    info = CompetitorInfo(
        direct_competitors=[_make_competitor(c) for c in direct_raw],
        indirect_competitors=[_make_competitor(c) for c in indirect_raw],
        industry_position=_safe_str(data.get("industry_position", "")),
        positioning_axis_x=_safe_str(data.get("positioning_axis_x", "")),
        positioning_axis_y=_safe_str(data.get("positioning_axis_y", "")),
    )

    try:
        tx = float(data.get("target_company_position_x", 5))
    except (ValueError, TypeError):
        tx = 5.0
    try:
        ty = float(data.get("target_company_position_y", 5))
    except (ValueError, TypeError):
        ty = 5.0
    info._target_position = (tx, ty)

    return info, sources


def research_customer(client: anthropic.Anthropic, company_name: str, industry: str) -> tuple[CustomerInfo, list[dict]]:
    """Customer/市場分析を実行"""
    prompt = f"""市場分析の専門家として、以下の企業が属する市場をWeb検索で分析し、JSON形式で出力してください。

企業名: {company_name}　業界: {industry}

```json
{{
  "market_size": "市場規模（金額・成長率）",
  "market_trend": "市場トレンド（200字）",
  "target_segments": ["セグメント1", "セグメント2"],
  "target_description": "ターゲット顧客層の説明（200字）",
  "similar_cases": [
    {{"company": "企業名", "industry": "業種", "description": "事例説明",
      "relevance": "{company_name}への参考ポイント"}}
  ]
}}
```
最新の市場データを検索し、類似事例は他業種から2〜3件含めてください。"""

    text, sources = _search_and_analyze(client, prompt)
    data = _parse_json_from_text(text)

    if not data:
        return CustomerInfo(), sources

    segments_raw = data.get("target_segments", [])
    if not isinstance(segments_raw, list):
        segments_raw = [_safe_str(segments_raw)]
    cases_raw = data.get("similar_cases", [])
    if not isinstance(cases_raw, list):
        cases_raw = []

    info = CustomerInfo(
        market_size=_safe_str(data.get("market_size", "")),
        market_trend=_safe_str(data.get("market_trend", "")),
        target_segments=[_safe_str(s) for s in segments_raw],
        target_description=_safe_str(data.get("target_description", "")),
        similar_cases=[_make_similar_case(c) for c in cases_raw],
    )
    return info, sources


def generate_key_findings(client: anthropic.Anthropic, result: ResearchResult) -> list[str]:
    """収集した情報からキーファインディングを生成"""
    summary = f"""以下の3C分析結果から、キーファインディングを3〜5つJSON形式で出力してください。

企業: {result.client_name}（{result.industry}）
事業: {result.company.business_overview}
勢い: {result.company.brand_momentum}
業界位置: {result.competitor.industry_position}
競合: {', '.join(_safe_str(c.name) for c in result.competitor.direct_competitors)}
市場: {result.customer.market_size}
トレンド: {result.customer.market_trend}

```json
{{"key_findings": ["ファインディング1", "ファインディング2", "ファインディング3"]}}
```"""

    def _api_call():
        return client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=2048,
            messages=[{"role": "user", "content": summary}],
        )

    response = _call_api_with_retry(_api_call)

    # レスポンスからテキストを安全に抽出
    text = ""
    content_list = _safe_get(response, "content", [])
    if isinstance(content_list, list):
        for block in content_list:
            block_type = _safe_get_type(block)
            if block_type == "text":
                text_val = _safe_get(block, "text", "")
                text += _safe_str(text_val)
    else:
        print(f"[DEBUG] generate_key_findings: content is {type(content_list).__name__}")

    data = _parse_json_from_text(text)

    findings = data.get("key_findings", ["分析データの詳細は各セクションをご確認ください。"])
    if not isinstance(findings, list):
        findings = [_safe_str(findings)]
    return [_safe_str(f) for f in findings]


def run_full_research(
    company_name: str,
    industry: str,
    orientation: str = "",
    on_progress=None,
) -> ResearchResult:
    """3C分析のフルリサーチを実行

    Args:
        company_name: クライアント名/ブランド名
        industry: 業種・業界
        orientation: オリエンシート情報（任意）
        on_progress: 進捗コールバック (phase: str, detail: str) -> None
    """
    client = _create_client()
    all_sources = []

    def _progress(phase, detail=""):
        if on_progress:
            on_progress(phase, detail)

    # 1. Company分析
    _progress("company", "企業・ブランド情報を収集中...")
    company_info, company_sources = research_company(client, company_name, industry, orientation)
    all_sources.extend(company_sources)

    # Company → Competitor の間にスリープ（レートリミット対策）
    time.sleep(15)

    # 2. Competitor分析
    _progress("competitor", "競合情報を分析中...")
    competitor_info, competitor_sources = research_competitor(client, company_name, industry)
    all_sources.extend(competitor_sources)

    # Competitor → Customer の間にスリープ（レートリミット対策）
    time.sleep(15)

    # 3. Customer分析
    _progress("customer", "市場・顧客情報を分析中...")
    customer_info, customer_sources = research_customer(client, company_name, industry)
    all_sources.extend(customer_sources)

    # 結果を組み立て
    result = ResearchResult(
        client_name=company_name,
        industry=industry,
        orientation_info=orientation,
        company=company_info,
        competitor=competitor_info,
        customer=customer_info,
        sources=[SourceInfo(url=s["url"], title=s.get("title", "")) for s in all_sources],
    )

    # 4. キーファインディング生成
    _progress("summary", "エグゼクティブサマリーを生成中...")
    result.key_findings = generate_key_findings(client, result)

    # 対象企業のポジション座標を保持
    if hasattr(competitor_info, "_target_position"):
        result.competitor._target_position = competitor_info._target_position

    _progress("done", "分析完了")
    return result
