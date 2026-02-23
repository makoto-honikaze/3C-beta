"""3C分析のデータモデル定義"""

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class TimelineEvent:
    """沿革・歴史の1イベント"""
    year: str
    description: str


@dataclass
class NewsItem:
    """ニュース・プレスリリース"""
    title: str
    date: str
    summary: str
    url: str = ""


@dataclass
class SNSInfo:
    """SNS分析結果"""
    platform: str              # "X", "Instagram" 等
    summary: str               # 全体的な傾向
    tone: str                  # ポジティブ/ネガティブ/ニュートラル
    key_topics: list[str] = field(default_factory=list)


@dataclass
class CompanyInfo:
    """Company（企業・ブランド）分析"""
    name: str = ""
    official_url: str = ""
    mission_vision: str = ""
    business_overview: str = ""
    products_services: str = ""
    timeline: list[TimelineEvent] = field(default_factory=list)
    recent_news: list[NewsItem] = field(default_factory=list)
    ir_summary: str = ""
    sns_analysis: list[SNSInfo] = field(default_factory=list)
    brand_momentum: str = ""   # ブランドの勢い・熱量の総合評価


@dataclass
class Competitor:
    """競合1社の情報"""
    name: str
    description: str = ""
    strengths: str = ""
    weaknesses: str = ""
    differentiation: str = ""  # 対象企業との差別化ポイント
    # ポジショニングマップ用座標（0〜10）
    position_x: float = 5.0
    position_y: float = 5.0


@dataclass
class CompetitorInfo:
    """Competitor（競合）分析"""
    direct_competitors: list[Competitor] = field(default_factory=list)
    indirect_competitors: list[Competitor] = field(default_factory=list)
    industry_position: str = ""    # 業界内でのポジション
    positioning_axis_x: str = ""   # ポジショニングマップX軸ラベル
    positioning_axis_y: str = ""   # ポジショニングマップY軸ラベル


@dataclass
class SimilarCase:
    """類似事例"""
    company: str
    industry: str
    description: str
    relevance: str = ""   # 参考になるポイント


@dataclass
class CustomerInfo:
    """Customer（顧客・市場）分析"""
    market_size: str = ""
    market_trend: str = ""
    target_segments: list[str] = field(default_factory=list)
    target_description: str = ""
    similar_cases: list[SimilarCase] = field(default_factory=list)


@dataclass
class SourceInfo:
    """情報ソース"""
    url: str
    title: str = ""
    accessed_at: str = ""


@dataclass
class PerspectiveView:
    """1つの立場からの分析"""
    needs: str = ""
    concerns: str = ""
    opportunities: str = ""  # 顧客視点の場合は "desires" として使う


@dataclass
class PerspectiveAnalysis:
    """立場別ニーズ分析（経営者/現場/顧客）"""
    executive: PerspectiveView = field(default_factory=PerspectiveView)
    frontline: PerspectiveView = field(default_factory=PerspectiveView)
    customer: PerspectiveView = field(default_factory=PerspectiveView)


@dataclass
class QuestionsAnalysis:
    """問いの自動生成結果"""
    role: str = "総合的なマーケティング担当者"
    questions: list[str] = field(default_factory=list)


@dataclass
class ResearchResult:
    """3C分析の全結果"""
    client_name: str
    industry: str
    orientation_info: str = ""
    company: CompanyInfo = field(default_factory=CompanyInfo)
    competitor: CompetitorInfo = field(default_factory=CompetitorInfo)
    customer: CustomerInfo = field(default_factory=CustomerInfo)
    perspective: PerspectiveAnalysis = field(default_factory=PerspectiveAnalysis)
    questions: QuestionsAnalysis = field(default_factory=QuestionsAnalysis)
    key_findings: list[str] = field(default_factory=list)
    sources: list[SourceInfo] = field(default_factory=list)
    created_at: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M"))
