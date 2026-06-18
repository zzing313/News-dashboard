import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import os
import random

# ==========================================
# PAGE CONFIG & CSS
# ==========================================
st.set_page_config(page_title="AI 홍보 뉴스 모니터링 대시보드", layout="wide", initial_sidebar_state="expanded")

# Custom CSS
st.markdown("""
<style>
    /* Global font and background */
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    * {
        font-family: 'Pretendard', sans-serif !important;
    }
    
    .stApp {
        background-color: #f8f9fa;
    }
    
    /* ── 사이드바 고정 스타일 ── */
    [data-testid="stSidebar"] {
        background-color: #1a2235;
        color: white;
        min-width: 220px !important;
        max-width: 220px !important;
    }
    [data-testid="stSidebar"] * {
        color: white !important;
    }

    /* 사이드바 접기 버튼 완전 숨김 (고정 사이드바라 불필요) */
    [data-testid="stSidebarCollapseButton"] {
        display: none !important;
    }
    [data-testid="collapsedControl"] {
        display: none !important;
    }

    /* Top metric cards */
    div[data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    /* Plot container cards */
    .plot-container {
        background-color: white;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
    }
    
    /* Header titles */
    h1, h2, h3 {
        color: #1e293b !important;
        font-weight: 700 !important;
    }
    
    /* Custom KPI styling */
    .kpi-card {
        background: white;
        border-radius: 12px;
        padding: 24px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        display: flex;
        flex-direction: column;
        justify-content: center;
        height: 100%;
    }
    .kpi-title {
        font-size: 14px;
        color: #64748b;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-size: 32px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 8px;
    }
    .kpi-sub {
        font-size: 13px;
        color: #10b981;
    }
    .kpi-sub.down {
        color: #ef4444;
    }
    .kpi-sub.neutral {
        color: #64748b;
    }
    
    /* News Cards */
    .news-card {
        background: white;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        padding: 12px;
        display: flex;
        gap: 12px;
        margin-bottom: 12px;
        cursor: pointer;
        transition: transform 0.2s;
    }
    .news-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .news-card img {
        width: 100px;
        height: 70px;
        object-fit: cover;
        border-radius: 4px;
    }
    .news-title {
        font-size: 14px;
        font-weight: 600;
        color: #1e293b;
        margin: 0 0 8px 0;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        overflow: hidden;
    }
    .news-meta {
        font-size: 12px;
        color: #64748b;
    }

    /* Article list table styling */
    .article-row {
        background: white;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        padding: 14px 16px;
        margin-bottom: 8px;
        display: flex;
        align-items: flex-start;
        gap: 14px;
        transition: box-shadow 0.2s;
    }
    .article-row:hover {
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }
    .badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 700;
        white-space: nowrap;
    }
    .badge-pos { background: #d1fae5; color: #065f46; }
    .badge-neu { background: #e2e8f0; color: #475569; }
    .badge-neg { background: #fee2e2; color: #991b1b; }

    /* Sidebar nav active item */
    .nav-active {
        background: #4f46e5;
        padding: 10px;
        border-radius: 6px;
        font-weight: bold;
        cursor: pointer;
    }
    .nav-item {
        padding: 10px;
        opacity: 0.8;
        cursor: pointer;
        border-radius: 6px;
        transition: background 0.15s;
    }
    .nav-item:hover {
        background: rgba(255,255,255,0.08);
    }

    /* Pagination button area */
    .page-info {
        font-size: 14px;
        color: #475569;
        text-align: center;
        padding: 8px 0;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# DUMMY DATA GENERATOR
# ==========================================
def generate_dummy_data(filename):
    keywords = ["신세계면세점", "롯데면세점", "신라면세점", "현대면세점", "인천공항", "한국관광공사", "명품", "패션", "K뷰티"]
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for kw in keywords:
            num_rows = random.randint(10, 50)
            if kw == "신세계면세점": num_rows = 12
            elif kw == "신라면세점": num_rows = 18
            elif kw == "롯데면세점": num_rows = 15
            elif kw == "현대면세점": num_rows = 4
            
            data = []
            for i in range(num_rows):
                sentiment = random.choices(['긍정', '중립', '부정'], weights=[0.7, 0.25, 0.05])[0]
                titles = [
                    f"{kw}, 외국인 관광객 맞춤 서비스 강화",
                    f"{kw} 매출 증가… 중국 FIT 수요 회복세",
                    f"면세점 업계 지각변동, {kw} 전략은?",
                    f"{kw}, K-뷰티·명품 라인업 대폭 확대",
                    f"인천공항 면세구역 경쟁 치열… {kw} 선전",
                ]
                data.append({
                    "헤드라인": random.choice(titles),
                    "매체": random.choice(["매일경제", "한국경제", "뉴스1", "파이낸셜뉴스", "아시아경제", "조선비즈", "머니투데이"]),
                    "기자명": random.choice(["김OO", "이OO", "박OO", "최OO", "정OO", "강OO", "윤OO"]),
                    "내용요약": "최근 중국관광객 및 FIT 수요가 증가함에 따라 면세점 업계가 활기를 띠고 있다...",
                    "기사 링크": "https://n.news.naver.com",
                    "기사 사진 링크": "https://via.placeholder.com/150",
                    "감성": sentiment,
                    "본문": f"중국관광객 FIT 면세점 항공 명품 K-뷰티 관광 호텔 {kw}",
                    "작성일": (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
                })
            pd.DataFrame(data).to_excel(writer, sheet_name=kw, index=False)

# ==========================================
# DATA LOADING
# ==========================================
@st.cache_data
def load_data(file):
    excel_file = pd.ExcelFile(file)
    sheets_dict = pd.read_excel(excel_file, sheet_name=None)
    all_data = []
    for sheet_name, df in sheets_dict.items():
        df['Search_Keyword'] = sheet_name
        all_data.append(df)
    combined_df = pd.concat(all_data, ignore_index=True)
    return sheets_dict, combined_df

def get_file_for_date(date_obj):
    """주어진 날짜의 엑셀 파일을 반환. 없으면 더미 생성."""
    date_str = date_obj.strftime("%Y-%m-%d")
    file_name = f"{date_str}_뉴스모니터링.xlsx"
    if not os.path.exists(file_name):
        # 오늘/어제 파일이 없으면 더미 생성
        generate_dummy_data(file_name)
    return file_name, date_str

# ==========================================
# SESSION STATE (페이지 관리)
# ==========================================
if "current_page" not in st.session_state:
    st.session_state.current_page = "대시보드"

# ==========================================
# SIDEBAR (고정)
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='color: #eab308 !important; text-align: center; margin-bottom: 0;'>SHINSEGAE</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; letter-spacing: 2px; font-size: 12px; margin-top: 0;'>DUTY FREE</p>", unsafe_allow_html=True)
    st.markdown("<hr style='border-color: #334155;'>", unsafe_allow_html=True)

    menus = {
        '🏠 대시보드': '대시보드',
        '📄 기사 목록': '기사 목록',
        '🔍 키워드 분석': '키워드 분석',
        '👤 기자 분석': '기자 분석',
        '📊 경쟁사 분석': '경쟁사 분석',
        '📑 리포트': '리포트',
        '⚙️ 설정': '설정',
    }
    for label, page_key in menus.items():
        is_active = st.session_state.current_page == page_key
        css_class = "nav-active" if is_active else "nav-item"
        if st.sidebar.button(label, key=f"nav_{page_key}", use_container_width=True):
            st.session_state.current_page = page_key
            st.rerun()

    st.markdown("<br>" * 5, unsafe_allow_html=True)
    st.markdown(
        f"<div style='font-size:12px; color:#94a3b8;'>데이터 기준<br>{datetime.now().strftime('%Y.%m.%d')} "
        f"({['월','화','수','목','금','토','일'][datetime.now().weekday()]})<br>전일 수집 데이터 기준</div>",
        unsafe_allow_html=True
    )

# ==========================================
# SIDEBAR 버튼 스타일 오버라이드 (색상 재정의)
# ==========================================
st.markdown("""
<style>
    [data-testid="stSidebar"] .stButton > button {
        background: transparent !important;
        border: none !important;
        text-align: left !important;
        padding: 10px 12px !important;
        border-radius: 6px !important;
        color: white !important;
        font-size: 14px !important;
        opacity: 0.85;
        transition: background 0.15s;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background: rgba(255,255,255,0.1) !important;
        opacity: 1;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# PAGE: 대시보드
# ==========================================
if st.session_state.current_page == "대시보드":
    yesterday = datetime.now() - timedelta(days=1)
    file_name, target_date_str = get_file_for_date(yesterday)
    sheets_dict, df_all = load_data(file_name)

    st.markdown("## AI 홍보 뉴스 모니터링 대시보드")
    st.markdown(f"<p style='color: #64748b;'>전일 수집 데이터 기준 ({target_date_str})</p>", unsafe_allow_html=True)

    # 1. KPIs
    col1, col2, col3, col4 = st.columns(4)
    total_articles = len(df_all)
    shinsegae_articles = len(df_all[df_all['Search_Keyword'] == '신세계면세점'])

    with col1:
        st.markdown(f"""
        <div class='kpi-card'>
            <div class='kpi-title'>전일 전체 기사 (키워드 전체 합계)</div>
            <div class='kpi-value'>{total_articles} 건</div>
            <div class='kpi-sub down'>↓ 12.4% 전일 대비</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class='kpi-card'>
            <div class='kpi-title'>전일 신세계면세점 기사</div>
            <div class='kpi-value'>{shinsegae_articles} 건</div>
            <div class='kpi-sub'>↑ 9.1% 전일 대비</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div class='kpi-card'>
            <div class='kpi-title'>기사 수 가장 많은 키워드</div>
            <div class='kpi-value' style='color: #10b981;'>FIT</div>
            <div class='kpi-sub neutral'>기사 수 42건 (22.6%)</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown("""
        <div class='kpi-card'>
            <div class='kpi-title'>기사 내 가장 언급 많이 된 키워드</div>
            <div class='kpi-value' style='color: #8b5cf6;'>중국관광객</div>
            <div class='kpi-sub neutral'>언급 횟수 156회</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # 2. Charts (Middle Row)
    m_col1, m_col2, m_col3 = st.columns([1.5, 1.5, 1.2])

    with m_col1:
        st.markdown("<div class='plot-container'>", unsafe_allow_html=True)
        st.markdown("<h4>키워드별 기사 수 TOP 10</h4>", unsafe_allow_html=True)
        bar_data = pd.DataFrame({
            '키워드': ['FIT', '중국관광', '항공', 'K-뷰티', '호텔', '면세점', '명품', '여행수요', 'MZ세대', '환율'][::-1],
            '기사 수': [42, 31, 24, 19, 15, 12, 10, 9, 8, 6][::-1]
        })
        fig_bar = px.bar(bar_data, x='기사 수', y='키워드', orientation='h', color_discrete_sequence=['#3b82f6'])
        fig_bar.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=300, plot_bgcolor='white', paper_bgcolor='white')
        fig_bar.update_xaxes(showgrid=True, gridcolor='#f1f5f9')
        st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
        st.markdown("</div>", unsafe_allow_html=True)

    with m_col2:
        st.markdown("<div class='plot-container'>", unsafe_allow_html=True)
        st.markdown("<h4>키워드 언급 빈도 (워드클라우드)</h4>", unsafe_allow_html=True)
        words = {'중국관광객': 156, '면세점': 120, '항공': 90, '명품': 85, 'FIT': 70, 'K-뷰티': 65, '관광': 60, '호텔': 40, '여행': 35}
        wc = WordCloud(font_path='MALGUN.TTF', background_color='white', width=400, height=300, colormap='tab10').generate_from_frequencies(words)
        fig_wc, ax_wc = plt.subplots(figsize=(5, 4))
        ax_wc.imshow(wc, interpolation='bilinear')
        ax_wc.axis('off')
        st.pyplot(fig_wc)
        st.markdown("</div>", unsafe_allow_html=True)

    with m_col3:
        st.markdown("<div class='plot-container'>", unsafe_allow_html=True)
        st.markdown("<h4>신세계면세점 기사 점유율 (SOV)</h4>", unsafe_allow_html=True)
        sov_data = pd.DataFrame({
            '브랜드': ['신세계면세점', '신라면세점', '롯데면세점', '현대면세점', '기타'],
            '기사 수': [12, 18, 15, 4, 137]
        })
        colors = ['#ef4444', '#3b82f6', '#10b981', '#8b5cf6', '#e2e8f0']
        fig_pie = go.Figure(data=[go.Pie(labels=sov_data['브랜드'], values=sov_data['기사 수'], hole=.6, marker_colors=colors)])
        fig_pie.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=300, showlegend=True, legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.05))
        fig_pie.add_annotation(text="점유율<br><b style='font-size:24px'>6.4%</b><br>(12 / 186)", x=0.5, y=0.5, font_size=14, showarrow=False)
        st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})
        st.markdown("</div>", unsafe_allow_html=True)

    # 3. Details (Bottom Row) — 감성분석 | 기자현황 | 홍보팀 KPI
    b_col1, b_col2, b_col3 = st.columns([1, 1, 1])

    with b_col1:
        st.markdown("<div class='plot-container' style='height: 350px;'>", unsafe_allow_html=True)
        st.markdown("<h4>기사 감성 분석</h4>", unsafe_allow_html=True)
        sent_counts = df_all['감성'].value_counts().reindex(['긍정', '중립', '부정'], fill_value=0)
        total_sent = len(df_all) if len(df_all) > 0 else 1
        pos_pct = int((sent_counts['긍정'] / total_sent) * 100)
        neu_pct = int((sent_counts['중립'] / total_sent) * 100)
        neg_pct = 100 - pos_pct - neu_pct
        fig_sent = go.Figure(data=[go.Pie(
            labels=['긍정', '중립', '부정'],
            values=[sent_counts['긍정'], sent_counts['중립'], sent_counts['부정']],
            hole=.6,
            marker_colors=['#10b981', '#94a3b8', '#ef4444'],
            textinfo='percent',
            textfont_size=12
        )])
        fig_sent.update_layout(
            margin=dict(l=0, r=20, t=10, b=0), height=220,
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=12))
        )
        fig_sent.add_annotation(
            text=f"<b>{pos_pct}%</b>",
            x=0.43, y=0.5, font_size=18, showarrow=False
        )
        st.plotly_chart(fig_sent, use_container_width=True, config={'displayModeBar': False})
        st.markdown(f"""
        <div style='font-size:13px; color:#475569; margin-top:4px;'>
            <span style='color:#10b981;'>● 긍정</span> {pos_pct}% ({sent_counts['긍정']}건)&nbsp;&nbsp;
            <span style='color:#94a3b8;'>● 중립</span> {neu_pct}% ({sent_counts['중립']}건)&nbsp;&nbsp;
            <span style='color:#ef4444;'>● 부정</span> {neg_pct}% ({sent_counts['부정']}건)
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with b_col2:
        st.markdown("<div class='plot-container' style='height: 350px; overflow: hidden;'>", unsafe_allow_html=True)
        st.markdown("<h4>기자별 기사 작성 현황 TOP 5</h4>", unsafe_allow_html=True)
        top_reporters = df_all['기자명'].value_counts().head(5).reset_index()
        top_reporters.columns = ['기자명', '기사 수(건)']
        top_reporters.index = range(1, 6)
        mock_media = ["매일경제", "한국경제", "뉴스1", "파이낸셜뉴스", "아시아경제"]
        top_reporters['소속 매체'] = [mock_media[i] for i in range(len(top_reporters))]
        st.dataframe(top_reporters[['기자명', '소속 매체', '기사 수(건)']], use_container_width=True, hide_index=False)
        st.markdown("</div>", unsafe_allow_html=True)

    with b_col3:
        current_month = datetime.now().strftime("%m")
        st.markdown(f"""
        <div class='plot-container' style='height: 350px;'>
            <h4 style='margin-bottom:16px;'>홍보팀 KPI ({current_month}월 누적)</h4>
            <div style='display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px;'>
                <div style='background:#f8fafc; border-radius:10px; padding:14px 10px; text-align:center; border:1px solid #e2e8f0;'>
                    <div style='font-size:12px; color:#64748b; margin-bottom:6px;'>보도자료 배포</div>
                    <div style='font-size:26px; font-weight:700; color:#1e293b;'>8</div>
                    <div style='font-size:12px; color:#94a3b8;'>건</div>
                </div>
                <div style='background:#f8fafc; border-radius:10px; padding:14px 10px; text-align:center; border:1px solid #e2e8f0;'>
                    <div style='font-size:12px; color:#64748b; margin-bottom:6px;'>기사화</div>
                    <div style='font-size:26px; font-weight:700; color:#1e293b;'>112</div>
                    <div style='font-size:12px; color:#94a3b8;'>건</div>
                </div>
                <div style='background:#f0fdf4; border-radius:10px; padding:14px 10px; text-align:center; border:1px solid #bbf7d0;'>
                    <div style='font-size:12px; color:#64748b; margin-bottom:6px;'>평균 기사화율</div>
                    <div style='font-size:26px; font-weight:700; color:#10b981;'>14.0</div>
                    <div style='font-size:12px; color:#94a3b8;'>배</div>
                </div>
            </div>
            <div style='display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; margin-top: 12px;'>
                <div style='background:#f8fafc; border-radius:10px; padding:14px 10px; text-align:center; border:1px solid #e2e8f0;'>
                    <div style='font-size:12px; color:#64748b; margin-bottom:6px;'>단독 기사</div>
                    <div style='font-size:26px; font-weight:700; color:#1e293b;'>3</div>
                    <div style='font-size:12px; color:#94a3b8;'>건</div>
                </div>
                <div style='background:#f8fafc; border-radius:10px; padding:14px 10px; text-align:center; border:1px solid #e2e8f0;'>
                    <div style='font-size:12px; color:#64748b; margin-bottom:6px;'>인터뷰</div>
                    <div style='font-size:26px; font-weight:700; color:#1e293b;'>2</div>
                    <div style='font-size:12px; color:#94a3b8;'>건</div>
                </div>
                <div style='background:#f8fafc; border-radius:10px; padding:14px 10px; text-align:center; border:1px solid #e2e8f0;'>
                    <div style='font-size:12px; color:#64748b; margin-bottom:6px;'>칼럼/기고</div>
                    <div style='font-size:26px; font-weight:700; color:#1e293b;'>1</div>
                    <div style='font-size:12px; color:#94a3b8;'>건</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # 4. 최근 신세계면세점 주요 기사 (전체 너비, 카드 4개 가로 배열)
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;'>
        <h4 style='margin:0;'>최근 신세계면세점 주요 기사</h4>
        <a href='#' style='font-size:13px; color:#4f46e5; text-decoration:none;'>더보기 ›</a>
    </div>
    """, unsafe_allow_html=True)

    ssg_news = df_all[df_all['Search_Keyword'] == '신세계면세점'].head(4)
    news_images = [
        "https://images.unsplash.com/photo-1555529733-0e670560f7e1?w=300&h=180&fit=crop",
        "https://images.unsplash.com/photo-1528360983277-13d401cdc186?w=300&h=180&fit=crop",
        "https://images.unsplash.com/photo-1607082349566-187342175e2f?w=300&h=180&fit=crop",
        "https://images.unsplash.com/photo-1441986300917-64674bd600d8?w=300&h=180&fit=crop",
    ]
    news_cols = st.columns(4)
    for i, (col, row) in enumerate(zip(news_cols, ssg_news.itertuples())):
        img_url = news_images[i % len(news_images)]
        link = getattr(row, '_5', '#')
        date_str = str(getattr(row, '작성일', ''))[:10]
        with col:
            st.markdown(f"""
            <div style='background:white; border-radius:10px; border:1px solid #e2e8f0;
                        box-shadow:0 1px 4px rgba(0,0,0,0.06); overflow:hidden; cursor:pointer;'
                 onclick="window.open('{link}', '_blank')">
                <img src='{img_url}' style='width:100%; height:130px; object-fit:cover;' alt='news'>
                <div style='padding:12px;'>
                    <p style='font-size:13px; font-weight:600; color:#1e293b; margin:0 0 8px 0;
                               display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;
                               overflow:hidden; line-height:1.5;'>{row.헤드라인}</p>
                    <p style='font-size:11px; color:#94a3b8; margin:0;'>{row.매체} | {date_str}</p>
                </div>
            </div>
            """, unsafe_allow_html=True)


# ==========================================
# PAGE: 기사 목록
# ==========================================
elif st.session_state.current_page == "기사 목록":
    st.markdown("## 📄 기사 목록")

    # ── 날짜 선택 및 필터 영역 ──
    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns([1.5, 1.5, 1.5, 2])

    with filter_col1:
        # 날짜 선택: 오늘부터 최대 30일 전까지
        max_date = datetime.now().date() - timedelta(days=1)
        min_date = datetime.now().date() - timedelta(days=30)
        selected_date = st.date_input(
            "📅 날짜 선택",
            value=max_date,
            min_value=min_date,
            max_value=max_date,
            key="article_date_picker"
        )

    # 선택된 날짜의 데이터 로드
    selected_file, selected_date_str = get_file_for_date(selected_date)
    _, df_articles = load_data(selected_file)
    all_keywords = ["전체"] + sorted(df_articles['Search_Keyword'].unique().tolist())

    with filter_col2:
        selected_kw = st.selectbox("🔍 키워드", all_keywords, key="article_kw_filter")

    with filter_col3:
        selected_sent = st.selectbox("😊 감성", ["전체", "긍정", "중립", "부정"], key="article_sent_filter")

    with filter_col4:
        search_text = st.text_input("🔎 제목/내용 검색", placeholder="검색어를 입력하세요...", key="article_search")

    st.markdown("<hr style='border-color: #e2e8f0; margin: 4px 0 16px 0;'>", unsafe_allow_html=True)

    # ── 필터링 ──
    df_filtered = df_articles.copy()
    if selected_kw != "전체":
        df_filtered = df_filtered[df_filtered['Search_Keyword'] == selected_kw]
    if selected_sent != "전체":
        df_filtered = df_filtered[df_filtered['감성'] == selected_sent]
    if search_text:
        mask = (
            df_filtered['헤드라인'].astype(str).str.contains(search_text, case=False, na=False) |
            df_filtered['내용요약'].astype(str).str.contains(search_text, case=False, na=False)
        )
        df_filtered = df_filtered[mask]

    df_filtered = df_filtered.reset_index(drop=True)
    total_count = len(df_filtered)

    # ── 페이지네이션 설정 ──
    ITEMS_PER_PAGE = 15
    total_pages = max(1, (total_count + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)

    if "article_page" not in st.session_state:
        st.session_state.article_page = 1
    # 필터 변경 시 첫 페이지로 리셋
    filter_key = f"{selected_date_str}_{selected_kw}_{selected_sent}_{search_text}"
    if "last_filter_key" not in st.session_state or st.session_state.last_filter_key != filter_key:
        st.session_state.article_page = 1
        st.session_state.last_filter_key = filter_key

    current_page_num = st.session_state.article_page
    start_idx = (current_page_num - 1) * ITEMS_PER_PAGE
    end_idx = min(start_idx + ITEMS_PER_PAGE, total_count)
    df_page = df_filtered.iloc[start_idx:end_idx]

    # ── 요약 헤더 ──
    info_col1, info_col2 = st.columns([3, 1])
    with info_col1:
        st.markdown(
            f"<p style='color:#475569; font-size:14px;'>"
            f"<b style='color:#1e293b;'>{selected_date_str}</b> 기준 &nbsp;|&nbsp; "
            f"총 <b style='color:#4f46e5;'>{total_count}건</b> 조회 &nbsp;|&nbsp; "
            f"{start_idx+1}–{end_idx}건 표시</p>",
            unsafe_allow_html=True
        )

    # ── 기사 카드 목록 ──
    if total_count == 0:
        st.info("조건에 맞는 기사가 없습니다.")
    else:
        for _, row in df_page.iterrows():
            sent = row.get('감성', '중립')
            badge_class = {'긍정': 'badge-pos', '부정': 'badge-neg'}.get(sent, 'badge-neu')
            link = row.get('기사 링크', '#')
            date_str = str(row.get('작성일', ''))[:10]

            st.markdown(f"""
            <div class='article-row'>
                <div style='flex:1; min-width:0;'>
                    <div style='display:flex; align-items:center; gap:8px; margin-bottom:6px; flex-wrap:wrap;'>
                        <span class='badge {badge_class}'>{sent}</span>
                        <span style='font-size:12px; color:#64748b; background:#f1f5f9; padding:2px 8px; border-radius:12px;'>{row.get('Search_Keyword','')}</span>
                        <span style='font-size:12px; color:#94a3b8;'>{row.get('매체','')}</span>
                        <span style='font-size:12px; color:#94a3b8;'>|</span>
                        <span style='font-size:12px; color:#94a3b8;'>{row.get('기자명','')}</span>
                        <span style='font-size:12px; color:#94a3b8; margin-left:auto;'>{date_str}</span>
                    </div>
                    <a href='{link}' target='_blank' style='font-size:15px; font-weight:600; color:#1e293b; text-decoration:none;'>
                        {row.get('헤드라인','')}
                    </a>
                    <p style='font-size:13px; color:#64748b; margin:6px 0 0 0; overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;'>
                        {row.get('내용요약','')}
                    </p>
                </div>
            </div>
            """, unsafe_allow_html=True)

    # ── 페이지네이션 버튼 ──
    st.markdown("<br>", unsafe_allow_html=True)
    pg_cols = st.columns([2, 1, 1, 1, 2])

    with pg_cols[1]:
        if st.button("◀ 이전", disabled=(current_page_num <= 1), use_container_width=True, key="prev_page"):
            st.session_state.article_page -= 1
            st.rerun()

    with pg_cols[2]:
        st.markdown(
            f"<div class='page-info'><b>{current_page_num}</b> / {total_pages}</div>",
            unsafe_allow_html=True
        )

    with pg_cols[3]:
        if st.button("다음 ▶", disabled=(current_page_num >= total_pages), use_container_width=True, key="next_page"):
            st.session_state.article_page += 1
            st.rerun()


# ==========================================
# PAGE: 기자 분석
# ==========================================
elif st.session_state.current_page == "기자 분석":
    import numpy as np

    st.markdown("## 👤 기자 분석")
    st.markdown("<p style='color:#64748b;'>출입 기자 현황 및 보도 트렌드 분석</p>", unsafe_allow_html=True)

    # ── 날짜 선택 ──
    r_col1, r_col2 = st.columns([1, 4])
    with r_col1:
        max_date = datetime.now().date() - timedelta(days=1)
        min_date = datetime.now().date() - timedelta(days=30)
        sel_date = st.date_input("📅 기준일 선택", value=max_date,
                                 min_value=min_date, max_value=max_date, key="reporter_date")

    # ── 데이터 로드 (기준일 + 전일) ──
    cur_file, cur_date_str = get_file_for_date(sel_date)
    prev_file, prev_date_str = get_file_for_date(sel_date - timedelta(days=1))
    _, df_cur = load_data(cur_file)
    _, df_prev = load_data(prev_file)

    # '알수없음' 제거
    df_cur_c  = df_cur[df_cur['기자명'] != '알수없음'].copy()
    df_prev_c = df_prev[df_prev['기자명'] != '알수없음'].copy()

    # ── 기자별 집계 ──
    cur_cnt  = df_cur_c.groupby(['매체','기자명']).size().reset_index(name='현재_기사수')
    prev_cnt = df_prev_c.groupby(['매체','기자명']).size().reset_index(name='과거_기사수')
    trend_df = pd.merge(cur_cnt, prev_cnt, on=['매체','기자명'], how='outer')
    trend_df['현재_기사수'] = trend_df['현재_기사수'].fillna(0).astype(int)
    trend_df['과거_기사수'] = trend_df['과거_기사수'].fillna(0).astype(int)
    trend_df['누적_기사수'] = trend_df['현재_기사수'] + trend_df['과거_기사수']

    def calc_trend(cur, prev):
        if prev == 0 and cur > 0:   return '✨ 신규'
        elif cur > prev:             return '🔺 상승'
        elif cur < prev:             return '🔻 하락'
        elif cur == prev and cur > 0:return '➖ 유지'
        else:                        return '🔻 하락'

    trend_df['트렌드'] = trend_df.apply(lambda r: calc_trend(r['현재_기사수'], r['과거_기사수']), axis=1)

    # ── 감성 집계 ──
    sent_agg = df_cur_c.groupby(['매체','기자명'])['감성'].agg(
        lambda x: x.value_counts().idxmax() if len(x) > 0 else '중립'
    ).reset_index(name='주요_감성')

    # ── 관심 분야 태그 ──
    tag_kw = {
        '#투자':    ['투자','인수','M&A','지분','펀딩','상장','매각'],
        '#ESG':     ['ESG','지속가능','친환경','사회공헌','탄소중립'],
        '#신제품':  ['신제품','출시','론칭','공개','신규'],
        '#프로모션':['프로모션','이벤트','할인','캠페인','팝업'],
        '#실적':    ['실적','매출','영업이익','흑자','순이익'],
        '#인사':    ['인사','선임','대표이사','부회장','임원'],
        '#부정':    ['논란','의혹','수사','하락','과징금','사고'],
    }
    def find_tags(texts):
        found = set()
        full = ' '.join(str(t) for t in texts)
        for tag, kws in tag_kw.items():
            if any(k in full for k in kws):
                found.add(tag)
        return ' '.join(sorted(found)) if found else '#일반'

    df_cur_c['본문'] = df_cur_c['본문'].fillna('')
    tag_agg = df_cur_c.groupby(['매체','기자명']).apply(
        lambda g: find_tags(g['헤드라인'].tolist() + g['본문'].tolist())
    ).reset_index(name='관심분야')

    # ── 최근 보도일 ──
    recent_date = df_cur_c.groupby(['매체','기자명'])['작성일'].max().reset_index(name='최근_보도일')
    recent_date['최근_보도일'] = recent_date['최근_보도일'].astype(str).str[:10]

    # ── 최종 병합 ──
    final_df = trend_df.merge(sent_agg, on=['매체','기자명'], how='left') \
                       .merge(tag_agg,  on=['매체','기자명'], how='left') \
                       .merge(recent_date, on=['매체','기자명'], how='left')
    final_df['관심분야']   = final_df['관심분야'].fillna('#일반')
    final_df['주요_감성']  = final_df['주요_감성'].fillna('중립')
    final_df['최근_보도일']= final_df['최근_보도일'].fillna(prev_date_str)
    final_df = final_df.sort_values('누적_기사수', ascending=False).reset_index(drop=True)

    # ────────────────────────────────────────
    # KPI 요약 카드 (상단)
    # ────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    total_reporters = len(final_df)
    new_reporters   = len(final_df[final_df['트렌드'] == '✨ 신규'])
    rising          = len(final_df[final_df['트렌드'] == '🔺 상승'])
    active_today    = len(final_df[final_df['현재_기사수'] > 0])

    for col, title, val, sub, color in [
        (k1, "전체 출입 기자 수",    f"{total_reporters}명", f"{cur_date_str} 기준", "#4f46e5"),
        (k2, "금일 활동 기자",        f"{active_today}명",   f"{cur_date_str} 기사 작성", "#10b981"),
        (k3, "신규 등장 기자",        f"{new_reporters}명",  "전일 대비 신규",           "#f59e0b"),
        (k4, "기사 수 상승 기자",     f"{rising}명",         "전일 대비 증가",            "#3b82f6"),
    ]:
        with col:
            st.markdown(f"""
            <div class='kpi-card'>
                <div class='kpi-title'>{title}</div>
                <div class='kpi-value' style='color:{color};'>{val}</div>
                <div class='kpi-sub neutral'>{sub}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ────────────────────────────────────────
    # 필터 + 검색
    # ────────────────────────────────────────
    f1, f2, f3, f4 = st.columns([1.5, 1.5, 1.5, 2])
    with f1:
        media_opts = ["전체"] + sorted(final_df['매체'].unique().tolist())
        sel_media = st.selectbox("🏢 매체", media_opts, key="rep_media")
    with f2:
        trend_opts = ["전체", "✨ 신규", "🔺 상승", "➖ 유지", "🔻 하락"]
        sel_trend = st.selectbox("📈 트렌드", trend_opts, key="rep_trend")
    with f3:
        sent_opts = ["전체", "긍정", "중립", "부정"]
        sel_sent2 = st.selectbox("😊 주요 감성", sent_opts, key="rep_sent")
    with f4:
        rep_search = st.text_input("🔎 기자명 / 매체 검색", placeholder="검색어를 입력하세요...", key="rep_search")

    st.markdown("<hr style='border-color:#e2e8f0; margin:4px 0 16px 0;'>", unsafe_allow_html=True)

    # ── 필터링 ──
    disp = final_df.copy()
    if sel_media != "전체":
        disp = disp[disp['매체'] == sel_media]
    if sel_trend != "전체":
        disp = disp[disp['트렌드'] == sel_trend]
    if sel_sent2 != "전체":
        disp = disp[disp['주요_감성'] == sel_sent2]
    if rep_search:
        mask = (disp['기자명'].str.contains(rep_search, na=False) |
                disp['매체'].str.contains(rep_search, na=False))
        disp = disp[mask]
    disp = disp.reset_index(drop=True)

    st.markdown(
        f"<p style='color:#475569; font-size:14px;'>총 <b style='color:#4f46e5;'>{len(disp)}명</b> 조회</p>",
        unsafe_allow_html=True
    )

    # ────────────────────────────────────────
    # 기자 카드 테이블
    # ────────────────────────────────────────
    trend_color = {'✨ 신규': '#f59e0b', '🔺 상승': '#10b981', '➖ 유지': '#64748b', '🔻 하락': '#ef4444'}
    sent_badge  = {'긍정': ('badge-pos','긍정'), '부정': ('badge-neg','부정'), '중립': ('badge-neu','중립')}

    # 헤더
    st.markdown("""
    <div style='display:grid; grid-template-columns:2fr 1.5fr 0.8fr 0.8fr 0.8fr 1.5fr 2fr;
                gap:0; background:#f1f5f9; border-radius:8px 8px 0 0;
                padding:10px 16px; font-size:12px; font-weight:700; color:#475569;
                border:1px solid #e2e8f0; border-bottom:none;'>
        <span>기자명 / 매체</span>
        <span style='text-align:center'>최근 보도일</span>
        <span style='text-align:center'>누적 기사</span>
        <span style='text-align:center'>금일</span>
        <span style='text-align:center'>트렌드</span>
        <span style='text-align:center'>주요 감성</span>
        <span>관심 분야</span>
    </div>
    """, unsafe_allow_html=True)

    ITEMS = 20
    if 'rep_page' not in st.session_state:
        st.session_state.rep_page = 1
    rep_filter_key = f"{sel_media}_{sel_trend}_{sel_sent2}_{rep_search}_{str(sel_date)}"
    if st.session_state.get('rep_filter_key') != rep_filter_key:
        st.session_state.rep_page = 1
        st.session_state.rep_filter_key = rep_filter_key

    total_rep_pages = max(1, (len(disp) + ITEMS - 1) // ITEMS)
    pg = st.session_state.rep_page
    page_df = disp.iloc[(pg-1)*ITEMS : pg*ITEMS]

    for _, row in page_df.iterrows():
        t_color = trend_color.get(row['트렌드'], '#64748b')
        s_cls, s_lbl = sent_badge.get(row['주요_감성'], ('badge-neu','중립'))
        tags_html = ' '.join(
            f"<span style='background:#ede9fe; color:#6d28d9; border-radius:20px; "
            f"padding:2px 8px; font-size:11px; font-weight:600;'>{t}</span>"
            for t in row['관심분야'].split()
        ) if row['관심분야'] != '#일반' else \
            "<span style='background:#f1f5f9; color:#94a3b8; border-radius:20px; padding:2px 8px; font-size:11px;'>#일반</span>"

        st.markdown(f"""
        <div style='display:grid; grid-template-columns:2fr 1.5fr 0.8fr 0.8fr 0.8fr 1.5fr 2fr;
                    gap:0; background:white; padding:12px 16px; border:1px solid #e2e8f0;
                    border-top:none; align-items:center; transition:background 0.15s;'
             onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='white'">
            <div>
                <div style='font-size:14px; font-weight:600; color:#1e293b;'>{row['기자명']}</div>
                <div style='font-size:12px; color:#64748b; margin-top:2px;'>{row['매체']}</div>
            </div>
            <div style='text-align:center; font-size:13px; color:#475569;'>{row['최근_보도일']}</div>
            <div style='text-align:center; font-size:15px; font-weight:700; color:#1e293b;'>{row['누적_기사수']}</div>
            <div style='text-align:center; font-size:15px; font-weight:700; color:#4f46e5;'>{row['현재_기사수']}</div>
            <div style='text-align:center; font-size:13px; font-weight:600; color:{t_color};'>{row['트렌드']}</div>
            <div style='text-align:center;'><span class='badge {s_cls}'>{s_lbl}</span></div>
            <div style='display:flex; flex-wrap:wrap; gap:4px;'>{tags_html}</div>
        </div>
        """, unsafe_allow_html=True)

    # 테이블 하단 경계
    st.markdown("<div style='border:1px solid #e2e8f0; border-top:none; border-radius:0 0 8px 8px; height:4px; background:#f8fafc;'></div>", unsafe_allow_html=True)

    # ── 페이지네이션 ──
    st.markdown("<br>", unsafe_allow_html=True)
    pg_c = st.columns([2, 1, 1, 1, 2])
    with pg_c[1]:
        if st.button("◀ 이전", disabled=(pg <= 1), use_container_width=True, key="rep_prev"):
            st.session_state.rep_page -= 1
            st.rerun()
    with pg_c[2]:
        st.markdown(f"<div class='page-info'><b>{pg}</b> / {total_rep_pages}</div>", unsafe_allow_html=True)
    with pg_c[3]:
        if st.button("다음 ▶", disabled=(pg >= total_rep_pages), use_container_width=True, key="rep_next"):
            st.session_state.rep_page += 1
            st.rerun()


# ==========================================
# PAGE: 기타 (미구현)
# ==========================================
else:
    page_name = st.session_state.current_page
    st.markdown(f"## {page_name}")
    st.markdown("<br>", unsafe_allow_html=True)
    st.info(f"🚧 **{page_name}** 페이지는 준비 중입니다.")
