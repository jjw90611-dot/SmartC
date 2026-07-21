import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path
from datetime import datetime
import json
import os

# ================================================================
# 페이지 기본 설정
# ================================================================
st.set_page_config(
    page_title="포스코퓨처엠 스마트 밀폐공간 안전설계 및 온열질환 예방 프로그램",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ================================================================
# Streamlit 기본 UI 요소 숨기기 (원본 HTML 디자인 유지)
# ================================================================
hide_streamlit_style = """
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {display: none;}
    .block-container {
        padding-top: 0rem;
        padding-bottom: 0rem;
        padding-left: 0rem;
        padding-right: 0rem;
        max-width: 100%;
    }
    .stApp {
        margin: 0;
        padding: 0;
    }
    iframe {
        width: 100%;
        border: none;
    }
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ================================================================
# 조회수 카운터 (파일 기반 간이 카운터)
# ================================================================
COUNTER_FILE = "view_count.json"

def load_counter():
    """조회수 데이터 로드"""
    if not os.path.exists(COUNTER_FILE):
        return {"total": 0, "daily": {}}
    try:
        with open(COUNTER_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"total": 0, "daily": {}}

def save_counter(data):
    """조회수 데이터 저장"""
    try:
        with open(COUNTER_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def update_counter():
    """세션당 1회만 조회수 증가"""
    if "counted" not in st.session_state:
        data = load_counter()
        today = datetime.now().strftime("%Y-%m-%d")
        data["total"] = data.get("total", 0) + 1
        data["daily"][today] = data.get("daily", {}).get(today, 0) + 1
        save_counter(data)
        st.session_state.counted = True
        st.session_state.today_count = data["daily"][today]
        st.session_state.total_count = data["total"]
    return st.session_state.today_count, st.session_state.total_count

today_cnt, total_cnt = update_counter()

# ================================================================
# 원본 HTML 파일 로드
# ================================================================
html_path = Path(__file__).parent / "index.html"
html_content = html_path.read_text(encoding="utf-8")

# 조회수 가짜 데이터를 실제 카운터로 치환
html_content = html_content.replace(
    'document.getElementById(\'todayCount\').innerText = "12";',
    f'document.getElementById(\'todayCount\').innerText = "{today_cnt}";'
)
html_content = html_content.replace(
    'document.getElementById(\'totalCount\').innerText = "345";',
    f'document.getElementById(\'totalCount\').innerText = "{total_cnt}";'
)

# ================================================================
# HTML 렌더링 (원본 100% 유지)
# ================================================================
components.html(
    html_content,
    height=2400,       # 결과 페이지 전체가 보이도록 넉넉히 설정
    scrolling=True
)
