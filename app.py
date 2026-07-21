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
# Streamlit 기본 UI 요소 숨기기 (원본 HTML 디자인 100% 유지)
# ================================================================
hide_streamlit_style = """
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {display: none;}
    .stAppDeployButton {display: none;}
    [data-testid="stToolbar"] {visibility: hidden;}
    [data-testid="stDecoration"] {visibility: hidden;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    .block-container {
        padding: 0 !important;
        max-width: 100% !important;
    }
    .stApp {
        margin: 0;
        padding: 0;
    }
    iframe {
        width: 100%;
        border: none;
    }
    section[data-testid="stSidebar"] {display: none;}
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

# ================================================================
# 조회수 표시 (footer 위에 삽입) - 원본 코드 훼손 없이 삽입만 진행
# ================================================================
view_counter_html = f'''
    <div style="text-align:center; padding: 16px 18px 0 18px; background: var(--bg);">
      <div style="display:inline-flex; align-items:center; background:white; border:1px solid var(--border); border-radius:12px; padding:10px 20px; box-shadow:0 2px 8px rgba(0,0,0,0.05); font-size:16px; color:var(--posco-navy);">
        <span style="margin:0 8px;">📊 조회수</span>
        <span style="margin:0 8px;">Today <b style="color:var(--posco-blue); font-size:18px;">{today_cnt}</b></span>
        <span style="color:#ccc;">|</span>
        <span style="margin:0 8px;">Total <b style="color:var(--posco-blue); font-size:18px;">{total_cnt}</b></span>
      </div>
    </div>
    <footer class="copyright">
'''

# footer 태그 앞에 조회수 삽입
html_content = html_content.replace(
    '<footer class="copyright">',
    view_counter_html
)

# ================================================================
# HTML 렌더링 (원본 100% 유지 + 조회수만 삽입)
# ================================================================
components.html(
    html_content,
    height=2600,       # 새 버전은 섹션이 늘어나 조금 더 여유 있게
    scrolling=True
)
