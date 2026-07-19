import streamlit as st
import math
import pandas as pd
from datetime import datetime

# --- [1] 페이지 설정 및 인쇄용 CSS ---
st.set_page_config(page_title="포스코퓨처엠 밀폐공간 작업 프로그램 자동화", layout="wide")

# 브라우저 인쇄(Ctrl+P) 시 사이드바와 불필요한 UI를 숨기는 CSS
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            @media print {
                section[data-testid="stSidebar"] {display: none;}
                .stButton {display: none;}
                div[data-testid="stToolbar"] {display: none;}
            }
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# --- [2] 데이터베이스 (작업 종류별 위험요인 및 대책) ---
HAZARD_DB = {
    "내화물 해체/시공": {
        "hazard": "건조 재료 분진 흡입, 피부/눈 접촉, 산소결핍",
        "solution": "작업 전/중 환기, 방진마스크 착용, 세안설비 위치 안내",
        "ppe": "방진마스크, 보안경, 안전모, 안전화",
        "scenario": "산소결핍 질식 및 분진 흡입 환자 구조 훈련"
    },
    "용접/사상/절단 (화기작업)": {
        "hazard": "용접 흄 발생, 일산화탄소(CO) 중독, 불티로 인한 화재/폭발",
        "solution": "불티비산방지포 설치, 소화기 비치, 고정식 가스검지기 운용",
        "ppe": "송기마스크 또는 방독마스크, 용접면, 방염복",
        "scenario": "밀폐공간 내 화기작업 중 화재 발생 대응 훈련"
    },
    "도장/코팅/세척 (유기용제)": {
        "hazard": "유기용제 증기 흡입, 인화성 가스(LEL) 폭발, 질식",
        "solution": "방폭형 환기팬 가동, 화기 반입 엄금, 연속 가스 측정",
        "ppe": "방독마스크(유기화합물용), 내화학장갑, 보호복",
        "scenario": "유기용제 중독 환자 발생 및 화재 대피 훈련"
    },
    "슬러지/퇴적물 제거": {
        "hazard": "황화수소(H2S) 가스 중독, 메탄 가스 폭발, 미끄러짐",
        "solution": "작업 전 체적 10배 이상 환기, 슬러지 교란 전 가스 재측정",
        "ppe": "송기마스크, 방수복, 안전장화, 구명줄",
        "scenario": "황화수소 중독 질식 환자 구출 및 심폐소생술 훈련"
    },
    "단순 점검/가설작업": {
        "hazard": "산소결핍, 어두운 조명으로 인한 추락/전도",
        "solution": "작업 전 가스측정, 방폭형 조명 설치, 안전대 체결",
        "ppe": "안전모, 안전화, 그네식 안전대",
        "scenario": "밀폐공간 내부 추락 환자 들것 구조 훈련"
    }
}

# --- [3] 사이드바 입력 폼 ---
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/8/8d/POSCO_Future_M_logo.svg/512px-POSCO_Future_M_logo.svg.png", width=200)
st.sidebar.header("📝 밀폐공간 프로그램 입력")

with st.sidebar.expander("1. 공사 기본 정보", expanded=True):
    project_name = st.text_input("공사명", "6기 CDQ Chamber 기둥연와 개선교체")
    start_date = st.date_input("작업 시작일")
    end_date = st.date_input("작업 종료일")
    manager = st.text_input("현장소장", "김국현")
    safety_officer = st.text_input("안전관리자", "임근범")
    watcher = st.text_input("감시인", "배명주")
    workers_count = st.number_input("동시 출입 근로자 수(명)", min_value=1, value=5)

with st.sidebar.expander("2. 밀폐공간 제원 및 환경", expanded=True):
    space_name = st.text_input("작업 장소명", "CDQ Chamber 내부")
    space_type = st.selectbox("공간 형태", ["사각형", "원통형"])
    
    if space_type == "사각형":
        width = st.number_input("가로 (m)", value=10.0)
        length = st.number_input("세로 (m)", value=10.0)
        height = st.number_input("높이 (m)", value=10.3)
        volume = width * length * height
    else:
        diameter = st.number_input("직경 (m)", value=10.0)
        height = st.number_input("높이 (m)", value=10.0)
        volume = math.pi * ((diameter/2)**2) * height
        
    temp = st.number_input("내부 예상 기온 (℃)", value=32.0)
    humidity = st.number_input("내부 예상 습도 (%)", value=60.0)

with st.sidebar.expander("3. 작업 내용 및 장비", expanded=True):
    task_types = st.multiselect("작업 종류 (다중 선택 가능)", list(HAZARD_DB.keys()), default=["내화물 해체/시공"])
    fan_capacity = st.number_input("환기팬 1대당 용량 (m³/hr)", value=4600)
    fan_count = st.number_input("환기팬 수량 (대)", min_value=1, value=4)
    gas_detector = st.text_input("가스측정기 모델명", "GX-3R (4종)")

generate_btn = st.sidebar.button("🚀 프로그램 자동 생성", use_container_width=True)

# --- [4] 메인 화면 (보고서 생성 로직) ---
if generate_btn:
    if not task_types:
        st.warning("작업 종류를 최소 1개 이상 선택해주세요.")
        st.stop()

    # 계산 로직
    total_fan_capacity_hr = fan_capacity * fan_count
    total_fan_capacity_min = total_fan_capacity_hr / 60
    
    # 환기 기준 (포스코퓨처엠 지침)
    pre_work_vol = volume * 5
    pre_work_time = pre_work_vol / total_fan_capacity_min if total_fan_capacity_min > 0 else 0
    during_work_vol_hr = volume * 20
    
    # 체감온도 계산 (간이 식 적용)
    heat_index = temp + 0.33 * humidity - 0.7 * 10 - 4.0
    if heat_index >= 38: heat_level, rest_rule = "위험", "매 시간 15분 작업 / 45분 휴식"
    elif heat_index >= 35: heat_level, rest_rule = "경고", "매 시간 45분 작업 / 15분 휴식"
    elif heat_index >= 33: heat_level, rest_rule = "주의", "매 시간 50분 작업 / 10분 휴식"
    else: heat_level, rest_rule = "관심", "주기적인 수분 섭취 및 자율 휴식"

    # --- 문서 출력 시작 ---
    st.title("📄 밀폐공간 작업 프로그램")
    st.markdown(f"**작성일:** {datetime.now().strftime('%Y년 %m월 %d일')} | **회사명:** ㈜포스코퓨처엠")
    st.divider()

    # 1. 기본 정보
    st.header("1. 공사 및 작업자 정보")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"- **공 사 명:** {project_name}")
        st.markdown(f"- **작업기간:** {start_date} ~ {end_date}")
        st.markdown(f"- **작업장소:** {space_name} ({space_type})")
    with col2:
        st.markdown(f"- **현장소장:** {manager}")
        st.markdown(f"- **안전관리자:** {safety_officer}")
        st.markdown(f"- **감시인:** {watcher} 외")
        st.markdown(f"- **출입인원:** {workers_count} 명")

    # 2. 유해위험요인 및 대책
    st.header("2. 유해·위험요인 파악 및 관리방안")
    hazard_data = []
    for task in task_types:
        hazard_data.append({
            "작업 종류": task,
            "유해·위험요인": HAZARD_DB[task]["hazard"],
            "관리 대책": HAZARD_DB[task]["solution"],
            "필수 보호구": HAZARD_DB[task]["ppe"]
        })
    st.table(pd.DataFrame(hazard_data))

    # 3. 환기 능력 계산서
    st.header("3. 환기 능력 계산서 (KOSHA GUIDE 및 사내 지침 기준)")
    st.info(f"**밀폐공간 체적(V):** {volume:,.1f} m³")
    
    v_col1, v_col2 = st.columns(2)
    with v_col1:
        st.subheader("작업 전 환기 (체적의 5배 이상)")
        st.markdown(f"- **필요 환기량:** {pre_work_vol:,.1f} m³")
        st.markdown(f"- **보유 환기능력:** {total_fan_capacity_min:,.1f} m³/min")
        st.markdown(f"- **환기 소요시간:** **약 {math.ceil(pre_work_time)}분 이상 가동 필요**")
        st.success("계산결과: 충족 (작업 전 반드시 위 시간 이상 환기 실시)")
        
    with v_col2:
        st.subheader("작업 중 환기 (시간당 체적의 20회 이상)")
        st.markdown(f"- **시간당 필요 환기량:** {during_work_vol_hr:,.1f} m³/hr")
        st.markdown(f"- **보유 환기능력:** {total_fan_capacity_hr:,.1f} m³/hr")
        if total_fan_capacity_hr >= during_work_vol_hr:
            st.success("계산결과: 충족 (작업 중 상시 가동)")
        else:
            st.error("계산결과: 부족 (환기팬 추가 투입 필요!)")

    # 4. 온열질환 예방 계획
    st.header("4. 온열질환 예방 및 국소냉방 계획")
    st.markdown(f"- **내부 예상 기온/습도:** {temp}℃ / {humidity}%")
    st.markdown(f"- **예상 체감온도:** **{heat_index:.1f}℃ ({heat_level})**")
    st.markdown(f"- **휴식 기준:** {rest_rule}")
    st.markdown(f"- **스포트쿨러 필요 냉방능력:** {workers_count * 1500:,} kcal/hr (중작업 기준)")

    # 5. 가스 측정 기록부 양식
    st.header("5. 산소 및 유해가스 농도 측정 기록부")
    st.markdown(f"- **사용 장비:** {gas_detector}")
    gas_df = pd.DataFrame({
        "구분": ["작업 전 (오전)", "작업 중 (오전)", "작업 전 (오후)", "작업 중 (오후)"],
        "측정시간": [" : ", " : ", " : ", " : "],
        "O2 (18~23.5%)": ["", "", "", ""],
        "CO (<30ppm)": ["", "", "", ""],
        "H2S (<10ppm)": ["", "", "", ""],
        "CO2 (<1.5%)": ["", "", "", ""],
        "LEL (<10%)": ["", "", "", ""],
        "측정자 서명": ["", "", "", ""]
    })
    st.table(gas_df)

    # 6. 비상상황 훈련 계획
    st.header("6. 비상상황 교육 및 훈련 계획")
    scenario = HAZARD_DB[task_types[0]]["scenario"] # 첫번째 작업 기준 시나리오
    st.markdown(f"- **훈련 시나리오:** **{scenario}**")
    st.markdown(f"- **훈련 주관:** {manager} (현장소장)")
    st.markdown(f"- **비상 연락망:**")
    st.markdown(f"  - 현장소장: {manager}")
    st.markdown(f"  - 안전관리자: {safety_officer}")
    st.markdown(f"  - 포스코퓨처엠 방재센터: 사내 비상연락망 참조")

    # 인쇄 안내
    st.divider()
    st.caption("💡 **Tip:** 키보드의 `Ctrl + P` (Mac은 `Cmd + P`)를 누르시면 이 보고서를 PDF로 저장하거나 바로 인쇄할 수 있습니다.")

else:
    st.info("👈 좌측 사이드바에 정보를 입력하고 **'프로그램 자동 생성'** 버튼을 눌러주세요.")
