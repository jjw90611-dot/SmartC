import streamlit as st
import pandas as pd
from datetime import date

# ===== 페이지 설정 =====
st.set_page_config(page_title="밀폐공간 작업 프로그램", page_icon="📋", layout="wide")

st.title("📋 밀폐공간 작업 프로그램 작성")
st.caption("POSCO FUTUREM : Company Use Only")
st.divider()

# ===== 1. 기본 정보 =====
st.header("1️⃣ 기본 정보")
col1, col2 = st.columns(2)
with col1:
    project_name = st.text_input("공사명", placeholder="예: OO 설비 정비공사")
    construction_period = st.text_input("공사기간", placeholder="예: 2025.01.01 ~ 2025.12.31")
    company_name = st.text_input("회사명")
with col2:
    confined_work_period = st.text_input("밀폐공간작업기간", placeholder="예: 2025.03.01 ~ 2025.03.15")
    work_detail = st.text_input("작업내용", placeholder="예: 반응기 내부 청소")
    site_manager = st.text_input("현장소장")

write_date = st.date_input("작성일", value=date.today())

st.divider()

# ===== 2. 밀폐공간 위치 및 위험요인 =====
st.header("2️⃣ 밀폐공간 위치 및 유해·위험요인")

st.subheader("(1) 작업장소별 위험물질 관리 현황")
hazard_df = st.data_editor(
    pd.DataFrame({
        "작업장소": [""],
        "공정명": [""],
        "유해위험물질": [""],
        "입조전_관리방안": [""],
        "입조중_관리방안": [""],
    }),
    num_rows="dynamic",
    use_container_width=True,
    key="hazard"
)

st.subheader("(2) Blind 설치 현황")
blind_df = st.data_editor(
    pd.DataFrame({
        "공정명": [""],
        "Blind설치개수(EA)": [0],
        "확인자": [""],
    }),
    num_rows="dynamic",
    use_container_width=True,
    key="blind"
)

st.divider()

# ===== 3. 유해위험물질 및 대책 =====
st.header("3️⃣ 유해·위험요인 및 대책")
material_df = st.data_editor(
    pd.DataFrame({
        "물질명": [""],
        "유해위험요인": [""],
        "대책": [""],
    }),
    num_rows="dynamic",
    use_container_width=True,
    key="material"
)

st.divider()

# ===== 4. 교육 =====
st.header("4️⃣ 안전보건교육 및 훈련")
col1, col2 = st.columns(2)
with col1:
    st.markdown("**특별안전보건교육**")
    special_edu_date = st.text_input("일정", key="s_date")
    special_edu_instructor = st.text_input("강사", key="s_ins")
with col2:
    st.markdown("**가스농도 측정자 교육**")
    gas_edu_date = st.text_input("일정", key="g_date")
    gas_edu_instructor = st.text_input("강사", key="g_ins")

st.divider()

# ===== 5. 장비 현황 =====
st.header("5️⃣ 장비 현황")

st.subheader("(1) 보유장비")
owned_df = st.data_editor(
    pd.DataFrame({
        "장비명": [""],
        "모델명": [""],
        "보유대수": [0],
        "최근교정일": [""],
        "교정주기": [""],
    }),
    num_rows="dynamic",
    use_container_width=True,
    key="owned"
)

st.subheader("(2) 대여 장비")
rental_df = st.data_editor(
    pd.DataFrame({
        "장비명": [""],
        "필요수량": [0],
        "대여일수": [""],
        "연락처": [""],
        "장비관리담당자": [""],
    }),
    num_rows="dynamic",
    use_container_width=True,
    key="rental"
)

st.subheader("(3) 비상용 구조장비")
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("**구조장비**")
    chk_crane = st.checkbox("이동식 크레인")
    chk_hoist = st.checkbox("호이스트")
    chk_tripod = st.checkbox("삼각대")
    chk_rope = st.checkbox("구명로프")
    chk_stretcher = st.checkbox("들것")
    chk_winch = st.checkbox("윈치")
    rescue_etc = st.text_input("기타 구조장비")
with col2:
    st.markdown("**호흡장비**")
    chk_airline = st.checkbox("송기마스크")
    chk_elsa = st.checkbox("ELSA(대피용)")
    breathing_etc = st.text_input("기타 호흡장비")
with col3:
    st.markdown("**통화장비**")
    chk_radio = st.checkbox("무전기")
    chk_speaker = st.checkbox("확성기")
    chk_whistle = st.checkbox("호각")
    comm_etc = st.text_input("기타 통화장비")

etc_equipment = st.text_area("기타 추가장비")

st.divider()

# ===== 6. 비상연락체계 =====
st.header("6️⃣ 비상연락체계 및 작업자 정보")

st.subheader("비상연락체계")
emergency_df = st.data_editor(
    pd.DataFrame({
        "구분": [""],
        "성명": [""],
        "연락처": [""],
        "사내소방담당부서": [""],
    }),
    num_rows="dynamic",
    use_container_width=True,
    key="emergency"
)

st.subheader("작업자 정보")
worker_df = st.data_editor(
    pd.DataFrame({
        "구분": [""],
        "성명": [""],
        "연락처": [""],
    }),
    num_rows="dynamic",
    use_container_width=True,
    key="worker"
)

st.divider()

# ===== 7. 환기시설 =====
st.header("7️⃣ 환기시설")
col1, col2 = st.columns(2)
with col1:
    vent_process = st.text_input("공정명", key="v_proc")
with col2:
    volume = st.number_input("작업장소 용적(㎥)", min_value=0.0, value=0.0, step=1.0)

st.subheader("환기팬 목록")
fan_df = st.data_editor(
    pd.DataFrame({
        "모델명": [""],
        "환기팬용량(㎥/h)": [0.0],
        "수량(EA)": [0],
        "장비관리담당자": [""],
    }),
    num_rows="dynamic",
    use_container_width=True,
    key="fan"
)

# 환기능력 자동계산
if not fan_df.empty and volume > 0:
    total_capacity = (fan_df["환기팬용량(㎥/h)"] * fan_df["수량(EA)"]).sum()
    total_count = fan_df["수량(EA)"].sum()
    
    st.info(f"**총 환기능력**: {total_capacity:,.1f} ㎥/h  |  **총 대수**: {int(total_count)} EA")
    
    pre_required = volume * 10
    fan_per_min = total_capacity / 60
    vent_time = pre_required / fan_per_min if fan_per_min > 0 else 0
    run_required = volume * 20
    
    st.markdown("**⏱ 작업 시작 전 (용적의 10배 이상 환기)**")
    st.write(f"- 필요 환기량: {pre_required:,.1f} ㎥")
    st.write(f"- 분당 환기능력: {fan_per_min:,.2f} ㎥/min")
    st.write(f"- 필요 가동시간: **{vent_time:,.1f} 분 이상**")
    
    st.markdown("**⏱ 작업 중 (시간당 용적의 20회 이상 환기)**")
    st.write(f"- 시간당 필요량: {run_required:,.1f} ㎥/hr")
    result = "✅ 적합" if total_capacity >= run_required else "❌ 부적합"
    st.write(f"- 판정: {result}")

vent_note = st.text_area("환기시설 특이사항")

st.divider()

# ===== 최종 미리보기 =====
st.header("📄 최종 문서 미리보기")

if st.button("📋 미리보기 생성", type="primary", use_container_width=True):
    st.markdown("---")
    st.markdown("### 밀폐공간 작업 프로그램")
    st.markdown(f"""
    - **공사명**: {project_name}
    - **공사기간**: {construction_period}
    - **밀폐공간작업기간**: {confined_work_period}
    - **작업내용**: {work_detail}
    - **작성일**: {write_date}
    - **회사명**: {company_name}
    - **현장소장**: {site_manager}
    """)
    
    st.markdown("#### 1. 밀폐공간 위치 및 위험물질")
    st.dataframe(hazard_df, use_container_width=True)
    
    st.markdown("#### 2. Blind 설치 현황")
    st.dataframe(blind_df, use_container_width=True)
    
    st.markdown("#### 3. 유해위험요인 및 대책")
    st.dataframe(material_df, use_container_width=True)
    
    st.markdown("#### 4. 교육")
    st.write(f"- 특별안전보건교육: {special_edu_date} / 강사: {special_edu_instructor}")
    st.write(f"- 가스농도 측정자 교육: {gas_edu_date} / 강사: {gas_edu_instructor}")
    
    st.markdown("#### 5. 보유장비")
    st.dataframe(owned_df, use_container_width=True)
    
    st.markdown("#### 6. 대여장비")
    st.dataframe(rental_df, use_container_width=True)
    
    st.markdown("#### 7. 비상연락체계")
    st.dataframe(emergency_df, use_container_width=True)
    
    st.markdown("#### 8. 작업자")
    st.dataframe(worker_df, use_container_width=True)
    
    st.markdown("#### 9. 환기팬")
    st.dataframe(fan_df, use_container_width=True)
    
    st.success("✅ 미리보기 완성! 인쇄(Ctrl+P)로 PDF 저장 가능합니다.")
