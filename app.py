import streamlit as st
from docxtpl import DocxTemplate
from datetime import date
from io import BytesIO
import pandas as pd

st.set_page_config(page_title="밀폐공간 작업 프로그램 자동생성", page_icon="📋", layout="wide")

st.title("📋 밀폐공간 작업 프로그램 자동 생성")
st.caption("포스코퓨처엠 | 항목을 입력하면 워드 파일이 자동 생성됩니다.")

# ================== 1. 기본 정보 ==================
st.header("1️⃣ 기본 정보")
col1, col2 = st.columns(2)
with col1:
    project_name = st.text_input("공사명", "R-30101A 정비공사")
    construction_period = st.text_input("공사기간", "2025.01.01 ~ 2025.03.31")
    confined_work_period = st.text_input("밀폐공간 작업기간", "2025.02.01 ~ 2025.02.15")
with col2:
    work_detail = st.text_input("작업내용", "반응기 내부 촉매 교체")
    company_name = st.text_input("회사명", "포스코퓨처엠")
    site_manager = st.text_input("현장소장", "홍길동")
write_date = st.date_input("작성일", date.today())

# ================== 2. 위험요인 관리 ==================
st.header("2️⃣ 밀폐공간 위치 및 위험물질 관리")
st.markdown("**(1) 사업장 내 밀폐공간 위치 및 위험물질 관리 현황**")
hazard_df = st.data_editor(
    pd.DataFrame([
        {"작업장소": "R-30101A", "공정명": "반응기 정비", "유해위험물질": "N2, 산화니켈", "입조전_관리방안": "N2 퍼지 후 Air 치환", "입조중_관리방안": "연속 가스측정"},
    ]),
    num_rows="dynamic", use_container_width=True, key="hazard"
)

st.markdown("**(2) Blind 설치 개수 및 위치**")
blind_df = st.data_editor(
    pd.DataFrame([{"공정명": "R-30101A", "Blind설치개수": "5", "확인자": "홍길동"}]),
    num_rows="dynamic", use_container_width=True, key="blind"
)

# ================== 3. 유해위험요인별 대책 ==================
st.header("3️⃣ 유해·위험요인별 대책")
material_df = st.data_editor(
    pd.DataFrame([
        {"물질명": "N2", "유해위험요인": "질식", "대책": "MSDS 교육, Air Line Mask 착용, 감시인 배치"},
    ]),
    num_rows="dynamic", use_container_width=True, key="material"
)

# ================== 4. 교육 ==================
st.header("4️⃣ 안전보건교육 및 훈련")
col1, col2 = st.columns(2)
with col1:
    st.subheader("특별안전보건교육")
    special_edu_date = st.text_input("교육 일정 (특별안전)", "2025.01.20")
    special_edu_instructor = st.text_input("강사 (특별안전)", "김안전")
with col2:
    st.subheader("가스농도 측정자 교육")
    gas_edu_date = st.text_input("교육 일정 (가스측정)", "2025.01.22")
    gas_edu_instructor = st.text_input("강사 (가스측정)", "이보건")

# ================== 5. 장비 ==================
st.header("5️⃣ 장비 현황")
st.markdown("**(1) 보유장비 현황**")
owned_df = st.data_editor(
    pd.DataFrame([{"장비명": "산소농도측정기", "모델명": "GX-2009", "보유대수": "3", "최근교정일": "2024.12.01", "교정주기": "1년"}]),
    num_rows="dynamic", use_container_width=True, key="owned"
)

st.markdown("**(2) 대여 장비 현황**")
rental_df = st.data_editor(
    pd.DataFrame([{"장비명": "송풍기", "필요수량": "3", "대여일수": "15", "연락처": "000-0000-0000", "장비관리담당자": "박대여"}]),
    num_rows="dynamic", use_container_width=True, key="rental"
)

st.markdown("**(3) 비상용 구조장비 비치현황**")
col1, col2, col3 = st.columns(3)
with col1:
    st.write("**구조장비**")
    rescue = {
        "이동식크레인": st.checkbox("이동식 크레인"),
        "호이스트": st.checkbox("호이스트"),
        "삼각대": st.checkbox("삼각대", value=True),
        "구명로프": st.checkbox("구명로프", value=True),
        "들것": st.checkbox("들것", value=True),
        "윈치": st.checkbox("윈치"),
    }
    rescue_etc = st.text_input("구조장비 기타", "")
with col2:
    st.write("**호흡장비**")
    breathing = {
        "송기마스크": st.checkbox("송기마스크(Air Line Mask)", value=True),
        "ELSA": st.checkbox("ELSA (대피용)", value=True),
    }
    breathing_etc = st.text_input("호흡장비 기타", "")
with col3:
    st.write("**통화장비**")
    comm = {
        "무전기": st.checkbox("무전기", value=True),
        "확성기": st.checkbox("확성기"),
        "호각": st.checkbox("호각", value=True),
    }
    comm_etc = st.text_input("통화장비 기타", "")

etc_equipment = st.text_area("기타 추가장비", "")

# ================== 6. 인원 ==================
st.header("6️⃣ 비상연락체계 및 작업자 정보")
st.markdown("**(1) 비상연락체계**")
emergency_df = st.data_editor(
    pd.DataFrame([
        {"구분": "현장소장", "성명": "홍길동", "연락처": "010-0000-0001", "사내소방담당부서": "안전방재팀"},
        {"구분": "안전관리자", "성명": "김안전", "연락처": "010-0000-0002", "사내소방담당부서": "안전방재팀"},
        {"구분": "관리감독자", "성명": "이감독", "연락처": "010-0000-0003", "사내소방담당부서": "안전방재팀"},
        {"구분": "작업반장", "성명": "박반장", "연락처": "010-0000-0004", "사내소방담당부서": "안전방재팀"},
    ]),
    num_rows="dynamic", use_container_width=True, key="emergency"
)

st.markdown("**(2) 작업자 정보**")
worker_df = st.data_editor(
    pd.DataFrame([
        {"구분": "공사감독", "성명": "", "연락처": ""},
        {"구분": "관리감독자 1", "성명": "", "연락처": ""},
        {"구분": "감시인 1", "성명": "", "연락처": ""},
        {"구분": "작업자 1", "성명": "", "연락처": ""},
    ]),
    num_rows="dynamic", use_container_width=True, key="worker"
)

# ================== 7. 환기시설 (자동계산) ==================
st.header("7️⃣ 환기시설 (자동 계산)")
col1, col2 = st.columns(2)
with col1:
    vent_process = st.text_input("공정명", "R-30101A")
with col2:
    volume = st.number_input("작업장소 용적 (㎥)", min_value=0.0, value=284.5, step=0.1)

st.markdown("**환기팬 목록** (모델명·용량·수량 입력 → 자동 합산)")
fan_df = st.data_editor(
    pd.DataFrame([
        {"모델명": "SMP-20", "환기팬용량(㎥/h)": 1800, "수량(EA)": 1, "장비관리담당자": "홍길동(000-0000-0000)"},
        {"모델명": "SMP-25", "환기팬용량(㎥/h)": 3120, "수량(EA)": 1, "장비관리담당자": "홍길동(000-0000-0000)"},
        {"모델명": "SMP-30", "환기팬용량(㎥/h)": 4620, "수량(EA)": 1, "장비관리담당자": "홍길동(000-0000-0000)"},
    ]),
    num_rows="dynamic", use_container_width=True, key="fan"
)

# 자동 계산
total_fan_capacity = int((fan_df["환기팬용량(㎥/h)"] * fan_df["수량(EA)"]).sum())
total_fan_count = int(fan_df["수량(EA)"].sum())
fan_per_min = round(total_fan_capacity / 60, 1) if total_fan_capacity else 0
vent_pre_required = round(volume * 10, 1)
vent_run_required = round(volume * 20, 1)
vent_time_min = round(vent_pre_required / fan_per_min, 1) if fan_per_min else 0
vent_pre_ok = "만족" if fan_per_min * vent_time_min >= vent_pre_required else "불만족"
vent_run_ok = "만족" if total_fan_capacity >= vent_run_required else "불만족"

st.info(f"""
**📊 자동 계산 결과**
- 환기팬 총 용량: **{total_fan_capacity:,} ㎥/h** ({total_fan_count}대)
- 분당 환기능력: **{fan_per_min} ㎥/min**
- 작업 전 필요환기량(체적×10): **{vent_pre_required:,} ㎥**
- 작업 중 필요환기량(체적×20): **{vent_run_required:,} ㎥/hr**
- 필요 가동시간: **약 {vent_time_min} 분**
- 작업 전 판정: **{vent_pre_ok}** / 작업 중 판정: **{vent_run_ok}**
""")

vent_note = st.text_area("환기시설 특이사항", "")

# ================== 8. 워드파일 생성 ==================
st.header("8️⃣ 워드 파일 생성")

def bool_to_check(b):
    return "☑" if b else "☐"

def build_context():
    return {
        # 기본
        "project_name": project_name,
        "construction_period": construction_period,
        "confined_work_period": confined_work_period,
        "work_detail": work_detail,
        "write_date": write_date.strftime("%Y.%m.%d"),
        "company_name": company_name,
        "site_manager": site_manager,
        # 표(반복)
        "hazard_rows": hazard_df.to_dict(orient="records"),
        "blind_rows": blind_df.to_dict(orient="records"),
        "material_rows": material_df.to_dict(orient="records"),
        "owned_equipment": owned_df.to_dict(orient="records"),
        "rental_equipment": rental_df.to_dict(orient="records"),
        "emergency_contacts": emergency_df.to_dict(orient="records"),
        "workers": worker_df.to_dict(orient="records"),
        "fan_rows": fan_df.to_dict(orient="records"),
        # 교육
        "special_edu_date": special_edu_date,
        "special_edu_instructor": special_edu_instructor,
        "gas_edu_date": gas_edu_date,
        "gas_edu_instructor": gas_edu_instructor,
        # 체크박스
        "chk_crane": bool_to_check(rescue["이동식크레인"]),
        "chk_hoist": bool_to_check(rescue["호이스트"]),
        "chk_tripod": bool_to_check(rescue["삼각대"]),
        "chk_rope": bool_to_check(rescue["구명로프"]),
        "chk_stretcher": bool_to_check(rescue["들것"]),
        "chk_winch": bool_to_check(rescue["윈치"]),
        "rescue_etc": rescue_etc,
        "chk_airline": bool_to_check(breathing["송기마스크"]),
        "chk_elsa": bool_to_check(breathing["ELSA"]),
        "breathing_etc": breathing_etc,
        "chk_radio": bool_to_check(comm["무전기"]),
        "chk_speaker": bool_to_check(comm["확성기"]),
        "chk_whistle": bool_to_check(comm["호각"]),
        "comm_etc": comm_etc,
        "etc_equipment": etc_equipment,
        # 환기 자동계산
        "vent_process": vent_process,
        "volume": volume,
        "total_fan_capacity": total_fan_capacity,
        "total_fan_count": total_fan_count,
        "fan_per_min": fan_per_min,
        "vent_pre_required": vent_pre_required,
        "vent_run_required": vent_run_required,
        "vent_time_min": vent_time_min,
        "vent_pre_result": f"환기팬({total_fan_count}대) {vent_time_min}분 이상 가동 시 환기능력이 {vent_pre_required}㎥보다 크므로 {vent_pre_ok}",
        "vent_run_result": f"환기팬({total_fan_count}대) 환기능력이 {total_fan_capacity}㎥/h으로 20회 환기 기준량({vent_run_required}㎥)보다 크므로 {vent_run_ok}",
        "vent_note": vent_note,
    }

if st.button("📥 워드 파일 생성 및 다운로드", type="primary", use_container_width=True):
    try:
        doc = DocxTemplate("template.docx")
        doc.render(build_context())
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.success("✅ 파일이 생성되었습니다!")
        st.download_button(
            label="💾 다운로드",
            data=buffer,
            file_name=f"밀폐공간작업프로그램_{project_name}_{write_date.strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except FileNotFoundError:
        st.error("⚠️ `template.docx` 파일이 없습니다. 같은 폴더에 워드 템플릿을 넣어주세요.")
    except Exception as e:
        st.error(f"⚠️ 오류 발생: {e}")
