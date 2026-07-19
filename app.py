import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from io import BytesIO
from datetime import datetime

# --- [1] 페이지 설정 ---
st.set_page_config(page_title="포스코퓨처엠 밀폐공간 & 비상훈련 자동화", layout="wide")

# --- [2] 데이터베이스 ---
HAZARD_DB = {
    "내화물 해체/시공": {
        "hazard": "건조 재료 분진 흡입, 피부/눈 접촉, 산소결핍",
        "solution": "작업 전/중 환기, 방진마스크 착용, 세안설비 위치 안내",
        "emergency_type": "산소결핍 질식",
        "scenario": "산소결핍 질식 및 분진 흡입 환자 구조 훈련"
    },
    "용접/사상/절단 (화기작업)": {
        "hazard": "용접 흄 발생, 일산화탄소(CO) 중독, 불티로 화재/폭발",
        "solution": "불티비산방지포 설치, 소화기 비치, 고정식 가스검지기 운용",
        "emergency_type": "밀폐공간 화재",
        "scenario": "밀폐공간 내 화기작업 중 화재 발생 대응 훈련"
    }
}

# --- [3] PPT 생성 함수 (회사 양식 유지) ---
def create_ppt(data):
    prs = Presentation()
    
    # 슬라이드 1: 밀폐공간 작업 프로그램 (표지)
    slide1 = prs.slides.add_slide(prs.slide_layouts[6]) # 빈 슬라이드
    tb1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    tb1.text_frame.text = "밀폐공간 작업 프로그램"
    tb1.text_frame.paragraphs[0].font.size = Pt(32)
    tb1.text_frame.paragraphs[0].font.bold = True
    tb1.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # 표지 정보 테이블
    table_shape = slide1.shapes.add_table(5, 2, Inches(1.5), Inches(2.5), Inches(7), Inches(2.5))
    table = table_shape.table
    table.columns[0].width = Inches(2)
    table.columns[1].width = Inches(5)
    
    info_data = [
        ["공 사 명", data['project_name']],
        ["공사기간", f"{data['start_date']} ~ {data['end_date']}"],
        ["밀폐공간 작업기간", f"{data['start_date']} ~ {data['end_date']}"],
        ["작성일", datetime.now().strftime('%Y.%m.%d')],
        ["회사명", f"㈜포스코퓨처엠    현장소장: {data['manager']} (인)"]
    ]
    for row_idx, row_data in enumerate(info_data):
        table.cell(row_idx, 0).text = row_data[0]
        table.cell(row_idx, 1).text = row_data[1]

    # 슬라이드 2: 비상상황 교육 및 훈련 계획서
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    tb2 = slide2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    tb2.text_frame.text = "비상상황 교육 및 훈련 계획서\n[관련근거:(PCR-L-331)비상상황대비및대응규정]"
    tb2.text_frame.paragraphs[0].font.size = Pt(24)
    tb2.text_frame.paragraphs[0].font.bold = True

    tb2_sub = slide2.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(0.5))
    tb2_sub.text_frame.text = f"1. 개요\n부서명: {data['dept_name']}      작성일: {datetime.now().strftime('%Y.%m.%d')}"
    
    tb2_sub2 = slide2.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(0.5))
    tb2_sub2.text_frame.text = "2. 교육 및 훈련 계획 (2.1 종류별 시나리오)"

    # 훈련 계획 테이블
    table2_shape = slide2.shapes.add_table(2, 4, Inches(0.5), Inches(3.2), Inches(9), Inches(1))
    table2 = table2_shape.table
    headers = ["순번", "비상상황 종류", "훈련 시나리오", "담당자 주관"]
    for col_idx, header in enumerate(headers):
        table2.cell(0, col_idx).text = header
    
    table2.cell(1, 0).text = "1"
    table2.cell(1, 1).text = data['emergency_type']
    table2.cell(1, 2).text = data['scenario']
    table2.cell(1, 3).text = f"{data['manager']} (합동)"

    # 슬라이드 3: 비상상황 교육 및 훈련 실시 결과서
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    tb3 = slide3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    tb3.text_frame.text = "비상상황 교육 및 훈련 실시 결과서"
    tb3.text_frame.paragraphs[0].font.size = Pt(24)
    tb3.text_frame.paragraphs[0].font.bold = True

    # 결과서 테이블
    table3_shape = slide3.shapes.add_table(5, 2, Inches(0.5), Inches(1.5), Inches(9), Inches(2.5))
    table3 = table3_shape.table
    table3.columns[0].width = Inches(2)
    table3.columns[1].width = Inches(7)
    
    res_data = [
        ["비상훈련명", f"{data['space_name']} 내부 밀폐 작업시 {data['emergency_type']} 대응훈련"],
        ["훈련장소(공정)", data['space_name']],
        ["상황 및 원인", data['scenario']],
        ["훈련 일시", f"{data['start_date']} 13:30 ~ 14:10 (40분간)"],
        ["훈련 주관자", f"{data['manager']} (현장소장)"]
    ]
    for row_idx, row_data in enumerate(res_data):
        table3.cell(row_idx, 0).text = row_data[0]
        table3.cell(row_idx, 1).text = row_data[1]

    tb3_sub = slide3.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(9), Inches(2))
    tb3_sub.text_frame.text = "3. 총평\n비상상황에 신속한 대피와 복구가 이루어졌음.\n밀폐공간작업 프로그램을 잘 수행해서 안전한 작업장이 이루어져야 겠습니다."

    # 메모리 버퍼에 PPT 저장
    ppt_stream = BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- [4] 사이드바 입력 폼 ---
st.sidebar.header("📝 밀폐공간 & 비상훈련 입력")

project_name = st.sidebar.text_input("공사명", "6기 CDQ Chamber 기둥연와 개선교체")
dept_name = st.sidebar.text_input("부서명", "플랜트공사그룹")
start_date = st.sidebar.date_input("작업 시작일")
end_date = st.sidebar.date_input("작업 종료일")
manager = st.sidebar.text_input("현장소장", "김국현")
space_name = st.sidebar.text_input("작업 장소명", "CDQ Chamber")
task_type = st.sidebar.selectbox("작업 종류", list(HAZARD_DB.keys()))

# --- [5] 메인 화면 (미리보기 및 PPT 다운로드) ---
st.title("📄 포스코퓨처엠 양식 자동 생성기")
st.info("좌측에 값을 입력하면 아래 양식이 자동으로 채워지며, PPT로 다운로드할 수 있습니다.")

# 데이터 취합
doc_data = {
    "project_name": project_name,
    "dept_name": dept_name,
    "start_date": start_date.strftime("%Y년 %m월 %d일"),
    "end_date": end_date.strftime("%Y년 %m월 %d일"),
    "manager": manager,
    "space_name": space_name,
    "emergency_type": HAZARD_DB[task_type]["emergency_type"],
    "scenario": HAZARD_DB[task_type]["scenario"]
}

# 화면 미리보기 (Streamlit UI)
st.divider()
st.subheader("1. 밀폐공간 작업 프로그램 (표지 미리보기)")
st.markdown(f"""
* **공 사 명:** {doc_data['project_name']}
* **공사기간:** {doc_data['start_date']} ~ {doc_data['end_date']}
* **회사명:** ㈜포스코퓨처엠
* **현장소장:** {doc_data['manager']} (인)
""")

st.subheader("2. 비상상황 교육 및 훈련 계획서 (미리보기)")
st.markdown(f"""
* **부서명:** {doc_data['dept_name']}
* **비상상황 종류:** {doc_data['emergency_type']}
* **훈련 시나리오:** {doc_data['scenario']}
* **담당자 주관:** {doc_data['manager']}
""")

st.subheader("3. 비상상황 교육 및 훈련 실시 결과서 (미리보기)")
st.markdown(f"""
* **비상훈련명:** {doc_data['space_name']} 내부 밀폐 작업시 {doc_data['emergency_type']} 대응훈련
* **훈련장소:** {doc_data['space_name']}
* **상황 및 원인:** {doc_data['scenario']}
""")

st.divider()

# PPT 생성 및 다운로드 버튼
ppt_file = create_ppt(doc_data)

st.download_button(
    label="📥 PPT 파일로 출력 및 다운로드",
    data=ppt_file,
    file_name=f"밀폐공간_비상훈련_{project_name}.pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    use_container_width=True
)
