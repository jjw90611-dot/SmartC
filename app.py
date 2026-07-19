import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from io import BytesIO
from datetime import datetime

# --- [1] 페이지 설정 ---
st.set_page_config(page_title="포스코퓨처엠 밀폐공간 & 비상훈련 자동화", layout="wide")

# --- [2] 데이터베이스 (위험요인 및 시나리오) ---
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

# --- [3] PPT 생성 함수 (원본 양식 100% 반영) ---
def create_ppt(data):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6] # 빈 슬라이드 레이아웃

    # ---------------------------------------------------------
    # [Slide 1] 밀폐공간 작업 프로그램 (표지)
    # ---------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_slide_layout)
    tb1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    p1 = tb1.text_frame.paragraphs[0]
    p1.text = "밀폐공간 작업 프로그램"
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.alignment = PP_ALIGN.CENTER

    table1_shape = slide1.shapes.add_table(5, 2, Inches(1.5), Inches(2.5), Inches(7), Inches(2.5))
    table1 = table1_shape.table
    table1.columns[0].width = Inches(2)
    table1.columns[1].width = Inches(5)
    
    info_data = [
        ["공 사 명", data['project_name']],
        ["공사기간", f"{data['start_date']} ~ {data['end_date']}"],
        ["밀폐공간 작업기간", f"{data['start_date']} ~ {data['end_date']}"],
        ["작성일", data['today']],
        ["회사명", f"㈜포스코퓨처엠    현장소장: {data['manager']} (인)"]
    ]
    for r, row in enumerate(info_data):
        table1.cell(r, 0).text = row[0]
        table1.cell(r, 1).text = row[1]

    # ---------------------------------------------------------
    # [Slide 2] 밀폐공간 작업 프로그램 (목차)
    # ---------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_slide_layout)
    tb2 = slide2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(6))
    tf2 = tb2.text_frame
    tf2.text = "목 차"
    tf2.paragraphs[0].font.size = Pt(24)
    tf2.paragraphs[0].font.bold = True
    
    contents = [
        "1. 사업장내 밀폐공간의 위치 파악 및 관리방안",
        "2. 밀폐공간내 질식·중독 등을 일으킬 수 있는 유해∙위험요인의 파악 및 관리 방안",
        "3. 밀폐공간작업 시 사전 확인이 필요한 사항에 대한 확인 절차",
        "4. 안전보건교육 및 훈련",
        "5. 그 밖에 밀폐공간 작업 근로자의 건강장해 예방에 관한 사항",
        "6. 작업 일시, 기간, 장소 및 내용 등 작업 정보",
        "7. 관리감독자, 근로자, 감시인 등 작업자 정보",
        "8. 산소 및 유해가스 농도의 측정결과 및 후속조치 사항",
        "9. 작업 중 불활성가스 또는 유해가스의 누출∙유입∙ 발생 가능성 검토 및 후속조치 사항",
        "10. 작업 시 착용하여야 할 보호구의 종류",
        "11. 비상연락체계",
        "12. 프로그램의 평가"
    ]
    for content in contents:
        p = tf2.add_paragraph()
        p.text = content
        p.font.size = Pt(14)

    # ---------------------------------------------------------
    # [Slide 3] 1. 사업장 내 밀폐공간의 위치 파악 및 관리방안
    # ---------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_slide_layout)
    tb3 = slide3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    tb3.text_frame.text = "1. 사업장 내 밀폐공간의 위치 파악 및 관리방안\n⊙ 사업장 내 밀폐공간 위치 파악"
    tb3.text_frame.paragraphs[0].font.size = Pt(20)
    tb3.text_frame.paragraphs[0].font.bold = True

    table3_shape = slide3.shapes.add_table(3, 6, Inches(0.5), Inches(1.5), Inches(9), Inches(1.5))
    table3 = table3_shape.table
    headers3 = ["연번", "공정명", "작업장소 명칭", "TYPE", "근로자수(명)", "비고"]
    for i, h in enumerate(headers3): table3.cell(0, i).text = h
    
    table3.cell(1, 0).text = "1"
    table3.cell(1, 1).text = data['task_type']
    table3.cell(1, 2).text = data['space_name']
    table3.cell(1, 3).text = "-"
    table3.cell(1, 4).text = str(data['workers'])
    table3.cell(1, 5).text = "-"

    # ---------------------------------------------------------
    # [Slide 4] 비상상황 교육 및 훈련 계획서
    # ---------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_slide_layout)
    tb4 = slide4.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    tb4.text_frame.text = "비상상황 교육 및 훈련 계획서\n[관련근거:(PCR-L-331)비상상황대비및대응규정]"
    tb4.text_frame.paragraphs[0].font.size = Pt(20)
    tb4.text_frame.paragraphs[0].font.bold = True

    tb4_sub = slide4.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(0.5))
    tb4_sub.text_frame.text = f"1. 개요\n부서명: {data['dept_name']}      작성일: {data['today']}"
    
    tb4_sub2 = slide4.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(0.5))
    tb4_sub2.text_frame.text = "2. 교육 및 훈련 계획\n2.1 종류별 시나리오"

    table4_shape = slide4.shapes.add_table(2, 4, Inches(0.5), Inches(3.5), Inches(9), Inches(1))
    table4 = table4_shape.table
    headers4 = ["순번", "비상상황 종류", "훈련 시나리오", "담당자 주관"]
    for i, h in enumerate(headers4): table4.cell(0, i).text = h
    
    table4.cell(1, 0).text = "1"
    table4.cell(1, 1).text = data['emergency_type']
    table4.cell(1, 2).text = data['scenario']
    table4.cell(1, 3).text = f"{data['manager']} (합동)"

    # ---------------------------------------------------------
    # [Slide 5] 비상상황 교육 및 훈련 실시 결과서
    # ---------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_slide_layout)
    tb5 = slide5.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    tb5.text_frame.text = "비상상황 교육 및 훈련 실시 결과서\n[관련근거:(PCR-L-331)비상상황대비및대응규정]"
    tb5.text_frame.paragraphs[0].font.size = Pt(20)
    tb5.text_frame.paragraphs[0].font.bold = True

    tb5_sub = slide5.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(0.5))
    tb5_sub.text_frame.text = "1. 개요"

    table5_shape = slide5.shapes.add_table(5, 2, Inches(0.5), Inches(2.0), Inches(9), Inches(2.5))
    table5 = table5_shape.table
    table5.columns[0].width = Inches(2)
    table5.columns[1].width = Inches(7)
    
    res_data = [
        ["비상훈련명", f"{data['space_name']} 내부 밀폐 작업시 {data['emergency_type']} 대응훈련"],
        ["훈련장소(공정)", data['space_name']],
        ["상황 및 원인", data['scenario']],
        ["훈련 일시", f"{data['start_date']} 13:30 ~ 14:10 (40분간)"],
        ["훈련 주관자", f"{data['manager']} (현장소장)"]
    ]
    for r, row in enumerate(res_data):
        table5.cell(r, 0).text = row[0]
        table5.cell(r, 1).text = row[1]

    tb5_sub2 = slide5.shapes.add_textbox(Inches(0.5), Inches(5.0), Inches(9), Inches(1.5))
    tb5_sub2.text_frame.text = "3. 총평\n비상상황에 신속한 대피와 복구가 이루어졌음.\n밀폐공간작업 프로그램을 잘 수행해서 안전한 작업장이 이루어져야 겠습니다."

    # 메모리 버퍼에 PPT 저장
    ppt_stream = BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- [4] 사이드바 입력 폼 ---
st.sidebar.header("📝 밀폐공간 & 비상훈련 입력")

project_name = st.sidebar.text_input("공사명", "포)6기 CDQ Chamber 기둥연와 개선교체")
dept_name = st.sidebar.text_input("부서명", "플랜트공사그룹")
start_date = st.sidebar.date_input("작업 시작일")
end_date = st.sidebar.date_input("작업 종료일")
manager = st.sidebar.text_input("현장소장", "김국현")
space_name = st.sidebar.text_input("작업 장소명", "CDQ Chamber")
workers = st.sidebar.number_input("근로자수(명)", value=20)
task_type = st.sidebar.selectbox("작업 종류", list(HAZARD_DB.keys()))

# --- [5] 메인 화면 (원본 양식 미리보기) ---
st.title("📄 포스코퓨처엠 양식 자동 생성기")
st.info("좌측에 값을 입력하면 아래 원본 양식에 자동으로 채워지며, PPT로 다운로드할 수 있습니다.")

# 데이터 취합
doc_data = {
    "project_name": project_name,
    "dept_name": dept_name,
    "start_date": start_date.strftime("%Y년 %m월 %d일"),
    "end_date": end_date.strftime("%Y년 %m월 %d일"),
    "today": datetime.now().strftime("%Y.%m.%d"),
    "manager": manager,
    "space_name": space_name,
    "workers": workers,
    "task_type": task_type,
    "emergency_type": HAZARD_DB[task_type]["emergency_type"],
    "scenario": HAZARD_DB[task_type]["scenario"]
}

# --- 화면 미리보기 (원본 양식 표 형태 유지) ---
st.divider()
st.subheader("📑 밀폐공간 작업 프로그램 (표지)")
df_cover = pd.DataFrame([
    ["공 사 명", doc_data['project_name']],
    ["공사기간", f"{doc_data['start_date']} ~ {doc_data['end_date']}"],
    ["밀폐공간 작업기간", f"{doc_data['start_date']} ~ {doc_data['end_date']}"],
    ["작성일", doc_data['today']],
    ["회사명", f"㈜포스코퓨처엠    현장소장: {doc_data['manager']} (인)"]
])
st.table(df_cover)

st.subheader("📑 1. 사업장 내 밀폐공간의 위치 파악 및 관리방안")
st.markdown("**⊙ 사업장 내 밀폐공간 위치 파악**")
df_location = pd.DataFrame({
    "연번": ["1"],
    "공정명": [doc_data['task_type']],
    "작업장소 명칭": [doc_data['space_name']],
    "TYPE": ["-"],
    "근로자수(명)": [f"{doc_data['workers']}명"],
    "비고": ["-"]
})
st.table(df_location)

st.subheader("📑 비상상황 교육 및 훈련 계획서")
st.caption("[관련근거:(PCR-L-331)비상상황대비및대응규정]")
st.markdown(f"**1. 개요** | 부서명: {doc_data['dept_name']} | 작성일: {doc_data['today']}")
st.markdown("**2. 교육 및 훈련 계획 (2.1 종류별 시나리오)**")
df_plan = pd.DataFrame({
    "순번": ["1"],
    "비상상황 종류": [doc_data['emergency_type']],
    "훈련 시나리오": [doc_data['scenario']],
    "담당자 주관": [f"{doc_data['manager']} (합동)"]
})
st.table(df_plan)

st.subheader("📑 비상상황 교육 및 훈련 실시 결과서")
st.caption("[관련근거:(PCR-L-331)비상상황대비및대응규정]")
df_result = pd.DataFrame([
    ["비상훈련명", f"{doc_data['space_name']} 내부 밀폐 작업시 {doc_data['emergency_type']} 대응훈련"],
    ["훈련장소(공정)", doc_data['space_name']],
    ["상황 및 원인", doc_data['scenario']],
    ["훈련 일시", f"{doc_data['start_date']} 13:30 ~ 14:10 (40분간)"],
    ["훈련 주관자", f"{doc_data['manager']} (현장소장)"]
])
st.table(df_result)

st.divider()

# PPT 생성 및 다운로드 버튼
ppt_file = create_ppt(doc_data)

st.download_button(
    label="📥 PPT 파일로 출력 및 다운로드 (원본 양식 적용)",
    data=ppt_file,
    file_name=f"밀폐공간_비상훈련_{project_name}.pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    use_container_width=True
)
