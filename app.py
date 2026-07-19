import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO
from datetime import datetime

# --- [1] 페이지 설정 ---
st.set_page_config(page_title="포스코퓨처엠 밀폐공간 서류 자동화", layout="wide")

# --- [2] 데이터베이스 ---
HAZARD_DB = {
    "내화물 해체/시공": {
        "emergency_type": "산소결핍 질식",
        "scenario": "산소결핍 질식 및 분진 흡입 환자 구조 훈련"
    },
    "용접/사상/절단 (화기작업)": {
        "emergency_type": "밀폐공간 화재",
        "scenario": "밀폐공간 내 화기작업 중 화재 발생 대응 훈련"
    }
}

# --- [3] PPT 서식 도우미 함수 (회사 양식 스타일 적용) ---
def format_cell(cell, text, bold=False, font_size=12, align=PP_ALIGN.CENTER):
    """표의 셀 텍스트와 정렬, 폰트를 회사 양식처럼 깔끔하게 맞추는 함수"""
    cell.text = str(text)
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = align
        for run in paragraph.runs:
            run.font.name = '맑은 고딕'
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor(0, 0, 0) # 검은색 글씨

def add_title(slide, text, top=0.5, size=24):
    """슬라이드 제목 추가 함수"""
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(top), Inches(9), Inches(1))
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.name = '맑은 고딕'
    p.font.size = Pt(size)
    p.font.bold = True
    return tb

# --- [4] PPT 생성 함수 (백지에서 표 직접 그리기) ---
def create_ppt_from_scratch(data):
    prs = Presentation()
    blank_layout = prs.slide_layouts[6] # 빈 슬라이드

    # ---------------------------------------------------------
    # [Slide 1] 표지
    # ---------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_layout)
    
    # 제목
    tb1 = slide1.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(1))
    p1 = tb1.text_frame.paragraphs[0]
    p1.text = "밀폐공간 작업 프로그램"
    p1.font.name = '맑은 고딕'
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.alignment = PP_ALIGN.CENTER

    # 표지 정보 표
    table1_shape = slide1.shapes.add_table(5, 2, Inches(1.5), Inches(3.5), Inches(7), Inches(2.5))
    table1 = table1_shape.table
    table1.columns[0].width = Inches(2.5)
    table1.columns[1].width = Inches(4.5)
    
    info_data = [
        ["공 사 명", data['project_name']],
        ["공사기간", f"{data['start_date']} ~ {data['end_date']}"],
        ["밀폐공간 작업기간", f"{data['start_date']} ~ {data['end_date']}"],
        ["작성일", data['today']],
        ["회사명", f"㈜포스코퓨처엠    현장소장: {data['manager']} (인)"]
    ]
    for r, row in enumerate(info_data):
        format_cell(table1.cell(r, 0), row[0], bold=True)
        format_cell(table1.cell(r, 1), row[1], align=PP_ALIGN.LEFT)

    # ---------------------------------------------------------
    # [Slide 2] 1. 사업장 내 밀폐공간의 위치 파악 및 관리방안
    # ---------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_layout)
    add_title(slide2, "1. 사업장 내 밀폐공간의 위치 파악 및 관리방안", top=0.5, size=20)
    add_title(slide2, "⊙ 사업장 내 밀폐공간 위치 파악", top=1.2, size=14)

    table2_shape = slide2.shapes.add_table(2, 6, Inches(0.5), Inches(2.0), Inches(9), Inches(1.0))
    table2 = table2_shape.table
    
    headers2 = ["연번", "공정명", "작업장소 명칭", "TYPE", "근로자수(명)", "비고"]
    widths2 = [0.8, 2.0, 2.5, 1.2, 1.5, 1.0]
    for i, w in enumerate(widths2): table2.columns[i].width = Inches(w)

    for i, h in enumerate(headers2): 
        format_cell(table2.cell(0, i), h, bold=True)
    
    row_data2 = ["1", data['task_type'], data['space_name'], "-", f"{data['workers']}명", "-"]
    for i, val in enumerate(row_data2):
        format_cell(table2.cell(1, i), val)

    # ---------------------------------------------------------
    # [Slide 3] 비상상황 교육 및 훈련 계획서
    # ---------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    add_title(slide3, "비상상황 교육 및 훈련 계획서", top=0.5, size=22)
    add_title(slide3, "[관련근거:(PCR-L-331)비상상황대비및대응규정]", top=1.0, size=12)
    
    add_title(slide3, f"1. 개요\n  - 부서명: {data['dept_name']}      - 작성일: {data['today']}", top=1.8, size=14)
    add_title(slide3, "2. 교육 및 훈련 계획\n  2.1 종류별 시나리오", top=2.8, size=14)

    table3_shape = slide3.shapes.add_table(2, 4, Inches(0.5), Inches(3.8), Inches(9), Inches(1.0))
    table3 = table3_shape.table
    
    headers3 = ["순번", "비상상황 종류", "훈련 시나리오", "담당자 주관"]
    widths3 = [1.0, 2.5, 4.0, 1.5]
    for i, w in enumerate(widths3): table3.columns[i].width = Inches(w)

    for i, h in enumerate(headers3): 
        format_cell(table3.cell(0, i), h, bold=True)
    
    row_data3 = ["1", data['emergency_type'], data['scenario'], f"{data['manager']} (합동)"]
    for i, val in enumerate(row_data3):
        format_cell(table3.cell(1, i), val)

    # ---------------------------------------------------------
    # [Slide 4] 비상상황 교육 및 훈련 실시 결과서
    # ---------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_layout)
    add_title(slide4, "비상상황 교육 및 훈련 실시 결과서", top=0.5, size=22)
    add_title(slide4, "[관련근거:(PCR-L-331)비상상황대비및대응규정]", top=1.0, size=12)
    add_title(slide4, "1. 개요", top=1.8, size=14)

    table4_shape = slide4.shapes.add_table(5, 2, Inches(0.5), Inches(2.3), Inches(9), Inches(2.5))
    table4 = table4_shape.table
    table4.columns[0].width = Inches(2.5)
    table4.columns[1].width = Inches(6.5)
    
    res_data = [
        ["비상훈련명", f"{data['space_name']} 내부 밀폐 작업시 {data['emergency_type']} 대응훈련"],
        ["훈련장소(공정)", data['space_name']],
        ["상황 및 원인", data['scenario']],
        ["훈련 일시", f"{data['start_date']} 13:30 ~ 14:10 (40분간)"],
        ["훈련 주관자", f"{data['manager']} (현장소장)"]
    ]
    for r, row in enumerate(res_data):
        format_cell(table4.cell(r, 0), row[0], bold=True)
        format_cell(table4.cell(r, 1), row[1], align=PP_ALIGN.LEFT)

    add_title(slide4, "3. 총평\n  - 비상상황에 신속한 대피와 복구가 이루어졌음.\n  - 밀폐공간작업 프로그램을 잘 수행해서 안전한 작업장이 이루어져야 겠습니다.", top=5.2, size=14)

    # 메모리 버퍼에 PPT 저장
    ppt_stream = BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- [5] 사이드바 입력 폼 ---
st.sidebar.header("📝 데이터 입력")

project_name = st.sidebar.text_input("공사명", "포)6기 CDQ Chamber 기둥연와 개선교체")
dept_name = st.sidebar.text_input("부서명", "플랜트공사그룹")
start_date = st.sidebar.date_input("작업 시작일")
end_date = st.sidebar.date_input("작업 종료일")
manager = st.sidebar.text_input("현장소장", "김국현")
space_name = st.sidebar.text_input("작업 장소명", "CDQ Chamber")
workers = st.sidebar.number_input("근로자수(명)", value=20)
task_type = st.sidebar.selectbox("작업 종류", list(HAZARD_DB.keys()))

# --- [6] 메인 화면 ---
st.title("📄 포스코퓨처엠 밀폐공간 서류 자동 생성기 (보안용)")
st.info("사내 보안 정책을 고려하여, 외부 파일 업로드 없이 파이썬이 직접 회사 양식의 표와 서식을 그려서 PPT를 생성합니다.")

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

# PPT 생성 및 다운로드 버튼
ppt_file = create_ppt_from_scratch(doc_data)

st.download_button(
    label="📥 회사 양식 PPT 다운로드 (파일 업로드 불필요)",
    data=ppt_file,
    file_name=f"밀폐공간_작업프로그램_{project_name}.pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    use_container_width=True
)
