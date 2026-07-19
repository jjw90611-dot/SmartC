import streamlit as st
from pptx import Presentation
from io import BytesIO
from datetime import datetime

# --- [1] 페이지 설정 ---
st.set_page_config(page_title="포스코퓨처엠 양식 자동화", layout="wide")

# --- [2] 데이터베이스 (위험요인 및 시나리오) ---
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

# --- [3] PPT 텍스트 치환 함수 ---
def replace_text_in_ppt(ppt_file, replacements):
    prs = Presentation(ppt_file)
    
    # 슬라이드 -> 도형(텍스트박스) 및 표(Table) 순회
    for slide in prs.slides:
        for shape in slide.shapes:
            # 1. 일반 텍스트 박스인 경우
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in replacements.items():
                            if key in run.text:
                                run.text = run.text.replace(key, str(value))
            
            # 2. 표(Table)인 경우
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                for key, value in replacements.items():
                                    if key in run.text:
                                        run.text = run.text.replace(key, str(value))
                                        
    # 변경된 PPT를 메모리에 저장
    ppt_stream = BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- [4] 사이드바 입력 폼 ---
st.sidebar.header("📝 데이터 입력")

project_name = st.sidebar.text_input("공사명", "포)6기 CDQ Chamber 기둥연와 개선교체")
dept_name = st.sidebar.text_input("부서명", "플랜트공사그룹")
start_date = st.sidebar.date_input("작업 시작일")
end_date = st.sidebar.date_input("작업 종료일")
manager = st.sidebar.text_input("현장소장", "김국현")
space_name = st.sidebar.text_input("작업 장소명", "CDQ Chamber")
workers = st.sidebar.number_input("근로자수(명)", value=20)
task_type = st.sidebar.selectbox("작업 종류", list(HAZARD_DB.keys()))

# --- [5] 메인 화면 ---
st.title("📄 포스코퓨처엠 PPT 템플릿 자동 완성기")
st.markdown("""
회사 원본 PPT 파일에 `{{공사명}}`, `{{현장소장}}` 등의 예약어를 입력해 둔 템플릿을 업로드하세요.
좌측에 입력한 데이터가 원본 양식 그대로 쏙쏙 채워져서 완성된 PPT로 다운로드됩니다.
""")

# 템플릿 파일 업로드
uploaded_template = st.file_uploader("1단계: 템플릿 PPTX 파일 업로드", type=["pptx"])

if uploaded_template is not None:
    st.success("템플릿이 성공적으로 업로드되었습니다! 아래 버튼을 눌러 완성된 파일을 다운로드하세요.")
    
    # 치환할 데이터 딕셔너리 생성
    replacements = {
        "{{공사명}}": project_name,
        "{{부서명}}": dept_name,
        "{{시작일}}": start_date.strftime("%Y년 %m월 %d일"),
        "{{종료일}}": end_date.strftime("%Y년 %m월 %d일"),
        "{{작성일}}": datetime.now().strftime("%Y.%m.%d"),
        "{{현장소장}}": manager,
        "{{작업장소}}": space_name,
        "{{근로자수}}": workers,
        "{{작업종류}}": task_type,
        "{{비상상황}}": HAZARD_DB[task_type]["emergency_type"],
        "{{시나리오}}": HAZARD_DB[task_type]["scenario"]
    }
    
    # PPT 생성
    final_ppt = replace_text_in_ppt(uploaded_template, replacements)
    
    # 다운로드 버튼
    st.download_button(
        label="📥 2단계: 완성된 PPT 파일 다운로드",
        data=final_ppt,
        file_name=f"자동완성_{project_name}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
else:
    st.info("👆 먼저 템플릿 PPT 파일을 업로드해 주세요.")
