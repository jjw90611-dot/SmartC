import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
from datetime import datetime
import math

st.set_page_config(page_title="포스코퓨처엠 밀폐공간 작업 프로그램 자동 작성기", layout="wide")

# =========================================================
# 서식 헬퍼
# =========================================================
def set_cell_bg(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)

def set_cell_border(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), '000000')
        tc_borders.append(border)
    tc_pr.append(tc_borders)

def set_font(run, name='맑은 고딕', size=10, bold=False, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:ascii'), name)
    run._element.rPr.rFonts.set(qn('w:hAnsi'), name)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def write_cell(cell, text, size=10, bold=False, align='center', v_align='center', bg=None, color=None):
    cell.text = ''
    if bg:
        set_cell_bg(cell, bg)
    set_cell_border(cell)
    if v_align == 'center':
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    elif v_align == 'top':
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    lines = str(text).split('\n')
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        if align == 'center':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == 'left':
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif align == 'right':
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(line)
        set_font(run, size=size, bold=bold, color=color)

def add_para(doc, text, size=10, bold=False, align='left', space_after=6, space_before=0, color=None):
    p = doc.add_paragraph()
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(text)
    set_font(run, size=size, bold=bold, color=color)
    return p

def add_header_marker(doc):
    add_para(doc, '포스코퓨처엠 : Company Use Only', size=9, align='right', space_after=8, color=(100, 100, 100))

def add_section_title(doc, text, size=16):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    set_font(run, size=size, bold=True, color=(0, 59, 112))  # 포스코 네이비
    return p

def set_col_widths(table, widths_cm):
    for row in table.rows:
        for i, width in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(width)

def set_row_height(row, height_cm):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight.set(qn('w:hRule'), 'atLeast')
    trPr.append(trHeight)

HEADER_BG = 'D9E1F2'
SUBHEADER_BG = 'F2F2F2'
YELLOW_BG = 'FFF2CC'
ORANGE_BG = 'FCE5CD'
RED_BG = 'F4CCCC'
GRAY_BG = 'F3F3F3'

# =========================================================
# 체감온도 계산 (Rothfusz Heat Index, 섭씨 변환식)
# =========================================================
def calc_heat_index(t_c, rh):
    """섭씨 기온과 습도(%)로 체감온도(℃) 산출"""
    if t_c is None or rh is None:
        return None
    try:
        # Steadman/Rothfusz 공식 근사 (JS 코드와 동일 로직)
        w = t_c * math.atan(0.151977 * math.sqrt(rh + 8.313659)) \
            + math.atan(t_c + rh) - math.atan(rh - 1.676331) \
            + 0.00391838 * (rh ** 1.5) * math.atan(0.023101 * rh) - 4.686035
        a = -0.2442 + 0.55399 * w + 0.45535 * t_c \
            - 0.0022 * (w ** 2) + 0.00278 * w * t_c + 3.0
        return round(a, 1)
    except Exception:
        return None

def heat_stage(hi):
    if hi is None:
        return "정상 작업"
    if hi >= 38:
        return "폭염중대경보 (38℃ 이상)"
    elif hi >= 35:
        return "폭염경보 (35℃ 이상)"
    elif hi >= 31:
        return "폭염주의보 (31℃ 이상)"
    return "정상 작업"

# =========================================================
# 체적/환기량 계산
# =========================================================
def calc_volume(shape, W, L, H, D):
    rect = W * L * H if shape in ('rect', 'combined') else 0
    cyl = (math.pi * (D ** 2) / 4) * H if shape in ('cyl', 'combined') else 0
    return rect + cyl

# =========================================================
# 메인 문서 생성
# =========================================================
def create_word_document(data):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)

    style = doc.styles['Normal']
    style.font.name = '맑은 고딕'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    style._element.rPr.rFonts.set(qn('w:ascii'), '맑은 고딕')
    style._element.rPr.rFonts.set(qn('w:hAnsi'), '맑은 고딕')
    style.font.size = Pt(10)

    # ================= Page 1 : 표지 =================
    add_header_marker(doc)
    for _ in range(4):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('밀폐공간 작업\n프로그램')
    set_font(run, size=40, bold=True, color=(0, 59, 112))

    for _ in range(3):
        doc.add_paragraph()

    add_para(doc,
        '스마트 밀폐공간 안전설계 및 온열질환 예방 프로그램',
        size=14, bold=True, align='center', space_after=20,
        color=(0, 91, 172))

    for label, value in [
        ('공사명', data['project_name']),
        ('공사기간', f"{data['start_date']} ~ {data['end_date']}"),
        ('밀폐공간작업기간', f"{data['start_date']} ~ {data['end_date']}"),
        ('작업내용', data['work_detail']),
    ]:
        add_para(doc, f" ‑ {label} : {value}", size=14, space_after=8)

    for _ in range(2):
        doc.add_paragraph()

    add_para(doc, f"작성일 : {data['today']}", size=13, space_after=6)
    add_para(doc, f"회사명 : {data['company_name']}          현장소장 : {data['manager']} (서명)", size=13)
    doc.add_page_break()

    # ================= Page 2 : 목차 =================
    add_header_marker(doc)
    add_section_title(doc, '목차', size=20)
    doc.add_paragraph()
    toc = [
        "1. 밀폐공간의 위치 파악 및 관리방안",
        "2. 밀폐공간 내 질식∙중독 등을 일으킬 수 있는 유해∙위험요인의 파악 및 관리 방안",
        "3. 안전보건교육 및 훈련",
        "4. 그 밖에 밀폐공간 작업 근로자의 건강장해 예방에 관한 사항",
        "5. 프로그램 평가",
        "6. 스마트 밀폐공간 환기량 계산 (Smart Ventilation Design)",
        "7. 하절기 온열질환 예방 프로그램 (Heat Stress Prevention)",
        "",
        "첨부: 밀폐공간 작업 시 사전 확인이 필요한 사항에 대한 확인 절차",
    ]
    for item in toc:
        add_para(doc, item, size=13, space_after=10)
    doc.add_page_break()

    # ================= Page 3 : 1. 위치 및 관리방안 =================
    add_header_marker(doc)
    add_section_title(doc, '1. 밀폐공간의 위치 및 유해∙위험요인의 관리방안', size=14)

    add_para(doc, '(1) 사업장 내 밀폐공간 위치 및 위험물질 관리 현황', size=11, bold=True, space_after=6)
    t = doc.add_table(rows=3, cols=5)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    write_cell(t.cell(0, 0), '작업장소', bold=True, bg=HEADER_BG)
    write_cell(t.cell(0, 1), '공정명', bold=True, bg=HEADER_BG)
    write_cell(t.cell(0, 2), '질식·중독 발생 위험\n유해위험 물질', bold=True, bg=HEADER_BG)
    write_cell(t.cell(0, 3), '관리방안', bold=True, bg=HEADER_BG)
    write_cell(t.cell(0, 4), '', bold=True, bg=HEADER_BG)
    t.cell(0, 3).merge(t.cell(0, 4))
    write_cell(t.cell(1, 0), '', bg=HEADER_BG)
    write_cell(t.cell(1, 1), '', bg=HEADER_BG)
    write_cell(t.cell(1, 2), '', bg=HEADER_BG)
    write_cell(t.cell(1, 3), '입조 전', bold=True, bg=HEADER_BG)
    write_cell(t.cell(1, 4), '입조 중', bold=True, bg=HEADER_BG)
    t.cell(0, 0).merge(t.cell(1, 0))
    t.cell(0, 1).merge(t.cell(1, 1))
    t.cell(0, 2).merge(t.cell(1, 2))
    write_cell(t.cell(2, 0), data['space_name'])
    write_cell(t.cell(2, 1), data['process_name'])
    write_cell(t.cell(2, 2), data['hazard_material'])
    write_cell(t.cell(2, 3), '환기 및 가스측정', align='left')
    write_cell(t.cell(2, 4), '지속 환기 / 감시인 배치', align='left')
    set_col_widths(t, [2.5, 2.5, 3.5, 3.5, 3.5])
    for row in t.rows:
        set_row_height(row, 0.9)

    add_para(doc, '', size=6)
    add_para(doc, '* 위치도', size=10, space_after=2)
    add_para(doc, '(위치도 이미지 삽입)', size=9, color=(120, 120, 120), space_after=12)

    add_para(doc, '(2) Blind 설치 개수 및 위치', size=11, bold=True, space_after=6)
    t2 = doc.add_table(rows=2, cols=3)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['공정명', 'Blind 설치 개수 (EA)', '확인자 서명']):
        write_cell(t2.cell(0, i), h, bold=True, bg=HEADER_BG)
    write_cell(t2.cell(1, 0), data['process_name'])
    write_cell(t2.cell(1, 1), str(data['blind_count']))
    write_cell(t2.cell(1, 2), data['manager'])
    set_col_widths(t2, [5, 5, 5])
    for row in t2.rows:
        set_row_height(row, 0.9)

    add_para(doc, '* Blind 설치 도면 및 List 작업허가서 첨부 확인', size=9, space_before=6)
    doc.add_page_break()

    # ================= Page 4 : 2. 유해∙위험요인 =================
    add_header_marker(doc)
    add_section_title(doc, '2. 밀폐공간 내 질식∙중독 등을 일으킬 수 있는\n     유해∙위험요인의 파악 및 관리방안', size=13)

    t3 = doc.add_table(rows=10, cols=3)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['물질명', '유해위험요인', '대책']):
        write_cell(t3.cell(0, i), h, bold=True, bg=HEADER_BG)

    write_cell(t3.cell(1, 0), '예시)\n산화니켈', bold=True)
    write_cell(t3.cell(1, 1), '질식')
    write_cell(t3.cell(1, 2),
        '‑ MSDS교육 실시\n'
        '‑ 감시인, 예비감시인 및 입조작업자는 특별안전교육\n'
        '  (N2 분위기 위험요소, Air Line Mask 사용법 등) 사전 이수\n'
        '‑ 감시인과 입조 작업자간 무전기/스키퍼 구비\n'
        '‑ 감시인 자리이탈 금지 및 출입 인원 점검표 실시간 작성',
        size=9, align='left', v_align='top')
    write_cell(t3.cell(2, 1), '눈에 접촉')
    write_cell(t3.cell(2, 2),
        '‑ 많은 양의 물을 사용하여 15분 이상 동안 눈을 씻어내고\n  즉시 의사의 진료를 받을 것',
        size=9, align='left', v_align='top')
    write_cell(t3.cell(3, 1), '피부 접촉')
    write_cell(t3.cell(3, 2),
        '‑ 오염된 의복 및 신발을 제거한 후 15분 동안 비누와 물로\n  씻어내고 필요시 의사의 진료를 받을 것',
        size=9, align='left', v_align='top')
    write_cell(t3.cell(4, 1), '흡입')
    write_cell(t3.cell(4, 2),
        '‑ 구토를 할 경우 구토물이 기도를 막는 것을 방지하기 위해\n  머리를 둔부보다 낮추도록 할 것\n‑ 즉시 의사의 진료를 받을 것',
        size=9, align='left', v_align='top')
    t3.cell(1, 0).merge(t3.cell(4, 0))

    write_cell(t3.cell(5, 0), '예시)\nN2', bold=True)
    write_cell(t3.cell(5, 1), '질식')
    write_cell(t3.cell(5, 2),
        '‑ MSDS교육 실시\n'
        '‑ 감시인, 예비감시인 및 입조작업자는 특별안전교육\n'
        '  (N2 분위기 위험요소, Air Line Mask 사용법 등) 사전 이수\n'
        '‑ 감시인과 입조 작업자간 무전기/스키퍼 구비\n'
        '‑ 감시인 자리이탈 금지 및 출입 인원 점검표 실시간 작성',
        size=9, align='left', v_align='top')
    write_cell(t3.cell(6, 1), '흡입 시')
    write_cell(t3.cell(6, 2),
        '‑ 호흡이 없으면 인공호흡 실시\n‑ 노출로 인한 영향이 나타나면 비오염지역으로 옮길 것\n‑ 즉시 의사의 진찰과 치료를 받을 것',
        size=9, align='left', v_align='top')
    write_cell(t3.cell(7, 1), '피부 접촉 시')
    write_cell(t3.cell(7, 2),
        '‑ 화학물질의 피부 접촉 즉시 의사의 진찰과 치료를 받을 것\n‑ 해당물질에 접촉시 동상을 입을 가능성이 있으니 주의할 것',
        size=9, align='left', v_align='top')
    write_cell(t3.cell(8, 1), '눈접촉 시')
    write_cell(t3.cell(8, 2),
        '‑ 눈에 들어간 경우 눈꺼풀을 들어올려 15분 동안 물로 충분히\n  씻어낼 것\n‑ 눈에 화학물질이 들어간 경우 즉시 의사의 진찰과 진료를\n  받을 것',
        size=9, align='left', v_align='top')
    write_cell(t3.cell(9, 1), '섭취 시')
    write_cell(t3.cell(9, 2), '많은 양의 화학물질을 섭취한 경우 의사의 진찰을 받을 것',
        size=9, align='left', v_align='top')
    t3.cell(5, 0).merge(t3.cell(9, 0))
    set_col_widths(t3, [2.5, 3, 10])
    doc.add_page_break()

    # ================= Page 5 : 3. 안전보건교육 =================
    add_header_marker(doc)
    add_section_title(doc, '3. 안전보건교육 및 훈련', size=14)
    t4 = doc.add_table(rows=3, cols=2)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    write_cell(t4.cell(0, 0), '특별안전보건교육내용\n(산업안전보건법 시행규칙 별표5)', bold=True, bg=HEADER_BG)
    write_cell(t4.cell(0, 1), '교육일정 및 강사', bold=True, bg=HEADER_BG)
    write_cell(t4.cell(1, 0),
        '교육내용\n'
        '○ 산소농도 측정 및 작업환경에 관한 사항\n'
        '○ 사고 시의 응급처지 및 비상 시 구출에 관한 사항\n'
        '○ 보호구 착용 및 사용방법에 관한 사항 요령\n'
        '○ 작업내용 ∙ 안전작업방법 및 절차에 관한 사항\n'
        '○ 장비 ∙ 설비 및 시설 등의 안전점검에 관한 사항\n'
        '○ 그 밖에 안전 ∙ 보건관리에 필요한 사항',
        align='left', v_align='top')
    write_cell(t4.cell(1, 1),
        f'○ 특별안전보건교육\n  ‑ 일정 : {data["edu_date"]}\n  ‑ 강사 : {data["instructor"]}',
        align='left', v_align='top')
    write_cell(t4.cell(2, 0),
        '○ 기타사항\n  ‑ 특별안전 교육 일지 및 보관\n  ‑ 교육일지 양식 활용',
        align='left', v_align='top')
    write_cell(t4.cell(2, 1), '* 교육사진\n\n(사진 부착 공간)', align='center', v_align='center')
    set_col_widths(t4, [8.5, 7])
    doc.add_page_break()

    # ================= Page 6 : 4. 건강장해 예방 =================
    add_header_marker(doc)
    add_section_title(doc, '4. 그 밖에 밀폐공간 작업 근로자의 건강\n     장해 예방에 관한 사항', size=13)

    add_para(doc, '(1) 보유장비 현황', size=11, bold=True, space_after=6)
    t5 = doc.add_table(rows=2, cols=5)
    t5.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['장비명', '모델명', '보유대수', '최근교정일', '교정주기']):
        write_cell(t5.cell(0, i), h, bold=True, bg=HEADER_BG)
    for i in range(5):
        write_cell(t5.cell(1, i), '')
    set_col_widths(t5, [3, 3, 2.5, 3, 3])

    add_para(doc, '', size=6)
    add_para(doc, '(2) 대여 장비 현황', size=11, bold=True, space_after=6)
    t6 = doc.add_table(rows=2, cols=5)
    t6.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['장비명', '필요수량', '대여일수', '연락처(협력업체)', '장비관리담당자']):
        write_cell(t6.cell(0, i), h, bold=True, bg=HEADER_BG)
    for i in range(5):
        write_cell(t6.cell(1, i), '')
    set_col_widths(t6, [3, 2.5, 2.5, 4, 3.5])

    add_para(doc, '', size=6)
    add_para(doc, '(3) 비상용 구조장비 비치현황', size=11, bold=True, space_after=6)
    t7 = doc.add_table(rows=5, cols=2)
    t7.alignment = WD_TABLE_ALIGNMENT.CENTER
    write_cell(t7.cell(0, 0), '분 류', bold=True, bg=HEADER_BG)
    write_cell(t7.cell(0, 1), '세부 항목', bold=True, bg=HEADER_BG)
    write_cell(t7.cell(1, 0), '구조장비', bold=True)
    write_cell(t7.cell(1, 1),
        '☐ 이동식 크레인   ☐ 호이스트   ☐ 삼각대   ☐ 구명로프   ☐ 들것\n☐ 윈치   ☐ 기타 (         )',
        align='left')
    write_cell(t7.cell(2, 0), '호흡장비', bold=True)
    write_cell(t7.cell(2, 1),
        '☐ 송기마스크(Air Line Mask)   ☐ ELSA (대피용 호흡장비)\n☐ 기타 (         )',
        align='left')
    write_cell(t7.cell(3, 0), '통화장비', bold=True)
    write_cell(t7.cell(3, 1), '☐ 무전기   ☐ 확성기   ☐ 호각   ☐ 기타 (         )', align='left')
    write_cell(t7.cell(4, 0), '기타\n추가장비', bold=True)
    write_cell(t7.cell(4, 1), '', align='left')
    set_col_widths(t7, [3, 12.5])
    doc.add_page_break()

    # ================= Page 7 : 비상연락, 작업자, 가스기준 =================
    add_header_marker(doc)
    add_para(doc, '(4) 비상연락체계', size=11, bold=True, space_after=6)
    t8 = doc.add_table(rows=6, cols=3)
    t8.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['성 명', '연 락 처', '안전대응팀']):
        write_cell(t8.cell(0, i), h, bold=True, bg=HEADER_BG)
    labels = ['현장소장', '안전관리자', '관리감독자', '작업반장', '기타']
    for i, lab in enumerate(labels):
        write_cell(t8.cell(i+1, 0), lab)
        write_cell(t8.cell(i+1, 1), data['manager'] if lab == '현장소장' else '')
    write_cell(t8.cell(1, 2), '포스코퓨처엠\n안전대응팀', bold=True)
    for r in range(2, 6):
        write_cell(t8.cell(r, 2), '')
    t8.cell(1, 2).merge(t8.cell(5, 2))
    set_col_widths(t8, [4, 6, 5.5])

    add_para(doc, '', size=6)
    add_para(doc, '(5) 작업자 정보', size=11, bold=True, space_after=6)
    t9 = doc.add_table(rows=7, cols=6)
    t9.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['구분', '성명', '연락처', '구분', '성명', '연락처']):
        write_cell(t9.cell(0, i), h, bold=True, bg=HEADER_BG)
    roles_L = ['공사감독', '관리감독자 1', '관리감독자 2', '감시인 1', '감시인 2', '감시인 3']
    roles_R = ['작업자 1', '작업자 2', '작업자 3', '작업자 4', '작업자 5', '작업자 6']
    for i in range(6):
        write_cell(t9.cell(i+1, 0), roles_L[i])
        write_cell(t9.cell(i+1, 3), roles_R[i])
        for c in [1, 2, 4, 5]:
            write_cell(t9.cell(i+1, c), '')
    set_col_widths(t9, [2.7, 2.3, 2.5, 2.7, 2.3, 2.5])

    add_para(doc, '', size=6)
    add_para(doc, '(6) 산소 및 유해가스 농도의 측정결과 평가 기준', size=11, bold=True, space_after=2)
    add_para(doc, '  ‑ 측정결과는 작업허가서 참조', size=9, space_after=6)
    t10 = doc.add_table(rows=2, cols=6)
    t10.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['구분', 'O₂', 'HC', 'H₂S', 'CO', 'CO₂']):
        write_cell(t10.cell(0, i), h, bold=True, bg=HEADER_BG)
    vals = ['작업허용\n농도', '18.0% 이상\n23.5% 미만', 'LEL 10% 미만*)', '10 ppm 미만', '30 ppm 미만', '1.5% 미만']
    for i, v in enumerate(vals):
        write_cell(t10.cell(1, i), v, bold=(i == 0))
    set_col_widths(t10, [2.5, 3, 2.5, 2.5, 2.5, 2.5])
    add_para(doc, '*) 용접/용단/화염사용 : 0% LEL', size=9, space_before=4)
    doc.add_page_break()

    # ================= Page 8 : 5. 프로그램 평가 =================
    add_header_marker(doc)
    add_section_title(doc, '5. 프로그램의 평가', size=14)
    add_para(doc, '밀폐공간 작업 프로그램 평가표', size=12, bold=True, align='center', space_after=4)
    add_para(doc, f'일자 : {data["today"]}', size=10, align='right', space_after=8)

    eval_items = [
        ("밀폐공간\n허가", 1, "밀폐공간 작업장소 보유현황 및 위치 등에 대한 자료가 작성되어 있는가?"),
        ("", 2, "밀폐공간 출입시 작업허가서를 작성하여 발급 받았는가?"),
        ("", 3, "작업허가서는 규정양식을 사용하여 올바르게 작성되었는가?"),
        ("", 4, "프로그램 추진팀(장)은 작업허가서를 적법한 절차에 의해 발급하였는가?"),
        ("산소 및\n유해가스\n농도측정", 5, "산소 및 유해가스 농도 측정대상 물질은 적정하게 선택되었으며 측정시 누락된 물질은 없는가?"),
        ("", 6, "측정장비의 신뢰성(교정 등)은 확보되었는가?"),
        ("", 7, "측정지점수, 측정방법 등은 정해진 규정을 준수하였는가?"),
        ("", 8, "측정결과에 대한 판정은 적합하게 이루어졌는가?"),
        ("환기대책", 9, "밀폐공간 작업장소에 따라 적합한 환기방법, 환기량 선정 등 환기대책은 적절하게 수립되었는가?"),
        ("", 10, "환기팬의 점검은 주기적으로 실시하였는가?"),
        ("보호구\n선정 및\n사용", 11, "보호구의 종류 및 수량은 충분한가?"),
        ("", 12, "보호구의 보유수량 및 대여필요장비 목록은 작성되어 있는가?"),
        ("", 13, "작업에 따라 적합한 보호구가 선정되어 사용되었는가?"),
        ("", 14, "누출검사를 매사용 시마다 시행하도록 하고 있는가?"),
        ("", 15, "보호구를 주기적으로 청소, 점검 등을 실시하는가?"),
        ("응급처치\n체계", 16, "응급상황 발생시 비상연락을 위한 체계는 구축되어 있는가?"),
        ("", 17, "응급전화, 무전기 등의 통신장비는 구비되어 있는가?"),
        ("교육 및\n훈련의\n적정성", 18, "프로그램관리자, 관리감독자, 작업자 등에 대한 교육계획을 수립하여 시행하고 있는가?"),
        ("", 19, "밀폐공간 작업시마다 작업자에게 교육을 실시하고 있는가?"),
        ("", 20, "관련교육을 실시하는 경우 교육내용 등을 기록하고 보존하는가?"),
        ("", 21, "교육내용, 자료 등은 적절하며 최신성을 유지하고 있는가?"),
        ("", 22, "교육받은 자는 교육내용을 충분히 숙지하여 작업에 올바르게 적용하고 있는가?"),
    ]
    t13 = doc.add_table(rows=len(eval_items) + 2, cols=4)
    t13.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['구분', '번호', '평가항목', '평가\n(O,X)']):
        write_cell(t13.cell(0, i), h, bold=True, bg=HEADER_BG)
    for i, item in enumerate(eval_items):
        write_cell(t13.cell(i+1, 0), item[0], bold=True)
        write_cell(t13.cell(i+1, 1), str(item[1]))
        write_cell(t13.cell(i+1, 2), item[2], size=9, align='left')
        write_cell(t13.cell(i+1, 3), '')
    last = len(eval_items) + 1
    write_cell(t13.cell(last, 0), '프로그램평가결과', bold=True, bg=SUBHEADER_BG)
    write_cell(t13.cell(last, 1), '', bg=SUBHEADER_BG)
    write_cell(t13.cell(last, 2), '판정수', bold=True, bg=SUBHEADER_BG)
    write_cell(t13.cell(last, 3), '', bg=SUBHEADER_BG)
    t13.cell(last, 0).merge(t13.cell(last, 1))
    t13.cell(1, 0).merge(t13.cell(4, 0))
    t13.cell(5, 0).merge(t13.cell(8, 0))
    t13.cell(9, 0).merge(t13.cell(10, 0))
    t13.cell(11, 0).merge(t13.cell(15, 0))
    t13.cell(16, 0).merge(t13.cell(17, 0))
    t13.cell(18, 0).merge(t13.cell(22, 0))
    set_col_widths(t13, [2.2, 1.2, 10.5, 1.6])
    doc.add_page_break()

    # ================= Page 9 : 6. 스마트 밀폐공간 환기량 계산 =================
    add_header_marker(doc)
    add_section_title(doc, '6. 스마트 밀폐공간 환기량 계산\n     (Smart Ventilation Design)', size=13)

    add_para(doc, '(1) 입력 정보', size=11, bold=True, space_after=6)
    t_in = doc.add_table(rows=2, cols=6)
    t_in.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = ['공간 형태', '가로 W(m)', '세로 L(m)', '높이 H(m)', '직경 D(m)', '동시 작업자(명)']
    for i, h in enumerate(hdr):
        write_cell(t_in.cell(0, i), h, bold=True, bg=HEADER_BG)
    shape_kor = {'rect': '사각형', 'cyl': '원통형', 'combined': '결합형'}[data['shape']]
    row_in = [shape_kor, f"{data['W']}", f"{data['L']}", f"{data['H']}", f"{data['D']}", str(data['workers'])]
    for i, v in enumerate(row_in):
        write_cell(t_in.cell(1, i), v)
    set_col_widths(t_in, [2.5, 2.3, 2.3, 2.3, 2.3, 3.0])

    add_para(doc, '', size=6)
    add_para(doc, '(2) 체적(Volume) 계산 공식', size=11, bold=True, space_after=4)
    add_para(doc, '  ○ 사각형 : V = W × L × H', size=10, space_after=2)
    add_para(doc, '  ○ 원통형 : V = (π × D² / 4) × H', size=10, space_after=2)
    add_para(doc, '  ○ 결합형 : V = (W × L × H) + (π × D² / 4) × H', size=10, space_after=8)

    volume = calc_volume(data['shape'], data['W'], data['L'], data['H'], data['D'])
    Q_need = volume * 0.4  # m³/min

    add_para(doc, '(3) 계산 결과', size=11, bold=True, space_after=6)
    t_out = doc.add_table(rows=4, cols=2)
    t_out.alignment = WD_TABLE_ALIGNMENT.CENTER
    write_cell(t_out.cell(0, 0), '최종 체적 V', bold=True, bg=HEADER_BG)
    write_cell(t_out.cell(0, 1), f'{volume:,.2f} m³', bold=True)
    write_cell(t_out.cell(1, 0), '법적 필요 환기량 Q\n( = V × 0.4 )', bold=True, bg=HEADER_BG)
    write_cell(t_out.cell(1, 1), f'{Q_need:,.1f} m³/min', bold=True, color=(217, 48, 37))
    write_cell(t_out.cell(2, 0), '작업 前 급기시간', bold=True, bg=HEADER_BG)
    write_cell(t_out.cell(2, 1), '최소 30분 이상 (체적의 10배 이상 환기)')
    write_cell(t_out.cell(3, 0), '작업 中 환기횟수', bold=True, bg=HEADER_BG)
    write_cell(t_out.cell(3, 1), '시간당 20회 이상 공기교환 유지')
    set_col_widths(t_out, [5, 10])

    add_para(doc, '', size=6)
    add_para(doc,
        f'※ 최소 {Q_need:,.1f} m³/min 이상의 용량을 가진 송풍기를 준비하십시오.',
        size=10, bold=True, color=(0, 91, 172), space_after=8)

    # 국소냉방 (Spot Cooling)
    add_para(doc, '(4) 국소냉방(Spot Cooling) 필요 열량 계산', size=11, bold=True, space_after=4)
    add_para(doc, '  Q_spot (kcal/h) = 동시 작업자 수 × 1인당 필요 부하', size=10, space_after=2)
    add_para(doc, '  ○ 일반작업 : 1,200 kcal/h·인    ○ 중작업(용접·축조·연마 등) : 1,500 kcal/h·인',
             size=10, space_after=6)

    heavy_keywords = ['용접', '축조', '조적', '절단', '연마', '해체', '그라인더', '중작업', '고열', '타설', '브레이커']
    is_heavy = any(k in data['task'] for k in heavy_keywords)
    per_load = 1500 if is_heavy else 1200
    Q_spot = data['workers'] * per_load if data['workers'] > 0 else 0

    t_spot = doc.add_table(rows=3, cols=2)
    t_spot.alignment = WD_TABLE_ALIGNMENT.CENTER
    write_cell(t_spot.cell(0, 0), '작업 내용', bold=True, bg=HEADER_BG)
    write_cell(t_spot.cell(0, 1), data['task'] if data['task'] else '(미입력)')
    write_cell(t_spot.cell(1, 0), '1인당 필요 부하', bold=True, bg=HEADER_BG)
    write_cell(t_spot.cell(1, 1), f"{per_load:,} kcal/h·인  ({'중작업' if is_heavy else '일반작업'})")
    write_cell(t_spot.cell(2, 0), '총 필요 냉방 능력', bold=True, bg=HEADER_BG)
    write_cell(t_spot.cell(2, 1), f'{Q_spot:,} kcal/h', bold=True, color=(217, 48, 37))
    set_col_widths(t_spot, [5, 10])

    add_para(doc, '', size=4)
    add_para(doc,
        '⚠️ 실무 가이드 : 스포트쿨러 본체는 외부에 두고 토출 덕트만 내부로 연장하여 Cool Zone을 형성하십시오.',
        size=10, bold=True, color=(154, 90, 0), space_after=8)

    add_para(doc, '(5) 환기장치 설치 및 점검 기준', size=11, bold=True, space_after=4)
    for line in [
        '  ○ 환기팬 정압 : 최소 40 mmAq 이상',
        '  ○ 송풍관(덕트) 길이 : 15m 이하 (제조사 권장 기준 준수)',
        '  ○ 급기구 위치 : 작업 근로자 가까이에서 근로자를 등지고 설치',
        '  ○ 방폭 위험구역 내에서는 반드시 방폭형 구조 사용',
    ]:
        add_para(doc, line, size=10, space_after=2)

    add_para(doc, '', size=6)
    add_para(doc, '(6) 작업내용 맞춤 추가 안전조치', size=11, bold=True, space_after=4)
    task_actions = []
    if any(k in data['task'] for k in ['축조', '조적', '분진', '연마', '절단', '해체', '브레이커', '그라인더', '천공', '샌딩']):
        task_actions.append('■ 분진 발생 작업 : 국소배기장치 제어풍속(포위식 0.7m/s, 외부식 1.0~1.2m/s) 준수')
    if any(k in data['task'] for k in ['용접', '절단', '화기', '그라인더']):
        task_actions.append('■ 용접·화기 작업 : 산소결핍/일산화탄소/화재·폭발 위험 대비, 화기작업허가·불티 비산방지·소화기 배치 병행')
    if any(k in data['task'] for k in ['도장', '페인트', '방수', '코팅', '유기용제', '신너', '희석제']):
        task_actions.append('■ 도장·유기용제 작업 : 방폭형 환기장치 사용 및 유기가스용 방독마스크/송기마스크 적용 검토')
    if not task_actions:
        task_actions.append('■ 일반 작업 : 작업 전·중 산소 및 유해가스를 반복 측정할 것')
    task_actions.append('■ 농도 급변 위험 시 공기호흡기/송기마스크 착용 필수')

    for line in task_actions:
        add_para(doc, '  ' + line, size=10, space_after=3)

    doc.add_page_break()

    # ================= Page 10 : 7. 온열질환 예방 프로그램 =================
    add_header_marker(doc)
    add_section_title(doc, '7. 하절기 온열질환 예방 프로그램\n     (Heat Stress Prevention Program)', size=13)

    add_para(doc, '* 운영기간 : \'26. 6. 1 ~ 8. 31 (기상 상황에 따라 연장 가능)',
             size=10, bold=True, space_after=2)
    add_para(doc, '* 폭염안전 5대 수칙 : 물, 그늘, 휴식 + 냉방/보냉장구, 119 즉시 신고',
             size=10, bold=True, color=(217, 48, 37), space_after=10)

    # 체감온도 산출
    hi = calc_heat_index(data.get('temp'), data.get('rh'))
    stage = heat_stage(hi)

    add_para(doc, '(1) 체감온도 산출 결과', size=11, bold=True, space_after=6)
    t_hi = doc.add_table(rows=4, cols=2)
    t_hi.alignment = WD_TABLE_ALIGNMENT.CENTER
    write_cell(t_hi.cell(0, 0), '입력 기온', bold=True, bg=HEADER_BG)
    write_cell(t_hi.cell(0, 1), f"{data['temp']} ℃" if data.get('temp') is not None else '미입력')
    write_cell(t_hi.cell(1, 0), '입력 습도', bold=True, bg=HEADER_BG)
    write_cell(t_hi.cell(1, 1), f"{data['rh']} %" if data.get('rh') is not None else '미입력')
    write_cell(t_hi.cell(2, 0), '예측 체감온도', bold=True, bg=HEADER_BG)
    write_cell(t_hi.cell(2, 1), f"약 {hi} ℃" if hi is not None else '미산출',
               bold=True, color=(217, 48, 37))
    write_cell(t_hi.cell(3, 0), '판정 단계', bold=True, bg=HEADER_BG)
    write_cell(t_hi.cell(3, 1), stage, bold=True, color=(0, 91, 172))
    set_col_widths(t_hi, [5, 10])

    add_para(doc, '', size=6)
    add_para(doc, '(2) 체감온도별 작업/휴식 기준', size=11, bold=True, space_after=6)
    t_heat = doc.add_table(rows=3, cols=4)
    t_heat.alignment = WD_TABLE_ALIGNMENT.CENTER
    write_cell(t_heat.cell(0, 0), '휴식기준', bold=True, bg=GRAY_BG)
    write_cell(t_heat.cell(0, 1), '체감온도\n31℃ 이상', bold=True, bg=YELLOW_BG, color=(180, 95, 6))
    write_cell(t_heat.cell(0, 2), '체감온도\n35℃ 이상', bold=True, bg=ORANGE_BG, color=(230, 145, 56))
    write_cell(t_heat.cell(0, 3), '체감온도\n38℃ 이상', bold=True, bg=RED_BG, color=(204, 0, 0))

    write_cell(t_heat.cell(1, 0), '폭염작업\n(작업 / 휴식시간)', bold=True, bg=GRAY_BG)
    write_cell(t_heat.cell(1, 1), '50분 작업 + 10분 휴식', bold=True, color=(0, 91, 172))
    write_cell(t_heat.cell(1, 2), '45분 작업 + 15분 휴식', bold=True, color=(0, 91, 172))
    write_cell(t_heat.cell(1, 3),
        '온열질환 고위험작업 중지\n(불가피한 경우\n부장·그룹장 승인 下 진행)',
        bold=True, color=(0, 91, 172))

    write_cell(t_heat.cell(2, 0), '온열질환 민감군\n고위험 작업', bold=True, bg=GRAY_BG)
    write_cell(t_heat.cell(2, 1),
        '+5 ~ +20분 추가 휴식\n(관리감독자 / 작업책임자 재량)', bold=True)
    # 1,2열 병합
    t_heat.cell(2, 1).merge(t_heat.cell(2, 2))
    # 상단 3열도 rowspan 대신 아래 셀 텍스트 유지 (이미 위에 표기)
    set_col_widths(t_heat, [3.5, 4, 4, 4])

    add_para(doc, '※ 40분 이내 단순 점검의 경우 작업중지 및 금지 대상 미적용',
             size=9, align='right', color=(100, 100, 100), space_before=4, space_after=8)

    # 필수 서류
    add_para(doc, '(3) 필수 작성 서류 (체감온도 무관, 당해년도 12.31까지 보관)',
             size=11, bold=True, space_after=4)
    for line in [
        '  ○ 위험성평가 Check List : 물, 그늘, 휴식, 건강확인 등 적정 여부 매일 체크',
        '  ○ 건강상태 자율진단표 : 3개 이상 해당자(작업배제), 1~2개 해당자(밀착관리 / 휴식연장)',
        '  ○ 체감온도 휴식시간 기록표 : 작업시작~종료까지 시간대별 기록',
        '    (31℃ 이상 시 휴식 / 물 섭취 여부 작성)',
    ]:
        add_para(doc, line, size=10, space_after=2)

    add_para(doc, '', size=6)
    add_para(doc, '(4) 열순응 프로그램 (신규 배치자 및 장기 휴가 복귀자)',
             size=11, bold=True, space_after=4)
    t_acc = doc.add_table(rows=2, cols=5)
    t_acc.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['1일차', '2일차', '3일차', '4일차', '5일차']):
        write_cell(t_acc.cell(0, i), h, bold=True, bg=HEADER_BG)
    for i, v in enumerate(['20%', '40%', '60%', '80%', '100%\n(완전적응)']):
        write_cell(t_acc.cell(1, i), v, bold=True)
    set_col_widths(t_acc, [3, 3, 3, 3, 3.5])
    add_para(doc, '  ※ 단계적 작업량 증가로 심혈관계 부담을 낮추고 열순응 유도', size=9, space_before=2, space_after=10)

    # TBM 대본
    add_para(doc, '(5) 하절기 全 작업 TBM 필수 안내 대본', size=11, bold=True, space_after=6)

    hi_disp = f"{hi} ℃" if hi is not None else "OO ℃"
    t_tbm = doc.add_table(rows=1, cols=1)
    t_tbm.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbm_cell = t_tbm.cell(0, 0)
    set_cell_bg(tbm_cell, 'F8FAFD')
    set_cell_border(tbm_cell)
    tbm_cell.text = ''
    lines_tbm = [
        (f'"현재 체감온도는 {hi_disp} 입니다."', True, (217, 48, 37)),
        (f'"『하절기 온열질환 예방수칙』의 {stage} 단계에 따라 작업해 주시기 바랍니다."', True, (0, 91, 172)),
        ('"온열질환 예방을 위해 틈틈이 물을 마시고, 충분히 휴식을 취합니다."', False, None),
        ('"작업 중 건강에 이상이 있으면 즉시 알리기 바랍니다."', False, None),
        ('', False, None),
        ('(선창) 물! (함께) 좋아!  /  (선창) 그늘! (함께) 좋아!  /  (선창) 휴식! (함께) 좋아!', True, (217, 48, 37)),
    ]
    for i, (txt, bold, color) in enumerate(lines_tbm):
        p = tbm_cell.paragraphs[0] if i == 0 else tbm_cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(txt)
        set_font(run, size=11, bold=bold, color=color)
    set_col_widths(t_tbm, [15.5])
    doc.add_page_break()

    # ================= Page 11 : 첨부 - 사전 확인 절차 =================
    add_header_marker(doc)
    add_para(doc, '■ 첨부서류', size=13, bold=True, space_after=4)
    add_section_title(doc, '밀폐공간 작업 시 사전 확인이 필요한 사항에\n     대한 확인 절차', size=13)

    t14 = doc.add_table(rows=3, cols=3)
    t14.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(['구분', '사전 확인 사항', '확인 절차']):
        write_cell(t14.cell(0, i), h, bold=True, bg=HEADER_BG)
    write_cell(t14.cell(1, 0), '공통', bold=True)
    write_cell(t14.cell(1, 1),
        '1. 작업내용, 작업방법, 순서에 대한 협의를 사전에 실시하였는가?\n'
        '2. 용기 내 위험물 제거 및 가스농도 측정 결과 양호한가?\n'
        '3. 개방 후 배관격리, 밸브의 맹판을 설치하였는가?\n'
        '4. Purge용 Steam/질소 Line분리하고 Air Hose 설치하였는가?\n'
        '5. 밸브잠금 표식(“조작 절대금지”) 및 시건 장치를 설치하였는가?\n'
        '6. 장치 내 모든 작동부분을 전기적, 기계적으로 차단하였는가?\n'
        '7. 경고문 표지판 및 인원점검표를 설치하였는가?\n'
        '8. 방폭용 전등 및 공기구를 사용하고 있는가?\n'
        '9. 비상시 연락장비 및 구조장비 등을 비치하고 있는가?\n'
        '10. 밀폐공간 내부의 적절한 환기 조치를 하였는가?',
        size=9, align='left', v_align='top')
    write_cell(t14.cell(1, 2),
        '관리감독자와 감시인이\n작업허가 시에 확인하고\n작업허가서에 서명함.',
        size=9, align='center', v_align='center')
    write_cell(t14.cell(2, 0), '질소충전\n또는\n산소농도\n부족작업', bold=True)
    write_cell(t14.cell(2, 1),
        '1. 작업계획서 별도 마련되어 있고 비상시 구조방법은 있는가?\n'
        '2. Plant Air압력이 6.5 이상임을 확인하였는가?\n'
        '3. Air Line Mask Flame Set의 이상여부를 확인하였는가?\n'
        '4. 공기호흡기를 용기출입구에 비치하고 있는가?',
        size=9, align='left', v_align='top')
    write_cell(t14.cell(2, 2),
        '관리감독자와 감시인이\n작업허가 시에 확인하고\n작업허가서에 서명함.',
        size=9, align='center', v_align='center')
    set_col_widths(t14, [2.5, 9.5, 3.5])

    add_para(doc, '', size=8)
    add_para(doc, '가. 산소 및 유해가스 농도 측정 방법 및 유의사항', size=11, bold=True, space_after=4)
    add_para(doc, '* 다음의 경우 반드시 측정할 것', size=10, bold=True, space_after=2)
    for line in [
        '1) 당일의 작업을 개시하기 전',
        '2) 교대제로 작업을 행할 경우 작업 당일 최초 교대가 행해져서 작업이 시작되기 전',
        '3) 작업에 종사하는 전체 근로자가 작업을 하고 있던 장소를 떠났다가 돌아와 다시 작업을 개시하기 전',
        '4) 근로자의 건강, 환기장치 등에 이상이 있을 때',
        '5) 작업을 하는 과정에서 유해가스가 발생할 가능성이 있을 경우(연속측정)',
        '6) 작업자 또는 작업승인부서에서 측정이 필요하다고 인정되는 경우',
        '7) 작업당일 밀폐공간 최초 출입전 가스측정 값은 1시간 이내일 것',
    ]:
        add_para(doc, line, size=10, space_after=2)
    doc.add_page_break()

    # ================= Page 12 : 응급처치 =================
    add_header_marker(doc)
    add_section_title(doc, '응급처치 시 관찰사항', size=14)
    add_para(doc, '응급처치 시에는 다음의 사항을 주의 깊게 관찰하고 그 내용을 담당자에게 정확히 전달',
             size=10, space_after=6)
    for line in [
        '○ 의식이 있는지 확인한다.',
        '○ 호흡하고 있는지 확인한다. 호흡이 정지되어 있으면 머리를 뒤로 젖히거나 아래턱을 밀어내어 기도를 열어주고 다시 확인한다.',
        '○ 출혈의 유무를 살펴본다.',
        '○ 맥을 짚어본다. 맥박이 뛰지 않는다고 느낄 때는 동공을 살펴본다. 동공이 크게 벌어져 있으면 위험하고 동공의 크기가 좌우 틀리면 뇌에 이상이 있는 경우이다.',
        '○ 손발이 움직이는가를 본다.',
        '○ 얼굴과 피부색, 체온을 살펴본다. 혀, 입술, 피부 등이 푸르스름한 색 또는 흑색이 되고 손톱은 암자색이 되었는지 살펴본다.',
        '○ 재해자의 체온을 유지하도록 보온한다.',
        '○ 협력자를 구한다.',
        '○ 재해자를 운반할 때는 서두르지 말고 재해자의 마음을 가라앉히고 되도록 재해자의 상처를 건드리지 않도록 주의하여 운반한다.',
    ]:
        add_para(doc, line, size=10, space_after=3)

    # 저장
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# =========================================================
# Streamlit UI
# =========================================================
st.sidebar.header("📝 작업 정보 입력")

project_name = st.sidebar.text_input("공사명", "포항공장 R-30101A 내부 정비공사")
start_date = st.sidebar.date_input("작업 시작일")
end_date = st.sidebar.date_input("작업 종료일")
work_detail = st.sidebar.text_area("작업내용", "내부 슬러지 제거 및 검사")
company_name = st.sidebar.text_input("회사명", "포스코퓨처엠")
manager = st.sidebar.text_input("현장소장 성명", "홍길동")

st.sidebar.markdown("---")
st.sidebar.subheader("🔧 공정 및 위험요인")
process_name = st.sidebar.text_input("공정명", "R-30101A")
space_name = st.sidebar.text_input("작업장소", "반응기 내부")
hazard_material = st.sidebar.text_input("유해위험 물질", "H2S, N2")
blind_count = st.sidebar.number_input("Blind 설치 개수", value=5, step=1)

st.sidebar.markdown("---")
st.sidebar.subheader("👨‍🏫 교육")
edu_date = st.sidebar.text_input("교육일정", "2026.06.15 08:00~09:00")
instructor = st.sidebar.text_input("강사", "안전관리자 김안전")

st.sidebar.markdown("---")
st.sidebar.subheader("📐 스마트 밀폐공간 환기 설계")
shape = st.sidebar.selectbox("공간 형태", options=['rect', 'cyl', 'combined'],
                              format_func=lambda x: {'rect': '사각형', 'cyl': '원통형', 'combined': '결합형'}[x])
W = st.sidebar.number_input("가로 W (m)", value=3.0, step=0.1)
L = st.sidebar.number_input("세로 L (m)", value=5.0, step=0.1)
H = st.sidebar.number_input("높이 H (m)", value=2.5, step=0.1)
D = st.sidebar.number_input("직경 D (m) - 원통형만", value=0.0, step=0.1)
workers = st.sidebar.number_input("동시 작업자 수 (명)", value=3, step=1)
task = st.sidebar.text_input("작업 내용 (예: 용접, 축조, 도장, 일반 점검)", "일반 점검")

st.sidebar.markdown("---")
st.sidebar.subheader("🌡️ 온열질환 예방")
temp = st.sidebar.number_input("내부 예상 기온 (℃)", value=33.0, step=0.1)
rh = st.sidebar.number_input("내부 예상 습도 (%)", value=65, min_value=0, max_value=100, step=1)

# =========================================================
# 메인 화면
# =========================================================
st.title("📄 포스코퓨처엠 스마트 밀폐공간 안전설계 및 온열질환 예방 프로그램")
st.markdown("""
**밀폐공간 작업 프로그램(17페이지 원본 양식) + 스마트 환기량 계산 + 하절기 온열질환 예방 프로토콜**을 하나의 완성형 Word 문서로 자동 생성합니다.

- ✅ 표지 회사명 : **포스코퓨처엠**
- ✅ **밀폐공간 환기량 계산 프로그램** (체적 → 필요 환기량 Q=V×0.4, 스포트쿨러 필요 열량)
- ✅ **온열질환 예방 프로그램** (체감온도 산출, 3단계 폭염 기준, 열순응, TBM 대본)
- ✅ 작업 내용에 따른 **맞춤 안전조치** 자동 추천
""")

# 실시간 계산 미리보기
vol_preview = calc_volume(shape, W, L, H, D)
q_preview = vol_preview * 0.4
hi_preview = calc_heat_index(temp, rh)

col1, col2, col3 = st.columns(3)
col1.metric("체적 (V)", f"{vol_preview:,.2f} m³")
col2.metric("필요 환기량 (Q)", f"{q_preview:,.1f} m³/min")
col3.metric("예측 체감온도", f"{hi_preview} ℃" if hi_preview is not None else "-",
            delta=heat_stage(hi_preview))

data = {
    "project_name": project_name,
    "start_date": start_date.strftime("%Y.%m.%d"),
    "end_date": end_date.strftime("%Y.%m.%d"),
    "work_detail": work_detail,
    "today": datetime.now().strftime("%Y. %m. %d"),
    "company_name": company_name,
    "manager": manager,
    "process_name": process_name,
    "space_name": space_name,
    "hazard_material": hazard_material,
    "blind_count": blind_count,
    "edu_date": edu_date,
    "instructor": instructor,
    # 스마트 환기 설계
    "shape": shape,
    "W": W, "L": L, "H": H, "D": D,
    "workers": workers,
    "task": task,
    # 온열질환
    "temp": temp,
    "rh": rh,
}

word_file = create_word_document(data)

st.download_button(
    label="📥 포스코퓨처엠 밀폐공간 작업 프로그램 (Word) 다운로드",
    data=word_file,
    file_name=f"포스코퓨처엠_밀폐공간_작업프로그램_{project_name}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True,
)

st.caption("© 2026 포스코퓨처엠 안전보건팀 | Smart Confined-space Safety Design & Heat Stress Prevention Program")
