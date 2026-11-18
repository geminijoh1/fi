# -*- coding: utf-8 -*-
"""
월간주택시장동향(Monthly R Briefing) 원천 데이터 입력 및 자동화 마스터 템플릿 엑셀 생성 스크립트
사용자가 각 시트에 데이터를 채워 넣고 에이전트에게 보고서 작성을 의뢰할 수 있는 턴키 템플릿
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

def create_full_data_template(filename="월간주택시장동향_필요데이터_체크리스트.xlsx"):
    wb = openpyxl.Workbook()
    
    # ---------------------------------------------------------
    # 공통 스타일 정의
    # ---------------------------------------------------------
    font_main_title = Font(name="맑은 고딕", size=15, bold=True, color="FFFFFF")
    font_sub_title = Font(name="맑은 고딕", size=9, italic=False, color="E0E8F5")
    font_section = Font(name="맑은 고딕", size=11, bold=True, color="1B365D")
    font_table_hdr = Font(name="맑은 고딕", size=9, bold=True, color="FFFFFF")
    font_table_hdr_sub = Font(name="맑은 고딕", size=9, bold=True, color="222222")
    font_data = Font(name="맑은 고딕", size=9, bold=False, color="222222")
    font_data_bold = Font(name="맑은 고딕", size=9, bold=True, color="1B365D")
    font_data_num = Font(name="맑은 고딕", size=9, bold=False, color="111111")
    font_note = Font(name="맑은 고딕", size=8, italic=True, color="777777")
    
    fill_navy = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # 메인 네이비
    fill_slate = PatternFill(start_color="2C4D75", end_color="2C4D75", fill_type="solid") # 세부 헤더 블루
    fill_steel = PatternFill(start_color="466D98", end_color="466D98", fill_type="solid") # 부록 헤더
    fill_light_gray = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid") # 교대행
    fill_light_blue = PatternFill(start_color="EAF2F8", end_color="EAF2F8", fill_type="solid") # 강조행/합계
    fill_light_yellow = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid") # 입력 셀 강조
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    thin_gray = Side(border_style="thin", color="D3D9E2")
    thick_bottom = Side(border_style="medium", color="1B365D")
    double_bottom = Side(border_style="double", color="1B365D")
    
    border_data = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    border_hdr = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thick_bottom)
    border_total = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=double_bottom)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    def draw_title_banner(ws, title_text, guide_text):
        ws.merge_cells("A1:K1")
        ws.merge_cells("A2:K2")
        ws["A1"] = f"  {title_text}"
        ws["A1"].font = font_main_title
        ws["A1"].fill = fill_navy
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        
        ws["A2"] = f"  ▶ 안내: {guide_text}"
        ws["A2"].font = font_sub_title
        ws["A2"].fill = fill_slate
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
        
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 20
        ws.row_dimensions[3].height = 10

    # =========================================================
    # SHEET 0: 00_총괄_마스터체크리스트
    # =========================================================
    ws0 = wb.active
    ws0.title = "00_총괄체크리스트"
    ws0.views.sheetView[0].showGridLines = True
    
    draw_title_banner(ws0, "부동산114 월간주택시장동향 데이터 입력 총괄 마스터 시트", 
                      "하단 각 탭(시트)에 당월 신규 원천 데이터를 입력한 후 작업을 의뢰해 주시면, 마크다운/PDF 보고서가 자동 생성됩니다.")
    
    # 기준 정보 입력 카드
    ws0["A4"] = "■ 보고서 발행 기본 정보 (작업 의뢰 시 필수 확인)"
    ws0["A4"].font = font_section
    
    meta_headers = ["보고서 대상 연월", "보고서 발간일", "데이터 담당자", "검증 완료 여부", "보고서 테마 / 주요 이슈 키워드"]
    meta_values = ["2026년 9월호 (예시)", "2026-10-05", "부동산114 데이터솔루션팀", "검증완료", "주담대 금리 급등에 따른 매수 관망 및 거래량 급감세"]
    
    for c_idx, (h, v) in enumerate(zip(meta_headers, meta_values), 1):
        cell_h = ws0.cell(row=5, column=c_idx, value=h)
        cell_h.font = font_table_hdr
        cell_h.fill = fill_slate
        cell_h.alignment = align_center
        cell_h.border = border_hdr
        
        cell_v = ws0.cell(row=6, column=c_idx, value=v)
        cell_v.font = font_data_bold
        cell_v.fill = fill_light_yellow
        cell_v.alignment = align_center
        cell_v.border = border_data
    
    ws0.row_dimensions[4].height = 24
    ws0.row_dimensions[5].height = 24
    ws0.row_dimensions[6].height = 24
    
    # 시트별 입력 현황 표
    ws0["A8"] = "■ 하단 개별 데이터 시트별 입력 현황 체크리스트"
    ws0["A8"].font = font_section
    
    sheet_checklist_headers = [
        "시트 번호", "시트명", "담당 보고서 섹션", "핵심 데이터 항목", "반영 시점(Lag)", 
        "원천 데이터 출처", "데이터 입력 상태", "검증 상태", "최종 확인일시", "비고 / 특이사항"
    ]
    
    ws0.row_dimensions[9].height = 26
    for c_idx, h in enumerate(sheet_checklist_headers, 1):
        cell = ws0.cell(row=9, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_hdr
        
    sheets_info = [
        ("01", "01_거시_유동성", "01. 거시경제 - 실물경기 및 유동성", "동행/선행지수, 순환변동치, 시중통화량(M2)", "T-1월", "통계청, 한국은행 ECOS", "입력완료", "검증완료", "2026-10-02", "동행지수 상승 vs 선행순환 하락전환"),
        ("02", "02_금융_금리여신", "01. 거시경제 - 금융시장 동향", "기준금리, 가계/주담대 금리, 가계/주담대 잔액", "당월/T-1월", "한국은행 ECOS", "입력완료", "검증완료", "2026-10-02", "주담대 금리 3년 9개월래 최고, 잔액 +3.4조"),
        ("03", "03_매매가격동향", "02. 주택시장 - 매매가격 동향", "전국/권역/17개 시도 변동률, 급등락 시군구", "당월(T월)", "부동산114 자체 DB (PropX)", "입력완료", "검증완료", "2026-10-03", "전국 +0.40% (경기 0.77%, 대전 0.61%)"),
        ("04", "04_매매거래동향", "02. 주택시장 - 매매거래 동향", "전국/17개 시도 거래량, 면적규모별 비중", "T-1월 확정", "국토교통부 실거래 신고", "입력완료", "검증완료", "2026-10-03", "전국 3.75만호 (-25.19% 급감)"),
        ("05", "05_분양_청약동향", "03. 분양시장 - 분양공급 및 청약", "전국/권역 분양실적, 청약경쟁률 상위 5단지", "당월(T월)", "부동산114 DB, 청약홈", "입력완료", "검증완료", "2026-10-04", "전국 2.42만호, 1위 브라운스톤(61.9:1)"),
        ("06", "06_미분양동향", "03. 분양시장 - 미분양 동향", "전국/시도 미분양 재고, 준공후 미분양", "T-1월 말", "국토교통부 주택통계", "입력완료", "검증완료", "2026-10-04", "미분양 6.91만호(↑), 준공후 2.81만호(↓)"),
        ("07", "07_부록_분양단지", "APPENDIX - 주요 분양 단지", "당월 분양 개시 전국 주요 단지 리스트", "당월(T월)", "부동산114 캘린더 DB", "입력완료", "검증완료", "2026-10-04", "총 13개 주요 단지 수록"),
        ("08", "08_부록_입주단지", "APPENDIX - 주요 입주 단지", "당월 입주 개시 전국 주요 단지 리스트", "당월(T월)", "부동산114 캘린더 DB", "입력완료", "검증완료", "2026-10-04", "디에이치방배 등 13개 주요 단지 수록")
    ]
    
    dv_status = DataValidation(type="list", formula1='"미착수,입력중,입력완료,검증완료"', allow_blank=True)
    ws0.add_data_validation(dv_status)
    dv_status.add("G10:H17")
    
    for r_idx, s_info in enumerate(sheets_info, start=10):
        ws0.row_dimensions[r_idx].height = 22
        fill_r = fill_light_gray if r_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(s_info, start=1):
            cell = ws0.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            cell.fill = fill_r
            if c_idx in [1, 5, 7, 8, 9]:
                cell.alignment = align_center
                if c_idx in [7, 8]:
                    cell.font = font_data_bold
            elif c_idx == 2:
                cell.alignment = align_left
                cell.font = font_data_bold
            else:
                cell.alignment = align_left

    col_w_ws0 = {"A": 10, "B": 18, "C": 30, "D": 36, "E": 14, "F": 22, "G": 14, "H": 14, "I": 14, "J": 38}
    for col, width in col_w_ws0.items():
        ws0.column_dimensions[col].width = width

    # =========================================================
    # SHEET 1: 01_거시_유동성
    # =========================================================
    ws1 = wb.create_sheet(title="01_거시_유동성")
    ws1.views.sheetView[0].showGridLines = True
    draw_title_banner(ws1, "01. 거시경제 현황 - 실물경기 지표 및 시중통화량(M2) 입력 시트",
                      "통계청 산업활동동향(경기종합지수) 및 한국은행 ECOS(M2 광의통화 평잔 계절조정)에서 T-1월 데이터를 입력하십시오.")
    
    ws1["A4"] = "■ [표 1] 실물경기 지표 (경기종합지수 및 순환변동치) - 출처: 통계청 산업활동동향"
    ws1["A4"].font = font_section
    
    headers_macro = ["구분", "지표명", "기준", "단위", "전월 수치\n(2026.07)", "당월 수치\n(2026.08)", "전월대비 증감\n(공식 자동)", "시장 평가 및 시사점 메모"]
    ws1.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers_macro, 1):
        cell = ws1.cell(row=5, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_hdr
        
    data_macro = [
        ("동행지수", "동행종합지수", "2020=100", "지수", 118.1, 118.8, "=+F6-E6", "현재 경기 국면의 완만한 개선세 지속 (전월대비 +0.6% 상승)"),
        ("동행지수", "동행순환변동치", "추세치=100", "p", 101.2, 101.7, "=+F7-E7", "경기 확장 국면 유지 (+0.5p)"),
        ("선행지수", "선행종합지수", "2020=100", "지수", 131.2, 131.7, "=+F8-E8", "완만한 상승세 유지 (전월대비 +0.3% 상승)"),
        ("선행지수", "선행순환변동치", "추세치=100", "p", 104.3, 104.2, "=+F9-E9", "건설수주액 감소 등으로 하락 전환 (-0.1p), 향후 경기회복 둔화 시그널")
    ]
    for r_idx, row_vals in enumerate(data_macro, start=6):
        ws1.row_dimensions[r_idx].height = 22
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            if c_idx in [1, 3, 4]:
                cell.alignment = align_center
            elif c_idx in [5, 6]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_yellow
            elif c_idx == 7:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_blue
                cell.number_format = "+0.0;-0.0;0.0"
            else:
                cell.alignment = align_left
                
    ws1["A12"] = "■ [표 2] 시중 유동성(M2 광의통화) 추이 - 출처: 한국은행 ECOS (계절조정계열 평잔 기준)"
    ws1["A12"].font = font_section
    
    headers_m2 = ["구분", "통계 지표명", "단위", "전월 수치\n(2026.07)", "당월 수치\n(2026.08)", "전월비 증감액(조원)\n(수식)", "전월비 증감률(%)\n(수식)", "연속 증가개월수", "부동산 시장 영향 및 유동성 분석 메모"]
    ws1.row_dimensions[13].height = 26
    for c_idx, h in enumerate(headers_m2, 1):
        cell = ws1.cell(row=13, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_slate
        cell.alignment = align_center
        cell.border = border_hdr
        
    ws1.row_dimensions[14].height = 24
    m2_vals = ["시중유동성", "M2 (광의통화, 평잔)", "조원", 4225.6, 4266.8, "=+E14-D14", "=+(E14-D14)/D14*100", "10개월 연속", "기업 자금 유입 확대로 시중 유동성 대폭 확충, 부동산 매수 대기 자금 하방 지지 압력"]
    for c_idx, val in enumerate(m2_vals, start=1):
        cell = ws1.cell(row=14, column=c_idx, value=val)
        cell.font = font_data
        cell.border = border_data
        if c_idx in [1, 3, 8]:
            cell.alignment = align_center
        elif c_idx in [4, 5]:
            cell.alignment = align_right
            cell.font = font_data_bold
            cell.fill = fill_light_yellow
        elif c_idx in [6, 7]:
            cell.alignment = align_right
            cell.font = font_data_bold
            cell.fill = fill_light_blue
            cell.number_format = "+0.00;-0.00;0.00"
        else:
            cell.alignment = align_left

    col_w_ws1 = {"A": 14, "B": 24, "C": 14, "D": 14, "E": 18, "F": 18, "G": 20, "H": 18, "I": 46}
    for col, width in col_w_ws1.items():
        ws1.column_dimensions[col].width = width

    # =========================================================
    # SHEET 2: 02_금융_금리여신
    # =========================================================
    ws2 = wb.create_sheet(title="02_금융_금리여신")
    ws2.views.sheetView[0].showGridLines = True
    draw_title_banner(ws2, "02. 금융시장 동향 - 금리 및 예금은행 가계대출 잔액 입력 시트",
                      "한국은행 기준금리(최신) 및 예금은행 신규 대출금리/월말 대출잔액(T-1월 기준, 조원 단위)을 입력하십시오.")
    
    ws2["A4"] = "■ [표 1] 기준금리 및 예금은행 대출금리 (신규취급액 가중평균 기준) - 출처: 한국은행 ECOS"
    ws2["A4"].font = font_section
    
    headers_rate = ["구분", "금리 지표명", "공표 주기", "단위", "전월 금리\n(2026.07)", "당월 금리\n(2026.08)", "등락폭(%p)\n(수식)", "금리 평가 및 시장 영향 메모"]
    ws2.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers_rate, 1):
        cell = ws2.cell(row=5, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_hdr
        
    rates_data = [
        ("정책금리", "한국은행 기준금리", "금통위", "연 %", 2.75, 3.00, "=+F6-E6", "긴축 기조 유지 (+0.25%p 인상 반영)"),
        ("가계여신", "가계대출금리 (신규)", "월간", "연 %", 4.64, 4.76, "=+F7-E7", "조달비용 상승 압력 반영 (+0.12%p 상승)"),
        ("주담대", "주택담보대출금리 (신규)", "월간", "연 %", 4.48, 4.66, "=+F8-E8", "3년 9개월 만에 최고치 경신 (+0.18%p 급등), 매수 심리 위축 요인")
    ]
    for r_idx, row_vals in enumerate(rates_data, start=6):
        ws2.row_dimensions[r_idx].height = 22
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            if c_idx in [1, 3, 4]:
                cell.alignment = align_center
            elif c_idx in [5, 6]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_yellow
                cell.number_format = "0.00"
            elif c_idx == 7:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_blue
                cell.number_format = "+0.00;-0.00;0.00"
            else:
                cell.alignment = align_left
                
    ws2["A11"] = "■ [표 2] 예금은행 가계대출 잔액 동향 (월말 잔액 기준) - 출처: 한국은행 금융시장동향"
    ws2["A11"].font = font_section
    
    headers_loan = ["구분", "대출 계정과목", "단위", "전월 잔액\n(2026.07)", "당월 잔액\n(2026.08)", "전월비 증감액\n(조원, 수식)", "연속 증가추세", "비고 및 리스크 검증 (총잔액=주담대+기타 검증)"]
    ws2.row_dimensions[12].height = 26
    for c_idx, h in enumerate(headers_loan, 1):
        cell = ws2.cell(row=12, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_slate
        cell.alignment = align_center
        cell.border = border_hdr
        
    loans_data = [
        ("총여신", "가계대출 총잔액", "조원", 1194.8, 1198.3, "=+E13-D13", "6개월 연속 증가", "가계대출 +3.4조원 증가 (원문 오탈자 검증 필수)"),
        ("주택자금", "  └ 주택담보대출 잔액", "조원", 948.4, 952.5, "=+E14-D14", "견조한 증가세", "주담대 +4.0조원 증가 (전체 가계대출 증가 주도)"),
        ("기타여신", "  └ 기타대출 잔액 (신용 등)", "조원", 246.4, 245.8, "=+E15-D15", "감소 전환", "기타대출 -0.6조원 순감소 (주식 투자 수요 둔화 영향)")
    ]
    for r_idx, row_vals in enumerate(loans_data, start=13):
        ws2.row_dimensions[r_idx].height = 22
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            if c_idx in [1, 3, 7]:
                cell.alignment = align_center
            elif c_idx in [4, 5]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_yellow
                cell.number_format = "#,##0.0"
            elif c_idx == 6:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_blue
                cell.number_format = "+0.0;-0.0;0.0"
            else:
                cell.alignment = align_left

    col_w_ws2 = {"A": 14, "B": 26, "C": 14, "D": 14, "E": 18, "F": 18, "G": 20, "H": 46}
    for col, width in col_w_ws2.items():
        ws2.column_dimensions[col].width = width

    # =========================================================
    # SHEET 3: 03_매매가격동향
    # =========================================================
    ws3 = wb.create_sheet(title="03_매매가격동향")
    ws3.views.sheetView[0].showGridLines = True
    draw_title_banner(ws3, "03. 주택시장 동향 - 아파트 매매가격 변동률 입력 시트",
                      "부동산114 자체 아파트 시세 DB(PropX)에서 당월(T월) 전국, 권역, 17개 시도별 매매가격 변동률을 입력하십시오.")
    
    ws3["A4"] = "■ [표 1] 전국 및 권역별 아파트 매매가격 변동률 - 출처: 부동산114 시세 DB"
    ws3["A4"].font = font_section
    
    headers_price_region = ["권역 구분", "지역명", "단위", "전월 변동률\n(2026.08)", "당월 변동률\n(2026.09)", "등락폭(%p)\n(수식)", "시장 추세 및 핵심 특징 요약"]
    ws3.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers_price_region, 1):
        cell = ws3.cell(row=5, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_hdr
        
    price_regions = [
        ("전국", "전국 종합", "%", 0.55, 0.40, "=+E6-D6", "2개월 연속 상승폭 둔화 (+0.40%)"),
        ("수도권", "수도권 (서울·경기·인천)", "%", 0.68, 0.48, "=+E7-D7", "경기 남부권 오름세 지탱, 전반적 상승폭 축소"),
        ("지방", "지방 5개광역시 및 도", "%", 0.11, 0.10, "=+E8-D8", "미분양 적체에 따른 정체 및 보합세 유지")
    ]
    for r_idx, row_vals in enumerate(price_regions, start=6):
        ws3.row_dimensions[r_idx].height = 22
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            if c_idx in [1, 2, 3]:
                cell.alignment = align_center
                if c_idx in [1, 2]:
                    cell.font = font_data_bold
            elif c_idx in [4, 5]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_yellow
                cell.number_format = "+0.00;-0.00;0.00"
            elif c_idx == 6:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_blue
                cell.number_format = "+0.00;-0.00;0.00"
            else:
                cell.alignment = align_left
                
    ws3["A11"] = "■ [표 2] 전국 17개 시도별 매매가격 변동률 (전수 입력)"
    ws3["A11"].font = font_section
    
    headers_sido = ["No", "시도명", "권역 구분", "전월 변동률\n(2026.08, %)", "당월 변동률\n(2026.09, %)", "전월비 등락폭\n(%p, 수식)", "당월 순위", "상승/하락 구분", "주요 지역 이슈 및 시장 코멘트"]
    ws3.row_dimensions[12].height = 26
    for c_idx, h in enumerate(headers_sido, 1):
        cell = ws3.cell(row=12, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_slate
        cell.alignment = align_center
        cell.border = border_hdr
        
    sido_price_data = [
        (1, "경기", "수도권", 0.95, 0.77, "=+E13-D13", 1, "상승", "경기 남부권 중심 견조한 오름세 유지 (전국 1위)"),
        (2, "대전", "지방", 0.24, 0.61, "=+E14-D14", 2, "상승", "서구 선도지구 정비 기대감(1.44%) 급등으로 전국 2위"),
        (3, "서울", "수도권", 0.55, 0.34, "=+E15-D15", 3, "상승", "대출금리 상승 및 규제 관망세로 상승폭 축소"),
        (4, "전북", "지방", 0.18, 0.31, "=+E16-D16", 4, "상승", "전주 등 일부 중심지역 위주 상승"),
        (5, "경남", "지방", 0.08, 0.12, "=+E17-D17", 5, "상승", "창원 등 주거 선호지역 중심 완만한 회복"),
        (6, "충북", "지방", 0.06, 0.08, "=+E18-D18", 6, "상승", "청주 중심 완만한 보합세"),
        (7, "대구", "지방", -0.05, 0.05, "=+E19-D19", 7, "상승", "하락세 진정 및 소폭 반등"),
        (8, "부산", "지방", 0.02, 0.04, "=+E20-D20", 8, "상승", "동래·해운대 등 일부 제외 보합"),
        (9, "울산", "지방", 0.03, 0.03, "=+E21-D21", 9, "상승", "남구 위주 강보합 유지"),
        (10, "경북", "지방", 0.01, 0.02, "=+E22-D22", 10, "상승", "포항·구미 등 보합 국면"),
        (11, "광주", "지방", 0.00, 0.01, "=+E23-D23", 11, "상승", "수요 정체 속 보합"),
        (12, "충남", "지방", 0.01, 0.00, "=+E24-D24", 12, "보합", "천안·아산 보합 유지"),
        (13, "전남", "지방", -0.01, 0.00, "=+E25-D25", 13, "보합", "순천·여수 등 매수 관망"),
        (14, "인천", "수도권", 0.19, -0.01, "=+E26-D26", 14, "하락", "검단 등 공급물량 부담으로 하락 전환"),
        (15, "강원", "지방", 0.02, -0.02, "=+E27-D27", 15, "하락", "원주·춘천 매물 적체로 약보합 하락"),
        (16, "제주", "지방", -0.04, -0.03, "=+E28-D28", 16, "하락", "침체 장기화 속 마이너스 변동률 지속"),
        (17, "세종", "지방", -0.12, -0.10, "=+E29-D29", 17, "하락", "매물 적체 및 수요 위축으로 전국 최대 낙폭")
    ]
    for r_idx, row_vals in enumerate(sido_price_data, start=13):
        ws3.row_dimensions[r_idx].height = 20
        fill_r = fill_light_gray if r_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            cell.fill = fill_r
            if c_idx in [1, 2, 3, 7, 8]:
                cell.alignment = align_center
                if c_idx in [2, 8]:
                    cell.font = font_data_bold
            elif c_idx in [4, 5]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_yellow
                cell.number_format = "+0.00;-0.00;0.00"
            elif c_idx == 6:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_blue
                cell.number_format = "+0.00;-0.00;0.00"
            else:
                cell.alignment = align_left

    col_w_ws3 = {"A": 6, "B": 12, "C": 12, "D": 16, "E": 16, "F": 16, "G": 12, "H": 12, "I": 46}
    for col, width in col_w_ws3.items():
        ws3.column_dimensions[col].width = width

    # =========================================================
    # SHEET 4: 04_매매거래동향
    # =========================================================
    ws4 = wb.create_sheet(title="04_매매거래동향")
    ws4.views.sheetView[0].showGridLines = True
    draw_title_banner(ws4, "04. 주택시장 동향 - 아파트 매매거래량 및 규모별 비중 입력 시트",
                      "국토교통부 실거래가 신고 데이터(T-1월 확정치, 30일 신고기한 반영)를 만 세대(또는 세대) 단위로 입력하십시오.")
    
    ws4["A4"] = "■ [표 1] 전국 및 주요 시도 아파트 매매거래량 추이 - 출처: 국토교통부 실거래 공개시스템"
    ws4["A4"].font = font_section
    
    headers_trade = ["No", "지역 구분", "전월 거래량\n(2026.07, 만호)", "당월 거래량\n(2026.08, 만호)", "전월비 증감량\n(만호, 수식)", "전월비 증감률\n(%, 수식)", "전국 점유비중\n(%, 수식)", "거래 규모(면적)별 주요 집중도 특성", "시장 평가 코멘트"]
    ws4.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers_trade, 1):
        cell = ws4.cell(row=5, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_hdr
        
    trades_data = [
        ("합계", "전국 합계", 5.01, 3.75, "=+D6-C6", "=+(D6-C6)/C6*100", 100.0, "전국 전용 60~85㎡ 중형 중심 (소형 41%, 중형 49%)", "전월 대비 -25.19% 급감, 거래 절벽 양상 가시화"),
        (1, "경기", 1.70, 1.22, "=+D7-C7", "=+(D7-C7)/C7*100", "=+(D7/$D$6)*100", "전용 60~85㎡ 중형 실수요 거래 집중", "전국 거래량 1위 (비중 32.53%)"),
        (2, "서울", 0.66, 0.46, "=+D8-C8", "=+(D8-C8)/C8*100", "=+(D8/$D$6)*100", "전용 60㎡ 이하 소형 아파트 거래 선호 뚜렷", "금리 급등 및 대출 규제로 매수 관망 확대 (-30.3%)"),
        (3, "전남광주", 0.32, 0.25, "=+D9-C9", "=+(D9-C9)/C9*100", "=+(D9/$D$6)*100", "중소형 실수요 중심 거래", "거래 감소 속 지방 상위권 유지"),
        (4, "부산", 0.28, 0.23, "=+D10-C10", "=+(D10-C10)/C10*100", "=+(D10/$D$6)*100", "해운대 등 주요지 제외 전반적 정체", "전월비 -17.86% 감소"),
        (5, "경남", 0.26, 0.23, "=+D11-C11", "=+(D11-C11)/C11*100", "=+(D11/$D$6)*100", "창원 중심 실수요 유지", "지방 상위권 거래량"),
        (6, "인천", 0.24, 0.18, "=+D12-C12", "=+(D12-C12)/C12*100", "=+(D12/$D$6)*100", "송도·검단 관망세 확산", "거래량 -25.0% 감소"),
        (7, "대구", 0.18, 0.14, "=+D13-C13", "=+(D13-C13)/C13*100", "=+(D13/$D$6)*100", "수성구 등 급매물 소진 후 정체", "침체 지속"),
        (8, "세종", 0.06, 0.03, "=+D14-C14", "=+(D14-C14)/C14*100", "=+(D14/$D$6)*100", "거래 극심한 침체", "전월비 -50.0% 반토막"),
        (9, "제주", 0.02, 0.01, "=+D15-C15", "=+(D15-C15)/C15*100", "=+(D15/$D$6)*100", "외지인 투자 매수 전면 위축", "전국 최저 수준 거래량")
    ]
    for r_idx, row_vals in enumerate(trades_data, start=6):
        ws4.row_dimensions[r_idx].height = 22
        is_total = (r_idx == 6)
        fill_r = fill_light_blue if is_total else (fill_light_gray if r_idx % 2 == 0 else fill_white)
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws4.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data_bold if is_total else font_data
            cell.border = border_total if is_total else border_data
            cell.fill = fill_r
            if c_idx in [1, 2]:
                cell.alignment = align_center
            elif c_idx in [3, 4]:
                cell.alignment = align_right
                cell.font = font_data_bold
                if not is_total:
                    cell.fill = fill_light_yellow
                cell.number_format = "#,##0.00"
            elif c_idx in [5, 6, 7]:
                cell.alignment = align_right
                cell.font = font_data_bold
                if not is_total:
                    cell.fill = fill_light_blue
                if c_idx == 5:
                    cell.number_format = "+0.00;-0.00;0.00"
                else:
                    cell.number_format = "0.00"
            else:
                cell.alignment = align_left

    col_w_ws4 = {"A": 8, "B": 14, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 36, "I": 42}
    for col, width in col_w_ws4.items():
        ws4.column_dimensions[col].width = width

    # =========================================================
    # SHEET 5: 05_분양_청약동향
    # =========================================================
    ws5 = wb.create_sheet(title="05_분양_청약동향")
    ws5.views.sheetView[0].showGridLines = True
    draw_title_banner(ws5, "05. 분양시장 동향 - 분양실적 및 청약경쟁률 상위 5개 단지 입력 시트",
                      "부동산114 분양 DB 및 청약홈(applyhome.co.kr)에서 당월(T월) 분양 공급물량과 청약경쟁률 상위 단지를 입력하십시오.")
    
    ws5["A4"] = "■ [표 1] 권역별 분양공급 물량 실적 - 출처: 부동산114 분양 DB"
    ws5["A4"].font = font_section
    
    headers_supply = ["권역 구분", "단위", "전월 분양실적\n(2026.08, 만호)", "당월 분양실적\n(2026.09, 만호)", "전월비 증감률\n(%, 수식)", "권역 공급비중\n(%, 수식)", "분양시장 특징 및 권역별 동향 코멘트"]
    ws5.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers_supply, 1):
        cell = ws5.cell(row=5, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_hdr
        
    supplies_data = [
        ("전국 합계", "만 세대", 1.93, 2.42, "=+(D6-C6)/C6*100", 100.0, "전월 대비 +25.2% 대폭 증가 (총 2.42만 세대 분양 개시)"),
        ("수도권", "만 세대", 1.30, 1.23, "=+(D7-C7)/C7*100", "=+(D7/$D$6)*100", "수도권 분양 1.23만 세대 (비중 50.73%), 정비사업 위주 분양"),
        ("지방", "만 세대", 0.63, 1.19, "=+(D8-C8)/C8*100", "=+(D8/$D$6)*100", "지방 분양 전월 대비 1.88배(88.9%) 급증 (비중 49.27%)")
    ]
    for r_idx, row_vals in enumerate(supplies_data, start=6):
        ws5.row_dimensions[r_idx].height = 22
        is_total = (r_idx == 6)
        fill_r = fill_light_blue if is_total else (fill_light_gray if r_idx % 2 == 0 else fill_white)
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws5.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data_bold if is_total else font_data
            cell.border = border_total if is_total else border_data
            cell.fill = fill_r
            if c_idx in [1, 2]:
                cell.alignment = align_center
            elif c_idx in [3, 4]:
                cell.alignment = align_right
                cell.font = font_data_bold
                if not is_total:
                    cell.fill = fill_light_yellow
                cell.number_format = "#,##0.00"
            elif c_idx in [5, 6]:
                cell.alignment = align_right
                cell.font = font_data_bold
                if not is_total:
                    cell.fill = fill_light_blue
                cell.number_format = "+0.00;-0.00;0.00" if c_idx == 5 else "0.00"
            else:
                cell.alignment = align_left

    ws5["A11"] = "■ [표 2] 당월(9월) 전국 아파트 청약경쟁률 상위 5개 단지 - 출처: 한국부동산원 청약홈"
    ws5["A11"].font = font_section
    
    headers_subscr = ["순위", "단지명", "소재지 (시도 시군구 읍면동)", "총 세대수", "일반공급 호수", "접수 건수", "청약경쟁률\n(:1, 수식)", "사업 형태", "시공사(건설사명)", "청약 흥행 핵심 원인 메모"]
    ws5.row_dimensions[12].height = 26
    for c_idx, h in enumerate(headers_subscr, 1):
        cell = ws5.cell(row=12, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_slate
        cell.alignment = align_center
        cell.border = border_hdr
        
    subscr_top5 = [
        (1, "브라운스톤월곡센트럴", "서울 성북구 하월곡동", 145, 27, 1672, "=+F13/E13", "재건축", "이수건설", "서울 도심 희소성 및 소형 일반공급 분양가 경쟁력으로 흥행 1위"),
        (2, "진주판문지구레이크써밋웰가", "경남 진주시 판문동", 690, 346, 12093, "=+F14/E14", "분양", "흥한주택종합건설", "진주 판문지구 택지개발 입지 및 호수공원 조망권 프리미엄"),
        (3, "검암역푸르지오프라베뉴", "인천 검단구 당하동", 441, 114, 2389, "=+F15/E15", "분양", "대우건설", "검암역세권 및 분양가상한제 적용에 따른 가격 메리트"),
        (4, "삼천하늘채라비엘", "전북 전주시 삼천동1가", 602, 51, 754, "=+F16/E16", "재건축", "코오롱글로벌", "전주 원도심 재건축 정비사업 대단지 브랜드 선호"),
        (5, "더샵분당하이스트", "경기 성남시분당구 정자동", 1149, 68, 790, "=+F17/E17", "리모델링", "포스코이앤씨", "분당 1기 신도시 역세권 리모델링 희소성 및 하이엔드 브랜드")
    ]
    for r_idx, row_vals in enumerate(subscr_top5, start=13):
        ws5.row_dimensions[r_idx].height = 22
        fill_r = fill_light_gray if r_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws5.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            cell.fill = fill_r
            if c_idx in [1, 8, 9]:
                cell.alignment = align_center
                if c_idx == 1:
                    cell.font = font_data_bold
            elif c_idx in [4, 5, 6]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_yellow
                cell.number_format = "#,##0"
            elif c_idx == 7:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_blue
                cell.number_format = "0.00"
            else:
                cell.alignment = align_left

    col_w_ws5 = {"A": 8, "B": 24, "C": 26, "D": 14, "E": 14, "F": 14, "G": 16, "H": 14, "I": 20, "J": 44}
    for col, width in col_w_ws5.items():
        ws5.column_dimensions[col].width = width

    # =========================================================
    # SHEET 6: 06_미분양동향
    # =========================================================
    ws6 = wb.create_sheet(title="06_미분양동향")
    ws6.views.sheetView[0].showGridLines = True
    draw_title_banner(ws6, "06. 분양시장 동향 - 전국 미분양 및 준공후(악성) 미분양 입력 시트",
                      "국토교통부 주택통계(통계누리 stat.molit.go.kr)에서 T-1월 말 확정 미분양 및 준공후 미분양 호수를 입력하십시오.")
    
    ws6["A4"] = "■ [표 1] 미분양 재고 총괄 및 준공 후 미분양 추이 - 출처: 국토교통부 정기 주택통계"
    ws6["A4"].font = font_section
    
    headers_unm = ["미분양 구분", "단위", "전월 수치\n(2026.07)", "당월 수치\n(2026.08)", "전월대비 증감량\n(수식)", "전월대비 증감률\n(%, 수식)", "주요 증감 시도 현황 및 리스크 요인", "건설업계 영향 분석"]
    ws6.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers_unm, 1):
        cell = ws6.cell(row=5, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_hdr
        
    unm_summary = [
        ("전국 미분양 총계", "만 세대", 6.82, 6.91, "=+D6-C6", "=+(D6-C6)/C6*100", "경남(+1,697호), 경북(+772호), 대전(+655호) 증가 주도", "4개월 연속 미분양 누적세, 지방 리스크 확대"),
        (" - 수도권", "만 세대", 1.95, 1.91, "=+D7-C7", "=+(D7-C7)/C7*100", "서울 0.09만호 안정세, 경기 1.45만호, 인천 0.37만호", "수도권은 소폭 완화 흐름 (-0.04만호)"),
        (" - 지방", "만 세대", 4.88, 5.00, "=+D8-C8", "=+(D8-C8)/C8*100", "경남, 경북, 대구 등 지방 악성 적체 지속", "지방 미분양 5.0만호 돌파, 양극화 심화"),
        ("준공 후(악성) 미분양", "만 세대", 2.92, 2.81, "=+D9-C9", "=+(D9-C9)/C9*100", "경기(-284호), 경남(-178호), 경북(-156호) 순 해소", "2개월 연속 감소 (-3.7%), 할인분양 및 임대자구책 효과")
    ]
    for r_idx, row_vals in enumerate(unm_summary, start=6):
        ws6.row_dimensions[r_idx].height = 22
        is_highlight = (r_idx in [6, 9])
        fill_r = fill_light_blue if is_highlight else (fill_light_gray if r_idx % 2 == 0 else fill_white)
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws6.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data_bold if is_highlight else font_data
            cell.border = border_total if is_highlight else border_data
            cell.fill = fill_r
            if c_idx in [1, 2]:
                cell.alignment = align_center if c_idx == 2 else align_left
            elif c_idx in [3, 4]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_yellow
                cell.number_format = "#,##0.00"
            elif c_idx in [5, 6]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.fill = fill_light_blue
                cell.number_format = "+0.00;-0.00;0.00"
            else:
                cell.alignment = align_left

    col_w_ws6 = {"A": 22, "B": 12, "C": 18, "D": 18, "E": 18, "F": 18, "G": 38, "H": 40}
    for col, width in col_w_ws6.items():
        ws6.column_dimensions[col].width = width

    # =========================================================
    # SHEET 7: 07_부록_분양단지
    # =========================================================
    ws7 = wb.create_sheet(title="07_부록_분양단지")
    ws7.views.sheetView[0].showGridLines = True
    draw_title_banner(ws7, "APPENDIX 1. 당월 전국 주요 분양 단지 리스트 입력 시트",
                      "부동산114 분양 캘린더 DB에서 당월(T월) 전국 분양 개시 주요 단지(10~15개)를 입력하십시오.")
    
    ws7["A4"] = "■ 당월 전국 주요 분양 단지 리스트 (총 13개 단지 예시) - 출처: 부동산114 분양 DB"
    ws7["A4"].font = font_section
    
    headers_app_bun = ["번호", "시도", "시군구", "읍면동", "단지명", "총 세대수", "일반분양 호수", "사업 형태", "건설사명(시공사)", "비고 및 특이사항"]
    ws7.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers_app_bun, 1):
        cell = ws7.cell(row=5, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_steel
        cell.alignment = align_center
        cell.border = border_hdr
        
    bun_list = [
        (1, "서울", "성북구", "하월곡동", "브라운스톤월곡센트럴", 145, 27, "재건축", "이수건설", "서울 분양 최고 경쟁률 (61.9:1)"),
        (2, "경기", "성남시분당구", "정자동", "더샵분당하이스트", 1149, 68, "리모델링", "포스코이앤씨", "분당 1기 신도시 하이엔드 리모델링"),
        (3, "경기", "광명시", "소하동", "광명시티프라디움에듀하임", 426, 426, "분양", "시티건설", "광명 소하동 역세권 교육특화"),
        (4, "경기", "부천시소사구", "송내동", "중동역다채로움더원", 128, 45, "재건축", "부성종합건설", "중동역세권 소규모 재건축"),
        (5, "인천", "미추홀구", "학익동", "시티오씨엘9단지오션파크뷰", 1949, 1420, "분양", "HDC현대산업개발 외", "용현학익 1블록 대규모 미니신도시"),
        (6, "인천", "미추홀구", "숭의동", "숭의역노르웨이숲더스카이", 324, 180, "조합", "유림E&C", "숭의역 역세권 주상복합"),
        (7, "충남", "천안시서북구", "부대동", "천안아이파크시티3단지", 896, 896, "분양", "HDC현대산업개발", "성성지구 배후 아이파크 브랜드타운"),
        (8, "전남광주", "북구", "임동", "올뉴챔피언스시티1차", 3216, 1200, "분양", "우미건설", "임동 전남·일신방직 부지 대단지 개발"),
        (9, "울산", "남구", "야음동", "그랑라크에일린의뜰", 1521, 803, "재개발", "아이에스동서", "야음동 B-08구역 재개발 대단지"),
        (10, "울산", "남구", "무거동", "문수로비스타더파크", 700, 310, "분양", "동원개발", "문수로 생활권 및 옥동 학원가 인접"),
        (11, "부산", "수영구", "망미동", "쌍용더플래티넘센텀", 490, 130, "조합", "쌍용건설", "망미2구역 인근 주택재개발"),
        (12, "경남", "진주시", "판문동", "진주판문지구레이크써밋웰가", 690, 346, "분양", "흥한주택종합건설", "진주 판문지구 택지개발 분양 2위 흥행"),
        (13, "전북", "전주시완산구", "삼천동1가", "삼천하늘채라비엘", 602, 51, "재건축", "코오롱글로벌", "삼천동 주공 재건축 정비사업")
    ]
    for r_idx, row_vals in enumerate(bun_list, start=6):
        ws7.row_dimensions[r_idx].height = 20
        fill_r = fill_light_gray if r_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws7.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            cell.fill = fill_r
            if c_idx in [1, 2, 8]:
                cell.alignment = align_center
            elif c_idx in [6, 7]:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.number_format = "#,##0"
            else:
                cell.alignment = align_left

    col_w_ws7 = {"A": 8, "B": 12, "C": 16, "D": 14, "E": 28, "F": 14, "G": 16, "H": 14, "I": 24, "J": 38}
    for col, width in col_w_ws7.items():
        ws7.column_dimensions[col].width = width

    # =========================================================
    # SHEET 8: 08_부록_입주단지
    # =========================================================
    ws8 = wb.create_sheet(title="08_부록_입주단지")
    ws8.views.sheetView[0].showGridLines = True
    draw_title_banner(ws8, "APPENDIX 2. 당월 전국 주요 입주 단지 리스트 입력 시트",
                      "부동산114 입주 캘린더 DB에서 당월(T월) 전국 입주 개시 주요 단지(10~15개)를 입력하십시오.")
    
    ws8["A4"] = "■ 당월 전국 주요 입주 단지 리스트 (총 13개 단지 예시) - 출처: 부동산114 입주 DB"
    ws8["A4"].font = font_section
    
    headers_app_ip = ["번호", "시도", "시군구", "읍면동", "단지명", "총 세대수", "입주 예정월", "사업 형태", "건설사명(시공사)", "비고 및 특이사항"]
    ws8.row_dimensions[5].height = 26
    for c_idx, h in enumerate(headers_app_ip, 1):
        cell = ws8.cell(row=5, column=c_idx, value=h)
        cell.font = font_table_hdr
        cell.fill = fill_steel
        cell.alignment = align_center
        cell.border = border_hdr
        
    ip_list = [
        (1, "서울", "서초구", "방배동", "디에이치방배", 3064, "2026-09", "재건축", "현대건설", "방배5구역 재건축 강남 초대형 랜드마크 입주"),
        (2, "서울", "광진구", "구의동", "강변역센트럴아이파크", 215, "2026-09", "재개발", "HDC현대산업개발", "강변역세권 한양연립가로주택정비"),
        (3, "서울", "성북구", "보문동1가", "보문센트럴아이파크", 199, "2026-09", "재개발", "HDC현대산업개발", "보문제10구역 재개발 주택정비"),
        (4, "경기", "의정부시", "신곡동", "e편한세상신곡시그니처뷰", 815, "2026-09", "재개발", "DL건설", "장암생활권3구역 재개발"),
        (5, "경기", "양주시", "회정동", "회천중앙역대광로제비앙", 621, "2026-09", "분양", "대광건영", "양주회천 택지개발지구 A12블록"),
        (6, "경기", "부천시소사구", "송내동", "송내역푸르지오센트비엔 1단지", 504, "2026-09", "재건축", "대우건설", "송내1-1구역 주택재건축"),
        (7, "경기", "부천시소사구", "송내동", "송내역푸르지오센트비엔 2단지", 541, "2026-09", "재건축", "대우건설", "송내1-1구역 주택재건축"),
        (8, "경기", "파주시", "동패동", "운정중앙역이지더원", 379, "2026-09", "분양", "라인산업", "운정신도시 A44블록 GTX-A 역세권"),
        (9, "인천", "부평구", "갈산동", "중앙하이츠갈산역센트럴", 126, "2026-09", "재건축", "동우개발", "갈산역세권 소규모 재건축"),
        (10, "경남", "창원시성산구", "신월동", "창원센트럴아이파크", 1509, "2026-09", "재건축", "HDC현대산업개발", "신월2구역 주택재건축 대단지"),
        (11, "강원", "춘천시", "동면", "춘천아테라더퍼스트", 543, "2026-09", "분양", "금호건설", "춘천 만천지구 공동주택"),
        (12, "경북", "포항시북구", "항구동", "포항자이디오션", 212, "2026-09", "분양", "GS건설", "영일대 해수욕장 조망 초고층 주상복합"),
        (13, "부산", "연제구", "거제동", "거제역양우내안애아시아드", 168, "2026-09", "분양", "양우건설", "거제동 법조타운 역세권 아파트")
    ]
    for r_idx, row_vals in enumerate(ip_list, start=6):
        ws8.row_dimensions[r_idx].height = 20
        fill_r = fill_light_gray if r_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws8.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_data
            cell.fill = fill_r
            if c_idx in [1, 2, 7, 8]:
                cell.alignment = align_center
            elif c_idx == 6:
                cell.alignment = align_right
                cell.font = font_data_bold
                cell.number_format = "#,##0"
            else:
                cell.alignment = align_left

    col_w_ws8 = {"A": 8, "B": 12, "C": 16, "D": 14, "E": 28, "F": 14, "G": 14, "H": 14, "I": 24, "J": 38}
    for col, width in col_w_ws8.items():
        ws8.column_dimensions[col].width = width

    # 파일 저장
    wb.save(filename)
    print(f"전체 다중 시트 템플릿 생성 완료: {filename}")

if __name__ == "__main__":
    create_full_data_template()
