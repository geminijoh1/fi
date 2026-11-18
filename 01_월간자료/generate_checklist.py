# -*- coding: utf-8 -*-
"""
월간주택시장동향(Monthly R Briefing) 작성을 위한 필요 데이터 마스터 체크리스트 엑셀 생성 스크립트
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

def create_checklist_excel(filename="월간주택시장동향_필요데이터_체크리스트.xlsx"):
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # 스타일 공통 정의
    # -------------------------------------------------------------
    font_title = Font(name="맑은 고딕", size=16, bold=True, color="1B365D")
    font_subtitle = Font(name="맑은 고딕", size=10, italic=True, color="555555")
    font_section = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    font_header = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="맑은 고딕", size=9, bold=False, color="222222")
    font_data_bold = Font(name="맑은 고딕", size=9, bold=True, color="222222")
    font_kpi = Font(name="맑은 고딕", size=9, bold=True, color="003366")
    
    # 배경색 채우기
    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # 메인 헤더 다크네이비
    fill_header_sub = PatternFill(start_color="2C4D75", end_color="2C4D75", fill_type="solid") # 보조 헤더
    fill_cat_macro = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid") # 거시/금융
    fill_cat_price = PatternFill(start_color="EAF2EC", end_color="EAF2EC", fill_type="solid") # 가격/거래
    fill_cat_supply = PatternFill(start_color="FFF3E6", end_color="FFF3E6", fill_type="solid") # 분양/미분양
    fill_cat_app = PatternFill(start_color="F3EBF6", end_color="F3EBF6", fill_type="solid") # 부록
    fill_zebra = PatternFill(start_color="F9FAFC", end_color="F9FAFC", fill_type="solid") # 얼룩무늬 행
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    # 상태별 배경색
    fill_status_wait = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 대기(노랑)
    fill_status_done = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid") # 완료(연녹)
    fill_status_prog = PatternFill(start_color="CFE2F3", end_color="CFE2F3", fill_type="solid") # 진행(연파랑)
    
    # 테두리
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thick_border_bottom = Side(border_style="medium", color="1B365D")
    double_border_bottom = Side(border_style="double", color="1B365D")
    
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_header = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_border_bottom)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # =============================================================
    # SHEET 1: 데이터_수집_체크리스트
    # =============================================================
    ws1 = wb.active
    ws1.title = "01_데이터_수집_체크리스트"
    ws1.views.sheetView[0].showGridLines = True
    
    # 제목 영역
    ws1["A1"] = "부동산114 월간주택시장동향(Monthly R Briefing) 필수 데이터 수집 체크리스트"
    ws1["A1"].font = font_title
    ws1["A2"] = "※ 본 체크리스트는 매월 '월간주택시장동향' 보고서 발간 시 누락 없는 원천 데이터 확보 및 교차 검증을 위한 표준 운영 시트입니다."
    ws1["A2"].font = font_subtitle
    ws1.row_dimensions[1].height = 28
    ws1.row_dimensions[2].height = 18
    
    headers_ws1 = [
        "No", "대분류", "중분류", "세부 수집 항목(지표명)", "필수 수집 변수(컬럼)", 
        "단위", "반영 시점\n(Time-Lag)", "원천 출처 기관", "공식 통계명 / 시스템", 
        "데이터 획득 경로 (URL/솔루션)", "팩트체크 & 산출공식 검증 기준", 
        "수집 상태", "담당자", "확인일자", "특이사항 및 이슈 메모"
    ]
    
    ws1.row_dimensions[4].height = 30
    for col_idx, header in enumerate(headers_ws1, 1):
        cell = ws1.cell(row=4, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header
        
    items_ws1 = [
        # 01. 거시경제 현황 - 실물경기
        (1, "01. 거시경제", "실물경기", "동행종합지수", "당월지수, 전월지수, 전월비 증감률(%)", "2020=100, %", "T-1월", "통계청", "산업활동동향 (경기종합지수)", "국가통계포털(KOSIS) / 보도자료", "전월 확정치 개정 여부 확인, 증감률 부호(+/-) 검증", "미착수", "", "", ""),
        (2, "01. 거시경제", "실물경기", "동행순환변동치", "당월치, 전월치, 전월비 증감폭(p)", "p (포인트)", "T-1월", "통계청", "산업활동동향 (경기종합지수)", "국가통계포털(KOSIS) / 보도자료", "순환변동치 추세(상승/하락/보합 국면) 판정, 100 기준 상회 여부", "미착수", "", "", ""),
        (3, "01. 거시경제", "실물경기", "선행종합지수", "당월지수, 전월지수, 전월비 증감률(%)", "2020=100, %", "T-1월", "통계청", "산업활동동향 (경기종합지수)", "국가통계포털(KOSIS) / 보도자료", "전월대비 등락폭 및 건설수주액 구성지표 동향 체크", "미착수", "", "", ""),
        (4, "01. 거시경제", "실물경기", "선행순환변동치", "당월치, 전월치, 전월비 증감폭(p)", "p (포인트)", "T-1월", "통계청", "산업활동동향 (경기종합지수)", "국가통계포털(KOSIS) / 보도자료", "향후 경기 국면 선행 신호(전환점 발생 여부) 분석", "미착수", "", "", ""),
        
        # 01. 거시경제 현황 - 시중유동성
        (5, "01. 거시경제", "시중유동성", "시중통화량 (M2 광의통화 평잔)", "M2 평잔 총액, 전월비 증감액, 전월비 증감률(%), 연속 증가개월수", "조원, %", "T-1월", "한국은행", "통화 및 유동성 지표 (계절조정계열 평잔)", "한국은행 ECOS (1.1.2.2) / 보도자료", "평잔 기준 여부 확인(말잔과 혼동 금지), 증감액 단위(조원) 검증", "미착수", "", "", ""),
        
        # 02. 금융시장 동향 - 금리
        (6, "02. 금융시장", "금리동향", "한국은행 기준금리", "당월 기준금리, 전월대비 변동폭(%p), 통화정책 기조", "연 %, %p", "당월(최신)", "한국은행", "한국은행 기준금리", "한국은행 ECOS (1.3.1) / 금통위 의결문", "보고서 작성일 기준 최신 기준금리 및 금통위 결정 반영", "미착수", "", "", ""),
        (7, "02. 금융시장", "금리동향", "예금은행 가계대출금리", "신규취급액 가중평균금리, 전월대비 등락폭(%p)", "연 %, %p", "T-1월", "한국은행", "금융기관 가중평균금리 (신규취급액)", "한국은행 ECOS (1.3.2.2) / 보도자료", "잔액 기준이 아닌 '신규취급액' 기준 데이터 추출 확인", "미착수", "", "", ""),
        (8, "02. 금융시장", "금리동향", "예금은행 주택담보대출금리", "신규취급액 주담대금리, 전월대비 등락폭(%p)", "연 %, %p", "T-1월", "한국은행", "금융기관 가중평균금리 (주택담보대출)", "한국은행 ECOS (1.3.2.2) / 보도자료", "과거 시계열 대비 고점/저점 비교 (예: N년 M개월 만의 최고치 여부)", "미착수", "", "", ""),
        
        # 02. 금융시장 동향 - 여신잔액
        (9, "02. 금융시장", "여신동향", "예금은행 가계대출 총잔액", "월말잔액, 전월대비 증감액(조원), 연속 증가개월수", "조원", "T-1월 말", "한국은행", "예금은행 대출금 (용도별/가계대출)", "한국은행 ECOS (1.2.1.2) / 금융시장 동향", "조 단위 검증 (원문 억원 오탈자 주의), 잔액 합산 검증", "미착수", "", "", ""),
        (10, "02. 금융시장", "여신동향", "예금은행 주택담보대출 잔액", "월말잔액, 전월대비 증감액(조원)", "조원", "T-1월 말", "한국은행", "예금은행 대출금 (주택담보대출)", "한국은행 ECOS (1.2.1.2) / 금융시장 동향", "가계대출 중 주담대 비중 및 전체 증가분 견인 기여도 계산", "미착수", "", "", ""),
        (11, "02. 금융시장", "여신동향", "예금은행 기타대출 잔액", "월말잔액, 전월대비 증감액(조원, 신용대출 등)", "조원", "T-1월 말", "한국은행", "예금은행 대출금 (일반신용대출 등)", "한국은행 ECOS (1.2.1.2) / 금융시장 동향", "가계대출 증감액 = 주담대 증감액 + 기타대출 증감액 일치 검증", "미착수", "", "", ""),
        
        # 03. 주택시장 동향 - 매매가격
        (12, "03. 매매가격", "가격동향", "전국 아파트 매매가격 변동률", "당월(T월) 변동률, 전월 변동률, 전월비 등락폭(%p)", "%", "당월(T월)", "부동산114", "부동산114 아파트 시세 DB", "PropX 솔루션 / 시세분석 메뉴 / Raw DB", "전월 변동률과 비교하여 상승폭 확대/축소/하락전환 여부 판정", "미착수", "", "", ""),
        (13, "03. 매매가격", "가격동향", "권역별 매매가격 변동률", "수도권 변동률, 지방 변동률, 전월대비 등락폭(%p)", "%", "당월(T월)", "부동산114", "부동산114 아파트 시세 DB", "PropX 솔루션 / 시세분석 / 권역별 비교", "수도권과 지방 간 양극화 및 갭 차이 분석", "미착수", "", "", ""),
        (14, "03. 매매가격", "가격동향", "17개 시도별 매매가격 변동률", "전국 17개 시도별 변동률 리스트, 순위(상위/하위)", "%", "당월(T월)", "부동산114", "부동산114 아파트 시세 DB", "PropX 솔루션 / 지역별 시세 랭킹", "상승 시도 개수 vs 하락 시도 개수 산출, 상위 3개 및 하위 3개 선정", "미착수", "", "", ""),
        (15, "03. 매매가격", "가격동향", "주요 시군구 특이 급등/급락 지역", "급등/급락 시군구 변동률, 국지적 원인(재건축/교통 등)", "%", "당월(T월)", "부동산114", "부동산114 아파트 시세 DB", "PropX 솔루션 / 시군구 상세분석", "기사 및 호재 교차 검증 (예: 노후계획도시 선도지구, 규제지역 등)", "미착수", "", "", ""),
        
        # 04. 주택시장 동향 - 매매거래
        (16, "04. 매매거래", "거래량", "전국 아파트 매매거래량", "전국 총 거래량, 전월 거래량, 전월비 증감률(%)", "세대 (만 세대), %", "T-1월 확정", "국토교통부", "국토교통부 실거래가 신고 데이터", "국토부 실거래가 공개시스템 / R-ONE / PropX", "실거래 신고기한(30일)에 따른 1개월 시차 적용 확인 (만 세대 단위 환산)", "미착수", "", "", ""),
        (17, "04. 매매거래", "거래량", "17개 시도별 매매거래량 및 비중", "시도별 거래량, 전월비 증감률(%), 전국 점유비중(%)", "세대 (만 세대), %", "T-1월 확정", "국토교통부", "국토교통부 실거래가 신고 데이터", "국토부 실거래가 공개시스템 / R-ONE / PropX", "17개 시도 거래량 합 = 전국 합계 일치 검증, 상위 시도 집중도 계산", "미착수", "", "", ""),
        (18, "04. 매매거래", "거래량", "주택 규모(전용면적)별 거래 비중", "소형(60㎡ 이하), 중형(60~85㎡), 대형(85㎡ 초과) 거래 비중", "%", "T-1월 확정", "국토교통부", "국토교통부 실거래가 신고 데이터", "국토부 실거래가 공개시스템 / R-ONE / PropX", "지역별 선호 평형(서울 소형 vs 경기/지방 중형) 차별화 특성 도출", "미착수", "", "", ""),
        
        # 05. 분양시장 동향 - 공급 및 청약
        (19, "05. 분양시장", "공급실적", "전국 아파트 분양물량", "전국 총 분양세대수, 전월 분양물량, 전월비 증감률(%)", "세대 (만 세대), %", "당월(T월)", "부동산114", "부동산114 분양 조사 DB", "PropX 솔루션 / 분양 및 공급 분석", "임대 제외 순수 분양(일반+특별) 집계 기준 확인", "미착수", "", "", ""),
        (20, "05. 분양시장", "공급실적", "권역별 분양물량 및 공급 비중", "수도권 분양호수, 지방 분양호수, 권역별 비중(%)", "세대 (만 세대), %", "당월(T월)", "부동산114", "부동산114 분양 조사 DB", "PropX 솔루션 / 분양 및 공급 분석", "수도권 비중 + 지방 비중 = 100.0% 합산 검증", "미착수", "", "", ""),
        (21, "05. 분양시장", "청약결과", "지역별 평균 청약경쟁률", "17개 시도별 1순위 청약경쟁률 순위 및 수치", "대 1 (:1)", "당월(T월)", "한국부동산원", "청약홈 청약접수 및 경쟁률 통계", "청약홈 (applyhome.co.kr) / 부동산114 분양 DB", "지역별 청약 양극화 현상(서울/수도권 vs 지방 미달 단지) 점검", "미착수", "", "", ""),
        (22, "05. 분양시장", "청약결과", "당월 청약경쟁률 상위 5개 단지", "단지명, 소재지, 총세대, 일반공급, 청약경쟁률, 사업형태, 시공사", "세대, 건, :1", "당월(T월)", "한국부동산원", "청약홈 아파트 분양결과", "청약홈 (applyhome.co.kr) / 부동산114 분양 DB", "일반공급 세대수 대비 접수건수 정확도 검증, 순위 산정", "미착수", "", "", ""),
        
        # 06. 분양시장 동향 - 미분양
        (23, "06. 미분양", "재고현황", "전국 아파트 미분양 총 물량", "전국 미분양 호수, 전월 물량, 전월비 증감량, 증감률(%)", "세대 (만 세대), %", "T-1월 말", "국토교통부", "국토교통부 주택통계 (미분양 주택 현황)", "국토교통부 통계누리 / 보도자료", "국토부 정기 주택통계(매월 말 발표) T-1월 확정치 사용", "미착수", "", "", ""),
        (24, "06. 미분양", "재고현황", "권역 및 시도별 미분양 현황", "수도권/지방 물량, 17개 시도별 미분양 및 증감량", "세대 (만 세대)", "T-1월 말", "국토교통부", "국토교통부 주택통계 (미분양 주택 현황)", "국토교통부 통계누리 / 보도자료", "미분양 누적 증가 시도(위험지역) vs 해소 시도 분리 추출", "미착수", "", "", ""),
        (25, "06. 미분양", "준공후", "준공 후(악성) 미분양 물량 및 증감", "전국 준공후미분양 호수, 전월비 증감량, 증감률(%), 시도별 현황", "세대 (만 세대), %", "T-1월 말", "국토교통부", "국토교통부 주택통계 (준공후 미분양 현황)", "국토교통부 통계누리 / 보도자료", "건설업계 유동성 리스크 핵심 지표, 감소 원인(할인분양/임대 전환) 파악", "미착수", "", "", ""),
        
        # 07. 부록 - 단지 리스트
        (26, "07. 부록", "분양단지", "당월 전국 주요 분양 단지 리스트", "번호, 시도, 시군구, 읍면동, 단지명, 총세대수, 사업형태, 건설사명", "목록 (10~15개)", "당월(T월)", "부동산114", "부동산114 분양 캘린더 DB", "PropX 솔루션 / 분양 캘린더", "대단지 및 브랜드 건설사 위주 선별(10~15개), 시공사명 정제", "미착수", "", "", ""),
        (27, "07. 부록", "입주단지", "당월 전국 주요 입주 단지 리스트", "번호, 시도, 시군구, 읍면동, 단지명, 총세대수, 사업형태, 건설사명", "목록 (10~15개)", "당월(T월)", "부동산114", "부동산114 입주 캘린더 DB", "PropX 솔루션 / 입주 캘린더", "당월 입주 예정 단지 중 권역별 대표 대단지 선별", "미착수", "", "", "")
    ]
    
    start_row = 5
    for r_idx, item in enumerate(items_ws1, start=start_row):
        ws1.row_dimensions[r_idx].height = 24
        cat = item[1]
        
        # 행 배경색 결정
        if "01. 거시경제" in cat or "02. 금융시장" in cat:
            cat_fill = fill_cat_macro
        elif "03. 매매가격" in cat or "04. 매매거래" in cat:
            cat_fill = fill_cat_price
        elif "05. 분양시장" in cat or "06. 미분양" in cat:
            cat_fill = fill_cat_supply
        else:
            cat_fill = fill_cat_app
            
        row_fill = fill_zebra if r_idx % 2 == 0 else fill_white
        
        for c_idx, val in enumerate(item, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_cell
            
            # 정렬 및 서식
            if c_idx in [1, 6, 7, 12, 13, 14]: # No, 단위, 시점, 상태, 담당자, 일자
                cell.alignment = align_center
            elif c_idx in [2, 3, 8]: # 대분류, 중분류, 출처
                cell.alignment = align_center
                if c_idx == 2:
                    cell.font = font_data_bold
            else:
                cell.alignment = align_left
                
            # 수집상태 셀 배경색
            if c_idx == 12:
                cell.fill = fill_status_wait
                cell.font = font_data_bold
            elif c_idx in [2, 3]:
                cell.fill = cat_fill
            else:
                cell.fill = row_fill

    # 데이터 유효성 검사 (드롭다운: 미착수, 수집중, 수집완료, 검증완료)
    dv = DataValidation(type="list", formula1='"미착수,수집중,수집완료,검증완료"', allow_blank=True)
    ws1.add_data_validation(dv)
    dv.add(f"L{start_row}:L{start_row + len(items_ws1) - 1}")
    
    # 열 너비 설정
    col_widths_ws1 = {
        "A": 6,   # No
        "B": 15,  # 대분류
        "C": 12,  # 중분류
        "D": 26,  # 지표명
        "E": 32,  # 필수 수집 변수
        "F": 14,  # 단위
        "G": 15,  # 반영 시점
        "H": 14,  # 원천 출처
        "I": 28,  # 통계명
        "J": 30,  # URL/경로
        "K": 38,  # 검증 기준
        "L": 14,  # 수집 상태
        "M": 12,  # 담당자
        "N": 14,  # 확인일자
        "O": 22   # 메모
    }
    for col_letter, width in col_widths_ws1.items():
        ws1.column_dimensions[col_letter].width = width

    # =============================================================
    # SHEET 2: 02_시점매핑_및_TimeLag
    # =============================================================
    ws2 = wb.create_sheet(title="02_시점매핑_및_TimeLag")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2["A1"] = "월간주택시장동향 지표별 공표 일정 및 시점 매핑 (Time-Lag 분석 가이드)"
    ws2["A1"].font = font_title
    ws2["A2"] = "※ 보고서 발간 기준월(T월) 대비 공표 주관기관의 데이터 수집 시점 및 시차(Lag) 발생 원인 정리"
    ws2["A2"].font = font_subtitle
    ws2.row_dimensions[1].height = 28
    ws2.row_dimensions[2].height = 18
    
    headers_ws2 = [
        "지표 구분", "세부 통계 지표", "보고서 반영 시점\n(T월호 기준)", 
        "공표 주기", "통상 공표 일시", "출처 시스템", "Time-Lag 발생 원인 및 실무 가이드라인"
    ]
    
    ws2.row_dimensions[4].height = 28
    for col_idx, header in enumerate(headers_ws2, 1):
        cell = ws2.cell(row=4, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header_sub
        cell.alignment = align_center
        cell.border = border_header
        
    timelag_data = [
        ("01. 거시경제", "경기종합지수 (동행/선행)", "T-1월", "월간", "익월 말 (통상 28~31일)", "통계청 산업활동동향", "생산·소비·투자 등 실물 경제활동 통계 집계에 약 1개월이 소요되므로 전월 확정치 반영"),
        ("01. 거시경제", "시중통화량 (M2 광의통화)", "T-1월", "월간", "익월 중순 (통상 13~16일)", "한국은행 ECOS", "시중 예금취급기관의 통화량 결산에 약 40일이 소요되므로 전월 평잔 수치 반영"),
        ("02. 금융시장", "한국은행 기준금리", "당월(T월 최신)", "연 8회", "금통위 회의일 (해당 월)", "한국은행 공시", "금융통화위원회 통화정책방향 결정 회의 당일 즉시 반영 (작성 시점 최신값)"),
        ("02. 금융시장", "예금은행 대출금리 (신규)", "T-1월", "월간", "익월 말 (통상 28~31일)", "한국은행 ECOS", "전체 시중은행의 신규 취급 대출 가중평균 금리 집계에 1개월 소요"),
        ("02. 금융시장", "예금은행 가계/주담대 잔액", "T-1월 말", "월간", "익월 초~중순 (통상 10~12일)", "한국은행 금융시장동향", "은행권 대출 잔액 결산에 약 10영업일 소요 (월말 확정 잔액 반영)"),
        ("03. 주택가격", "아파트 매매가격 변동률", "당월(T월)", "월간/주간", "당월 말 (익월 1일 확정)", "부동산114 시세 DB / PropX", "부동산114 자체 시세 조사망(현장 중개업소 조사)을 활용하여 당월 시세 변동 즉시 반영"),
        ("04. 매매거래", "아파트 매매거래량", "T-1월 확정", "월간", "익월 말 (신고 완료 후)", "국토부 실거래가 / R-ONE", "부동산 거래신고 등에 관한 법률상 실거래 신고기한이 계약일로부터 '30일'이므로 1개월 시차 필수"),
        ("05. 분양시장", "분양공급 물량 실적", "당월(T월)", "월간", "당월 말 집계", "부동산114 분양 DB / PropX", "당월 모집공고 승인 및 입주자모집을 개시한 현장 전수 집계치 즉시 반영"),
        ("05. 분양시장", "청약 경쟁률 (단지별)", "당월(T월)", "수시/월간", "청약접수 종료 익일", "한국부동산원 청약홈", "당월 청약 일정이 마감된 단지의 1순위 접수건수 및 일반공급 호수 기반 산출"),
        ("06. 미분양", "전국 및 시도별 미분양", "T-1월 말", "월간", "익월 말 (통상 28~31일)", "국토교통부 주택통계", "전국 지자체가 관할 구역 내 미분양 현황을 취합하여 국토부에 보고하는 행정 시차 발생"),
        ("06. 미분양", "준공 후(악성) 미분양", "T-1월 말", "월간", "익월 말 (통상 28~31일)", "국토교통부 주택통계", "미분양 통계와 동일하게 행정 취합 시차로 인해 T-1월 말 수치가 최종 확정치임"),
        ("07. 부록", "주요 분양/입주 단지", "당월(T월)", "월간", "당월 말 기준", "부동산114 캘린더 DB", "당월 분양 개시 단지 및 입주 개시 단지 리스트 추출")
    ]
    
    start_row_ws2 = 5
    for r_idx, item in enumerate(timelag_data, start=start_row_ws2):
        ws2.row_dimensions[r_idx].height = 24
        row_fill = fill_zebra if r_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(item, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_cell
            cell.fill = row_fill
            if c_idx in [1, 3, 4, 5]:
                cell.alignment = align_center
                if c_idx == 3:
                    cell.font = font_data_bold
            elif c_idx == 2:
                cell.alignment = align_left
                cell.font = font_data_bold
            else:
                cell.alignment = align_left
                
    col_widths_ws2 = {
        "A": 16,  # 구분
        "B": 24,  # 통계지표
        "C": 18,  # 반영시점
        "D": 12,  # 공표주기
        "E": 20,  # 공표일시
        "F": 26,  # 출처
        "G": 48   # 원인/가이드
    }
    for col_letter, width in col_widths_ws2.items():
        ws2.column_dimensions[col_letter].width = width

    # =============================================================
    # SHEET 3: 03_산출공식_및_데이터검증룰
    # =============================================================
    ws3 = wb.create_sheet(title="03_산출공식_및_데이터검증룰")
    ws3.views.sheetView[0].showGridLines = True
    
    ws3["A1"] = "월간주택시장동향 데이터 무결성 검증 룰(Rule) 및 산출 공식"
    ws3["A1"].font = font_title
    ws3["A2"] = "※ 보고서 작성 시 데이터 오기, 단위 불일치, 합계 불일치 등의 오류를 사전에 차단하는 5대 검증 수칙"
    ws3["A2"].font = font_subtitle
    ws3.row_dimensions[1].height = 28
    ws3.row_dimensions[2].height = 18
    
    headers_ws3 = [
        "검증 번호", "검증 영역", "검증 룰 명칭", "표준 산출 공식 / 검증 로직", 
        "오류 발생 예시 (주의사항)", "사전 차단 및 조치 가이드"
    ]
    
    ws3.row_dimensions[4].height = 28
    for col_idx, header in enumerate(headers_ws3, 1):
        cell = ws3.cell(row=4, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header_sub
        cell.alignment = align_center
        cell.border = border_header
        
    rules_data = [
        ("RULE-01", "증감 표기", "증감률(%) vs 등락폭(%p) 구분", "증감률(%) = (당월값 - 전월값) / 전월값 * 100\n등락폭(%p) = 당월 변동률(%) - 전월 변동률(%)", "변동률(%) 차이를 %로 잘못 표기하거나 포인트(p)를 %로 혼용", "지수/금리/변동률의 단순 차이는 반드시 '%p' 또는 'p'로 표기하고, 일반 거래량/공급량 증감은 '%'로 엄격히 구분"),
        ("RULE-02", "합계 정합성", "권역별 점유 비중 100% 검증", "수도권 비중(%) + 지방 비중(%) = 100.0%\n지역별 거래량/분양량 비중 합계 = 100.0%", "소수점 반올림 오차로 합계가 99.9% 또는 100.1%가 되는 현상 방치", "소수점 첫째자리까지 합산하여 100.0%가 되는지 검증하고 불일치 시 가장 비중이 큰 지역의 잔여치를 조정하여 100.0% 맞춤"),
        ("RULE-03", "단위 무결성", "조원 vs 억원 단위 누락 검증", "가계대출 총잔액 = 주담대 잔액 + 기타대출 잔액\n단위: 조원 (예: 1,198.3조원)", "보도자료 원문의 오탈자 '-0.6억원'을 그대로 전재하여 6조원 단위의 심각한 오류 발생", "금융 데이터 입력 시 총잔액 = 주담대 + 기타대출의 등식이 성립하는지 확인. 억원/조원 단위 스케일 점검"),
        ("RULE-04", "시점 일관성", "지표별 Time-Lag 혼동 방지", "매매가격: T월 / 거래량: T-1월 / 미분양: T-1월 / 거시금융: T-1월", "9월호 보고서에 8월 매매가격을 넣거나 9월 거래량(미확정치)을 사용하는 오류", "표준 가이드라인의 시점 매핑표를 기준으로 헤더에 반드시 해당 데이터의 실제 기준월을 명시"),
        ("RULE-05", "결측치 처리", "미발표 데이터 임의 추정 금지", "데이터 미발표 또는 수집 불가 시 '[데이터 확인 필요]' 표기", "발표 지연된 통계를 이전 수치로 임의 복사하거나 추정치로 기입하여 팩트 훼손", "통계 발표일이 보고서 마감일 이후인 경우 전월 수치를 유지하며 주석에 'T월 수치 발표 지연으로 T-1월치 인용' 명시")
    ]
    
    start_row_ws3 = 5
    for r_idx, item in enumerate(rules_data, start=start_row_ws3):
        ws3.row_dimensions[r_idx].height = 42
        row_fill = fill_zebra if r_idx % 2 == 0 else fill_white
        for c_idx, val in enumerate(item, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_cell
            cell.fill = row_fill
            if c_idx in [1, 2]:
                cell.alignment = align_center
                cell.font = font_data_bold
            elif c_idx == 3:
                cell.alignment = align_left
                cell.font = font_data_bold
            else:
                cell.alignment = align_left
                
    col_widths_ws3 = {
        "A": 14,  # 번호
        "B": 14,  # 영역
        "C": 26,  # 룰 명칭
        "D": 38,  # 공식/로직
        "E": 34,  # 오류 예시
        "F": 46   # 조치 가이드
    }
    for col_letter, width in col_widths_ws3.items():
        ws3.column_dimensions[col_letter].width = width

    # 파일 저장
    wb.save(filename)
    print(f"엑셀 파일 생성 완료: {filename}")

if __name__ == "__main__":
    create_checklist_excel()
