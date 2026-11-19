# -*- coding: utf-8 -*-
"""
부동산 경매 매물 분석 및 엑셀 매물장 자동 생성 스크립트
15년 차 부동산 경매·매매 실무 애널리스트 모델 적용
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_auction_excel(filename="경매매물_정리_분석표_20261119.xlsx"):
    wb = openpyxl.Workbook()
    
    # -----------------------------
    # Styles & Palettes
    # -----------------------------
    navy_dark = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    navy_mid = PatternFill(start_color="2C4D75", end_color="2C4D75", fill_type="solid")
    blue_header = PatternFill(start_color="335C8D", end_color="335C8D", fill_type="solid")
    blue_light = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    gray_header = PatternFill(start_color="595959", end_color="595959", fill_type="solid")
    gray_light = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    gray_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    
    green_safe = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    yellow_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    red_danger = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    accent_gold = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")
    
    font_title = Font(name="맑은 고딕", size=15, bold=True, color="1B365D")
    font_subtitle = Font(name="맑은 고딕", size=9, color="595959")
    font_header_w = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    font_header_b = Font(name="맑은 고딕", size=10, bold=True, color="1B365D")
    font_sec_title = Font(name="맑은 고딕", size=11, bold=True, color="1B365D")
    font_bold = Font(name="맑은 고딕", size=10, bold=True)
    font_normal = Font(name="맑은 고딕", size=10)
    font_small = Font(name="맑은 고딕", size=9, color="595959")
    font_alert_red = Font(name="맑은 고딕", size=10, bold=True, color="C00000")
    font_alert_green = Font(name="맑은 고딕", size=10, bold=True, color="276A3C")
    
    thin = Side(border_style="thin", color="D9D9D9")
    thick_bottom = Side(border_style="medium", color="1B365D")
    double_bottom = Side(border_style="double", color="1B365D")
    
    border_data = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_header = Border(left=thin, right=thin, top=thin, bottom=thick_bottom)
    border_total = Border(left=thin, right=thin, top=thin, bottom=double_bottom)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    
    # =========================================================================
    # SHEET 1: 종합_매물비교_대시보드
    # =========================================================================
    ws1 = wb.active
    ws1.title = "종합_매물비교_대시보드"
    ws1.views.sheetView[0].showGridLines = True
    
    # 타이틀
    ws1.merge_cells("A1:Q1")
    ws1["A1"] = "🏢 부동산 경매 매물 종합 분석 및 적정 입찰가 산정표"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[1].height = 32
    
    ws1.merge_cells("A2:Q2")
    ws1["A2"] = "분석기준일: 2026-11-19 | 대상: 화성봉담 우방아이유쉘1단지(2025타경56772), 안성대덕 신라아파트(2026타경50882) | 15년차 경매 실무 모델"
    ws1["A2"].font = font_subtitle
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[2].height = 18
    
    # 1행 Header 배치 (행 4)
    headers1 = [
        "사건번호", "소재지 / 물건명", "구분", "전용면적\n(㎡/평)", "층수", "매각기일", 
        "감정평가액\n(A)", "최저매각가\n(비율)", "권리분석\n안전도", "인수권리\n/보증금",
        "최근실거래\n(평균시세)", "보수적\n예상매도가(B)", "추천 입찰가\n(C)", 
        "총 취득원가\n(D, 수식연동)", "세전 예상순익\n(E = B - D)", "예상 수익률\n(E / D)", "투자판단 및 핵심 리스크"
    ]
    
    ws1.row_dimensions[4].height = 30
    for col_idx, h in enumerate(headers1, start=1):
        cell = ws1.cell(row=4, column=col_idx, value=h)
        cell.fill = navy_dark
        cell.font = font_header_w
        cell.alignment = align_center
        cell.border = border_header
        
    # 데이터 행 5: 화성 봉담 우방아이유쉘
    row5 = [
        "2025타경56772",
        "화성시 봉담읍 수영리 674\n우방아이유쉘1단지 102동 1301호",
        "아파트",
        "85.00㎡\n(25.71평)",
        "13 / 23층",
        "2026-11-25",
        405000000,
        "283,500,000\n(70%, 1회유찰)",
        "안전",
        "0원\n(인수권리 없음)",
        420000000,
        410000000,
        330000000,
        "='상세_화성봉담우방아이유쉘'!D22", # D22: 화성 시트 적정시나리오 실질 총 취득원가
        "=L5-N5",
        "=O5/N5",
        "★ 1순위 강력추천\n- 시세(4.2억) 대비 1.2억 저렴한 70% 최저가\n- 등기/임차 권리 깨끗(소유자점유)\n- 목표순익 7,200만원대 확보 가능"
    ]
    
    ws1.row_dimensions[5].height = 55
    for col_idx, val in enumerate(row5, start=1):
        cell = ws1.cell(row=5, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = border_data
        if col_idx in [1, 3, 5, 6, 8, 9, 10]:
            cell.alignment = align_center
        elif col_idx in [2, 17]:
            cell.alignment = align_left
        else:
            cell.alignment = align_right
            
    # 데이터 행 6: 안성 대덕 신라아파트
    row6 = [
        "2026타경50882",
        "안성시 대덕면 내리 675\n신라아파트 5층 나505호",
        "아파트",
        "35.28㎡\n(10.67평)",
        "5 / 15층",
        "2026-11-23",
        85000000,
        "85,000,000\n(100%, 신건)",
        "주의\n(명도/고평가)",
        "0원\n(대항력 없음)",
        74500000,
        72000000,
        55000000,
        "='상세_안성대덕신라아파트'!D22", # D22: 안성 시트 1회 유찰시 추천가 실질 총 취득원가
        "=L6-N6",
        "=O6/N6",
        "⚠️ 1차 신건 입찰금지(패스 권고)\n- 실거래(7,450만)보다 감정가(8,500만) 고평가\n- 신건 입찰시 -1,400만원 역마진\n- 1~2회 유찰 후 5,500만 이하 공략"
    ]
    
    ws1.row_dimensions[6].height = 55
    for col_idx, val in enumerate(row6, start=1):
        cell = ws1.cell(row=6, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = border_data
        if col_idx in [1, 3, 5, 6, 8, 9, 10]:
            cell.alignment = align_center
        elif col_idx in [2, 17]:
            cell.alignment = align_left
        else:
            cell.alignment = align_right
            
    # 금액 및 백분율 서식
    for r in [5, 6]:
        for c in [7, 11, 12, 13, 14, 15]:
            ws1.cell(row=r, column=c).number_format = "#,##0"
        ws1.cell(row=r, column=16).number_format = "0.0%"
        
    # 권리 안전도 컬러링
    ws1.cell(row=5, column=9).fill = green_safe
    ws1.cell(row=5, column=9).font = font_alert_green
    ws1.cell(row=6, column=9).fill = yellow_warn
    ws1.cell(row=6, column=9).font = font_alert_red
    
    # 하단 총평 및 권고사항 박스
    ws1.merge_cells("A8:Q8")
    ws1["A8"] = "📌 애널리스트 실무 핵심 요약 및 전략 브리핑"
    ws1["A8"].font = font_sec_title
    ws1["A8"].fill = blue_light
    ws1["A8"].border = border_header
    ws1.row_dimensions[8].height = 25
    
    briefing_text = (
        "1. [최우선 추천 물건] 2025타경56772 (화성 봉담 우방아이유쉘 1단지 1301호)\n"
        "   - 감정가 4억 500만원 대비 1회 유찰되어 최저가 2억 8,350만원(70%)으로 가격 메리트 극대화.\n"
        "   - 최근 실거래가 4억 2,000만원, 전세가 2억 6,250만~3억원으로 실질 전세가 수준에서 경매 낙찰 가능.\n"
        "   - 등기부상 말소기준권리(2016.12.01 한국주택금융공사) 이후 모든 권리 소멸, 조사된 임차인 없이 소유자 전입세대로 인수보증금 0원.\n"
        "   - 적정 입찰가 3억 3,000만원 산정 시 총취득원가 3억 3,744만원 대비 세전 순이익 약 7,256만원(수익률 21.5%) 확보 가능.\n\n"
        "2. [주의/보류 물건] 2026타경50882 (안성 대덕 신라아파트 나505호)\n"
        "   - 현재 1차 신건(8,500만원)은 실거래 시세(7,450만원)보다 약 1,000만원 이상 비싸게 감정평가된 고평가 물건으로 절대 입찰 금지.\n"
        "   - 외국인 임차인 2명 점유 추정으로 대항력은 없으나 명도 인도명령 및 이사 협의 과정에서 지연 리스크 존재.\n"
        "   - 최소 1회 유찰(최저가 5,950만원) 이후 또는 2회 유찰(최저가 4,165만원) 시점에 월세 수익형(보증금 1,000/월 60만원)으로 공략 추천."
    )
    ws1.merge_cells("A9:Q13")
    ws1["A9"] = briefing_text
    ws1["A9"].font = font_normal
    ws1["A9"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for r in range(9, 14):
        for c in range(1, 18):
            ws1.cell(row=r, column=c).border = border_data
            ws1.cell(row=r, column=c).fill = gray_zebra
            
    # 열 너비 자동 조정
    col_widths1 = [15, 32, 8, 14, 10, 12, 14, 16, 12, 14, 14, 15, 14, 16, 15, 12, 40]
    for idx, w in enumerate(col_widths1, start=1):
        ws1.column_dimensions[get_column_letter(idx)].width = w

    # =========================================================================
    # SHEET 2: 상세_화성봉담우방아이유쉘
    # =========================================================================
    ws2 = wb.create_sheet("상세_화성봉담우방아이유쉘")
    ws2.views.sheetView[0].showGridLines = True
    
    # 타이틀
    ws2.merge_cells("A1:G1")
    ws2["A1"] = "📋 [물건 1] 2025타경56772 화성 봉담 우방아이유쉘 102동 1301호 수지분석표"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[1].height = 30
    
    # 물건 개요 및 권리분석 표
    ws2.merge_cells("A3:G3")
    ws2["A3"] = "1. 매물 기본정보 및 권리분석 요약"
    ws2["A3"].font = font_sec_title
    ws2["A3"].fill = blue_light
    ws2.row_dimensions[3].height = 22
    
    info_table_1 = [
        ("사건번호", "2025타경56772 (수원지법 본원 경매4계)", "매각기일", "2026-11-25 (수)", "물건종류", "아파트 (23층 중 13층)"),
        ("소재지", "경기도 화성시 봉담읍 수영리 674 우방아이유쉘1단지 102동 1301호", "전용면적", "84.9964㎡ (25.71평)", "대지권", "55.3644㎡ (16.75평)"),
        ("감정평가액", "405,000,000원 (100%)", "최저매각가", "283,500,000원 (70%, 1회유찰)", "입찰보증금", "28,350,000원 (10%)"),
        ("말소기준권리", "2016.12.01 한국주택금융공사 근저당 (159,154,000원)", "임차인현황", "조사된 임차인 없음 (소유자 장OO 세대 전입)", "인수권리", "없음 (안전, 등기상 모든 권리 소멸)"),
        ("최근 실거래가", "매매 42,000만원 (2026.10, 20층) / 전세 26,250~30,000만원", "보수적예상가", "410,000,000원", "적정예상가", "420,000,000원"),
        ("특이사항/리스크", "소유자 점유 추정으로 통상적 명도(이사비 협의 150만원) 필요. 체납관리비는 임장 확인 필요(0원 반영)", "", "", "", "")
    ]
    
    for r_idx, row_data in enumerate(info_table_1, start=4):
        ws2.row_dimensions[r_idx].height = 20
        ws2.cell(row=r_idx, column=1, value=row_data[0]).font = font_bold
        ws2.cell(row=r_idx, column=1).fill = gray_light
        ws2.cell(row=r_idx, column=1).alignment = align_center
        ws2.cell(row=r_idx, column=1).border = border_data
        
        if row_data[0] == "특이사항/리스크":
            ws2.merge_cells(f"B{r_idx}:G{r_idx}")
            ws2.cell(row=r_idx, column=2, value=row_data[1]).font = font_normal
            ws2.cell(row=r_idx, column=2).alignment = align_left
            for c in range(2, 8):
                ws2.cell(row=r_idx, column=c).border = border_data
        else:
            ws2.merge_cells(f"B{r_idx}:C{r_idx}")
            ws2.cell(row=r_idx, column=2, value=row_data[1]).font = font_normal
            ws2.cell(row=r_idx, column=2).alignment = align_left
            ws2.cell(row=r_idx, column=2).border = border_data
            ws2.cell(row=r_idx, column=3).border = border_data
            
            ws2.cell(row=r_idx, column=4, value=row_data[2]).font = font_bold
            ws2.cell(row=r_idx, column=4).fill = gray_light
            ws2.cell(row=r_idx, column=4).alignment = align_center
            ws2.cell(row=r_idx, column=4).border = border_data
            
            ws2.cell(row=r_idx, column=5, value=row_data[3]).font = font_normal
            ws2.cell(row=r_idx, column=5).alignment = align_left
            ws2.cell(row=r_idx, column=5).border = border_data
            
            ws2.cell(row=r_idx, column=6, value=row_data[4]).font = font_bold
            ws2.cell(row=r_idx, column=6).fill = gray_light
            ws2.cell(row=r_idx, column=6).alignment = align_center
            ws2.cell(row=r_idx, column=6).border = border_data
            
            ws2.cell(row=r_idx, column=7, value=row_data[5]).font = font_normal
            ws2.cell(row=r_idx, column=7).alignment = align_left
            ws2.cell(row=r_idx, column=7).border = border_data

    # 2. 수식 연동형 수익률 및 적정 입찰가 시뮬레이션
    ws2.merge_cells("A11:G11")
    ws2["A11"] = "2. 수식 연동형 취득원가 및 3단계 입찰 시나리오 분석기 (Excel Formula)"
    ws2["A11"].font = font_sec_title
    ws2["A11"].fill = blue_light
    ws2.row_dimensions[11].height = 22
    
    sim_headers = ["비용 구분 및 계산 항목", "산출 기준 및 수식", "시나리오1: 공격적(최저가)", "시나리오2: 적정 추천가", "시나리오3: 보수적 입찰가", "비고 및 가이드라인", "데이터 검증"]
    ws2.row_dimensions[12].height = 26
    for c_idx, h in enumerate(sim_headers, start=1):
        cell = ws2.cell(row=12, column=c_idx, value=h)
        cell.fill = navy_mid
        cell.font = font_header_w
        cell.alignment = align_center
        cell.border = border_header
        
    sim_rows = [
        # 행 13: 낙찰가
        ("낙찰(매수)예정가", "사용자 입력", 283500000, 330000000, 310000000, "시나리오1은 최저가, 시나리오2는 추천가", "확정"),
        # 행 14: 인수보증금
        ("인수 권리 보증금", "권리분석 결과", 0, 0, 0, "대항력 임차인 없음 (0원)", "확정"),
        # 행 15: 취득세
        ("취득세 및 교육세", "=낙찰가 * 1.1%", "=C13*0.011", "=D13*0.011", "=E13*0.011", "주택가액 6억이하 1주택 기준 (1.1%)", "수식"),
        # 행 16: 법무비
        ("법무사 등기보수", "정액 산출", 800000, 800000, 800000, "이전등기, 국민주택채권 등 대행료", "확정"),
        # 행 17: 명도비
        ("명도비 (이사비/집행)", "협의 기준", 1500000, 1500000, 1500000, "소유자 자진퇴거 협의 이사비", "확정"),
        # 행 18: 수리비
        ("수리 및 인테리어비", "현장확인 필요", 0, 0, 0, "자료 미기재: 확인필요(임장) -> 0원 처리", "임장필요"),
        # 행 19: 체납관리비
        ("체납 관리비", "관리사무소 확인필요", 0, 0, 0, "자료 미기재: 확인필요(임장) -> 0원 처리", "임장필요"),
        # 행 20: 매도 중개수수료 (매도가 C24 * 0.4%)
        ("매도 중개수수료", "=예상매도가 * 0.4%", "=C24*0.004", "=D24*0.004", "=E24*0.004", "부동산 중개보수 요율 (0.4%)", "수식"),
        # 행 21: 금융비용
        ("대출이자 및 금융비용", "대출 70%, 4%, 3개월", 1984500, 2310000, 2170000, "매도 시점까지 3개월 단기 보유 가정", "추정"),
        # 행 22: 실질 총 취득원가 합계
        ("■ 실질 총 취득원가", "=SUM(취득비용 전체)", "=SUM(C13:C21)", "=SUM(D13:D21)", "=SUM(E13:E21)", "낙찰가+인수금+제반부대비용 총합계", "수식(SUM)"),
        # 행 23: 구분선
        ("-----------------------", "-------------------", "-------------------", "-------------------", "-------------------", "-------------------", "-------"),
        # 행 24: 예상 매도가격
        ("예상 매도가격 (기준)", "실거래가 기준", 420000000, 420000000, 410000000, "최근 20층 4.2억 실거래 반영", "확정"),
        # 행 25: 세전 순수익
        ("■ 세전 예상 순수익", "=예상매도가 - 총취득원가", "=C24-C22", "=D24-D22", "=E24-E22", "단기 매매 차익 (양도세 전)", "수식"),
        # 행 26: 총투자금 대비 수익률
        ("■ 총투자금 대비 수익률", "=순수익 / 총취득원가", "=C25/C22", "=D25/D22", "=E25/E22", "수익률 (%)", "수식"),
        # 행 27: 실투자금 (자기자본)
        ("실투자금 (대출70% 가정)", "=총취득원가 - (낙찰가*70%)", "=C22-(C13*0.7)", "=D22-(D13*0.7)", "=E22-(E13*0.7)", "자기자본 투입 예상액", "수식"),
        # 행 28: 실투자금 대비 수익률
        ("■ 실투자금 대비 수익률", "=순수익 / 실투자금", "=C25/C27", "=D25/D27", "=E25/E27", "ROE (자기자본 수익률)", "수식")
    ]
    
    for idx, r_data in enumerate(sim_rows, start=13):
        ws2.row_dimensions[idx].height = 20
        c1 = ws2.cell(row=idx, column=1, value=r_data[0])
        c2 = ws2.cell(row=idx, column=2, value=r_data[1])
        c3 = ws2.cell(row=idx, column=3, value=r_data[2])
        c4 = ws2.cell(row=idx, column=4, value=r_data[3])
        c5 = ws2.cell(row=idx, column=5, value=r_data[4])
        c6 = ws2.cell(row=idx, column=6, value=r_data[5])
        c7 = ws2.cell(row=idx, column=7, value=r_data[6])
        
        c1.border = border_data
        c2.border = border_data
        c3.border = border_data
        c4.border = border_data
        c5.border = border_data
        c6.border = border_data
        c7.border = border_data
        
        c1.font = font_bold if "■" in str(r_data[0]) else font_normal
        c2.font = font_small
        c6.font = font_small
        c7.font = font_small
        
        c1.alignment = align_left
        c2.alignment = align_center
        c6.alignment = align_left
        c7.alignment = align_center
        
        c3.alignment = align_right
        c4.alignment = align_right
        c5.alignment = align_right
        
        # 강조 색상
        if "■ 실질 총 취득원가" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = blue_light
                c.font = font_bold
                c.border = border_total
        elif "■ 세전 예상 순수익" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = accent_gold
                c.font = font_alert_red
        elif "■ 총투자금 대비 수익률" in str(r_data[0]) or "■ 실투자금 대비 수익률" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = green_safe
                c.font = font_alert_green
                
        # 포맷
        if idx in [13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 27]:
            c3.number_format = "#,##0"
            c4.number_format = "#,##0"
            c5.number_format = "#,##0"
        elif idx in [26, 28]:
            c3.number_format = "0.0%"
            c4.number_format = "0.0%"
            c5.number_format = "0.0%"
            
    # 열 너비 설정
    col_widths2 = [24, 25, 22, 22, 22, 35, 14]
    for idx, w in enumerate(col_widths2, start=1):
        ws2.column_dimensions[get_column_letter(idx)].width = w

    # =========================================================================
    # SHEET 3: 상세_안성대덕신라아파트
    # =========================================================================
    ws3 = wb.create_sheet("상세_안성대덕신라아파트")
    ws3.views.sheetView[0].showGridLines = True
    
    # 타이틀
    ws3.merge_cells("A1:G1")
    ws3["A1"] = "📋 [물건 2] 2026타경50882 안성 대덕 신라아파트 나505호 수지분석표"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws3.row_dimensions[1].height = 30
    
    # 매물 기본정보
    ws3.merge_cells("A3:G3")
    ws3["A3"] = "1. 매물 기본정보 및 권리·시세 리스크 분석 요약"
    ws3["A3"].font = font_sec_title
    ws3["A3"].fill = yellow_warn
    ws3.row_dimensions[3].height = 22
    
    info_table_2 = [
        ("사건번호", "2026타경50882 (수원지법 평택지원 경매3계)", "매각기일", "2026-11-23 (월)", "물건종류", "아파트 (15층 중 5층)"),
        ("소재지", "경기도 안성시 대덕면 내리 675 신라아파트 나505호", "전용면적", "35.28㎡ (10.67평)", "대지권", "11.253㎡ (3.40평)"),
        ("감정평가액", "85,000,000원 (100%, 고평가)", "최저매각가", "85,000,000원 (신건, 100%)", "입찰보증금", "8,500,000원 (10%)"),
        ("말소기준권리", "2013.06.28 청주축산업협동조합 근저당 (29,400,000원)", "임차인현황", "외국인 임차인 2인 등재 (JUMINGFENG, JUYIWAN)", "인수권리", "없음 (후순위로 대항력 소멸)"),
        ("최근 실거래가", "매매 7,450~7,500만원 (감정가보다 1,000만원 낮음!)", "월세 시세", "보증금 500~1,500만 / 월 55~70만원 (중앙대 배후)", "시세 리스크", "신건(8,500만) 입찰 시 확정 역마진 발생"),
        ("특이사항/리스크", "외국인 임차인 점유로 인도명령/명도 난이도 '주의'. 신건 입찰 절대 금지, 1~2회 유찰 후 5,500만원 이하 입찰 추천", "", "", "", "")
    ]
    
    for r_idx, row_data in enumerate(info_table_2, start=4):
        ws3.row_dimensions[r_idx].height = 20
        ws3.cell(row=r_idx, column=1, value=row_data[0]).font = font_bold
        ws3.cell(row=r_idx, column=1).fill = gray_light
        ws3.cell(row=r_idx, column=1).alignment = align_center
        ws3.cell(row=r_idx, column=1).border = border_data
        
        if row_data[0] == "특이사항/리스크":
            ws3.merge_cells(f"B{r_idx}:G{r_idx}")
            ws3.cell(row=r_idx, column=2, value=row_data[1]).font = font_alert_red
            ws3.cell(row=r_idx, column=2).alignment = align_left
            for c in range(2, 8):
                ws3.cell(row=r_idx, column=c).border = border_data
        else:
            ws3.merge_cells(f"B{r_idx}:C{r_idx}")
            ws3.cell(row=r_idx, column=2, value=row_data[1]).font = font_normal
            ws3.cell(row=r_idx, column=2).alignment = align_left
            ws3.cell(row=r_idx, column=2).border = border_data
            ws3.cell(row=r_idx, column=3).border = border_data
            
            ws3.cell(row=r_idx, column=4, value=row_data[2]).font = font_bold
            ws3.cell(row=r_idx, column=4).fill = gray_light
            ws3.cell(row=r_idx, column=4).alignment = align_center
            ws3.cell(row=r_idx, column=4).border = border_data
            
            ws3.cell(row=r_idx, column=5, value=row_data[3]).font = font_normal
            ws3.cell(row=r_idx, column=5).alignment = align_left
            ws3.cell(row=r_idx, column=5).border = border_data
            
            ws3.cell(row=r_idx, column=6, value=row_data[4]).font = font_bold
            ws3.cell(row=r_idx, column=6).fill = gray_light
            ws3.cell(row=r_idx, column=6).alignment = align_center
            ws3.cell(row=r_idx, column=6).border = border_data
            
            ws3.cell(row=r_idx, column=7, value=row_data[5]).font = font_normal
            ws3.cell(row=r_idx, column=7).alignment = align_left
            ws3.cell(row=r_idx, column=7).border = border_data

    # 수식 연동 시뮬레이션
    ws3.merge_cells("A11:G11")
    ws3["A11"] = "2. 수식 연동형 시나리오 분석: 1차 신건(역마진) vs 유찰 후 목표 입찰가 비교 (Excel Formula)"
    ws3["A11"].font = font_sec_title
    ws3["A11"].fill = yellow_warn
    ws3.row_dimensions[11].height = 22
    
    sim_headers3 = ["비용 구분 및 계산 항목", "산출 기준 및 수식", "시나리오1: 1차 신건(비추천)", "시나리오2: 1회 유찰시(추천)", "시나리오3: 2회 유찰시(공격)", "비고 및 가이드라인", "데이터 검증"]
    ws3.row_dimensions[12].height = 26
    for c_idx, h in enumerate(sim_headers3, start=1):
        cell = ws3.cell(row=12, column=c_idx, value=h)
        cell.fill = gray_header
        cell.font = font_header_w
        cell.alignment = align_center
        cell.border = border_header
        
    sim_rows3 = [
        # 행 13: 낙찰가
        ("낙찰(매수)예정가", "사용자 입력", 85000000, 55000000, 42000000, "신건은 100%, 시나리오2·3은 유찰가", "확정"),
        # 행 14: 인수보증금
        ("인수 권리 보증금", "권리분석 결과", 0, 0, 0, "후순위 임차인(대항력 없음, 0원)", "확정"),
        # 행 15: 취득세
        ("취득세 및 교육세", "=낙찰가 * 1.1%", "=C13*0.011", "=D13*0.011", "=E13*0.011", "주택가액 6억이하 1주택 기준 (1.1%)", "수식"),
        # 행 16: 법무비
        ("법무사 등기보수", "정액 산출", 800000, 800000, 800000, "소유권이전 등기 대행료", "확정"),
        # 행 17: 명도비
        ("명도비 (이사비/집행)", "협의 기준", 1500000, 1500000, 1500000, "외국인 임차인 퇴거 협의비용", "확정"),
        # 행 18: 수리비
        ("수리 및 인테리어비", "현장확인 필요", 0, 0, 0, "자료 미기재: 확인필요(임장) -> 0원 처리", "임장필요"),
        # 행 19: 체납관리비
        ("체납 관리비", "관리사무소 확인필요", 0, 0, 0, "자료 미기재: 확인필요(임장) -> 0원 처리", "임장필요"),
        # 행 20: 매도 중개수수료 (매도가 C24 * 0.5%)
        ("매도 중개수수료", "=예상매도가 * 0.5%", "=C24*0.005", "=D24*0.005", "=E24*0.005", "5천~2억 미만 주택 중개요율 (0.5%)", "수식"),
        # 행 21: 금융비용
        ("대출이자 및 금융비용", "대출 70%, 4%, 3개월", 595000, 385000, 294000, "매도 시점까지 3개월 단기 보유 가정", "추정"),
        # 행 22: 실질 총 취득원가 합계
        ("■ 실질 총 취득원가", "=SUM(취득비용 전체)", "=SUM(C13:C21)", "=SUM(D13:D21)", "=SUM(E13:E21)", "낙찰가+인수금+제반부대비용 총합계", "수식(SUM)"),
        # 행 23: 구분선
        ("-----------------------", "-------------------", "-------------------", "-------------------", "-------------------", "-------------------", "-------"),
        # 행 24: 예상 매도가격
        ("예상 매도가격 (기준)", "실거래가 기준", 74500000, 74500000, 74500000, "최근 실거래가(7,450만원) 기준", "확정"),
        # 행 25: 세전 순수익
        ("■ 세전 예상 순수익", "=예상매도가 - 총취득원가", "=C24-C22", "=D24-D22", "=E24-E22", "단기 매매 차익 (양도세 전)", "수식"),
        # 행 26: 총투자금 대비 수익률
        ("■ 총투자금 대비 수익률", "=순수익 / 총취득원가", "=C25/C22", "=D25/D22", "=E25/E22", "수익률 (%)", "수식"),
        # 행 27: 실투자금 (자기자본)
        ("실투자금 (대출70% 가정)", "=총취득원가 - (낙찰가*70%)", "=C22-(C13*0.7)", "=D22-(D13*0.7)", "=E22-(E13*0.7)", "자기자본 투입 예상액", "수식"),
        # 행 28: 실투자금 대비 수익률
        ("■ 실투자금 대비 수익률", "=순수익 / 실투자금", "=C25/C27", "=D25/D27", "=E25/E27", "ROE (자기자본 수익률)", "수식")
    ]
    
    for idx, r_data in enumerate(sim_rows3, start=13):
        ws3.row_dimensions[idx].height = 20
        c1 = ws3.cell(row=idx, column=1, value=r_data[0])
        c2 = ws3.cell(row=idx, column=2, value=r_data[1])
        c3 = ws3.cell(row=idx, column=3, value=r_data[2])
        c4 = ws3.cell(row=idx, column=4, value=r_data[3])
        c5 = ws3.cell(row=idx, column=5, value=r_data[4])
        c6 = ws3.cell(row=idx, column=6, value=r_data[5])
        c7 = ws3.cell(row=idx, column=7, value=r_data[6])
        
        c1.border = border_data
        c2.border = border_data
        c3.border = border_data
        c4.border = border_data
        c5.border = border_data
        c6.border = border_data
        c7.border = border_data
        
        c1.font = font_bold if "■" in str(r_data[0]) else font_normal
        c2.font = font_small
        c6.font = font_small
        c7.font = font_small
        
        c1.alignment = align_left
        c2.alignment = align_center
        c6.alignment = align_left
        c7.alignment = align_center
        
        c3.alignment = align_right
        c4.alignment = align_right
        c5.alignment = align_right
        
        # 강조 색상
        if "■ 실질 총 취득원가" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = gray_light
                c.font = font_bold
                c.border = border_total
        elif "■ 세전 예상 순수익" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = red_danger if c == c3 else accent_gold
                c.font = font_alert_red
        elif "■ 총투자금 대비 수익률" in str(r_data[0]) or "■ 실투자금 대비 수익률" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = green_safe if c != c3 else red_danger
                c.font = font_alert_green if c != c3 else font_alert_red
                
        # 포맷
        if idx in [13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 27]:
            c3.number_format = "#,##0"
            c4.number_format = "#,##0"
            c5.number_format = "#,##0"
        elif idx in [26, 28]:
            c3.number_format = "0.0%"
            c4.number_format = "0.0%"
            c5.number_format = "0.0%"
            
    # 열 너비 설정
    for idx, w in enumerate(col_widths2, start=1):
        ws3.column_dimensions[get_column_letter(idx)].width = w
        
    wb.save(filename)
    print(f"Successfully updated auction excel workbook: {filename}")

if __name__ == "__main__":
    build_auction_excel()
