# -*- coding: utf-8 -*-
"""
경매 매물 자동화 마스터 시스템 (Master Auction DB Engine)
- 기능:
  1. '경매매물' 폴더 내 모든 경매 PDF 스캔
  2. 기등록 사건번호 중복 체크 후 신규 매물만 '경매_마스터_통합관리장.xlsx'의 [매물_통합목록]에 자동 추가(Append)
  3. [상세_수지계산기] 시트의 드롭다운 및 VLOOKUP 수식 연동
"""
import os
import re
import glob
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

MASTER_EXCEL_FILE = "경매_마스터_통합관리장.xlsx"
AUCTION_DIR = "경매매물"

# -------------------------------------------------------------
# 공통 디자인 팔레트 및 서식
# -------------------------------------------------------------
navy_dark = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
navy_mid = PatternFill(start_color="2C4D75", end_color="2C4D75", fill_type="solid")
blue_light = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
gray_header = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
gray_light = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
gray_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
yellow_accent = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
green_safe = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
red_danger = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

font_title = Font(name="맑은 고딕", size=15, bold=True, color="1B365D")
font_subtitle = Font(name="맑은 고딕", size=9, color="595959")
font_header_w = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
font_sec_title = Font(name="맑은 고딕", size=11, bold=True, color="1B365D")
font_bold = Font(name="맑은 고딕", size=10, bold=True)
font_normal = Font(name="맑은 고딕", size=10)
font_small = Font(name="맑은 고딕", size=9, color="595959")
font_green_b = Font(name="맑은 고딕", size=10, bold=True, color="276A3C")
font_red_b = Font(name="맑은 고딕", size=10, bold=True, color="C00000")

thin = Side(border_style="thin", color="D9D9D9")
thick_b = Side(border_style="medium", color="1B365D")
double_b = Side(border_style="double", color="1B365D")

border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
border_header = Border(left=thin, right=thin, top=thin, bottom=thick_b)
border_total = Border(left=thin, right=thin, top=thin, bottom=double_b)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")


def extract_auction_pdf_data(pdf_path):
    """PDF 파일에서 핵심 경매 팩트 데이터를 추출하여 딕셔너리로 반환"""
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += "\n" + t

    # 1. 사건번호
    case_no_match = re.search(r"(\d{4}\s*타경\s*\d+)", full_text)
    case_no = case_no_match.group(1).replace(" ", "") if case_no_match else "사건번호확인필요"

    # 2. 법원 / 계
    court_match = re.search(r"(\d{4}\s*타경\s*\d+)\s+([^\n\r]+?계)", full_text)
    court = court_match.group(2).strip() if court_match else "법원확인필요"

    # 3. 매각기일
    date_match = re.search(r"매각기일은\s*(\d{4}-\d{2}-\d{2})", full_text)
    if not date_match:
        date_match = re.search(r"매각기일\s*[:\s]*(\d{4}[-.]\d{2}[-.]\d{2})", full_text)
    sale_date = date_match.group(1).replace(".", "-") if date_match else "확인필요"

    # 4. 소재지
    loc_match = re.search(r"소재지\s*([^\n\r]+?)(?:지도보기|\n|$)", full_text)
    location = loc_match.group(1).strip() if loc_match else "소재지확인필요"

    # 5. 전용면적 / 대지권
    area_match = re.search(r"건물면적\s*([\d.]+)\s*㎡\s*\(\s*([\d.]+)\s*평\)", full_text)
    if area_match:
        area_m2 = float(area_match.group(1))
        area_py = float(area_match.group(2))
    else:
        area_m2, area_py = 0.0, 0.0

    # 6. 감정가 및 최저가
    appraisal_match = re.search(r"감정가\s*([\d,]+)\s*원", full_text)
    appraisal_val = int(appraisal_match.group(1).replace(",", "")) if appraisal_match else 0

    min_bid_match = re.search(r"최저가\s*\([^\)]*\)\s*([\d,]+)\s*원", full_text)
    if not min_bid_match:
        min_bid_match = re.search(r"최저가\s*([\d,]+)\s*원", full_text)
    min_bid_val = int(min_bid_match.group(1).replace(",", "")) if min_bid_match else appraisal_val

    # 최저가 비율
    ratio = round((min_bid_val / appraisal_val) * 100) if appraisal_val > 0 else 100
    status_str = f"{ratio}% (신건)" if ratio >= 100 else f"{ratio}% ({1 if ratio==70 else 2}회유찰)"

    # 7. 층수
    floor_match = re.search(r"(\d+)층(?:제?\d+호|나\d+호|\d+호)", location)
    total_floor_match = re.search(r"(\d+)층\s*건\s*내\s*제?(\d+)층", full_text)
    if total_floor_match:
        floor_str = f"{total_floor_match.group(2)} / {total_floor_match.group(1)}층"
    elif floor_match:
        floor_str = f"{floor_match.group(1)}층"
    else:
        floor_str = "확인필요"

    # 8. 권리분석 및 임차인
    safe_status = "안전"
    deposit_to_take = 0
    risk_summary = "등기부상 말소기준권리 이후 소멸, 인수권리 없음"

    if "외국인체류확인서" in full_text or "외국인 2인" in full_text:
        safe_status = "주의"
        risk_summary = "외국인 임차인 등재(대항력 소멸 추정), 명도 난이도 주의"
    elif "조사된 임차내역이 없습니다" in full_text:
        safe_status = "안전"
        risk_summary = "조사된 임차인 없음 (소유자 점유 추정, 이사비 협의 필요)"

    # 9. 실거래가 / 매도가 추정
    trade_match = re.search(r"최근실거래가\s*:\s*([\d,]+)\s*원", full_text)
    if trade_match:
        real_trade_val = int(trade_match.group(1).replace(",", ""))
    else:
        # 국토부 실거래 텍스트에서 검색
        recent_trades = re.findall(r"매\s*([\d,]+)", full_text)
        if recent_trades:
            real_trade_val = int(recent_trades[0].replace(",", "")) * 10000
        else:
            real_trade_val = appraisal_val

    # 보수적 예상매도가 (최근 실거래가의 97~98% 또는 적정가)
    expected_sale_val = int(real_trade_val * 0.98)

    # 추천 입찰가
    if ratio >= 100 and real_trade_val < appraisal_val:
        # 고평가 신건 -> 유찰 권고 (1회 유찰가 이하)
        recommended_bid = int(appraisal_val * 0.65)
        decision = "⚠️ 1차 입찰금지 (시세 역마진, 유찰 대기)"
        stage = "유찰대기"
    else:
        # 적정 유찰 상태
        recommended_bid = int(min_bid_val * 1.15) if min_bid_val * 1.15 < expected_sale_val * 0.85 else min_bid_val
        decision = "★ 1순위 추천 (실거래 대비 안전마진 확보)"
        stage = "신규검토"

    return {
        "case_no": case_no,
        "court": court,
        "location": location,
        "category": "아파트",
        "area_m2": area_m2,
        "area_py": area_py,
        "floor": floor_str,
        "sale_date": sale_date,
        "appraisal_val": appraisal_val,
        "min_bid_val": min_bid_val,
        "status_str": status_str,
        "safe_status": safe_status,
        "deposit_to_take": deposit_to_take,
        "real_trade_val": real_trade_val,
        "expected_sale_val": expected_sale_val,
        "recommended_bid": recommended_bid,
        "decision": decision,
        "stage": stage,
        "risk_summary": risk_summary
    }


def create_or_update_master_db():
    """마스터 DB 엑셀을 생성하거나 기존 파일에 신규 매물을 누적 추가"""
    is_new = not os.path.exists(MASTER_EXCEL_FILE)
    
    if is_new:
        wb = openpyxl.Workbook()
        ws_master = wb.active
        ws_master.title = "매물_통합목록"
        ws_calc = wb.create_sheet("상세_수지계산기")
    else:
        wb = openpyxl.load_workbook(MASTER_EXCEL_FILE)
        ws_master = wb["매물_통합목록"]
        ws_calc = wb["상세_수지계산기"]

    ws_master.views.sheetView[0].showGridLines = True
    ws_calc.views.sheetView[0].showGridLines = True

    # ---------------------------------------------------------
    # 1. [매물_통합목록] 시트 헤더 초기화 (신규 생성 시)
    # ---------------------------------------------------------
    headers_master = [
        "순번", "사건번호", "법원/계", "소재지 및 단지명", "구분",
        "전용면적(㎡)", "전용(평)", "층수", "매각기일",
        "감정평가액", "최저매각가", "진행상태", "권리안전도", "인수권리금",
        "최근실거래가", "예상매도가", "추천입찰가",
        "총 취득원가\n(수식)", "세전 예상순익\n(수식)", "수익률\n(수식)",
        "진행단계", "투자판단 및 핵심리스크"
    ]

    if is_new:
        ws_master.merge_cells("A1:V1")
        ws_master["A1"] = "📋 경매 매물 통합 관리 마스터 DB (가계부 목록)"
        ws_master["A1"].font = font_title
        ws_master.row_dimensions[1].height = 30

        ws_master.merge_cells("A2:V2")
        ws_master["A2"] = "신규 PDF가 추가될 때마다 아래 행으로 자동 누적됩니다. 상세 시뮬레이션은 [상세_수지계산기] 탭을 이용하세요."
        ws_master["A2"].font = font_subtitle
        ws_master.row_dimensions[2].height = 18

        ws_master.row_dimensions[4].height = 28
        for col_idx, h in enumerate(headers_master, start=1):
            cell = ws_master.cell(row=4, column=col_idx, value=h)
            cell.fill = navy_dark
            cell.font = font_header_w
            cell.alignment = align_center
            cell.border = border_header

    # 기등록 사건번호 목록 파악
    existing_case_nos = set()
    max_row = ws_master.max_row
    if max_row >= 5:
        for r in range(5, max_row + 1):
            v = ws_master.cell(row=r, column=2).value
            if v:
                existing_case_nos.add(str(v).strip())

    # ---------------------------------------------------------
    # 2. PDF 스캔 및 신규 데이터 누적 (Append)
    # ---------------------------------------------------------
    pdf_files = glob.glob(os.path.join(AUCTION_DIR, "*.pdf"))
    added_count = 0
    current_row = 5
    while ws_master.cell(row=current_row, column=2).value:
        current_row += 1

    for pdf_path in pdf_files:
        try:
            data = extract_auction_pdf_data(pdf_path)
            case_no = data["case_no"]
            if case_no in existing_case_nos:
                continue  # 이미 등록된 물건 건너뜀 (중복 방지)

            seq = current_row - 4
            ws_master.row_dimensions[current_row].height = 25

            # 수식 설정
            # Q열: 추천입찰가, N열: 인수권리금, P열: 예상매도가
            # 취득원가(R열) = Q + N + (Q*1.1%) + 800000(법무) + 1500000(명도) + (P*0.4%)(중개)
            cost_formula = f"=Q{current_row}+N{current_row}+(Q{current_row}*0.011)+800000+1500000+(P{current_row}*0.004)"
            profit_formula = f"=P{current_row}-R{current_row}"
            roi_formula = f"=S{current_row}/R{current_row}"

            row_values = [
                seq, data["case_no"], data["court"], data["location"], data["category"],
                data["area_m2"], data["area_py"], data["floor"], data["sale_date"],
                data["appraisal_val"], data["min_bid_val"], data["status_str"],
                data["safe_status"], data["deposit_to_take"], data["real_trade_val"],
                data["expected_sale_val"], data["recommended_bid"],
                cost_formula, profit_formula, roi_formula,
                data["stage"], data["decision"] + " | " + data["risk_summary"]
            ]

            for c_idx, val in enumerate(row_values, start=1):
                c = ws_master.cell(row=current_row, column=c_idx, value=val)
                c.border = border_cell
                c.font = font_normal

                # 정렬
                if c_idx in [1, 2, 3, 5, 8, 9, 12, 13, 21]:
                    c.alignment = align_center
                elif c_idx in [4, 22]:
                    c.alignment = align_left
                else:
                    c.alignment = align_right

                # 숫자 서식
                if c_idx in [6, 7]:
                    c.number_format = "#,##0.00"
                elif c_idx in [10, 11, 14, 15, 16, 17, 18, 19]:
                    c.number_format = "#,##0"
                elif c_idx == 20:
                    c.number_format = "0.0%"

                # 안전도 색상
                if c_idx == 13:
                    if val == "안전":
                        c.fill = green_safe
                        c.font = font_green_b
                    else:
                        c.fill = yellow_accent
                        c.font = font_red_b

            existing_case_nos.add(case_no)
            current_row += 1
            added_count += 1
        except Exception as e:
            print(f"Error parsing {pdf_path}: {e}")

    # 열 너비 자동 조정 (마스터DB)
    col_widths = [6, 16, 14, 34, 8, 12, 10, 12, 12, 14, 14, 14, 12, 12, 14, 14, 14, 16, 15, 12, 12, 45]
    for idx, w in enumerate(col_widths, start=1):
        ws_master.column_dimensions[get_column_letter(idx)].width = w

    # ---------------------------------------------------------
    # 3. [상세_수지계산기] 템플릿 구성 (단일 만능 영수증)
    # ---------------------------------------------------------
    total_items = max(current_row - 1, 5)

    ws_calc.merge_cells("A1:G1")
    ws_calc["A1"] = "🧾 경매 매물 1:1 정밀 수지계산기 (만능 영수증 템플릿)"
    ws_calc["A1"].font = font_title
    ws_calc.row_dimensions[1].height = 30

    ws_calc.merge_cells("A2:G2")
    ws_calc["A2"] = "사건번호를 선택하시면 [매물_통합목록]에서 데이터를 자동 조회하여 3단계 입찰 시나리오를 계산합니다."
    ws_calc["A2"].font = font_subtitle
    ws_calc.row_dimensions[2].height = 18

    # 드롭다운 선택 영역
    ws_calc.merge_cells("A3:B3")
    ws_calc["A3"] = "▶ 분석 대상 사건번호 선택 :"
    ws_calc["A3"].font = font_sec_title
    ws_calc["A3"].fill = yellow_accent
    ws_calc["A3"].alignment = align_center
    ws_calc["A3"].border = border_header

    ws_calc.merge_cells("C3:D3")
    first_case_no = ws_master.cell(row=5, column=2).value or "사건번호없음"
    ws_calc["C3"] = first_case_no
    ws_calc["C3"].font = Font(name="맑은 고딕", size=12, bold=True, color="1B365D")
    ws_calc["C3"].alignment = align_center
    ws_calc["C3"].fill = yellow_accent
    ws_calc["C3"].border = border_header
    ws_calc.row_dimensions[3].height = 26

    # 데이터 유효성 검사 (드롭다운 목록 추가)
    dv = DataValidation(type="list", formula1=f"='매물_통합목록'!$B$5:$B${total_items}", allow_blank=True)
    ws_calc.add_data_validation(dv)
    dv.add("C3")

    # 매물 기본 정보 자동 조회 테이블 (VLOOKUP)
    calc_info_headers = [
        ("소재지 및 단지명", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,3,FALSE),\"-\")", "매각기일", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,8,FALSE),\"-\")"),
        ("전용면적 / 층수", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,5,FALSE)&\"㎡ / \"&VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,7,FALSE),\"-\")", "권리안전도", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,12,FALSE),\"-\")"),
        ("감정평가액 (A)", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,9,FALSE),0)", "최저매각가", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,10,FALSE),0)"),
        ("최근실거래가 (시세)", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,14,FALSE),0)", "예상매도가 (B)", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,15,FALSE),0)")
    ]

    for idx, (k1, v1, k2, v2) in enumerate(calc_info_headers, start=5):
        ws_calc.row_dimensions[idx].height = 20
        ws_calc.cell(row=idx, column=1, value=k1).font = font_bold
        ws_calc.cell(row=idx, column=1).fill = gray_light
        ws_calc.cell(row=idx, column=1).border = border_cell
        ws_calc.cell(row=idx, column=1).alignment = align_center

        ws_calc.merge_cells(f"B{idx}:D{idx}")
        c_v1 = ws_calc.cell(row=idx, column=2, value=v1)
        c_v1.font = font_normal
        c_v1.alignment = align_left if idx in [5, 6] else align_right
        for col in range(2, 5):
            ws_calc.cell(row=idx, column=col).border = border_cell
        if idx in [7, 8]:
            c_v1.number_format = "#,##0"

        ws_calc.cell(row=idx, column=5, value=k2).font = font_bold
        ws_calc.cell(row=idx, column=5).fill = gray_light
        ws_calc.cell(row=idx, column=5).border = border_cell
        ws_calc.cell(row=idx, column=5).alignment = align_center

        ws_calc.merge_cells(f"F{idx}:G{idx}")
        c_v2 = ws_calc.cell(row=idx, column=6, value=v2)
        c_v2.font = font_normal
        c_v2.alignment = align_center if idx in [5, 6] else align_right
        for col in range(6, 8):
            ws_calc.cell(row=idx, column=col).border = border_cell
        if idx in [7, 8]:
            c_v2.number_format = "#,##0"

    # 3단계 시나리오 시뮬레이터 테이블
    ws_calc.merge_cells("A10:G10")
    ws_calc["A10"] = "■ 3단계 입찰 시나리오 및 실질 취득원가 정밀 계산"
    ws_calc["A10"].font = font_sec_title
    ws_calc["A10"].fill = blue_light
    ws_calc.row_dimensions[10].height = 22

    sim_headers = ["비용 구분 및 수지 항목", "산출 기준", "시나리오1 (공격적/최저가)", "시나리오2 (적정 추천가)", "시나리오3 (보수적 안전가)", "비고 및 검증", "현장확인"]
    ws_calc.row_dimensions[11].height = 25
    for c_idx, h in enumerate(sim_headers, start=1):
        c = ws_calc.cell(row=11, column=c_idx, value=h)
        c.fill = navy_mid
        c.font = font_header_w
        c.alignment = align_center
        c.border = border_header

    calc_sim_rows = [
        # 행 12: 낙찰가
        ("낙찰(매수)예정가", "자동 / 입력", "=F7", "=IF(F7>=B7, F7*0.8, F7*1.15)", "=IF(F7>=B7, F7*0.7, F7*1.08)", "시나리오별 입찰가 시뮬레이션", "확정"),
        # 행 13: 인수보증금
        ("인수 권리 보증금", "VLOOKUP", "=IFERROR(VLOOKUP($C$3,'매물_통합목록'!$B$5:$V$100,13,FALSE),0)", "=C13", "=C13", "대항력 미배당 보증금", "확정"),
        # 행 14: 취득세 (1.1%)
        ("취득세 및 교육세", "=낙찰가 * 1.1%", "=C12*0.011", "=D12*0.011", "=E12*0.011", "주택가액 6억 이하 기준 (1.1%)", "수식"),
        # 행 15: 법무사 보수
        ("법무사 등기보수", "정액", 800000, 800000, 800000, "이전등기 수수료 및 채권할인 등", "확정"),
        # 행 16: 명도비 (이사비)
        ("명도비 (이사비/집행)", "협의 기준", 1500000, 1500000, 1500000, "점유자 퇴거 협의 이사비", "확정"),
        # 행 17: 수리 및 인테리어비
        ("수리 및 인테리어비", "현장 확인", 0, 0, 0, "자료 미기재: 현장확인 필요 (0원 처리)", "임장필요"),
        # 행 18: 체납 관리비
        ("체납 관리비", "관리실 확인", 0, 0, 0, "자료 미기재: 공용관리실 확인 필요 (0원 처리)", "임장필요"),
        # 행 19: 매도 중개수수료
        ("매도 중개수수료", "=예상매도가 * 0.4%", "=$F$8*0.004", "=$F$8*0.004", "=$F$8*0.004", "매도 시 공인중개사 법정수수료", "수식"),
        # 행 20: 실질 총 취득원가 합계
        ("■ 실질 총 취득원가", "=SUM(취득비용)", "=SUM(C12:C19)", "=SUM(D12:D19)", "=SUM(E12:E19)", "낙찰가+인수금+모든 취득부대비용 합계", "수식(SUM)"),
        # 행 21: 구분선
        ("-------------------", "-------------", "-------------", "-------------", "-------------", "--------------------------", "-------"),
        # 행 22: 예상 매도가
        ("예상 매도가격 (B)", "실거래가 기준", "=$F$8", "=$F$8", "=$F$8", "목표 매도 기준가", "확정"),
        # 행 23: 세전 예상 순수익
        ("■ 세전 예상 순수익", "=매도가 - 원가", "=C22-C20", "=D22-D20", "=E22-E20", "단기 매매 차익", "수식"),
        # 행 24: 수익률
        ("■ 총투자금 대비 수익률", "=순수익 / 총원가", "=C23/C20", "=D23/D20", "=E23/E20", "투자 원금 대비 세전 수익률", "수식")
    ]

    for idx, r_data in enumerate(calc_sim_rows, start=12):
        ws_calc.row_dimensions[idx].height = 20
        c1 = ws_calc.cell(row=idx, column=1, value=r_data[0])
        c2 = ws_calc.cell(row=idx, column=2, value=r_data[1])
        c3 = ws_calc.cell(row=idx, column=3, value=r_data[2])
        c4 = ws_calc.cell(row=idx, column=4, value=r_data[3])
        c5 = ws_calc.cell(row=idx, column=5, value=r_data[4])
        c6 = ws_calc.cell(row=idx, column=6, value=r_data[5])
        c7 = ws_calc.cell(row=idx, column=7, value=r_data[6])

        for c in [c1, c2, c3, c4, c5, c6, c7]:
            c.border = border_cell

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

        # 강조 서식
        if "■ 실질 총 취득원가" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = blue_light
                c.font = font_bold
                c.border = border_total
        elif "■ 세전 예상 순수익" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = yellow_accent
                c.font = font_red_b
        elif "■ 총투자금 대비 수익률" in str(r_data[0]):
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = green_safe
                c.font = font_green_b

        # 포맷
        if idx in [12, 13, 14, 15, 16, 17, 18, 19, 20, 22, 23]:
            c3.number_format = "#,##0"
            c4.number_format = "#,##0"
            c5.number_format = "#,##0"
        elif idx == 24:
            c3.number_format = "0.0%"
            c4.number_format = "0.0%"
            c5.number_format = "0.0%"

    calc_col_widths = [24, 20, 24, 24, 24, 32, 12]
    for idx, w in enumerate(calc_col_widths, start=1):
        ws_calc.column_dimensions[get_column_letter(idx)].width = w

    wb.save(MASTER_EXCEL_FILE)
    print(f"Master Excel DB updated successfully: {MASTER_EXCEL_FILE} (Newly added: {added_count} items)")
    return added_count


if __name__ == "__main__":
    create_or_update_master_db()
