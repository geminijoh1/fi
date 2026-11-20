# -*- coding: utf-8 -*-
"""
경매 및 일반 급매 통합 자동화 마스터 시스템 (v2.1 견고화 버전)
- 01_경매매물 : e옥션/법원 경매 PDF 스캔 -> [유형: 🏛️경매]
- 02_일반급매 : 급매물 브리핑 PDF 및 텍스트 파일(.txt) 스캔 -> [유형: 🏢일반급매]
- 결과물: '경매_급매_통합관리장.xlsx' (단일 통합 마스터 DB + 수식 자동 분기)
"""
import os
import sys
import re
import glob
import shutil
import warnings
import logging

# pdfminer 관련 경고 및 stderr 로그 숨김 처리 (콘솔을 깨끗하게 유지)
warnings.filterwarnings("ignore")
logging.getLogger("pdfminer").setLevel(logging.ERROR)

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

MASTER_EXCEL = "경매_급매_통합관리장.xlsx"
DIR_AUCTION = "01_경매매물"
DIR_EXPRESS = "02_일반급매"
DIR_AUCTION_DONE = os.path.join(DIR_AUCTION, "_완료")
DIR_EXPRESS_DONE = os.path.join(DIR_EXPRESS, "_완료")


def move_to_done(file_path, done_dir):
    """
    분석이 완료된 파일을 '_완료' 보관 폴더로 안전하게 이동
    - 이미 동일한 파일명이 보관 폴더에 존재할 경우 덮어쓰기(갱신) 처리
    - 에러 발생 시에도 전체 자동화 파이프라인이 멈추지 않도록 예외 처리
    """
    try:
        os.makedirs(done_dir, exist_ok=True)
        base_name = os.path.basename(file_path)
        dest_path = os.path.join(done_dir, base_name)
        if os.path.exists(dest_path):
            os.remove(dest_path)
        shutil.move(file_path, dest_path)
        return True
    except Exception as e:
        print(f"    [보관 이동 실패] {os.path.basename(file_path)}: {e}")
        return False

# 팔레트 및 서식
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


def parse_auction_pdf(pdf_path):
    """경매 PDF 텍스트 파싱"""
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += "\n" + t

    case_match = re.search(r"(\d{4}\s*타경\s*\d+)", full_text)
    case_no = case_match.group(1).replace(" ", "") if case_match else "사건번호확인"

    court_match = re.search(r"(\d{4}\s*타경\s*\d+)\s+([^\n\r]+?계)", full_text)
    court = court_match.group(2).strip() if court_match else "법원확인"

    date_match = re.search(r"매각기일은\s*(\d{4}-\d{2}-\d{2})", full_text)
    if not date_match:
        date_match = re.search(r"매각기일\s*[:\s]*(\d{4}[-.]\d{2}[-.]\d{2})", full_text)
    sale_date = date_match.group(1).replace(".", "-") if date_match else "확인필요"

    loc_match = re.search(r"소재지\s*([^\n\r]+?)(?:지도보기|\n|$)", full_text)
    location = loc_match.group(1).strip() if loc_match else "소재지확인"

    area_match = re.search(r"건물면적\s*([\d.]+)\s*㎡\s*\(\s*([\d.]+)\s*평\)", full_text)
    area_m2 = float(area_match.group(1)) if area_match else 0.0
    area_py = float(area_match.group(2)) if area_match else 0.0

    app_match = re.search(r"감정가\s*([\d,]+)\s*원", full_text)
    appraisal = int(app_match.group(1).replace(",", "")) if app_match else 0

    min_match = re.search(r"최저가\s*\([^\)]*\)\s*([\d,]+)\s*원", full_text)
    if not min_match:
        min_match = re.search(r"최저가\s*([\d,]+)\s*원", full_text)
    min_bid = int(min_match.group(1).replace(",", "")) if min_match else appraisal

    ratio = round((min_bid / appraisal) * 100) if appraisal > 0 else 100
    status_str = f"{ratio}% (신건)" if ratio >= 100 else f"{ratio}% ({1 if ratio==70 else 2}회유찰)"

    floor_match = re.search(r"(\d+)층건\s*내\s*제?(\d+)층", full_text)
    if floor_match:
        floor_str = f"{floor_match.group(2)} / {floor_match.group(1)}층"
    else:
        f_simple = re.search(r"(\d+)층", location)
        floor_str = f"{f_simple.group(1)}층" if f_simple else "확인필요"

    # 권리분석
    safe_status = "안전"
    risk_text = "말소기준권리 이후 소멸, 인수권리 없음"

    if "외국인" in full_text:
        safe_status = "주의"
        risk_text = "외국인 임차인 등재(대항력 소멸 추정), 명도 난이도 주의"
    elif "조사된 임차내역이 없습니다" in full_text:
        safe_status = "안전"
        risk_text = "조사된 임차인 없음 (소유자 점유 추정, 이사비 150만원 반영)"

    # 실거래가
    trade_match = re.search(r"최근실거래가\s*:\s*([\d,]+)\s*원", full_text)
    if trade_match:
        real_trade = int(trade_match.group(1).replace(",", ""))
    else:
        recent_trades = re.findall(r"매\s*([\d,]+)", full_text)
        real_trade = int(recent_trades[0].replace(",", "")) * 10000 if recent_trades else appraisal

    expected_sale = int(real_trade * 0.98)

    # 추천 입찰가
    if ratio >= 100 and real_trade < appraisal:
        rec_bid = int(appraisal * 0.65)
        decision = "⚠️ 1차 입찰금지 (시세 역마진, 유찰 대기)"
        stage = "유찰대기"
    else:
        rec_bid = int(min_bid * 1.15) if min_bid * 1.15 < expected_sale * 0.85 else min_bid
        decision = "★ 1순위 추천 (안전마진 확보)"
        stage = "신규검토"

    return {
        "prop_type": "🏛️경매",
        "case_no": case_no,
        "source": court,
        "location": location,
        "area_m2": area_m2,
        "area_py": area_py,
        "floor": floor_str,
        "target_date": sale_date,
        "base_price": appraisal,
        "min_or_nego": min_bid,
        "status_str": status_str,
        "safe_status": safe_status,
        "deposit_or_evict": 1500000,
        "real_trade": real_trade,
        "expected_sale": expected_sale,
        "rec_acquire_price": rec_bid,
        "stage": stage,
        "memo": decision + " | " + risk_text
    }


def parse_express_file(file_path):
    """일반 급매 텍스트(.txt) 또는 PDF 파싱"""
    if file_path.endswith(".txt"):
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
    else:
        content = ""
        with pdfplumber.open(file_path) as pdf:
            for p in pdf.pages:
                content += "\n" + (p.extract_text() or "")

    def get_val(key, default=""):
        m = re.search(rf"[■\-\*]?\s*{key}\s*[:=]\s*([^\n\r]+)", content)
        return m.group(1).strip() if m else default

    fname = os.path.basename(file_path).replace(".txt", "").replace(".pdf", "")
    case_no = f"급매-{fname[:12]}"
    prop_name = get_val("매물명", fname)
    source = get_val("중개업소", "현지 공인중개사")
    location = get_val("소재지", prop_name)

    area_m2, area_py = 84.99, 25.71
    area_str = get_val("전용면적")
    m_m2 = re.search(r"([\d.]+)\s*㎡", area_str)
    if m_m2: area_m2 = float(m_m2.group(1))
    m_py = re.search(r"([\d.]+)\s*평", area_str)
    if m_py: area_py = float(m_py.group(1))

    floor = get_val("층수", "로얄층")
    target_date = get_val("목표기일", "협의가능")

    def parse_money(key, default_val):
        s = get_val(key)
        digits = re.sub(r"[^\d]", "", s)
        return int(digits) if digits else default_val

    base_price = parse_money("최초호가", 350000000)
    nego_target = parse_money("네고목표가", int(base_price * 0.95))
    real_trade = parse_money("최근실거래가", int(base_price * 1.15))
    expected_sale = parse_money("예상매도가", int(real_trade * 0.98))
    safe_status = get_val("권리상태", "안전")
    special_memo = get_val("특이사항", "정상 매매계약")

    return {
        "prop_type": "🏢일반급매",
        "case_no": case_no,
        "source": source,
        "location": f"{location} ({prop_name})",
        "area_m2": area_m2,
        "area_py": area_py,
        "floor": floor,
        "target_date": target_date,
        "base_price": base_price,
        "min_or_nego": nego_target,
        "status_str": "급매네고중",
        "safe_status": safe_status,
        "deposit_or_evict": 0,
        "real_trade": real_trade,
        "expected_sale": expected_sale,
        "rec_acquire_price": nego_target,
        "stage": "현장임장",
        "memo": f"소장님 브리핑 급매 | {special_memo}"
    }


def build_integrated_master():
    """경매와 급매를 모두 통합한 마스터 엑셀 생성/갱신"""
    print("\n=======================================================================")
    print("  [경매 & 일반급매 자동화] 신규 매물 스캔 및 분석을 시작합니다.")
    print("=======================================================================")

    is_new = not os.path.exists(MASTER_EXCEL)
    if is_new:
        wb = openpyxl.Workbook()
        ws_master = wb.active
        ws_master.title = "통합_매물목록"
        ws_calc = wb.create_sheet("상세_수지계산기")
    else:
        try:
            wb = openpyxl.load_workbook(MASTER_EXCEL)
            ws_master = wb["통합_매물목록"]
            ws_calc = wb["상세_수지계산기"]
        except PermissionError:
            print(f"\n[오류] '{MASTER_EXCEL}' 파일이 현재 열려있습니다!")
            print(">> 엑셀 파일을 닫으신 후 다시 실행해주세요.\n")
            return -1

    ws_master.views.sheetView[0].showGridLines = True
    ws_calc.views.sheetView[0].showGridLines = True

    # 1. 통합 헤더 구성
    headers = [
        "순번", "취득유형", "관리/사건번호", "법원 / 중개업소", "소재지 및 단지명",
        "전용면적(㎡)", "전용(평)", "층수", "목표기일\n(매각/잔금)",
        "감정가/호가", "최저가/네고목표", "진행상태", "권리안전도", "명도비/인수금",
        "최근실거래가", "예상매도가", "추천취득가",
        "총 취득원가\n(수식자동분기)", "세전 예상순익\n(수식)", "수익률\n(수식)",
        "진행단계", "투자판단 및 핵심메모"
    ]

    ws_master.merge_cells("A1:V1")
    ws_master["A1"] = "🏢 경매 & 일반 급매 통합 관리 마스터 DB (All-in-One)"
    ws_master["A1"].font = font_title
    ws_master.row_dimensions[1].height = 30

    ws_master.merge_cells("A2:V2")
    ws_master["A2"] = "01_경매매물(PDF)과 02_일반급매(PDF/TXT)가 통합 집계됩니다. 취득유형에 따라 중개료·명도비 수식이 자동 분기됩니다."
    ws_master["A2"].font = font_subtitle
    ws_master.row_dimensions[2].height = 18

    ws_master.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, start=1):
        cell = ws_master.cell(row=4, column=col_idx, value=h)
        cell.fill = navy_dark
        cell.font = font_header_w
        cell.alignment = align_center
        cell.border = border_header

    # 기존 등록 식별자 파악
    existing_keys = set()
    for r in range(5, ws_master.max_row + 1):
        val = ws_master.cell(row=r, column=3).value
        if val:
            existing_keys.add(str(val).strip())

    current_row = 5
    while ws_master.cell(row=current_row, column=3).value:
        current_row += 1

    # 2. 경매 매물 스캔 (01_경매매물) - '_완료' 하위 폴더를 제외하고 신규 PDF 파일만 수집
    os.makedirs(DIR_AUCTION, exist_ok=True)
    os.makedirs(DIR_AUCTION_DONE, exist_ok=True)
    auction_files = [
        os.path.join(DIR_AUCTION, f) for f in os.listdir(DIR_AUCTION)
        if os.path.isfile(os.path.join(DIR_AUCTION, f)) and f.lower().endswith(".pdf")
    ]
    print(f"\n[1/2] '01_경매매물' 폴더 스캔: 신규 대상 {len(auction_files)}개 PDF 발견")
    added_count = 0

    for f in auction_files:
        try:
            fname = os.path.basename(f)
            # 1단계 고속화: 파일명에서 사건번호 사전 추출 -> 이미 등록된 건은 무거운 PDF 열기 없이 즉시 보관함 이동
            fname_case_match = re.search(r"(\d{4}\s*타경\s*\d+)", fname)
            if fname_case_match:
                quick_case = fname_case_match.group(1).replace(" ", "")
                if quick_case in existing_keys:
                    print(f"  - 이미 등록됨 (사전 필터링 건너뜀 -> '_완료' 이동): {quick_case}")
                    move_to_done(f, DIR_AUCTION_DONE)
                    continue

            # 2단계: 신규 파일만 정밀 PDF 파싱 실행
            d = parse_auction_pdf(f)
            if d["case_no"] in existing_keys:
                print(f"  - 이미 등록됨 (파싱 확인 후 '_완료' 이동): {d['case_no']}")
                move_to_done(f, DIR_AUCTION_DONE)
                continue

            formula_cost = f'=IF(B{current_row}="🏛️경매", Q{current_row}+N{current_row}+(Q{current_row}*0.011)+800000+(P{current_row}*0.004), Q{current_row}+N{current_row}+(Q{current_row}*0.011)+800000+(Q{current_row}*0.004)+(P{current_row}*0.004))'
            formula_profit = f"=P{current_row}-R{current_row}"
            formula_roi = f"=S{current_row}/R{current_row}"

            seq = current_row - 4
            row_v = [
                seq, d["prop_type"], d["case_no"], d["source"], d["location"],
                d["area_m2"], d["area_py"], d["floor"], d["target_date"],
                d["base_price"], d["min_or_nego"], d["status_str"],
                d["safe_status"], d["deposit_or_evict"], d["real_trade"],
                d["expected_sale"], d["rec_acquire_price"],
                formula_cost, formula_profit, formula_roi,
                d["stage"], d["memo"]
            ]

            ws_master.row_dimensions[current_row].height = 25
            for c_idx, val in enumerate(row_v, start=1):
                c = ws_master.cell(row=current_row, column=c_idx, value=val)
                c.border = border_cell
                c.font = font_normal
                if c_idx in [1, 2, 3, 4, 8, 9, 12, 13, 21]:
                    c.alignment = align_center
                elif c_idx in [5, 22]:
                    c.alignment = align_left
                else:
                    c.alignment = align_right

                if c_idx in [6, 7]: c.number_format = "#,##0.00"
                elif c_idx in [10, 11, 14, 15, 16, 17, 18, 19]: c.number_format = "#,##0"
                elif c_idx == 20: c.number_format = "0.0%"

                if c_idx == 13:
                    c.fill = green_safe if val == "안전" else yellow_accent
                    c.font = font_green_b if val == "안전" else font_red_b
                elif c_idx == 2:
                    c.fill = blue_light if "경매" in val else yellow_accent

            existing_keys.add(d["case_no"])
            print(f"  + [신규 추가] {d['case_no']} | {d['location'][:25]}")
            current_row += 1
            added_count += 1

            # 등록 완료 후 완료 보관 폴더로 이동하여 다음 실행 시 불필요한 스캔 방지
            move_to_done(f, DIR_AUCTION_DONE)
        except Exception as e:
            print(f"  [경고] 파일 파싱 실패: {os.path.basename(f)} ({e})")

    # 3. 일반 급매물 스캔 (02_일반급매) - '_완료' 하위 폴더를 제외하고 신규 파일만 수집
    os.makedirs(DIR_EXPRESS, exist_ok=True)
    os.makedirs(DIR_EXPRESS_DONE, exist_ok=True)
    express_files = [
        os.path.join(DIR_EXPRESS, f) for f in os.listdir(DIR_EXPRESS)
        if os.path.isfile(os.path.join(DIR_EXPRESS, f)) and (f.lower().endswith(".txt") or f.lower().endswith(".pdf"))
    ]
    print(f"\n[2/2] '02_일반급매' 폴더 스캔: 신규 대상 {len(express_files)}개 파일 발견")

    for f in express_files:
        try:
            fname = os.path.basename(f).replace(".txt", "").replace(".pdf", "")
            quick_case = f"급매-{fname[:12]}"
            # 사전 필터링
            if quick_case in existing_keys:
                print(f"  - 이미 등록됨 (사전 필터링 건너뜀 -> '_완료' 이동): {quick_case}")
                move_to_done(f, DIR_EXPRESS_DONE)
                continue

            d = parse_express_file(f)
            if d["case_no"] in existing_keys:
                print(f"  - 이미 등록됨 (파싱 확인 후 '_완료' 이동): {d['case_no']}")
                move_to_done(f, DIR_EXPRESS_DONE)
                continue

            formula_cost = f'=IF(B{current_row}="🏛️경매", Q{current_row}+N{current_row}+(Q{current_row}*0.011)+800000+(P{current_row}*0.004), Q{current_row}+N{current_row}+(Q{current_row}*0.011)+800000+(Q{current_row}*0.004)+(P{current_row}*0.004))'
            formula_profit = f"=P{current_row}-R{current_row}"
            formula_roi = f"=S{current_row}/R{current_row}"

            seq = current_row - 4
            row_v = [
                seq, d["prop_type"], d["case_no"], d["source"], d["location"],
                d["area_m2"], d["area_py"], d["floor"], d["target_date"],
                d["base_price"], d["min_or_nego"], d["status_str"],
                d["safe_status"], d["deposit_or_evict"], d["real_trade"],
                d["expected_sale"], d["rec_acquire_price"],
                formula_cost, formula_profit, formula_roi,
                d["stage"], d["memo"]
            ]

            ws_master.row_dimensions[current_row].height = 25
            for c_idx, val in enumerate(row_v, start=1):
                c = ws_master.cell(row=current_row, column=c_idx, value=val)
                c.border = border_cell
                c.font = font_normal
                if c_idx in [1, 2, 3, 4, 8, 9, 12, 13, 21]:
                    c.alignment = align_center
                elif c_idx in [5, 22]:
                    c.alignment = align_left
                else:
                    c.alignment = align_right

                if c_idx in [6, 7]: c.number_format = "#,##0.00"
                elif c_idx in [10, 11, 14, 15, 16, 17, 18, 19]: c.number_format = "#,##0"
                elif c_idx == 20: c.number_format = "0.0%"

                if c_idx == 13:
                    c.fill = green_safe if val == "안전" else yellow_accent
                    c.font = font_green_b if val == "안전" else font_red_b
                elif c_idx == 2:
                    c.fill = yellow_accent

            existing_keys.add(d["case_no"])
            print(f"  + [신규 추가] {d['case_no']} | {d['location'][:25]}")
            current_row += 1
            added_count += 1

            # 등록 완료 후 완료 보관 폴더로 이동하여 다음 실행 시 불필요한 스캔 방지
            move_to_done(f, DIR_EXPRESS_DONE)
        except Exception as e:
            print(f"  [경고] 급매 파일 파싱 실패: {os.path.basename(f)} ({e})")

    # 마스터 열 너비
    col_widths = [6, 12, 16, 18, 36, 12, 10, 12, 12, 14, 14, 14, 12, 12, 14, 14, 14, 17, 15, 12, 12, 45]
    for idx, w in enumerate(col_widths, start=1):
        ws_master.column_dimensions[get_column_letter(idx)].width = w

    # 4. [상세_수지계산기] 템플릿 구성
    total_items = max(current_row - 1, 5)

    ws_calc.merge_cells("A1:G1")
    ws_calc["A1"] = "🧾 경매 & 일반 급매 1:1 정밀 수지계산기"
    ws_calc["A1"].font = font_title
    ws_calc.row_dimensions[1].height = 30

    ws_calc.merge_cells("A2:G2")
    ws_calc["A2"] = "관리/사건번호를 선택하시면 해당 물건의 취득유형(경매/급매)에 맞춰 부대비용 및 입찰/매수가 시나리오가 자동 계산됩니다."
    ws_calc["A2"].font = font_subtitle
    ws_calc.row_dimensions[2].height = 18

    # 선택 영역
    ws_calc.merge_cells("A3:B3")
    ws_calc["A3"] = "▶ 관리/사건번호 선택 :"
    ws_calc["A3"].font = font_sec_title
    ws_calc["A3"].fill = yellow_accent
    ws_calc["A3"].alignment = align_center
    ws_calc["A3"].border = border_header

    ws_calc.merge_cells("C3:D3")
    first_key = ws_master.cell(row=5, column=3).value or "사건번호없음"
    ws_calc["C3"] = first_key
    ws_calc["C3"].font = Font(name="맑은 고딕", size=12, bold=True, color="1B365D")
    ws_calc["C3"].alignment = align_center
    ws_calc["C3"].fill = yellow_accent
    ws_calc["C3"].border = border_header
    ws_calc.row_dimensions[3].height = 26

    dv = DataValidation(type="list", formula1=f"='통합_매물목록'!$C$5:$C${total_items}", allow_blank=True)
    ws_calc.add_data_validation(dv)
    dv.add("C3")

    # 기본정보 VLOOKUP
    calc_info = [
        ("취득유형", "=IFERROR(VLOOKUP($C$3,'통합_매물목록'!$C$5:$V$100,2,FALSE),\"-\")", "목표기일", "=IFERROR(VLOOKUP($C$3,'통합_매물목록'!$C$5:$V$100,9,FALSE),\"-\")"),
        ("소재지 및 단지명", "=IFERROR(VLOOKUP($C$3,'통합_매물목록'!$C$5:$V$100,5,FALSE),\"-\")", "권리안전도", "=IFERROR(VLOOKUP($C$3,'통합_매물목록'!$C$5:$V$100,13,FALSE),\"-\")"),
        ("감정가 / 최초호가", "=IFERROR(VLOOKUP($C$3,'통합_매물목록'!$C$5:$V$100,10,FALSE),0)", "최저가 / 네고목표", "=IFERROR(VLOOKUP($C$3,'통합_매물목록'!$C$5:$V$100,11,FALSE),0)"),
        ("최근실거래가 (시세)", "=IFERROR(VLOOKUP($C$3,'통합_매물목록'!$C$5:$V$100,15,FALSE),0)", "예상매도가", "=IFERROR(VLOOKUP($C$3,'통합_매물목록'!$C$5:$V$100,16,FALSE),0)")
    ]

    for idx, (k1, v1, k2, v2) in enumerate(calc_info, start=5):
        ws_calc.row_dimensions[idx].height = 20
        ws_calc.cell(row=idx, column=1, value=k1).font = font_bold
        ws_calc.cell(row=idx, column=1).fill = gray_light
        ws_calc.cell(row=idx, column=1).border = border_cell
        ws_calc.cell(row=idx, column=1).alignment = align_center

        ws_calc.merge_cells(f"B{idx}:D{idx}")
        c_v1 = ws_calc.cell(row=idx, column=2, value=v1)
        c_v1.font = font_normal
        c_v1.alignment = align_left if idx in [5, 6] else align_right
        for col in range(2, 5): ws_calc.cell(row=idx, column=col).border = border_cell
        if idx in [7, 8]: c_v1.number_format = "#,##0"

        ws_calc.cell(row=idx, column=5, value=k2).font = font_bold
        ws_calc.cell(row=idx, column=5).fill = gray_light
        ws_calc.cell(row=idx, column=5).border = border_cell
        ws_calc.cell(row=idx, column=5).alignment = align_center

        ws_calc.merge_cells(f"F{idx}:G{idx}")
        c_v2 = ws_calc.cell(row=idx, column=6, value=v2)
        c_v2.font = font_normal
        c_v2.alignment = align_center if idx in [5, 6] else align_right
        for col in range(6, 8): ws_calc.cell(row=idx, column=col).border = border_cell
        if idx in [7, 8]: c_v2.number_format = "#,##0"

    # 수지 시뮬레이션
    ws_calc.merge_cells("A10:G10")
    ws_calc["A10"] = "■ 3단계 취득가(입찰/네고) 시나리오 및 취득원가 정밀 계산"
    ws_calc["A10"].font = font_sec_title
    ws_calc["A10"].fill = blue_light
    ws_calc.row_dimensions[10].height = 22

    sim_headers = ["비용 구분 및 수지 항목", "산출 기준", "시나리오1 (공격적/최저가)", "시나리오2 (적정 추천가)", "시나리오3 (보수적 안전가)", "비고 및 검증", "데이터출처"]
    ws_calc.row_dimensions[11].height = 25
    for c_idx, h in enumerate(sim_headers, start=1):
        c = ws_calc.cell(row=11, column=c_idx, value=h)
        c.fill = navy_mid
        c.font = font_header_w
        c.alignment = align_center
        c.border = border_header

    calc_sim_rows = [
        ("취득(입찰/매수)예정가", "자동 / 입력", "=F7", "=IF(F7>=B7, F7*0.8, F7*1.15)", "=IF(F7>=B7, F7*0.7, F7*1.08)", "시나리오별 취득가", "확정"),
        ("명도비 / 인수보증금", "=IF(B5=\"🏛️경매\", 1500000, 0)", "=IF(B5=\"🏛️경매\", 1500000, 0)", "=C13", "=C13", "경매는 명도비 150만, 급매는 0원", "수식"),
        ("취득세 및 교육세", "=취득가 * 1.1%", "=C12*0.011", "=D12*0.011", "=E12*0.011", "주택 취득세율 (1.1%)", "수식"),
        ("법무사 등기보수", "정액", 800000, 800000, 800000, "소유권이전 등기 대행료", "확정"),
        ("수리 및 인테리어비", "현장 확인", 0, 0, 0, "자료 미기재: 현장확인 필요 (0원 처리)", "임장필요"),
        ("체납 관리비", "관리실 확인", 0, 0, 0, "자료 미기재: 공용관리실 확인 필요 (0원 처리)", "임장필요"),
        ("중개수수료 합계", "유형별 자동계산", "=IF(B5=\"🏛️경매\", $F$8*0.004, (C12*0.004)+($F$8*0.004))", "=IF(B5=\"🏛️경매\", $F$8*0.004, (D12*0.004)+($F$8*0.004))", "=IF(B5=\"🏛️경매\", $F$8*0.004, (E12*0.004)+($F$8*0.004))", "경매 1회 / 급매 매수+매도 2회 반영", "수식"),
        ("■ 실질 총 취득원가", "=SUM(취득비용)", "=SUM(C12:C18)", "=SUM(D12:D18)", "=SUM(E12:E18)", "모든 취득부대비용 합계", "수식(SUM)"),
        ("-------------------", "-------------", "-------------", "-------------", "-------------", "--------------------------", "-------"),
        ("예상 매도가격 (B)", "실거래가 기준", "=$F$8", "=$F$8", "=$F$8", "목표 매도 기준가", "확정"),
        ("■ 세전 예상 순수익", "=매도가 - 원가", "=C21-C19", "=D21-D19", "=E21-E19", "단기 매매 차익", "수식"),
        ("■ 총투자금 대비 수익률", "=순수익 / 총원가", "=C22/C19", "=D22/D19", "=E22/E19", "원금 대비 세전 수익률", "수식")
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

        if idx in [12, 13, 14, 15, 16, 17, 18, 19, 21, 22]:
            c3.number_format = "#,##0"
            c4.number_format = "#,##0"
            c5.number_format = "#,##0"
        elif idx == 23:
            c3.number_format = "0.0%"
            c4.number_format = "0.0%"
            c5.number_format = "0.0%"

    calc_col_widths = [24, 20, 24, 24, 24, 32, 12]
    for idx, w in enumerate(calc_col_widths, start=1):
        ws_calc.column_dimensions[get_column_letter(idx)].width = w

    try:
        wb.save(MASTER_EXCEL)
        print("\n=======================================================================")
        print(f"  [성공] 마스터 엑셀 갱신 완료: '{MASTER_EXCEL}'")
        print(f"  - 신규 등록된 매물: {added_count}건")
        print(f"  - 완료된 파일은 각 폴더 내 '_완료' 보관함으로 자동 격리되었습니다.")
        print(f"  - 다음 실행 시에는 '_완료'를 건너뛰고 신규 파일만 초고속 스캔합니다.")
        print("=======================================================================\n")
        return added_count
    except PermissionError:
        print(f"\n[오류] 저장 실패: '{MASTER_EXCEL}' 파일이 열려있습니다.")
        print(">> 엑셀 파일을 닫으신 후 다시 실행해주세요.\n")
        return -1


if __name__ == "__main__":
    build_integrated_master()
