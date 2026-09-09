# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ARIAL = "Arial"
BLUE = Font(name=ARIAL, size=10, color="0000FF")           # 입력
BLACK = Font(name=ARIAL, size=10)                           # 수식
HDR = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
TITLE = Font(name=ARIAL, size=13, bold=True)
NOTE = Font(name=ARIAL, size=9, italic=True, color="595959")
BOLD = Font(name=ARIAL, size=10, bold=True)
HDRFILL = PatternFill("solid", fgColor="404040")
YELLOW = PatternFill("solid", fgColor="FFFF00")
GREYFILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

WON = '#,##0;(#,##0);-'
PCT1 = '0.0%'
PCT2 = '0.00%'
MULT = '0.00"배"'

wb = Workbook()

# ============================================================
# 시트 4: 기준표  (먼저 만들어 다른 시트가 참조)
# ============================================================
ref = wb.active
ref.title = "기준표"
ref["A1"] = "기준표 — 채널별 수수료율 및 게이트 임계값"; ref["A1"].font = TITLE
ref["A2"] = "노란 셀은 정책·목표가 바뀌면 수정하세요. 다른 모든 시트가 이 값을 참조합니다."; ref["A2"].font = NOTE

ref["A4"] = "채널별 판매수수료율"; ref["A4"].font = BOLD
fees = [
    ("네이버 스마트스토어 (일반유입)", 0.06633, "판매 3.003% + 주문관리 3.63%. 2025-06 개편 기준"),
    ("네이버 (마케팅링크 유입)",        0.04631, "판매 1.001% + 주문관리 3.63%"),
    ("네이버 (스타트 제로수수료 초기)", 0.0,     "주문관리 12개월 / 매출연동 6개월 면제 [면제항목 매핑 미검증]"),
    ("토스쇼핑 (일반)",                 0.08,    "저수수료 정책 8% 수준"),
    ("토스쇼핑 (광고 경유)",            0.0,     "광고 클릭 후 7일 내 구매는 판매수수료 0%"),
    ("쿠팡 (생활용품·가전디지털)",      0.078,   "카테고리별 4~10.9% 중"),
    ("쿠팡 (주방·가구·반려·스포츠)",    0.108,   ""),
    ("11번가 (상한)",                   0.13,    "7~13% + 서버이용료 월 77,000원(전월 500만↑)"),
    ("카카오 톡스토어",                 0.0,     "판매수수료 무료. 단 노출 채널별 추가 수수료 가능 [미검증]"),
    ("오늘의집",                        None,    "카테고리별. 파트너센터 로그인 후 직접 입력 [미확인]"),
    ("올웨이즈",                        0.015,   "+ 월 서버이용료 49,000원(월매출 100만↑)"),
    ("알리 K-venue",                    0.08,    "신규 입점 90일 면제"),
]
ref["A5"] = "채널"; ref["B5"] = "수수료율"; ref["C5"] = "비고"
for c in "ABC":
    ref[f"{c}5"].font = HDR; ref[f"{c}5"].fill = HDRFILL; ref[f"{c}5"].border = BOX
r = 6
for name, rate, memo in fees:
    ref[f"A{r}"] = name; ref[f"A{r}"].font = BLACK; ref[f"A{r}"].border = BOX
    ref[f"B{r}"] = rate; ref[f"B{r}"].font = BLUE; ref[f"B{r}"].number_format = PCT2
    ref[f"B{r}"].fill = YELLOW; ref[f"B{r}"].border = BOX
    ref[f"C{r}"] = memo; ref[f"C{r}"].font = NOTE; ref[f"C{r}"].border = BOX
    r += 1

ref["A20"] = "게이트 임계값"; ref["A20"].font = BOLD
ref["A21"] = "항목"; ref["B21"] = "값"; ref["C21"] = "근거"
for c in "ABC":
    ref[f"{c}21"].font = HDR; ref[f"{c}21"].fill = HDRFILL; ref[f"{c}21"].border = BOX
gates = [
    ("게이트1 최소 기여이익", 18000, WON, "목표(건당 순익 1만원) 역산. 일반 기준은 8,000원"),
    ("목표 건당 순이익",      10000, WON, "사용자 설정 목표"),
    ("기준 CPC",                250,  WON, "쿠팡 통상 200~300원대. 토스 최소입찰 200원"),
    ("최소 월간 검색량",      15000, '#,##0', "월 100건 목표 ÷ 전환율 3% ÷ CTR 20% [CTR 미검증]"),
    ("리뷰 14일 최소 증가",      20, '#,##0', "죽은 시장 배제선 [미검증 휴리스틱]"),
    ("1위 집중도 상한",         0.70, PCT1, "초과 시 승자독식 [미검증 휴리스틱]"),
    ("최저가 갭 상한",          0.15, PCT1, "초과 시 최저가 경쟁 진행 중 [미검증 휴리스틱]"),
    ("반품률 경고선",           0.15, PCT1, "20% 초과 시 손절"),
    ("주간 취소율 상한",        0.02, PCT1, "초과 시 해당 도매사 노출 중단"),
]
r = 22
for name, val, fmt, memo in gates:
    ref[f"A{r}"] = name; ref[f"A{r}"].font = BLACK; ref[f"A{r}"].border = BOX
    ref[f"B{r}"] = val; ref[f"B{r}"].font = BLUE; ref[f"B{r}"].number_format = fmt
    ref[f"B{r}"].fill = YELLOW; ref[f"B{r}"].border = BOX
    ref[f"C{r}"] = memo; ref[f"C{r}"].font = NOTE; ref[f"C{r}"].border = BOX
    r += 1

ref["A32"] = "출처: 본 저장소 docs/consignment-platform-strategy.md (2026-09 기준). [미검증] 표기 항목은 1차 출처 확인 실패."
ref["A32"].font = NOTE
for col, w in zip("ABC", [34, 14, 66]):
    ref.column_dimensions[col].width = w

# 참조 주소 상수
F_NAVER   = "기준표!$B$6"
F_TOSSAD  = "기준표!$B$10"
F_TOSSGEN = "기준표!$B$9"
F_CPGEN   = "기준표!$B$11"
G_CONTRIB = "기준표!$B$22"
G_TARGET  = "기준표!$B$23"
G_CPC     = "기준표!$B$24"
G_SEARCH  = "기준표!$B$25"
G_RVGROW  = "기준표!$B$26"
G_CONC    = "기준표!$B$27"

# ============================================================
# 시트 1: 소싱계산기
# ============================================================
ws = wb.create_sheet("소싱계산기", 0)
ws["A1"] = "소싱 게이트 계산기"; ws["A1"].font = TITLE
ws["A2"] = "파란 글씨 셀(C·D·E열)만 입력하세요. 나머지는 자동 계산됩니다."; ws["A2"].font = NOTE
ws["A3"] = "게이트1 판정은 네이버(주력 채널) 기여이익 기준입니다. 다른 채널 기여이익은 참고용으로 나란히 표시됩니다."; ws["A3"].font = NOTE
ws["A4"] = "7행은 입력 형식을 보여주는 예시입니다. 실제 사용 시 덮어쓰세요."; ws["A4"].font = NOTE

cols = [
    ("A", "상품명", 22), ("B", "키워드", 16), ("C", "판매가", 11), ("D", "공급가", 11),
    ("E", "배송비", 10), ("F", "원가율", 10), ("G", "네이버\n수수료", 11),
    ("H", "네이버\n기여이익", 12), ("I", "토스\n(광고경유)", 12), ("J", "토스\n(일반)", 11),
    ("K", "쿠팡\n(생활용품)", 12), ("L", "게이트1\n판정", 11), ("M", "허용 CPA", 11),
    ("N", "필요전환율\n@CPC250", 13),
]
for c, h, w in cols:
    cell = ws[f"{c}6"]; cell.value = h; cell.font = HDR; cell.fill = HDRFILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BOX
    ws.column_dimensions[c].width = w
ws.row_dimensions[6].height = 30

FIRST, LAST = 7, 46
for r in range(FIRST, LAST + 1):
    for c in "AB":
        ws[f"{c}{r}"].font = BLUE; ws[f"{c}{r}"].border = BOX
    for c in "CDE":
        cell = ws[f"{c}{r}"]; cell.font = BLUE; cell.number_format = WON; cell.border = BOX
    ws[f"F{r}"] = f'=IF(OR(C{r}="",D{r}=""),"",D{r}/C{r})'
    ws[f"F{r}"].number_format = PCT1
    ws[f"G{r}"] = f'=IF(C{r}="","",C{r}*{F_NAVER})'
    ws[f"H{r}"] = f'=IF(OR(C{r}="",D{r}="",E{r}=""),"",C{r}-D{r}-E{r}-G{r})'
    ws[f"I{r}"] = f'=IF(OR(C{r}="",D{r}="",E{r}=""),"",C{r}-D{r}-E{r}-C{r}*{F_TOSSAD})'
    ws[f"J{r}"] = f'=IF(OR(C{r}="",D{r}="",E{r}=""),"",C{r}-D{r}-E{r}-C{r}*{F_TOSSGEN})'
    ws[f"K{r}"] = f'=IF(OR(C{r}="",D{r}="",E{r}=""),"",C{r}-D{r}-E{r}-C{r}*{F_CPGEN})'
    ws[f"L{r}"] = f'=IF(H{r}="","",IF(H{r}>={G_CONTRIB},"PASS","FAIL"))'
    ws[f"M{r}"] = f'=IF(H{r}="","",H{r}-{G_TARGET})'
    ws[f"N{r}"] = f'=IF(OR(M{r}="",M{r}<=0),"",{G_CPC}/M{r})'
    for c in "GHIJKM":
        ws[f"{c}{r}"].font = BLACK; ws[f"{c}{r}"].number_format = WON; ws[f"{c}{r}"].border = BOX
    ws[f"F{r}"].font = BLACK; ws[f"F{r}"].border = BOX
    ws[f"L{r}"].font = BOLD; ws[f"L{r}"].border = BOX
    ws[f"L{r}"].alignment = Alignment(horizontal="center")
    ws[f"N{r}"].font = BLACK; ws[f"N{r}"].number_format = PCT2; ws[f"N{r}"].border = BOX

# 예시 행
ws["A7"] = "[예시] 차량용 무선청소기"; ws["B7"] = "차량용청소기"
ws["C7"] = 45000; ws["D7"] = 20250; ws["E7"] = 3000
for c in "ABCDE":
    ws[f"{c}7"].fill = GREYFILL

ws["A48"] = "판정 읽는 법"; ws["A48"].font = BOLD
ws["A49"] = "· 게이트1 PASS = 네이버 기여이익이 기준표 임계값(기본 18,000원) 이상"; ws["A49"].font = NOTE
ws["A50"] = "· 허용 CPA = 목표 순이익을 남기고 광고에 쓸 수 있는 건당 상한. 음수면 광고비 0원이어도 목표 미달"; ws["A50"].font = NOTE
ws["A51"] = "· 필요전환율 = CPC 250원 기준. 국내 이커머스 평균 1.33~2%, 상위 20% 3.3%. 4% 넘으면 사실상 불가"; ws["A51"].font = NOTE
ws["A52"] = "· 토스(광고경유)는 판매수수료 0%지만 CPC 광고비가 별도이므로 다른 채널과 직접 비교하지 말 것"; ws["A52"].font = NOTE
ws.freeze_panes = "C7"

# ============================================================
# 시트 2: 수요스캔
# ============================================================
d = wb.create_sheet("수요스캔", 1)
d["A1"] = "수요 스캔 (1층 필터)"; d["A1"].font = TITLE
d["A2"] = "파란 글씨 셀만 입력. 월간검색량은 네이버 검색광고 키워드도구, 진폭·YoY는 네이버 데이터랩에서."; d["A2"].font = NOTE
d["A3"] = "계절진폭 = 최근 1년 피크월 ÷ 저점월. YoY는 소수로 입력 (예: -12% → -0.12)."; d["A3"].font = NOTE

dcols = [("A","키워드",18),("B","월간검색량",13),("C","검색량\n판정",10),("D","계절진폭\n(피크÷저점)",13),
         ("E","YoY",10),("F","유형",11),("G","종합판정",11),("H","메모",34)]
for c,h,w in dcols:
    cell = d[f"{c}5"]; cell.value=h; cell.font=HDR; cell.fill=HDRFILL
    cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=BOX
    d.column_dimensions[c].width=w
d.row_dimensions[5].height=30

for r in range(6, 46):
    d[f"A{r}"].font = BLUE; d[f"A{r}"].border = BOX
    d[f"B{r}"].font = BLUE; d[f"B{r}"].number_format='#,##0'; d[f"B{r}"].border = BOX
    d[f"D{r}"].font = BLUE; d[f"D{r}"].number_format=MULT; d[f"D{r}"].border = BOX
    d[f"E{r}"].font = BLUE; d[f"E{r}"].number_format=PCT1; d[f"E{r}"].border = BOX
    d[f"H{r}"].font = BLUE; d[f"H{r}"].border = BOX
    d[f"C{r}"] = f'=IF(B{r}="","",IF(B{r}>={G_SEARCH},"PASS","FAIL"))'
    d[f"F{r}"] = f'=IF(D{r}="","",IF(D{r}>=3,"시즌형",IF(D{r}<=1.5,"상시형","중간")))'
    d[f"G{r}"] = (f'=IF(OR(B{r}="",D{r}="",E{r}=""),"",'
                  f'IF(AND(B{r}>={G_SEARCH},OR(D{r}>=3,D{r}<=1.5),E{r}>=-0.1),"PASS","FAIL"))')
    for c in "CFG":
        d[f"{c}{r}"].font = BOLD if c in "CG" else BLACK
        d[f"{c}{r}"].alignment = Alignment(horizontal="center"); d[f"{c}{r}"].border = BOX

d["A6"]="[예시] 차량용청소기"; d["B6"]=22000; d["D6"]=1.4; d["E6"]=-0.05; d["H6"]="상시형, 시장 유지 중"
for c in "ABDEH": d[f"{c}6"].fill = GREYFILL

d["A48"]="판정 기준"; d["A48"].font=BOLD
d["A49"]="· 검색량 PASS = 기준표 최소 월간검색량 이상. 검색량은 검색 유입의 물리적 상한임"; d["A49"].font=NOTE
d["A50"]="· 유형: 진폭 3배 이상 시즌형 / 1.5배 이하 상시형 / 그 사이는 애매 → 제외 권장"; d["A50"].font=NOTE
d["A51"]="· 종합판정 PASS = 검색량 통과 AND (시즌형 또는 상시형) AND YoY -10% 이내"; d["A51"].font=NOTE
d["A52"]="· 예외: 토스쇼핑은 홈피드 추천(pCTR×pCVR) 기반이라 검색량 기준이 적용되지 않음"; d["A52"].font=NOTE
d.freeze_panes = "B6"

# ============================================================
# 시트 3: 리뷰증가트래커
# ============================================================
t = wb.create_sheet("리뷰증가트래커", 2)
t["A1"] = "리뷰 증가 속도 트래커 (실판매 신호)"; t["A1"].font = TITLE
t["A2"] = "타깃 키워드로 검색해 상위 10개 상품의 리뷰 수를 오늘 기록하고, 14일 뒤 같은 상품을 재측정하세요."; t["A2"].font = NOTE
t["A3"] = "리뷰는 실제 구매의 결과물이며, 직접 셀 수 있는 유일한 공개 판매 흔적입니다."; t["A3"].font = NOTE
t["A4"] = "키워드 열에는 아래 요약표와 정확히 같은 문자열을 입력하세요 (SUMIF 매칭)."; t["A4"].font = NOTE

tcols=[("A","키워드",18),("B","순위",7),("C","상품명",30),("D","1차측정일",12),
       ("E","1차리뷰수",11),("F","2차측정일",12),("G","2차리뷰수",11),("H","증가분",10)]
for c,h,w in tcols:
    cell=t[f"{c}6"]; cell.value=h; cell.font=HDR; cell.fill=HDRFILL
    cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=BOX
    t.column_dimensions[c].width=w

for r in range(7, 47):
    for c in "ABCDEFG":
        t[f"{c}{r}"].font = BLUE; t[f"{c}{r}"].border = BOX
    t[f"B{r}"].alignment = Alignment(horizontal="center")
    t[f"E{r}"].number_format='#,##0'; t[f"G{r}"].number_format='#,##0'
    t[f"D{r}"].number_format='yyyy-mm-dd'; t[f"F{r}"].number_format='yyyy-mm-dd'
    t[f"H{r}"] = f'=IF(OR(E{r}="",G{r}=""),0,G{r}-E{r})'
    t[f"H{r}"].font = BLACK; t[f"H{r}"].number_format='#,##0;-#,##0;'; t[f"H{r}"].border = BOX

demo = [(1,"OO 무선 차량용 청소기",1240,1310),
        (2,"XX 차량용 진공청소기",860,905),
        (3,"YY 미니 카클리너",412,442)]
for i,(rank,nm,r1,r2) in enumerate(demo):
    rr = 7+i
    t[f"A{rr}"]="[예시] 차량용청소기"; t[f"B{rr}"]=rank; t[f"C{rr}"]=nm
    t[f"E{rr}"]=r1; t[f"G{rr}"]=r2
    for c in "ABCEG": t[f"{c}{rr}"].fill = GREYFILL

t["A49"]="키워드별 요약"; t["A49"].font=BOLD
scols=[("A","키워드",18),("B","총 증가분",11),("C","최대 증가분",12),("D","1위 집중도",11),("E","판정",10),("F","해석",40)]
for c,h,w in scols:
    cell=t[f"{c}50"]; cell.value=h; cell.font=HDR; cell.fill=HDRFILL
    cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=BOX
for r in range(51, 59):
    t[f"A{r}"].font = BLUE; t[f"A{r}"].border = BOX
    t[f"B{r}"] = f'=IF(A{r}="","",SUMIF($A$7:$A$46,A{r},$H$7:$H$46))'
    t[f"C{r}"] = f'=IF(A{r}="","",SUMPRODUCT(MAX(($A$7:$A$46=A{r})*($H$7:$H$46))))'
    t[f"D{r}"] = f'=IF(OR(B{r}="",B{r}=0),"",C{r}/B{r})'
    t[f"E{r}"] = (f'=IF(A{r}="","",IF(AND(B{r}>={G_RVGROW},D{r}<{G_CONC}),"PASS","FAIL"))')
    t[f"F{r}"] = (f'=IF(A{r}="","",IF(B{r}<{G_RVGROW},"시장 너무 작음",'
                  f'IF(D{r}>={G_CONC},"승자독식 — 진입 불리","분산 시장 — 진입 가능")))')
    t[f"B{r}"].number_format='#,##0'; t[f"C{r}"].number_format='#,##0'; t[f"D{r}"].number_format=PCT1
    for c in "BCDEF":
        t[f"{c}{r}"].font = BOLD if c=="E" else BLACK; t[f"{c}{r}"].border = BOX
    t[f"E{r}"].alignment = Alignment(horizontal="center")
t["A51"]="[예시] 차량용청소기"; t["A51"].fill=GREYFILL

t["A61"]="읽는 법"; t["A61"].font=BOLD
t["A62"]="· 총 증가분 = 14일간 상위 10개 리뷰 증가 합계. 기준표 임계값 미만이면 죽은 시장"; t["A62"].font=NOTE
t["A63"]="· 1위 집중도 = 최대 증가분 ÷ 총 증가분. 높을수록 한 셀러가 독식 중"; t["A63"].font=NOTE
t["A64"]="· 절대 판매량은 알 수 없음 (리뷰 작성률 미공개). 상대 비교와 추세 판단에만 사용할 것"; t["A64"].font=NOTE
t["A65"]="· 내 상품 등록 후 같은 방식으로 내 리뷰 증가 속도를 재면 점유율을 추정할 수 있음"; t["A65"].font=NOTE
t.freeze_panes = "B7"

out = "/home/user/-/tools/sourcing-gate-calculator.xlsx"
wb.save(out)
print("saved:", out)
