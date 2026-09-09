# 소싱 게이트 계산기

`sourcing-gate-calculator.xlsx` — 도매사이트를 직접 보면서 숫자만 입력하면
`docs/consignment-platform-strategy.md`에서 정한 게이트를 자동 판정하는 워크북.

## 시트

| 시트 | 용도 | 입력 |
|---|---|---|
| 소싱계산기 | 게이트 1 (단위경제) | 상품명·키워드·판매가·공급가·배송비 |
| 수요스캔 | 1층 필터 (수요) | 키워드·월간검색량·계절진폭·YoY |
| 리뷰증가트래커 | 실판매 신호 측정 | 상위 10개 상품의 1차/2차 리뷰 수 |
| 기준표 | 채널별 수수료율·게이트 임계값 | (노란 셀 수정 시 전 시트에 반영) |

파란 글씨 셀만 입력한다. 회색 배경 행은 입력 형식을 보여주는 예시이므로 덮어쓴다.

## 재생성

```bash
pip install openpyxl
python3 build_calculator.py
```

## 검증

이 컨테이너의 LibreOffice가 동작하지 않아(1개 수식 파일도 타임아웃) `recalc.py` 대신
`formulas` 패키지로 독립 평가했다. **1,038개 셀 전수 평가, 오류 0건.**
예시 행 결과가 문서의 수기 계산과 일치함을 확인했다
(네이버 기여이익 18,765원, 허용 CPA 8,765원, 필요전환율 2.85%).

수식은 전부 2007년 이전 함수(IF/OR/AND/SUMIF/SUMPRODUCT/MAX)만 사용한다.
`MAXIFS`는 이 환경에서 검증할 수 없어 `SUMPRODUCT(MAX(...))`로 대체했다.

```bash
pip install formulas
python3 -c "
import formulas, warnings; warnings.filterwarnings('ignore')
sol = formulas.ExcelModel().loads('sourcing-gate-calculator.xlsx').finish().calculate()
ERR = ('#REF!','#VALUE!','#NAME?','#DIV/0!','#N/A','#NUM!','#NULL!')
bad = [k for k,v in sol.items() if \"'!\" in k and isinstance(getattr(v,'value',[[None]])[0,0] if hasattr(v,'value') else None, str) and str(v.value[0,0]).strip() in ERR]
print('오류 셀:', len(bad))"
```
