"""
SG&A 분기별 분석 설정 파일
============================================================
매 분기 작업 순서:
  1. data/raw/ 아래에 분기 폴더 생성  예) 2025_Q4
  2. 해당 폴더에 법인 Excel 파일 복사
  3. data/exchange_rates.xlsx 에 새 분기 환율 행 추가
  4. 아래 TARGET_PERIOD 를 새 분기로 변경
  5. run.bat 실행

폴더 이름 규칙: YYYY_Q#  (예: 2025_Q4, 2026_Q1)
  → 스크립트가 자동으로 분기 코드로 변환합니다.
============================================================
"""

from pathlib import Path

# ===========================================================
# ★ 매 분기 수정 항목 ★
# ===========================================================

# 분석 대상 분기 (data/raw/ 폴더명 기준, 예: "2025_Q4")
TARGET_PERIOD = "2025_Q4"

# 환율은 data/exchange_rates.xlsx 에서 관리합니다.
# 새 분기 추가 시 해당 파일을 Excel 로 열고 행을 추가하세요.

# ===========================================================
# 고정 설정 (변경 빈도 낮음)
# ===========================================================

# 프로젝트 루트 경로
BASE_DIR = Path(__file__).parent

# 중간 처리 파일 저장 경로 (분기별 통합 데이터)
PROCESSED_FOLDER = BASE_DIR / "data" / "processed"

# 최종 분석 결과 파일
OUTPUT_FOLDER = BASE_DIR / "output"
OUTPUT_FILE = OUTPUT_FOLDER / f"InBody_SGA_Analysis_{TARGET_PERIOD}.xlsx"

# 법인 전체명 → 약칭 매핑
ENTITY_NAME_MAPPING = {
    "BIOSPACE LATIN AMERICA S DE RL DE CV(*)": "Mexico",
    "BWA": "BWA",
    "Biospace Co.,Ltd.": "China",
    "Biospace Inc. DBA INBODY": "USA",
    "InBody India Pvt. Ltd": "India",
    "InBody Japan Inc.": "Japan",
    "InBody Oceania ": "Oceania",
    "㈜코르트": "KOROT",
    "Inbody Europe B.V.": "Europe",
    "㈜인바디헬스케어": "헬스케어",
    "(주)삼한정공": "삼한정공",
    "INBODY TURKEY MEDİKAL TİCARET LİMİTED ŞİRKETİ": "Turkey",
    "INBODY ASIA SDN. BHD.": "Asia",
    "주식회사 인바디": "HQ",
    "케이오씨피 프로젝트 제2호 벤처투자조합": "KOCP",
    "InBody Vietnam": "Vietnam",
    "ADJ": "연결조정",
}

# 법인별 통화 매핑
ENTITY_CURRENCY_MAP = {
    "BIOSPACE LATIN AMERICA S DE RL DE CV(*)": "MXN",
    "BWA": "USD",
    "Biospace Co.,Ltd.": "CNH",
    "Biospace Inc. DBA INBODY": "USD",
    "InBody India Pvt. Ltd": "INR",
    "InBody Japan Inc.": "JPY",
    "InBody Oceania ": "AUD",
    "㈜코르트": "KRW",
    "Inbody Europe B.V.": "EUR",
    "㈜인바디헬스케어": "KRW",
    "(주)삼한정공": "KRW",
    "INBODY TURKEY MEDİKAL TİCARET LİMİTED ŞİRKETİ": "TRY",
    "INBODY ASIA SDN. BHD.": "MYR",
    "주식회사 인바디": "KRW",
    "케이오씨피 프로젝트 제2호 벤처투자조합": "KRW",
    "InBody Vietnam": "VND",
    "ADJ": "KRW",
}

# 법인 표시 순서
ENTITY_ORDER = [
    "HQ", "USA", "Japan", "China", "Europe", "Asia",
    "India", "Mexico", "Oceania", "BWA", "Vietnam",
    "Turkey", "KOROT", "헬스케어", "삼한정공", "KOCP", "연결조정",
]

# 원본 계정명 → 표시 계정명 정규화 매핑
# (processed 파일은 그대로 유지, build_pivot 시 적용)
ACCOUNT_NAME_MAP = {
    "매출액": "매출",
    "기타수익": "영업외수익",
    "기타비용": "영업외비용",
}

# 인건비 구성 계정 (합산 → 인건비 표시)
LABOR_ACCOUNTS = [
    "임원급여", "급여", "잡급", "제수당", "상여금", "임원상여", "퇴직급여",
]

# 기타 판관비 구성 계정 (합산 → 기타 표시)
MISC_SGA_ACCOUNTS = [
    "접대비", "무형자산상각비", "소모품비", "세금과공과", "차량유지비",
    "전력비", "판매보증비", "보험료", "도서인쇄비", "통신비", "잡비",
    "지급임차료", "교육훈련비", "수도광열비", "수선유지비", "회의비", "제회비",
]

# 계정과목 표시 순서 (전체 PL)
REV_ACCOUNT_ORDER = [
    "매출",
    "매출원가",
    "인건비",
    "복리후생비",
    "광고선전비",
    "지급수수료",
    "운반비",
    "경상연구개발비",
    "판매수수료",
    "여비교통비",
    "대손상각비",
    "감가상각비",
    "사용권자산상각비",
    "기타",
    "영업외수익",
    "영업외비용",
    "금융수익",
    "금융비용",
    "지분법손익",
    "법인세비용",
    "세후중단영업손익",
    "당기순이익",
    "포괄손익",
    "총포괄손익",
]

# SG&A 분석 대상 계정 (SG&A 시트용)
SGA_TARGET_ACCOUNTS = [
    "매출",
    "매출원가",
    "매출총이익",   # 계산: 매출 - 매출원가
    "판관비",       # 계산: 아래 항목 합계
    "인건비",       # 계산: LABOR_ACCOUNTS 합산
    "복리후생비",
    "광고선전비",
    "지급수수료",
    "운반비",
    "경상연구개발비",
    "판매수수료",
    "여비교통비",
    "대손상각비",
    "감가상각비",
    "사용권자산상각비",
    "기타",         # 계산: MISC_SGA_ACCOUNTS 합산
    "영업이익",     # 계산: 매출총이익 - 판관비
]

# 판관비 구성 계정 - 개별 계정 flat list (인건비+기타 포함)
SGA_COMPONENTS = (
    LABOR_ACCOUNTS
    + [
        "복리후생비", "광고선전비", "지급수수료", "운반비",
        "경상연구개발비", "판매수수료", "여비교통비", "대손상각비",
        "감가상각비", "사용권자산상각비",
    ]
    + MISC_SGA_ACCOUNTS
)

# 계산 항목 (Pivot에서 읽지 않고 계산)
CALCULATED_ACCOUNTS = ["매출총이익", "인건비", "기타", "판관비", "영업이익"]
