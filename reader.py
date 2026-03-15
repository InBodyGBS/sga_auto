"""
법인 원본 Excel 파일 읽기 및 분기별 통합 데이터 생성

폴더 명명 규칙: data/raw/YYYY_Q#  (예: 2025_Q4, 2026_Q1)
스크립트가 폴더를 자동 감지하여 분기 코드로 변환합니다.
"""

import re
import pandas as pd
from pathlib import Path

from config import (
    ENTITY_CURRENCY_MAP,
    ENTITY_NAME_MAPPING,
    PROCESSED_FOLDER,
    BASE_DIR,
    ACCOUNT_NAME_MAP,
)

RAW_FOLDER   = BASE_DIR / "data" / "raw"
RATES_FILE   = BASE_DIR / "data" / "exchange_rates.xlsx"

# 폴더명 패턴: 2025_Q4, 2024_Q3, 2026_Q1 등
_FOLDER_PATTERN = re.compile(r"^(\d{4})_Q([1-4])$")

# ── 환율 테이블 (Excel 파일에서 로드) ──────────────────────

def load_exchange_rates() -> dict[str, dict[str, float]]:
    """
    data/exchange_rates.xlsx → {period: {currency: rate}} 딕셔너리 반환
    컬럼 구조: Period | Currency | Rate
    """
    if not RATES_FILE.exists():
        raise FileNotFoundError(
            f"환율 파일을 찾을 수 없습니다: {RATES_FILE}\n"
            "data/exchange_rates.xlsx 파일이 있는지 확인하세요."
        )

    df = pd.read_excel(RATES_FILE, header=1)

    # 컬럼명 공백 제거 + Rate 컬럼 정규화 (예: "Rate (1외화 = ?원)" → "Rate")
    df.columns = df.columns.str.strip()
    df.columns = [c if not c.startswith("Rate") else "Rate" for c in df.columns]

    required = {"Period", "Currency", "Rate"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"exchange_rates.xlsx 에 필수 컬럼이 없습니다: {missing}\n"
            f"현재 컬럼: {list(df.columns)}"
        )

    # 빈 행 제거
    df = df.dropna(subset=["Period", "Currency", "Rate"])
    df["Period"]   = df["Period"].astype(str).str.strip()
    df["Currency"] = df["Currency"].astype(str).str.strip()
    df["Rate"]     = pd.to_numeric(df["Rate"], errors="coerce")

    rates: dict[str, dict[str, float]] = {}
    for _, row in df.iterrows():
        period   = row["Period"]
        currency = row["Currency"]
        rate     = float(row["Rate"])
        rates.setdefault(period, {})[currency] = rate

    return rates


# 모듈 로드 시 1회만 읽음
_EXCHANGE_RATES: dict[str, dict[str, float]] | None = None

def get_exchange_rates() -> dict[str, dict[str, float]]:
    global _EXCHANGE_RATES
    if _EXCHANGE_RATES is None:
        _EXCHANGE_RATES = load_exchange_rates()
    return _EXCHANGE_RATES


def folder_to_period(folder_name: str) -> str | None:
    """
    폴더명 → 분기 코드 변환
    '2025_Q4' → '2025_Q4'  (EXCHANGE_RATES 키와 동일하게 유지)
    유효하지 않은 이름이면 None 반환
    """
    if _FOLDER_PATTERN.match(folder_name):
        return folder_name
    return None


def period_to_filename(period: str) -> str:
    """
    분기 코드 → 저장 파일명 변환
    '2025_Q4' → 'final_financial_data_2025_Q4.xlsx'
    """
    return f"final_financial_data_{period}.xlsx"


def discover_raw_folders() -> dict[str, Path]:
    """
    data/raw/ 아래의 분기 폴더를 자동 감지
    반환: {period: folder_path}  (시간순 정렬)
    """
    if not RAW_FOLDER.exists():
        return {}

    result = {}
    for folder in sorted(RAW_FOLDER.iterdir()):
        if not folder.is_dir():
            continue
        period = folder_to_period(folder.name)
        if period:
            result[period] = folder
        else:
            print(f"  ⚠️  무시된 폴더 (이름 규칙 불일치): {folder.name}")

    return result


def _get_exchange_rate(period: str, currency: str) -> float:
    """기간과 통화로 환율 조회 (exchange_rates.xlsx 기반)"""
    rates = get_exchange_rates().get(period, {})
    rate = rates.get(currency, None)
    if rate is None:
        print(f"    ⚠️  환율 미설정: {period} / {currency} → 0 처리")
        print(f"        data/exchange_rates.xlsx 에 [{period}] [{currency}] 행을 추가하세요.")
        return 0.0
    return rate


def _preprocess_raw_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """H~J → K, M~O → P 이동 전처리 (원본 파일 형식 대응)"""
    df = df.copy()
    for col in [7, 8, 9]:       # H, I, J
        mask = df[col].notna()
        df.loc[mask, 10] = df.loc[mask, col]
    for col in [12, 13, 14]:    # M, N, O
        mask = df[col].notna()
        df.loc[mask, 15] = df.loc[mask, col]
    return df


def _extract_sheet(file_path: Path, sheet_name: str, period: str, row_end: int) -> pd.DataFrame | None:
    """단일 시트(BS 또는 IS)에서 데이터 추출"""
    try:
        raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        print(f"    ✗ {sheet_name} 시트 읽기 실패: {e}")
        return None

    entity_full = raw.iloc[0, 6]      # G1: 법인 전체명
    currency = ENTITY_CURRENCY_MAP.get(entity_full, "Unknown")
    entity_short = ENTITY_NAME_MAPPING.get(entity_full, entity_full)
    pl_rate = _get_exchange_rate(period, currency)

    raw = _preprocess_raw_sheet(raw)

    # 6행 = 컬럼명, 7행~ = 데이터  (G=6, K=10, P=15, W=22)
    data = raw.iloc[6:row_end, [6, 10, 15, 22]].copy()
    # 법인마다 컬럼명이 다를 수 있으므로 표준 이름으로 고정
    data.columns = ["Account_Code", "Rev_Account", "Classification", "Amount_Raw"]

    data["Entity"] = entity_full
    data["Entity_Short"] = entity_short
    data["Statement"] = sheet_name
    data["Currency"] = currency
    data["Period"] = period
    data["Exchange_Rate"] = pl_rate

    # Rev_Account 없는 행 제거
    data = data[data["Rev_Account"].notna()].copy()

    # Amount 숫자 변환 및 KRW 환산
    data["Amount_Raw"] = pd.to_numeric(data["Amount_Raw"], errors="coerce").fillna(0)
    data["Amount"] = data["Amount_Raw"]
    data["Amount(KRW)"] = data["Amount"] * pl_rate
    data = data.drop(columns=["Amount_Raw"])

    return data


def consolidate_quarter(period: str, folder: Path) -> pd.DataFrame | None:
    """
    분기 폴더의 모든 법인 파일을 읽어 통합 후
    data/processed/final_financial_data_{period}.xlsx 저장
    """
    excel_files = [
        f for f in sorted(folder.glob("*.xlsx"))
        if not f.name.startswith("~")
    ]

    if not excel_files:
        print(f"  ⚠️  xlsx 파일 없음: {folder}")
        return None

    print(f"  총 {len(excel_files)}개 파일 발견")
    all_records = []

    for file_path in excel_files:
        for sheet_name, row_end in [("BS", 231), ("IS", 129)]:
            df = _extract_sheet(file_path, sheet_name, period, row_end)
            if df is not None:
                all_records.append(df)
                entity_short = df["Entity_Short"].iloc[0]
                currency = df["Currency"].iloc[0]
                print(f"    ✓ {file_path.name} [{sheet_name}] - {entity_short} ({currency})")

    if not all_records:
        return None

    merged = pd.concat(all_records, ignore_index=True)

    PROCESSED_FOLDER.mkdir(parents=True, exist_ok=True)
    out_path = PROCESSED_FOLDER / period_to_filename(period)
    merged.to_excel(out_path, index=False)
    print(f"  ✓ 저장: {out_path.name}")

    return merged


def sync_all_quarters() -> dict[str, str]:
    """
    data/raw/ 의 모든 분기 폴더를 스캔하여
    아직 처리되지 않은 분기만 consolidate_quarter() 실행
    반환: {period: 상태}  ("처리됨" / "신규처리" / "파일없음" / "환율누락")
    """
    folders = discover_raw_folders()
    if not folders:
        print("  ⚠️  data/raw/ 에 분기 폴더가 없습니다.")
        print("      폴더 이름 규칙: YYYY_Q#  (예: 2025_Q4)")
        return {}

    status = {}
    for period, folder in folders.items():
        processed_file = PROCESSED_FOLDER / period_to_filename(period)

        if period not in get_exchange_rates():
            print(f"  ⚠️  [{period}] config.py 에 환율 미설정 → 건너뜀")
            status[period] = "환율누락"
            continue

        if processed_file.exists():
            print(f"  ⏭  [{period}] 이미 처리됨 → {processed_file.name}")
            status[period] = "처리됨"
            continue

        print(f"\n  📂 [{period}] 신규 처리: {folder}")
        result = consolidate_quarter(period, folder)
        status[period] = "신규처리" if result is not None else "파일없음"

    return status


def load_all_pl_data() -> pd.DataFrame | None:
    """
    data/processed/ 의 모든 final_financial_data_*.xlsx 로드 후 병합
    """
    files = sorted([
        f for f in PROCESSED_FOLDER.glob("final_financial_data_*.xlsx")
        if not f.name.startswith("~")
    ])

    if not files:
        print("❌ 처리된 데이터가 없습니다.")
        print("   data/raw/YYYY_Q# 폴더에 법인 파일을 넣고 다시 실행하세요.")
        return None

    print(f"\n📂 분기 데이터 {len(files)}개 로드:")
    dfs = []
    for f in files:
        df = pd.read_excel(f)
        print(f"  ✓ {f.name} ({len(df):,}행)")
        dfs.append(df)

    merged = pd.concat(dfs, ignore_index=True)
    print(f"\n  총 {len(merged):,}행 병합 완료")
    return merged


def build_pivot(df: pd.DataFrame, amount_col: str) -> pd.DataFrame:
    """
    IS 데이터 → Period × Entity × Rev_Account 피벗테이블
    amount_col: "Amount(KRW)" 또는 "Amount"
    """
    df_is = df[df["Statement"] == "IS"].copy()
    # 계정명 정규화 (매출액→매출, 기타수익→영업외수익 등)
    df_is["Rev_Account"] = df_is["Rev_Account"].map(
        lambda x: ACCOUNT_NAME_MAP.get(x, x) if pd.notna(x) else x
    )

    pivot = df_is.pivot_table(
        index="Rev_Account",
        columns=["Period", "Entity_Short"],
        values=amount_col,
        aggfunc="sum",
        fill_value=0,
    )

    # 분기별 법인 합계 추가
    for period in pivot.columns.get_level_values("Period").unique():
        period_cols = [c for c in pivot.columns if c[0] == period]
        pivot[(period, "합계")] = pivot[period_cols].sum(axis=1)

    pivot[("전체", "합계")] = pivot.sum(axis=1)
    return pivot.reset_index()
