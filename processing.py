# -*- coding: utf-8 -*-
"""
산지로드 결산서 처리 핵심 로직.
- 취합 결산서(xlsx)를 읽어 사이트별 수수료 규칙 / 상품별 고정 정산단가를 적용해
  매입가·정산가·마진·마진률을 계산한다.
- 규칙에 해당하지 않는 행은 '확인 필요' 목록으로 분리해 사용자가 직접 값을 넣도록 한다.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

RAW_COLUMNS = [
    "판매사", "고객선택옵션", "주문수량", "공급사배송비", "결제금액",
    "매입단가", "매입가", "정산가", "마진", "마진률", "비고",
    "판매사주문번호", "수령인연락처",
]

# ---------------------------------------------------------------------------
# 설정값 (사이트 수수료 규칙 / 상품별 고정 정산단가) 로드 & 저장
# ---------------------------------------------------------------------------

DEFAULT_RATE_RULES = [
    # keyword: 판매사명에 포함되면 매칭 (대소문자 무시). 먼저 등록된 규칙이 우선 적용됨.
    {"keyword": "카카오", "type": "percent", "rate": 0.90, "note": "결제금액의 90% (수수료 10%)"},
    {"keyword": "알리익스프레스", "type": "percent", "rate": 0.91, "note": ""},
    {"keyword": "자사몰", "type": "percent", "rate": 0.96, "note": ""},
    {"keyword": "네이버 스마트스토어", "type": "percent", "rate": 0.90, "note": ""},
    {"keyword": "티딜", "type": "percent", "rate": 0.85, "note": ""},
    {"keyword": "토스", "type": "percent", "rate": 0.967, "note": ""},
    {"keyword": "지마켓", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "옥션", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "제이슨딜", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "GS SHOP", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "NS홈쇼핑", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "11번가", "type": "same_as_payment", "rate": 1.0, "note": ""},
]

# 캐시딜: 기본 90%, 특정 상품만 예외 요율 적용
CASHDEAL_KEYWORD = "캐시딜"
CASHDEAL_DEFAULT_RATE = 0.90
DEFAULT_CASHDEAL_EXCEPTIONS = [
    {"상품명": "홍주부아카시아향사양꿀2.4kg", "rate": 0.80},
]

# 고정 정산단가(개당) 방식 사이트 - 결제금액과 무관하게 상품별로 정해진 단가 사용
FIXED_PRICE_SITE_KEYWORDS = ["케이딜", "꿈꾸는이웃", "홈앤쇼핑", "LG 복지몰", "제트언스"]

DATA_DIR = Path(__file__).parent / "data"
RATE_RULES_PATH = DATA_DIR / "site_rules.json"
CASHDEAL_PATH = DATA_DIR / "cashdeal_exceptions.json"
FIXED_PRICE_PATH = DATA_DIR / "fixed_price_table.csv"


def load_rate_rules() -> list[dict]:
    if RATE_RULES_PATH.exists():
        return json.loads(RATE_RULES_PATH.read_text(encoding="utf-8"))
    return DEFAULT_RATE_RULES


def save_rate_rules(rules: list[dict]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    RATE_RULES_PATH.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")


def load_cashdeal_exceptions() -> list[dict]:
    if CASHDEAL_PATH.exists():
        return json.loads(CASHDEAL_PATH.read_text(encoding="utf-8"))
    return DEFAULT_CASHDEAL_EXCEPTIONS


def save_cashdeal_exceptions(rows: list[dict]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CASHDEAL_PATH.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")


def load_fixed_price_table() -> pd.DataFrame:
    if FIXED_PRICE_PATH.exists():
        return pd.read_csv(FIXED_PRICE_PATH, encoding="utf-8-sig")
    return pd.DataFrame(columns=["판매사", "상품명", "정산단가(개당)"])


def save_fixed_price_table(df: pd.DataFrame) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(FIXED_PRICE_PATH, index=False, encoding="utf-8-sig")


# ---------------------------------------------------------------------------
# 취합 결산서(xlsx) 읽기
# ---------------------------------------------------------------------------

def read_raw_settlement(file) -> pd.DataFrame:
    """취합 결산서의 '결산서' 시트에서 상세 표(9행 헤더, 10행부터 데이터)를 읽는다."""
    df_full = pd.read_excel(file, sheet_name="결산서", header=None)
    header_row_idx = None
    for idx in range(min(20, len(df_full))):
        row_vals = df_full.iloc[idx].astype(str).tolist()
        if "판매사" in row_vals and "결제금액" in row_vals:
            header_row_idx = idx
            break
    if header_row_idx is None:
        raise ValueError("'결산서' 시트에서 헤더 행(판매사/결제금액 등)을 찾지 못했습니다.")

    header = df_full.iloc[header_row_idx]
    col_map = {}
    for col_name in ["판매사", "고객선택옵션", "주문수량", "공급사배송비", "결제금액",
                      "매입단가", "매입가", "정산가", "마진", "마진률", "비고",
                      "판매사주문번호", "수령인연락처"]:
        matches = header[header == col_name].index.tolist()
        if matches:
            col_map[col_name] = matches[0]

    required = ["판매사", "고객선택옵션", "주문수량", "공급사배송비", "결제금액", "매입단가"]
    missing = [c for c in required if c not in col_map]
    if missing:
        raise ValueError(f"필수 컬럼을 찾지 못했습니다: {missing}")

    data = df_full.iloc[header_row_idx + 1:].copy()
    out = pd.DataFrame()
    for col_name, col_idx in col_map.items():
        out[col_name] = data[col_idx].values
    out = out.dropna(subset=["판매사"]).reset_index(drop=True)

    for num_col in ["주문수량", "공급사배송비", "결제금액", "매입단가"]:
        out[num_col] = pd.to_numeric(out[num_col], errors="coerce").fillna(0)

    return out


# ---------------------------------------------------------------------------
# 계산 로직
# ---------------------------------------------------------------------------

@dataclass
class ProcessResult:
    df: pd.DataFrame                 # 계산이 끝난 전체 데이터 (확인 필요 행 포함)
    needs_review: pd.DataFrame       # 규칙이 없어 정산가를 계산하지 못한 행
    rate_rules: list = field(default_factory=list)
    fixed_price_table: pd.DataFrame = field(default_factory=pd.DataFrame)


def _match_keyword(site: str, keyword: str) -> bool:
    return keyword.lower() in str(site).lower()


def process_settlement(
    raw: pd.DataFrame,
    rate_rules: list[dict] | None = None,
    cashdeal_exceptions: list[dict] | None = None,
    fixed_price_table: pd.DataFrame | None = None,
) -> ProcessResult:
    rate_rules = rate_rules if rate_rules is not None else load_rate_rules()
    cashdeal_exceptions = cashdeal_exceptions if cashdeal_exceptions is not None else load_cashdeal_exceptions()
    fixed_price_table = fixed_price_table if fixed_price_table is not None else load_fixed_price_table()

    fixed_price_lookup = {
        (row["판매사"], row["상품명"]): row["정산단가(개당)"]
        for _, row in fixed_price_table.iterrows()
    }
    cashdeal_ex_lookup = {row["상품명"]: row["rate"] for row in cashdeal_exceptions}

    df = raw.copy()
    df["매입가"] = df["매입단가"] * df["주문수량"] + df["공급사배송비"]

    settlement_prices = []
    rule_applied = []

    for _, row in df.iterrows():
        site = row["판매사"]
        product = row["고객선택옵션"]
        payment = row["결제금액"]
        qty = row["주문수량"]

        # 1) 고정 정산단가 사이트 우선 확인
        if any(_match_keyword(site, kw) for kw in FIXED_PRICE_SITE_KEYWORDS):
            unit_price = fixed_price_lookup.get((site, product))
            if unit_price is not None:
                settlement_prices.append(unit_price * qty)
                rule_applied.append(f"고정단가({unit_price:,.1f}/개)")
            else:
                settlement_prices.append(None)
                rule_applied.append("고정단가 미등록")
            continue

        # 2) 캐시딜 예외 처리
        if _match_keyword(site, CASHDEAL_KEYWORD):
            rate = cashdeal_ex_lookup.get(product, CASHDEAL_DEFAULT_RATE)
            settlement_prices.append(payment * rate)
            rule_applied.append(f"캐시딜 {rate:.0%}")
            continue

        # 3) 일반 퍼센트/동일금액 규칙
        matched = None
        for rule in rate_rules:
            if _match_keyword(site, rule["keyword"]):
                matched = rule
                break
        if matched is not None:
            settlement_prices.append(payment * matched["rate"])
            rule_applied.append(f'{matched["keyword"]} {matched["rate"]:.1%}')
        else:
            settlement_prices.append(None)
            rule_applied.append("규칙 없음")

    df["정산가"] = settlement_prices
    df["적용규칙"] = rule_applied
    df["마진"] = df["정산가"] - df["매입가"]
    df["마진률"] = df.apply(
        lambda r: (r["마진"] / r["정산가"]) if pd.notna(r["정산가"]) and r["정산가"] not in (0, None) else None,
        axis=1,
    )

    needs_review = df[df["정산가"].isna()].copy()

    return ProcessResult(
        df=df,
        needs_review=needs_review,
        rate_rules=rate_rules,
        fixed_price_table=fixed_price_table,
    )


# ---------------------------------------------------------------------------
# 요약(대시보드용) 집계
# ---------------------------------------------------------------------------

def summarize_by_site(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("판매사", dropna=False).agg(
        정산가=("정산가", "sum"),
        마진=("마진", "sum"),
        주문수량=("주문수량", "sum"),
    ).reset_index()
    g["마진률"] = g.apply(lambda r: (r["마진"] / r["정산가"]) if r["정산가"] else None, axis=1)
    return g.sort_values("정산가", ascending=False, na_position="last")


def summarize_by_product(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("고객선택옵션", dropna=False).agg(
        주문수량=("주문수량", "sum"),
        정산가=("정산가", "sum"),
        마진=("마진", "sum"),
    ).reset_index()
    return g.sort_values("주문수량", ascending=False, na_position="last")
