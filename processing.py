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
import sys
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
    {"keyword": "토스", "type": "percent", "rate": 0.967,
     "note": "주문배송관리 파일 미제공 시 fallback 요율. 실제로는 주문마다 수수료가 달라 "
             "TOSS_INCENTIVE_RATE/TOSS_DEFAULT_RATE로 개별 계산됨 (아래 참고)"},
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

# 토스: 상품(주문)마다 실제 수수료가 달라 정률 규칙 하나로는 계산할 수 없다.
# 토스 '주문배송관리' 파일의 F열('받은 혜택')에 "수수료 0원 적용" 문구가 있으면 수수료 3.3%,
# 빈칸이면 수수료 11%가 적용된다 (2026-09 확인, 사용자 확인 사항). 이 파일이 없으면 위
# DEFAULT_RATE_RULES의 "토스" 규칙(96.7%)으로 fallback 처리된다.
TOSS_KEYWORD = "토스"
TOSS_INCENTIVE_TEXT = "수수료 0원 적용"
TOSS_INCENTIVE_RATE = 1 - 0.033   # 0.967 — F열에 "수수료 0원 적용"이 있는 경우
TOSS_DEFAULT_RATE = 1 - 0.11      # 0.89  — F열이 빈칸인 경우

# 실제로 결산 대상인 사이트 목록 (2026-09 확정). 이름이 정확히 일치하는 행만 정산하고,
# 나머지(맛장군 계열, 쿠팡 개인/법인, 테무, 농가살리기 등 기타 유통 실험 채널)는 결산에서
# 통째로 제외한다. 부분 키워드가 아니라 "정확히 일치"로 비교한다 — "카카오" 같은 키워드로
# 하면 "맛장군 카카오"까지 걸려버리기 때문.
SETTLEMENT_SITES = [
    "꿈꾸는이웃(산지로드)",
    "더로드 네이버 스마트스토어",
    "더로드 자사몰",
    "산지로드 11번가 sjroad_cop",
    "산지로드 LG 복지몰",
    "산지로드 GS SHOP",
    "산지로드 NS홈쇼핑",
    "산지로드 알리익스프레스",
    "산지로드 옥션",
    "산지로드 제이슨",
    "산지로드 지마켓",
    "산지로드 카카오",
    "산지로드 캐시딜",
    "산지로드 케이딜",
    "산지로드 토스",
    "산지로드 티딜",
    "산지로드 홈앤쇼핑",
    "제트언스(산지로드)",
]

def _get_base_dir() -> Path:
    """일반 파이썬 스크립트로 실행할 때와, PyInstaller로 exe로 묶었을 때 모두
    올바른 기준 폴더를 반환한다. exe로 묶인 경우 exe 파일이 있는 폴더를 기준으로
    'data' 폴더를 만들어 데이터가 exe 옆에 저장/유지되도록 한다."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


DATA_DIR = _get_base_dir() / "data"
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


def read_toss_delivery_file(file) -> dict[str, float]:
    """토스 '주문배송관리' 엑셀(주문내역 시트)에서 주문번호별 실제 수수료율을 읽어온다.

    F열 '받은 혜택'에 "수수료 0원 적용" 문구가 있으면 수수료 3.3%(요율 0.967),
    빈칸이면 수수료 11%(요율 0.89)가 적용된다. 반환값은 {주문번호(str): 요율} 딕셔너리로,
    process_settlement(..., toss_delivery_rates=...)에 그대로 넘기면 된다.
    """
    df_full = pd.read_excel(file, sheet_name=0, header=None)
    header_row_idx = None
    for idx in range(min(10, len(df_full))):
        row_vals = df_full.iloc[idx].astype(str).tolist()
        if "주문번호" in row_vals and "받은 혜택" in row_vals:
            header_row_idx = idx
            break
    if header_row_idx is None:
        raise ValueError("주문배송관리 파일에서 헤더 행(주문번호/받은 혜택)을 찾지 못했습니다.")

    header = df_full.iloc[header_row_idx]
    col_order = header[header == "주문번호"].index[0]
    col_benefit = header[header == "받은 혜택"].index[0]

    data = df_full.iloc[header_row_idx + 1:]
    rates: dict[str, float] = {}
    for _, row in data.iterrows():
        order_no = row[col_order]
        if pd.isna(order_no):
            continue
        order_no = str(order_no).strip()
        # 헤더 바로 아래 '수정 가능/수정 불가' 안내 행 등 숫자가 아닌 잡음 행은 건너뜀
        if not order_no or not order_no.replace(".", "", 1).isdigit():
            continue
        benefit = row[col_benefit]
        benefit = "" if pd.isna(benefit) else str(benefit)
        rate = TOSS_INCENTIVE_RATE if TOSS_INCENTIVE_TEXT in benefit else TOSS_DEFAULT_RATE
        rates[order_no] = rate
    return rates


# ---------------------------------------------------------------------------
# 계산 로직
# ---------------------------------------------------------------------------

@dataclass
class ProcessResult:
    df: pd.DataFrame                 # 계산이 끝난 전체 데이터 (확인 필요 행 포함)
    needs_review: pd.DataFrame       # 규칙이 없어 정산가를 계산하지 못한 행
    excluded: pd.DataFrame = field(default_factory=pd.DataFrame)  # SETTLEMENT_SITES 밖이라 제외된 행
    rate_rules: list = field(default_factory=list)
    fixed_price_table: pd.DataFrame = field(default_factory=pd.DataFrame)


def _match_keyword(site: str, keyword: str) -> bool:
    return keyword.lower() in str(site).lower()


def process_settlement(
    raw: pd.DataFrame,
    rate_rules: list[dict] | None = None,
    cashdeal_exceptions: list[dict] | None = None,
    fixed_price_table: pd.DataFrame | None = None,
    toss_delivery_rates: dict[str, float] | None = None,
) -> ProcessResult:
    rate_rules = rate_rules if rate_rules is not None else load_rate_rules()
    cashdeal_exceptions = cashdeal_exceptions if cashdeal_exceptions is not None else load_cashdeal_exceptions()
    fixed_price_table = fixed_price_table if fixed_price_table is not None else load_fixed_price_table()

    fixed_price_lookup = {
        (row["판매사"], row["상품명"]): row["정산단가(개당)"]
        for _, row in fixed_price_table.iterrows()
    }
    cashdeal_ex_lookup = {row["상품명"]: row["rate"] for row in cashdeal_exceptions}

    # 0) 결산 대상 사이트만 남기고 나머지(맛장군, 쿠팡, 기타 실험 채널 등)는 제외
    in_scope_mask = raw["판매사"].isin(SETTLEMENT_SITES)
    excluded = raw[~in_scope_mask].copy()
    df = raw[in_scope_mask].copy()
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

        # 2.5) 토스: 주문배송관리 파일이 있으면 주문번호별 실제 요율을 우선 적용
        #      (없으면 아래 3)의 DEFAULT_RATE_RULES "토스" 96.7% 규칙으로 fallback)
        if toss_delivery_rates is not None and _match_keyword(site, TOSS_KEYWORD):
            order_no = row.get("판매사주문번호")
            order_no = "" if pd.isna(order_no) else str(order_no).strip()
            rate = toss_delivery_rates.get(order_no)
            if rate is not None:
                settlement_prices.append(payment * rate)
                rule_applied.append(f"토스 배송관리 매칭 ({rate:.1%})")
            else:
                settlement_prices.append(None)
                rule_applied.append("토스 배송관리 미매칭 - 확인 필요")
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
        excluded=excluded,
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


# ---------------------------------------------------------------------------
# 담당자별 원가 검토/수정 반영
#
# 사이트 담당자가 나뉘어 있어, "결산서 최종본"을 각자 받아 매입단가(원가)만 확인/수정한
# 뒤 다시 올리면 그 사람이 담당하는 사이트 행만 반영한다. 실수로 다른 담당자 사이트 값을
# 건드렸더라도, 그 사이트가 이 담당자 소관이 아니면 통째로 무시한다 (아래 keyword 매칭 기준).
# ---------------------------------------------------------------------------

REVIEWER_SITE_KEYWORDS: dict[str, list[str]] = {
    "박정배": ["더로드 네이버 스마트스토어", "더로드 자사몰", "산지로드 LG 복지몰",
              "산지로드 GS SHOP", "산지로드 알리익스프레스", "산지로드 옥션",
              "산지로드 제이슨", "산지로드 지마켓", "산지로드 카카오", "산지로드 캐시딜"],
    "김슬기": ["산지로드 NS홈쇼핑", "산지로드 케이딜", "산지로드 티딜", "산지로드 홈앤쇼핑"],
    "이재성": ["산지로드 11번가 sjroad_cop", "꿈꾸는이웃(산지로드)", "산지로드 토스",
              "제트언스(산지로드)"],
}
# 위 셋을 합치면 SETTLEMENT_SITES 18개와 정확히 일치한다 (미배정 사이트가 없어야 정상).


def unassigned_sites(df: pd.DataFrame) -> list[str]:
    """세 담당자 키워드 중 어디에도 매칭되지 않는 사이트 이름 목록."""
    all_sites = [s for s in df["판매사"].dropna().unique().tolist()]
    owned = set()
    for keywords in REVIEWER_SITE_KEYWORDS.values():
        owned |= {s for s in all_sites if any(_match_keyword(s, kw) for kw in keywords)}
    return sorted(set(all_sites) - owned)


def apply_reviewer_corrections(
    df: pd.DataFrame,
    reviewed: pd.DataFrame,
    site_keywords: list[str],
) -> tuple[pd.DataFrame, dict]:
    """담당자가 검토/수정한 파일(reviewed)의 '매입단가'를 df에 반영한다.

    site_keywords에 매칭되는 사이트 행만 반영 대상이며, reviewed에 다른 사이트 행이
    섞여 있어도 전부 무시한다(다른 담당자 실수 방지). 매칭은 (판매사, 고객선택옵션,
    판매사주문번호)로 하고, 판매사주문번호가 없으면 (판매사, 고객선택옵션)로 대체한다.
    매입단가가 바뀐 행은 매입가/마진/마진률을 다시 계산한다. df, reviewed 원본은
    바꾸지 않고 새 DataFrame을 반환한다.
    """

    def is_owned(site) -> bool:
        return any(_match_keyword(site, kw) for kw in site_keywords)

    use_order_no = "판매사주문번호" in df.columns and "판매사주문번호" in reviewed.columns

    def make_key(row) -> tuple:
        if use_order_no:
            return (row["판매사"], row["고객선택옵션"], str(row.get("판매사주문번호", "")))
        return (row["판매사"], row["고객선택옵션"])

    owned_reviewed = reviewed[reviewed["판매사"].apply(is_owned)]
    ignored_rows = len(reviewed) - len(owned_reviewed)

    correction_map = {make_key(r): r["매입단가"] for _, r in owned_reviewed.iterrows()}

    out = df.copy()
    owned_mask = out["판매사"].apply(is_owned)
    changed = 0
    for idx, row in out[owned_mask].iterrows():
        new_cost = correction_map.get(make_key(row))
        if new_cost is not None and pd.notna(new_cost) and new_cost != row["매입단가"]:
            out.at[idx, "매입단가"] = new_cost
            changed += 1

    out.loc[owned_mask, "매입가"] = (
        out.loc[owned_mask, "매입단가"] * out.loc[owned_mask, "주문수량"]
        + out.loc[owned_mask, "공급사배송비"]
    )
    out.loc[owned_mask, "마진"] = out.loc[owned_mask, "정산가"] - out.loc[owned_mask, "매입가"]
    out.loc[owned_mask, "마진률"] = out.loc[owned_mask].apply(
        lambda r: (r["마진"] / r["정산가"]) if pd.notna(r["정산가"]) and r["정산가"] not in (0, None) else None,
        axis=1,
    )

    stats = {
        "owned_rows": int(owned_mask.sum()),
        "reviewed_owned_rows": int(len(owned_reviewed)),
        "ignored_rows": int(ignored_rows),
        "changed_rows": int(changed),
    }
    return out, stats


def summarize_by_product(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("고객선택옵션", dropna=False).agg(
        주문수량=("주문수량", "sum"),
        정산가=("정산가", "sum"),
        마진=("마진", "sum"),
    ).reset_index()
    return g.sort_values("주문수량", ascending=False, na_position="last")
