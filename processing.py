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

DEFAULT_RATE_RULES = [
    {"keyword": "카카오", "type": "percent", "rate": 0.90, "note": "결제금액의 90% (수수료 10%)"},
    {"keyword": "알리익스프레스", "type": "percent", "rate": 0.91, "note": ""},
    {"keyword": "자사몰", "type": "percent", "rate": 0.96, "note": ""},
    {"keyword": "네이버 스마트스토어", "type": "percent", "rate": 0.90, "note": ""},
    {"keyword": "티딜", "type": "percent", "rate": 0.85, "note": ""},
    {"keyword": "홈앤쇼핑", "type": "percent", "rate": 0.68, "note": "수수료 32%"},
    {"keyword": "토스", "type": "percent", "rate": 0.967,
     "note": "주문배송관리 파일 미제공 시 fallback 요율."},
    {"keyword": "지마켓", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "옥션", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "제이슨딜", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "GS SHOP", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "NS홈쇼핑", "type": "same_as_payment", "rate": 1.0, "note": ""},
    {"keyword": "11번가", "type": "same_as_payment", "rate": 1.0, "note": ""},
]

CASHDEAL_KEYWORD = "캐시딜"
CASHDEAL_DEFAULT_RATE = 0.90
DEFAULT_CASHDEAL_EXCEPTIONS = [
    {"상품명": "홍주부아카시아향사양꿀2.4kg", "rate": 0.80},
]

FIXED_PRICE_SITE_KEYWORDS = ["케이딜", "꿈꾸는이웃", "LG 복지몰", "제트언스"]

TOSS_KEYWORD = "토스"
# 토스 '주문배송관리' 파일의 '받은 혜택' 열(F열) 기준, 2026-09 정책 변경 이후 3가지 경우:
#   1) "수수료 6% (배송 인센티브)" -> 배송비 6% + 기본수수료 3.3% = 총 9.3%
#   2) "수수료 0% (상품 광고)"     -> 광고 참여로 배송비 0%, 기본수수료 3.3%만 부과
#   3) 공란(혜택 없음)             -> 기본수수료 11%
TOSS_SHIPPING_INCENTIVE_TEXT = "수수료 6%"   # 배송 인센티브 (6%+3.3%=9.3%)
TOSS_AD_INCENTIVE_TEXT = "수수료 0%"          # 상품 광고 (3.3%만)
TOSS_SHIPPING_INCENTIVE_RATE = 1 - 0.093      # 0.907
TOSS_AD_INCENTIVE_RATE = 1 - 0.033            # 0.967
TOSS_DEFAULT_RATE = 1 - 0.11                  # 0.89 (공란)

SETTLEMENT_SITES = [
    "꿈꾸는이웃(산지로드)", "더로드 네이버 스마트스토어", "더로드 자사몰",
    "산지로드 11번가 sjroad_cop", "산지로드 LG 복지몰", "산지로드 GS SHOP",
    "산지로드 NS홈쇼핑", "산지로드 알리익스프레스", "산지로드 옥션", "산지로드 제이슨",
    "산지로드 지마켓", "산지로드 카카오", "산지로드 캐시딜", "산지로드 케이딜",
    "산지로드 토스", "산지로드 티딜", "산지로드 홈앤쇼핑", "제트언스(산지로드)",
]


def _get_base_dir() -> Path:
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
        df = pd.read_csv(FIXED_PRICE_PATH, encoding="utf-8-sig")
    else:
        df = pd.DataFrame(columns=["판매사", "상품명", "정산단가(개당)"])
    # 홈앤쇼핑은 더 이상 '고정 정산단가' 방식이 아니라 32% 정률 방식으로 바뀌었으므로,
    # 예전에 등록해둔 홈앤쇼핑 행이 데이터 파일에 남아있어도(빈 행 포함) 항상 걸러내
    # 계산에 전혀 영향을 주지 않도록 한다.
    if "판매사" in df.columns and len(df) > 0:
        df = df[~df["판매사"].astype(str).str.contains("홈앤쇼핑", na=False)].reset_index(drop=True)
    return df


def save_fixed_price_table(df: pd.DataFrame) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(FIXED_PRICE_PATH, index=False, encoding="utf-8-sig")


def read_raw_settlement(file) -> pd.DataFrame:
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

    if "판매사주문번호" in out.columns:
        out["판매사주문번호"] = out["판매사주문번호"].apply(_normalize_order_no)

    return out


def _normalize_order_no(value) -> str:
    """주문번호를 비교 가능한 공통 형태로 정규화한다.

    엑셀에서 주문번호 컬럼이 '텍스트'로 저장된 파일과 '숫자'로 저장된 파일이 섞여 있으면,
    같은 주문번호라도 파이썬에서 읽었을 때 "1234567" 대 "1234567.0" 처럼 서로 다른
    문자열이 되어 매칭이 실패한다. 정수로 떨어지는 숫자는 소수점을 떼어내 정규화하고,
    앞뒤 공백도 제거해 이런 매칭 실패를 방지한다.
    """
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if not s:
        return ""
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return s
    except ValueError:
        return s


def read_toss_delivery_file(file) -> dict[str, float]:
    """토스 '주문배송관리' 엑셀(주문내역 시트)에서 주문번호별 실제 수수료율을 읽어온다.

    F열 '받은 혜택'에 "수수료 6%"가 포함되면 배송 인센티브(6%+기본 3.3%=9.3%, 요율 0.907),
    "수수료 0%"가 포함되면 상품 광고(기본 3.3%만, 요율 0.967), 공란이면 기본 수수료
    11%(요율 0.89)가 적용된다. 반환값은 {주문번호(str): 요율} 딕셔너리로,
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
        order_no = _normalize_order_no(order_no)
        if not order_no:
            continue
        # 헤더 바로 아래 '수정 가능/수정 불가' 안내 행 등 한글이 섞인 잡음 행은 건너뜀.
        # (예전엔 "숫자로만 이루어진 값"만 허용했는데, 토스 주문번호가 문자+숫자 조합일
        #  수도 있어 정상 주문번호까지 걸러지는 문제가 있었음 — 한글 포함 여부로만 거른다.)
        if any("\uac00" <= ch <= "\ud7a3" for ch in order_no):
            continue
        benefit = row[col_benefit]
        benefit = "" if pd.isna(benefit) else str(benefit)
        if TOSS_SHIPPING_INCENTIVE_TEXT in benefit:
            rate = TOSS_SHIPPING_INCENTIVE_RATE
        elif TOSS_AD_INCENTIVE_TEXT in benefit:
            rate = TOSS_AD_INCENTIVE_RATE
        else:
            rate = TOSS_DEFAULT_RATE
        rates[order_no] = rate
    return rates


@dataclass
class ProcessResult:
    df: pd.DataFrame
    needs_review: pd.DataFrame
    excluded: pd.DataFrame = field(default_factory=pd.DataFrame)
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

        if any(_match_keyword(site, kw) for kw in FIXED_PRICE_SITE_KEYWORDS):
            unit_price = fixed_price_lookup.get((site, product))
            if unit_price is not None:
                settlement_prices.append(unit_price * qty)
                rule_applied.append(f"고정단가({unit_price:,.1f}/개)")
            else:
                settlement_prices.append(None)
                rule_applied.append("고정단가 미등록")
            continue

        if _match_keyword(site, CASHDEAL_KEYWORD):
            rate = cashdeal_ex_lookup.get(product, CASHDEAL_DEFAULT_RATE)
            settlement_prices.append(payment * rate)
            rule_applied.append(f"캐시딜 {rate:.0%}")
            continue

        if toss_delivery_rates is not None and _match_keyword(site, TOSS_KEYWORD):
            order_no = _normalize_order_no(row.get("판매사주문번호"))
            rate = toss_delivery_rates.get(order_no)
            if rate is not None:
                settlement_prices.append(payment * rate)
                rule_applied.append(f"토스 배송관리 매칭 ({rate:.1%})")
            else:
                settlement_prices.append(None)
                rule_applied.append("토스 배송관리 미매칭 - 확인 필요")
            continue

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
        df=df, needs_review=needs_review, excluded=excluded,
        rate_rules=rate_rules, fixed_price_table=fixed_price_table,
    )


def summarize_by_site(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("판매사", dropna=False).agg(
        정산가=("정산가", "sum"), 마진=("마진", "sum"), 주문수량=("주문수량", "sum"),
    ).reset_index()
    g["마진률"] = g.apply(lambda r: (r["마진"] / r["정산가"]) if r["정산가"] else None, axis=1)
    return g.sort_values("정산가", ascending=False, na_position="last")


REVIEWER_SITE_KEYWORDS: dict[str, list[str]] = {
    "박정배": ["더로드 네이버 스마트스토어", "더로드 자사몰", "산지로드 LG 복지몰",
              "산지로드 GS SHOP", "산지로드 알리익스프레스", "산지로드 옥션",
              "산지로드 제이슨", "산지로드 지마켓", "산지로드 카카오", "산지로드 캐시딜"],
    "김슬기": ["산지로드 NS홈쇼핑", "산지로드 케이딜", "산지로드 티딜", "산지로드 홈앤쇼핑"],
    "이재성": ["산지로드 11번가 sjroad_cop", "꿈꾸는이웃(산지로드)", "산지로드 토스",
              "제트언스(산지로드)"],
}


def unassigned_sites(df: pd.DataFrame) -> list[str]:
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
        주문수량=("주문수량", "sum"), 정산가=("정산가", "sum"), 마진=("마진", "sum"),
    ).reset_index()
    return g.sort_values("주문수량", ascending=False, na_position="last")
