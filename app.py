# -*- coding: utf-8 -*-
"""
산지로드 결산 자동화 대시보드
------------------------------------------------
실행: streamlit run app.py

기능
1. 취합 결산서(xlsx) 업로드 -> 사이트별 규칙 적용 -> 최종 결산서 다운로드
2. 파이차트 대시보드 (사이트별 정산가/마진 비중, 상품 Top10)
3. 누적 데이터 검색 (사이트/상품/주문번호/날짜)
"""
import io
import sqlite3
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

import processing as P

DB_PATH = Path(__file__).parent / "data" / "settlement.db"

st.set_page_config(page_title="산지로드 결산 대시보드", layout="wide")


# ---------------------------------------------------------------------------
# 누적 DB
# ---------------------------------------------------------------------------

def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS records (
            결산일자 TEXT, 판매사 TEXT, 고객선택옵션 TEXT, 주문수량 REAL,
            공급사배송비 REAL, 결제금액 REAL, 매입단가 REAL, 매입가 REAL,
            정산가 REAL, 마진 REAL, 마진률 REAL, 적용규칙 TEXT,
            비고 TEXT, 판매사주문번호 TEXT, 수령인연락처 TEXT
        )
        """
    )
    return conn


def save_to_db(df: pd.DataFrame, settlement_date: date):
    conn = get_conn()
    date_str = settlement_date.isoformat()
    conn.execute("DELETE FROM records WHERE 결산일자 = ?", (date_str,))
    out = df.copy()
    out.insert(0, "결산일자", date_str)
    out.to_sql("records", conn, if_exists="append", index=False)
    conn.commit()
    conn.close()


def load_all_from_db() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql("SELECT * FROM records", conn)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df


# ---------------------------------------------------------------------------
# 최종 결산서(xlsx) 생성
# ---------------------------------------------------------------------------

def build_output_excel(df: pd.DataFrame) -> bytes:
    site_summary = P.summarize_by_site(df)
    product_summary = P.summarize_by_product(df)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="결산서", index=False)
        site_summary.to_excel(writer, sheet_name="최종정산가 및 마진", index=False)
        product_summary.to_excel(writer, sheet_name="품목별 판매수량", index=False)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.title("📊 산지로드 결산 자동화 대시보드")

tab_upload, tab_dashboard, tab_search, tab_rules = st.tabs(
    ["📥 결산서 업로드/처리", "🥧 대시보드", "🔍 검색", "⚙️ 규칙 관리"]
)

# ---------------------------------------------------------------------------
# 탭 1: 업로드/처리
# ---------------------------------------------------------------------------
with tab_upload:
    st.subheader("1. 취합 결산서 업로드")
    settlement_date = st.date_input("결산 일자", value=date.today())
    uploaded = st.file_uploader("취합 결산서(xlsx) 파일을 올려주세요", type=["xlsx"])

    if uploaded is not None:
        try:
            raw = P.read_raw_settlement(uploaded)
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            raw = None

        if raw is not None:
            st.success(f"{len(raw):,}건의 주문 데이터를 읽었습니다.")

            result = P.process_settlement(raw)
            n_total = len(result.df)
            n_ok = n_total - len(result.needs_review)

            c1, c2, c3 = st.columns(3)
            c1.metric("전체 행", f"{n_total:,}")
            c2.metric("정산가 계산 완료", f"{n_ok:,}")
            c3.metric("확인 필요", f"{len(result.needs_review):,}")

            if len(result.needs_review) > 0:
                st.warning(
                    "아래 사이트/상품은 정산 규칙이 없어 정산가를 계산하지 못했습니다. "
                    "'규칙 관리' 탭에서 요율 또는 고정 정산단가를 등록한 뒤 다시 처리해 주세요."
                )
                st.dataframe(
                    result.needs_review[["판매사", "고객선택옵션", "결제금액", "주문수량"]]
                    .groupby(["판매사", "고객선택옵션"], as_index=False)
                    .agg(건수=("결제금액", "count"), 결제금액합계=("결제금액", "sum")),
                    use_container_width=True,
                )

            st.subheader("2. 계산 결과")
            st.dataframe(result.df, use_container_width=True, height=350)

            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("💾 누적 DB에 저장 (이 날짜로 반영)", type="primary"):
                    save_to_db(result.df, settlement_date)
                    st.success(f"{settlement_date} 데이터를 누적 DB에 저장했습니다.")
                    st.session_state["_reload_db"] = True

            with col_b:
                excel_bytes = build_output_excel(result.df)
                st.download_button(
                    "⬇️ 최종 결산서 다운로드 (.xlsx)",
                    data=excel_bytes,
                    file_name=f"{settlement_date.strftime('%m%d')}_산지로드_결산_최종.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            st.session_state["last_result_df"] = result.df

# ---------------------------------------------------------------------------
# 탭 2: 대시보드
# ---------------------------------------------------------------------------
with tab_dashboard:
    st.subheader("누적 대시보드")
    db_df = load_all_from_db()

    if db_df.empty:
        st.info("아직 저장된 누적 데이터가 없습니다. '결산서 업로드/처리' 탭에서 먼저 저장해 주세요.")
    else:
        dates = sorted(db_df["결산일자"].unique())
        sel_dates = st.multiselect("기간(결산일자) 선택 - 비워두면 전체", dates, default=dates)
        view_df = db_df[db_df["결산일자"].isin(sel_dates)] if sel_dates else db_df

        total_i = view_df["정산가"].sum()
        total_j = view_df["마진"].sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("합계 정산가", f"{total_i:,.0f}")
        c2.metric("합계 마진", f"{total_j:,.0f}")
        c3.metric("마진률", f"{(total_j/total_i):.1%}" if total_i else "-")

        site_summary = P.summarize_by_site(view_df)
        product_summary = P.summarize_by_product(view_df)

        col1, col2 = st.columns(2)
        with col1:
            fig1 = px.pie(
                site_summary.dropna(subset=["정산가"]),
                names="판매사", values="정산가",
                title="사이트별 정산가 비중",
            )
            st.plotly_chart(fig1, use_container_width=True)

        with col2:
            fig2 = px.pie(
                site_summary.dropna(subset=["마진"]),
                names="판매사", values="마진",
                title="사이트별 마진 비중",
            )
            st.plotly_chart(fig2, use_container_width=True)

        top10 = product_summary.head(10).copy()
        etc_qty = product_summary["주문수량"].iloc[10:].sum()
        if etc_qty > 0:
            top10 = pd.concat(
                [top10, pd.DataFrame([{"고객선택옵션": "기타", "주문수량": etc_qty}])],
                ignore_index=True,
            )
        fig3 = px.pie(top10, names="고객선택옵션", values="주문수량",
                      title="상품별 주문수량 비중 (Top10 + 기타)")
        st.plotly_chart(fig3, use_container_width=True)

        st.subheader("월별 추이")
        monthly = view_df.copy()
        monthly["월"] = monthly["결산일자"].str.slice(0, 7)
        monthly_summary = monthly.groupby("월").agg(
            결제금액=("결제금액", "sum"), 정산가=("정산가", "sum"), 마진=("마진", "sum")
        ).reset_index()
        fig4 = px.bar(monthly_summary, x="월", y=["결제금액", "정산가", "마진"], barmode="group",
                      title="월별 결제금액/정산가/마진")
        st.plotly_chart(fig4, use_container_width=True)

# ---------------------------------------------------------------------------
# 탭 3: 검색
# ---------------------------------------------------------------------------
with tab_search:
    st.subheader("누적 데이터 검색")
    db_df = load_all_from_db()

    if db_df.empty:
        st.info("아직 저장된 누적 데이터가 없습니다.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        site_q = c1.text_input("판매사(부분검색)")
        product_q = c2.text_input("상품명(부분검색)")
        order_q = c3.text_input("주문번호(부분검색)")
        date_q = c4.text_input("결산일자 (YYYY-MM-DD)")

        filtered = db_df.copy()
        if site_q:
            filtered = filtered[filtered["판매사"].str.contains(site_q, case=False, na=False)]
        if product_q:
            filtered = filtered[filtered["고객선택옵션"].str.contains(product_q, case=False, na=False)]
        if order_q:
            filtered = filtered[filtered["판매사주문번호"].astype(str).str.contains(order_q, case=False, na=False)]
        if date_q:
            filtered = filtered[filtered["결산일자"].astype(str).str.contains(date_q, na=False)]

        st.write(f"검색 결과: {len(filtered):,}건")
        st.dataframe(filtered, use_container_width=True, height=450)

        if len(filtered) > 0:
            csv_bytes = filtered.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ 검색 결과 CSV 다운로드", data=csv_bytes, file_name="검색결과.csv")

# ---------------------------------------------------------------------------
# 탭 4: 규칙 관리
# ---------------------------------------------------------------------------
with tab_rules:
    st.subheader("사이트별 정산 규칙")
    st.caption(
        "판매사명에 '키워드'가 포함되면 해당 규칙이 적용됩니다. "
        "type: percent(결제금액×rate) / same_as_payment(결제금액과 동일, rate=1.0)"
    )
    rules = P.load_rate_rules()
    rules_df = pd.DataFrame(rules)
    edited_rules = st.data_editor(rules_df, num_rows="dynamic", use_container_width=True, key="rules_editor")
    if st.button("규칙 저장", key="save_rules"):
        P.save_rate_rules(edited_rules.to_dict("records"))
        st.success("사이트 규칙을 저장했습니다.")

    st.divider()
    st.subheader("캐시딜 예외 상품")
    ce = P.load_cashdeal_exceptions()
    ce_df = pd.DataFrame(ce)
    edited_ce = st.data_editor(ce_df, num_rows="dynamic", use_container_width=True, key="ce_editor")
    if st.button("캐시딜 예외 저장", key="save_ce"):
        P.save_cashdeal_exceptions(edited_ce.to_dict("records"))
        st.success("캐시딜 예외 상품을 저장했습니다.")

    st.divider()
    st.subheader("고정 정산단가 상품 (케이딜 / 꿈꾸는이웃 / 홈앤쇼핑 / LG복지몰 / 제트언스)")
    st.caption(
        "이 사이트들은 결제금액이 아니라 상품별로 정해진 '개당 정산단가'를 사용합니다. "
        "정산가 = 정산단가 × 주문수량. 새 상품이 나오면 이 표에 추가해 주세요."
    )
    fp = P.load_fixed_price_table()
    edited_fp = st.data_editor(fp, num_rows="dynamic", use_container_width=True, key="fp_editor")
    if st.button("고정 정산단가 저장", key="save_fp"):
        P.save_fixed_price_table(edited_fp)
        st.success("고정 정산단가 표를 저장했습니다.")
