# -*- coding: utf-8 -*-
"""
산지로드 결산 대시보드 - 데스크톱(프로그램) 버전
브라우저 없이 창(윈도우) 형태로 실행됩니다.

실행: python app_desktop.py
"""
import sqlite3
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd

import processing as P

DB_PATH = Path(__file__).parent / "data" / "settlement.db"


# ---------------------------------------------------------------------------
# DB 헬퍼 (app.py의 스트림릿 버전과 동일한 스키마/로직)
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


def save_to_db(df: pd.DataFrame, settlement_date: str):
    conn = get_conn()
    conn.execute("DELETE FROM records WHERE 결산일자 = ?", (settlement_date,))
    out = df.copy()
    out.insert(0, "결산일자", settlement_date)
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
# 공용: DataFrame을 ttk.Treeview에 채우는 헬퍼
# ---------------------------------------------------------------------------

def fill_treeview(tree: ttk.Treeview, df: pd.DataFrame, max_rows: int = 2000):
    tree.delete(*tree.get_children())
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=110, anchor="center")
    for _, row in df.head(max_rows).iterrows():
        values = [
            f"{v:,.1f}" if isinstance(v, float) else v
            for v in row.tolist()
        ]
        tree.insert("", "end", values=values)


# ---------------------------------------------------------------------------
# 메인 애플리케이션
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("산지로드 결산 대시보드")
        self.geometry("1200x800")

        self.current_df: pd.DataFrame | None = None
        self.toss_rates: dict | None = None
        self.working_df: pd.DataFrame | None = None  # 담당자 검토 반영본 (4/5단계)

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.tab_upload = ttk.Frame(notebook)
        self.tab_dashboard = ttk.Frame(notebook)
        self.tab_search = ttk.Frame(notebook)
        self.tab_rules = ttk.Frame(notebook)

        notebook.add(self.tab_upload, text="📥 업로드/처리")
        notebook.add(self.tab_dashboard, text="🥧 대시보드")
        notebook.add(self.tab_search, text="🔍 검색")
        notebook.add(self.tab_rules, text="⚙️ 규칙 관리")

        self._build_upload_tab()
        self._build_dashboard_tab()
        self._build_search_tab()
        self._build_rules_tab()

    # ------------------------------------------------------------------
    # 탭 1: 업로드/처리
    # ------------------------------------------------------------------
    def _build_upload_tab(self):
        top = ttk.Frame(self.tab_upload)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Label(top, text="결산 일자 (YYYY-MM-DD):").pack(side="left")
        self.date_entry = ttk.Entry(top, width=14)
        self.date_entry.insert(0, date.today().isoformat())
        self.date_entry.pack(side="left", padx=(4, 20))

        ttk.Button(top, text="1) 취합 결산서 열기", command=self.on_open_file).pack(side="left")
        ttk.Button(top, text="2) 누적 DB에 저장", command=self.on_save_db).pack(side="left", padx=8)
        ttk.Button(top, text="3) 최종 결산서 저장(xlsx)", command=self.on_export_excel).pack(side="left")

        toss_row = ttk.Frame(self.tab_upload)
        toss_row.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(
            toss_row, text="🚚 토스 주문배송관리 파일 열기 (선택)", command=self.on_open_toss_file
        ).pack(side="left")
        self.toss_label = ttk.Label(
            toss_row,
            text="미등록 시 토스는 기본 96.7%로 계산됩니다. 이 파일은 결산서 열기 '전에' 먼저 열어주세요.",
            foreground="gray",
        )
        self.toss_label.pack(side="left", padx=8)

        self.status_label = ttk.Label(self.tab_upload, text="파일을 열어주세요.", foreground="blue")
        self.status_label.pack(anchor="w", padx=10)

        self.review_label = ttk.Label(self.tab_upload, text="", foreground="red")
        self.review_label.pack(anchor="w", padx=10)

        ttk.Separator(self.tab_upload, orient="horizontal").pack(fill="x", padx=10, pady=8)

        reviewer_row = ttk.Frame(self.tab_upload)
        reviewer_row.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Label(reviewer_row, text="4) 담당자별 검토 파일 열기:").pack(side="left")
        for name in P.REVIEWER_SITE_KEYWORDS:
            ttk.Button(
                reviewer_row, text=name,
                command=lambda n=name: self.on_open_reviewer_file(n),
            ).pack(side="left", padx=4)
        ttk.Button(
            reviewer_row, text="5) 검토 반영 최종본 저장(xlsx)", command=self.on_export_reviewed_excel
        ).pack(side="left", padx=(16, 0))

        self.reviewer_label = ttk.Label(
            self.tab_upload,
            text="담당자가 매입단가만 확인/수정한 파일을 열면, 본인 담당 사이트 행만 반영되고 "
                 "다른 담당자 사이트는 무시됩니다.",
            foreground="gray",
        )
        self.reviewer_label.pack(anchor="w", padx=10, pady=(0, 8))

        tree_frame = ttk.Frame(self.tab_upload)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.upload_tree = ttk.Treeview(tree_frame)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.upload_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.upload_tree.xview)
        self.upload_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.upload_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

    def on_open_toss_file(self):
        path = filedialog.askopenfilename(
            title="토스 주문배송관리 파일 선택", filetypes=[("Excel 파일", "*.xlsx")]
        )
        if not path:
            return
        try:
            self.toss_rates = P.read_toss_delivery_file(path)
        except Exception as e:
            messagebox.showerror("오류", f"토스 배송관리 파일을 읽는 중 오류가 발생했습니다:\n{e}")
            return
        self.toss_label.config(
            text=f"토스 배송관리 파일 등록됨: 주문 {len(self.toss_rates):,}건의 수수료 정보 "
                 "(이제 '1) 취합 결산서 열기'를 누르면 이 정보로 토스가 계산됩니다)",
            foreground="green",
        )

    def on_open_file(self):
        path = filedialog.askopenfilename(
            title="취합 결산서 선택", filetypes=[("Excel 파일", "*.xlsx")]
        )
        if not path:
            return
        try:
            raw = P.read_raw_settlement(path)
            result = P.process_settlement(raw, toss_delivery_rates=self.toss_rates)
        except Exception as e:
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다:\n{e}")
            return

        self.current_df = result.df
        self.working_df = None  # 새 결산서를 열면 이전 검토 반영 상태는 초기화
        self.reviewer_label.config(
            text="담당자가 매입단가만 확인/수정한 파일을 열면, 본인 담당 사이트 행만 반영되고 "
                 "다른 담당자 사이트는 무시됩니다.",
            foreground="gray",
        )
        n_total = len(result.df)
        n_review = len(result.needs_review)
        n_excluded = len(result.excluded)
        self.status_label.config(
            text=f"총 {n_total:,}건 처리 완료 ({n_total - n_review:,}건 정산가 계산됨, "
                 f"결산 대상 18개 사이트 밖이라 제외한 행 {n_excluded:,}건)"
        )
        review_msgs = []
        if n_review > 0:
            review_msgs.append(
                f"⚠ 규칙이 없어 정산가를 계산 못한 행 {n_review:,}건 "
                "(규칙 관리 탭에서 요율/고정단가를 추가한 뒤 다시 열어주세요)"
            )
        if self.toss_rates is not None:
            toss_mask = result.df["판매사"].astype(str).str.contains("토스", na=False)
            toss_total = int(toss_mask.sum())
            toss_unmatched = int((toss_mask & result.df["정산가"].isna()).sum())
            if toss_total:
                review_msgs.append(
                    f"토스 주문 {toss_total:,}건 중 {toss_total - toss_unmatched:,}건이 "
                    "배송관리 파일과 매칭되었습니다 (미매칭은 위 확인 필요 건수에 포함됨)"
                )
        self.review_label.config(text="  |  ".join(review_msgs))

        fill_treeview(self.upload_tree, result.df.head(500))

    def on_save_db(self):
        if self.current_df is None:
            messagebox.showwarning("알림", "먼저 취합 결산서를 열어주세요.")
            return
        d = self.date_entry.get().strip()
        try:
            date.fromisoformat(d)
        except ValueError:
            messagebox.showerror("오류", "결산 일자를 YYYY-MM-DD 형식으로 입력해주세요.")
            return
        save_to_db(self.current_df, d)
        messagebox.showinfo("완료", f"{d} 데이터를 누적 DB에 저장했습니다.")

    def on_export_excel(self):
        if self.current_df is None:
            messagebox.showwarning("알림", "먼저 취합 결산서를 열어주세요.")
            return
        path = filedialog.asksaveasfilename(
            title="최종 결산서 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile="산지로드_결산_최종.xlsx",
        )
        if not path:
            return
        site_summary = P.summarize_by_site(self.current_df)
        product_summary = P.summarize_by_product(self.current_df)
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            self.current_df.to_excel(writer, sheet_name="결산서", index=False)
            site_summary.to_excel(writer, sheet_name="최종정산가 및 마진", index=False)
            product_summary.to_excel(writer, sheet_name="품목별 판매수량", index=False)
        messagebox.showinfo("완료", f"저장되었습니다:\n{path}")

    def on_open_reviewer_file(self, name: str):
        if self.current_df is None:
            messagebox.showwarning("알림", "먼저 취합 결산서를 열어주세요.")
            return
        path = filedialog.askopenfilename(
            title=f"{name} 검토 파일 선택", filetypes=[("Excel 파일", "*.xlsx")]
        )
        if not path:
            return
        try:
            reviewed_raw = P.read_raw_settlement(path)
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽는 중 오류가 발생했습니다:\n{e}")
            return

        base_df = self.working_df if self.working_df is not None else self.current_df
        keywords = P.REVIEWER_SITE_KEYWORDS[name]
        self.working_df, stats = P.apply_reviewer_corrections(base_df, reviewed_raw, keywords)
        self.reviewer_label.config(
            text=f"{name}: 담당 {stats['owned_rows']:,}행 중 매입단가 {stats['changed_rows']:,}건 반영"
                 f" (다른 담당자 사이트 행 {stats['ignored_rows']:,}건은 무시함)",
            foreground="green",
        )

    def on_export_reviewed_excel(self):
        df = self.working_df if self.working_df is not None else self.current_df
        if df is None:
            messagebox.showwarning("알림", "먼저 취합 결산서를 열어주세요.")
            return
        path = filedialog.asksaveasfilename(
            title="검토 반영 최종 결산서 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile="산지로드_결산_최종_검토반영.xlsx",
        )
        if not path:
            return
        site_summary = P.summarize_by_site(df)
        product_summary = P.summarize_by_product(df)
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="결산서", index=False)
            site_summary.to_excel(writer, sheet_name="최종정산가 및 마진", index=False)
            product_summary.to_excel(writer, sheet_name="품목별 판매수량", index=False)
        messagebox.showinfo("완료", f"저장되었습니다:\n{path}")

    # ------------------------------------------------------------------
    # 탭 2: 대시보드 (파이차트)
    # ------------------------------------------------------------------
    def _build_dashboard_tab(self):
        top = ttk.Frame(self.tab_dashboard)
        top.pack(fill="x", padx=10, pady=10)
        ttk.Button(top, text="🔄 대시보드 새로고침 (누적 DB 기준)", command=self.refresh_dashboard).pack(side="left")
        self.dash_summary_label = ttk.Label(top, text="")
        self.dash_summary_label.pack(side="left", padx=20)

        self.chart_frame = ttk.Frame(self.tab_dashboard)
        self.chart_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def refresh_dashboard(self):
        for widget in self.chart_frame.winfo_children():
            widget.destroy()

        db_df = load_all_from_db()
        if db_df.empty:
            ttk.Label(self.chart_frame, text="누적된 데이터가 없습니다. 먼저 업로드 후 'DB에 저장'을 눌러주세요.").pack()
            return

        total_i = db_df["정산가"].sum()
        total_j = db_df["마진"].sum()
        margin_rate = (total_j / total_i) if total_i else 0
        self.dash_summary_label.config(
            text=f"합계 정산가 {total_i:,.0f}  |  합계 마진 {total_j:,.0f}  |  마진률 {margin_rate:.1%}"
        )

        site_summary = P.summarize_by_site(db_df).dropna(subset=["정산가"])
        product_summary = P.summarize_by_product(db_df)

        fig = plt.Figure(figsize=(13, 5.5))
        ax1 = fig.add_subplot(1, 3, 1)
        ax2 = fig.add_subplot(1, 3, 2)
        ax3 = fig.add_subplot(1, 3, 3)

        if len(site_summary) > 0:
            ax1.pie(site_summary["정산가"], labels=site_summary["판매사"], autopct="%1.0f%%", textprops={"fontsize": 7})
        ax1.set_title("사이트별 정산가 비중")

        site_j = site_summary.dropna(subset=["마진"])
        if len(site_j) > 0:
            ax2.pie(site_j["마진"], labels=site_j["판매사"], autopct="%1.0f%%", textprops={"fontsize": 7})
        ax2.set_title("사이트별 마진 비중")

        top10 = product_summary.head(10).copy()
        etc = product_summary["주문수량"].iloc[10:].sum()
        labels = top10["고객선택옵션"].tolist()
        values = top10["주문수량"].tolist()
        if etc > 0:
            labels.append("기타")
            values.append(etc)
        ax3.pie(values, labels=labels, autopct="%1.0f%%", textprops={"fontsize": 6})
        ax3.set_title("상품별 주문수량 비중 (Top10+기타)")

        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # ------------------------------------------------------------------
    # 탭 3: 검색
    # ------------------------------------------------------------------
    def _build_search_tab(self):
        top = ttk.Frame(self.tab_search)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Label(top, text="판매사:").grid(row=0, column=0, padx=4)
        self.search_site = ttk.Entry(top, width=16)
        self.search_site.grid(row=0, column=1, padx=4)

        ttk.Label(top, text="상품명:").grid(row=0, column=2, padx=4)
        self.search_product = ttk.Entry(top, width=16)
        self.search_product.grid(row=0, column=3, padx=4)

        ttk.Label(top, text="주문번호:").grid(row=0, column=4, padx=4)
        self.search_order = ttk.Entry(top, width=16)
        self.search_order.grid(row=0, column=5, padx=4)

        ttk.Label(top, text="결산일자:").grid(row=0, column=6, padx=4)
        self.search_date = ttk.Entry(top, width=12)
        self.search_date.grid(row=0, column=7, padx=4)

        ttk.Button(top, text="검색", command=self.on_search).grid(row=0, column=8, padx=10)

        tree_frame = ttk.Frame(self.tab_search)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.search_tree = ttk.Treeview(tree_frame)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.search_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.search_tree.xview)
        self.search_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.search_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

    def on_search(self):
        db_df = load_all_from_db()
        if db_df.empty:
            messagebox.showinfo("알림", "누적된 데이터가 없습니다.")
            return
        filtered = db_df.copy()
        if self.search_site.get():
            filtered = filtered[filtered["판매사"].str.contains(self.search_site.get(), case=False, na=False)]
        if self.search_product.get():
            filtered = filtered[filtered["고객선택옵션"].str.contains(self.search_product.get(), case=False, na=False)]
        if self.search_order.get():
            filtered = filtered[filtered["판매사주문번호"].astype(str).str.contains(self.search_order.get(), na=False)]
        if self.search_date.get():
            filtered = filtered[filtered["결산일자"].astype(str).str.contains(self.search_date.get(), na=False)]

        fill_treeview(self.search_tree, filtered)

    # ------------------------------------------------------------------
    # 탭 4: 규칙 관리
    # ------------------------------------------------------------------
    def _build_rules_tab(self):
        ttk.Label(
            self.tab_rules,
            text="사이트별 정산 규칙 (판매사명에 키워드가 포함되면 적용됩니다)",
        ).pack(anchor="w", padx=10, pady=(10, 0))

        tree_frame = ttk.Frame(self.tab_rules)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.rules_tree = ttk.Treeview(tree_frame, columns=["keyword", "type", "rate", "note"], show="headings")
        for col, w in [("keyword", 150), ("type", 150), ("rate", 100), ("note", 300)]:
            self.rules_tree.heading(col, text=col)
            self.rules_tree.column(col, width=w)
        self.rules_tree.pack(fill="both", expand=True)
        self.rules_tree.bind("<Double-1>", self.on_edit_rule)
        self._reload_rules_tree()

        btns = ttk.Frame(self.tab_rules)
        btns.pack(fill="x", padx=10, pady=6)
        ttk.Button(btns, text="+ 규칙 추가", command=self.on_add_rule).pack(side="left")
        ttk.Button(btns, text="선택 규칙 삭제", command=self.on_delete_rule).pack(side="left", padx=8)
        ttk.Button(btns, text="더블클릭하면 수정할 수 있어요", state="disabled").pack(side="left", padx=8)

        ttk.Label(
            self.tab_rules,
            text="고정 정산단가 상품 (케이딜/꿈꾸는이웃/홈앤쇼핑/LG복지몰/제트언스: 정산가 = 정산단가 × 주문수량)",
        ).pack(anchor="w", padx=10, pady=(16, 0))

        fp_frame = ttk.Frame(self.tab_rules)
        fp_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.fp_tree = ttk.Treeview(fp_frame, columns=["판매사", "상품명", "정산단가(개당)"], show="headings")
        for col, w in [("판매사", 180), ("상품명", 300), ("정산단가(개당)", 130)]:
            self.fp_tree.heading(col, text=col)
            self.fp_tree.column(col, width=w)
        self.fp_tree.pack(fill="both", expand=True)
        self.fp_tree.bind("<Double-1>", self.on_edit_fixed_price)
        self._reload_fp_tree()

        fp_btns = ttk.Frame(self.tab_rules)
        fp_btns.pack(fill="x", padx=10, pady=6)
        ttk.Button(fp_btns, text="+ 상품 추가", command=self.on_add_fixed_price).pack(side="left")
        ttk.Button(fp_btns, text="선택 상품 삭제", command=self.on_delete_fixed_price).pack(side="left", padx=8)

    def _reload_rules_tree(self):
        self.rules = P.load_rate_rules()
        self.rules_tree.delete(*self.rules_tree.get_children())
        for r in self.rules:
            self.rules_tree.insert("", "end", values=[r["keyword"], r["type"], r["rate"], r.get("note", "")])

    def on_add_rule(self):
        self._rule_dialog()

    def on_edit_rule(self, event):
        sel = self.rules_tree.selection()
        if not sel:
            return
        idx = self.rules_tree.index(sel[0])
        self._rule_dialog(self.rules[idx], idx)

    def on_delete_rule(self):
        sel = self.rules_tree.selection()
        if not sel:
            return
        idx = self.rules_tree.index(sel[0])
        del self.rules[idx]
        P.save_rate_rules(self.rules)
        self._reload_rules_tree()

    def _rule_dialog(self, existing=None, idx=None):
        win = tk.Toplevel(self)
        win.title("규칙 추가/수정")
        fields = {}
        labels = ["keyword (판매사명에 포함될 키워드)", "type (percent 또는 same_as_payment)", "rate (예: 0.9)", "note (설명, 선택)"]
        keys = ["keyword", "type", "rate", "note"]
        for i, (lab, key) in enumerate(zip(labels, keys)):
            ttk.Label(win, text=lab).grid(row=i, column=0, sticky="w", padx=6, pady=4)
            e = ttk.Entry(win, width=30)
            e.grid(row=i, column=1, padx=6, pady=4)
            if existing:
                e.insert(0, str(existing.get(key, "")))
            fields[key] = e

        def on_ok():
            try:
                rate = float(fields["rate"].get())
            except ValueError:
                messagebox.showerror("오류", "rate는 숫자로 입력해주세요 (예: 0.9)")
                return
            new_rule = {
                "keyword": fields["keyword"].get(),
                "type": fields["type"].get() or "percent",
                "rate": rate,
                "note": fields["note"].get(),
            }
            if idx is None:
                self.rules.append(new_rule)
            else:
                self.rules[idx] = new_rule
            P.save_rate_rules(self.rules)
            self._reload_rules_tree()
            win.destroy()

        ttk.Button(win, text="저장", command=on_ok).grid(row=4, column=0, columnspan=2, pady=8)

    def _reload_fp_tree(self):
        self.fp_df = P.load_fixed_price_table()
        self.fp_tree.delete(*self.fp_tree.get_children())
        for _, row in self.fp_df.iterrows():
            self.fp_tree.insert("", "end", values=[row["판매사"], row["상품명"], row["정산단가(개당)"]])

    def on_add_fixed_price(self):
        self._fp_dialog()

    def on_edit_fixed_price(self, event):
        sel = self.fp_tree.selection()
        if not sel:
            return
        idx = self.fp_tree.index(sel[0])
        row = self.fp_df.iloc[idx]
        self._fp_dialog(row, idx)

    def on_delete_fixed_price(self):
        sel = self.fp_tree.selection()
        if not sel:
            return
        idx = self.fp_tree.index(sel[0])
        self.fp_df = self.fp_df.drop(self.fp_df.index[idx]).reset_index(drop=True)
        P.save_fixed_price_table(self.fp_df)
        self._reload_fp_tree()

    def _fp_dialog(self, existing=None, idx=None):
        win = tk.Toplevel(self)
        win.title("고정 정산단가 추가/수정")
        labels = ["판매사", "상품명", "정산단가(개당)"]
        entries = {}
        for i, lab in enumerate(labels):
            ttk.Label(win, text=lab).grid(row=i, column=0, sticky="w", padx=6, pady=4)
            e = ttk.Entry(win, width=35)
            e.grid(row=i, column=1, padx=6, pady=4)
            if existing is not None:
                e.insert(0, str(existing[lab]))
            entries[lab] = e

        def on_ok():
            try:
                price = float(entries["정산단가(개당)"].get())
            except ValueError:
                messagebox.showerror("오류", "정산단가는 숫자로 입력해주세요.")
                return
            new_row = {"판매사": entries["판매사"].get(), "상품명": entries["상품명"].get(), "정산단가(개당)": price}
            if idx is None:
                self.fp_df = pd.concat([self.fp_df, pd.DataFrame([new_row])], ignore_index=True)
            else:
                self.fp_df.loc[idx] = new_row
            P.save_fixed_price_table(self.fp_df)
            self._reload_fp_tree()
            win.destroy()

        ttk.Button(win, text="저장", command=on_ok).grid(row=3, column=0, columnspan=2, pady=8)


if __name__ == "__main__":
    app = App()
    app.mainloop()
