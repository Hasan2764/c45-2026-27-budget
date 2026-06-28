# ============================================================
# USMAN PUBLIC SCHOOL CAMPUS 45
# BUDGET DASHBOARD 2026-27
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(
    page_title="Budget Dashboard 2026-27",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Usman Public School Campus 45")
st.subheader("Budget Dashboard FY 2026-27")

# ------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------
st.sidebar.header("📂 Upload Files")

income_file = st.sidebar.file_uploader(
    "Upload Income File",
    type=["xlsx", "xls", "csv"]
)

expense_file = st.sidebar.file_uploader(
    "Upload Expense File",
    type=["xlsx", "xls", "csv"]
)

# ------------------------------------------------------------
# FUNCTIONS
# ------------------------------------------------------------
def read_excel_file(file, actual_sheet, budget_sheet):

    actual = pd.read_excel(file, sheet_name=actual_sheet)
    budget = pd.read_excel(file, sheet_name=budget_sheet)

    actual.columns = actual.columns.str.strip()
    budget.columns = budget.columns.str.strip()

    actual["Data Type"] = "Actual FY 2025-26"
    budget["Data Type"] = "Budget FY 2026-27"

    return actual, budget


def clean_data(df, category):

    df.columns = df.columns.str.strip()

    if "Amount" not in df.columns:
        numeric_cols = df.select_dtypes(include=np.number).columns

        if len(numeric_cols) > 0:
            df = df.rename(columns={numeric_cols[-1]: "Amount"})

    if "Month" not in df.columns:
        df["Month"] = "Unknown"

    if "Particular" not in df.columns:
        if len(df.columns) > 0:
            df["Particular"] = df.iloc[:, 0]

    df["Category"] = category
    df["Amount"] = pd.to_numeric(
        df["Amount"],
        errors="coerce"
    ).fillna(0)

    return df


def convert_df_to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
            output,
            engine="openpyxl"
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Profit Loss"
        )

    output.seek(0)
    return output


# ------------------------------------------------------------
# PROCESS FILES
# ------------------------------------------------------------
if income_file and expense_file:

    try:
        income_actual, income_budget = read_excel_file(
            income_file,
            "Actual FY 2025-26",
            "Budget FY 2026-27"
        )

        expense_actual, expense_budget = read_excel_file(
            expense_file,
            "Actual FY 2025-26",
            "Budget FY 2026-27"
        )

        income_actual = clean_data(
            income_actual,
            "Income"
        )

        income_budget = clean_data(
            income_budget,
            "Income"
        )

        expense_actual = clean_data(
            expense_actual,
            "Expense"
        )

        expense_budget = clean_data(
            expense_budget,
            "Expense"
        )

        df = pd.concat([
            income_actual,
            income_budget,
            expense_actual,
            expense_budget
        ], ignore_index=True)

        # ----------------------------------------------------
        # FILTERS
        # ----------------------------------------------------
        st.sidebar.header("🎯 Filters")

        months = sorted(
            df["Month"].dropna().astype(str).unique()
        )

        month_filter = st.sidebar.multiselect(
            "Select Month",
            months,
            default=months
        )

        category_filter = st.sidebar.multiselect(
            "Select Category",
            ["Income", "Expense"],
            default=["Income", "Expense"]
        )

        filtered_df = df[
            (df["Month"].astype(str).isin(month_filter))
            &
            (df["Category"].isin(category_filter))
        ]

        # ----------------------------------------------------
        # BUDGET DATA ONLY FOR KPI
        # ----------------------------------------------------
        budget_df = filtered_df[
            filtered_df["Data Type"]
            == "Budget FY 2026-27"
        ]

        total_income = budget_df[
            budget_df["Category"] == "Income"
        ]["Amount"].sum()

        total_expense = budget_df[
            budget_df["Category"] == "Expense"
        ]["Amount"].sum()

        profit = total_income - total_expense

        if total_income != 0:
            profit_pct = (
                    profit
                    / total_income
                    * 100
            )
        else:
            profit_pct = 0

        # ----------------------------------------------------
        # KPI CARDS
        # ----------------------------------------------------
        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Total Income 2026-27",
            f"{total_income:,.0f}"
        )

        c2.metric(
            "Total Expense 2026-27",
            f"{total_expense:,.0f}"
        )

        c3.metric(
            "Profit / Loss",
            f"{profit:,.0f}"
        )

        c4.metric(
            "% Profit / Loss",
            f"{profit_pct:.2f}%"
        )

        st.divider()

        # ----------------------------------------------------
        # CHARTS
        # ----------------------------------------------------
        st.subheader("📈 Financial Overview")

        chart1, chart2 = st.columns(2)

        summary = (
            filtered_df
            .groupby(
                ["Data Type", "Category"]
            )["Amount"]
            .sum()
            .reset_index()
        )

        fig1 = px.bar(
            summary,
            x="Category",
            y="Amount",
            color="Data Type",
            barmode="group",
            title="Actual vs Budget"
        )

        chart1.plotly_chart(
            fig1,
            use_container_width=True
        )

        monthly = (
            filtered_df
            .groupby(
                ["Month", "Category"]
            )["Amount"]
            .sum()
            .reset_index()
        )

        fig2 = px.line(
            monthly,
            x="Month",
            y="Amount",
            color="Category",
            markers=True,
            title="Monthly Trend"
        )

        chart2.plotly_chart(
            fig2,
            use_container_width=True
        )

        # ----------------------------------------------------
        # INCOME VS EXPENSE PIE
        # ----------------------------------------------------
        st.subheader("📊 Budget Composition")

        pie_df = (
            budget_df
            .groupby("Category")["Amount"]
            .sum()
            .reset_index()
        )

        fig3 = px.pie(
            pie_df,
            names="Category",
            values="Amount",
            hole=0.5,
            title="Income vs Expense Share"
        )

        st.plotly_chart(
            fig3,
            use_container_width=True
        )

        # ----------------------------------------------------
        # PROFIT & LOSS STATEMENT
        # ----------------------------------------------------
        st.subheader("📑 Profit & Loss Statement")

        pnl = (
            budget_df
            .groupby(
                ["Category", "Particular"]
            )["Amount"]
            .sum()
            .reset_index()
        )

        st.dataframe(
            pnl,
            use_container_width=True
        )

        # ----------------------------------------------------
        # EXPORT
        # ----------------------------------------------------
        excel_file = convert_df_to_excel(pnl)

        st.download_button(
            "⬇ Download Profit & Loss Statement",
            data=excel_file,
            file_name="Profit_Loss_2026_27.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error : {e}")

else:
    st.info(
        "Please upload both Income and Expense files to begin."
    )
