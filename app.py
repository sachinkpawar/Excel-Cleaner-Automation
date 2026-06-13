import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

st.set_page_config(
    page_title="Smart Excel Cleaner",
    page_icon="📊",
    layout="wide"
)

def generate_pdf_report(total_rows, total_columns, duplicate_rows, missing_values, quality_score, missing_df):
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)

    width, height = A4
    y = height - 50

    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(50, y, "Smart Excel Data Quality Report")

    y -= 40
    pdf.setFont("Helvetica", 11)
    pdf.drawString(50, y, "This report summarizes the data quality of the uploaded Excel file.")

    y -= 40
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(50, y, "Dataset Summary")

    y -= 25
    pdf.setFont("Helvetica", 11)
    pdf.drawString(50, y, f"Total Rows: {total_rows}")
    y -= 20
    pdf.drawString(50, y, f"Total Columns: {total_columns}")
    y -= 20
    pdf.drawString(50, y, f"Duplicate Rows: {duplicate_rows}")
    y -= 20
    pdf.drawString(50, y, f"Missing Values: {missing_values}")
    y -= 20
    pdf.drawString(50, y, f"Data Quality Score: {quality_score}/100")

    y -= 40
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(50, y, "Recommendations")

    y -= 25
    pdf.setFont("Helvetica", 11)

    recommendations = []

    if duplicate_rows > 0:
        recommendations.append("Remove duplicate rows to avoid repeated records.")
    else:
        recommendations.append("No duplicate rows found.")

    if missing_values > 0:
        recommendations.append("Review columns with missing values before analysis.")
    else:
        recommendations.append("No missing values found.")

    if quality_score >= 90:
        recommendations.append("Dataset quality is strong and suitable for analysis.")
    elif quality_score >= 70:
        recommendations.append("Dataset quality is moderate. Clean missing and duplicate data.")
    else:
        recommendations.append("Dataset quality is low. Data cleaning is strongly recommended.")

    for i, rec in enumerate(recommendations, start=1):
        pdf.drawString(50, y, f"{i}. {rec}")
        y -= 20

    y -= 25
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(50, y, "Top Columns with Missing Values")

    y -= 25
    pdf.setFont("Helvetica", 10)

    top_missing = missing_df.sort_values("Missing Values", ascending=False).head(10)

    for _, row in top_missing.iterrows():
        text = f"{row['Column']}: {row['Missing Values']} missing ({row['Missing %']}%)"
        pdf.drawString(50, y, text[:90])
        y -= 16

        if y < 60:
            pdf.showPage()
            y = height - 50

    pdf.save()
    buffer.seek(0)
    return buffer


st.title("📊 Smart Excel Data Cleaner & Analyzer")
st.write("Upload an Excel file, analyze data quality, clean it, and download the cleaned file/report.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Original Data Preview")
    st.dataframe(df.head(20), use_container_width=True)

    total_rows = len(df)
    total_columns = len(df.columns)
    duplicate_rows = df.duplicated().sum()
    missing_values = df.isnull().sum().sum()

    quality_score = max(
        0,
        round(
            100
            - ((duplicate_rows / max(total_rows, 1)) * 40)
            - ((missing_values / max(total_rows * total_columns, 1)) * 60),
            2,
        ),
    )

    st.subheader("Data Quality Dashboard")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Rows", total_rows)
    col2.metric("Columns", total_columns)
    col3.metric("Duplicates", duplicate_rows)
    col4.metric("Missing Values", missing_values)
    col5.metric("Quality Score", f"{quality_score}/100")

    st.subheader("Column-wise Missing Values")

    missing_df = pd.DataFrame({
        "Column": df.columns,
        "Missing Values": df.isnull().sum().values,
        "Missing %": (df.isnull().sum().values / max(len(df), 1) * 100).round(2)
    })

    st.dataframe(missing_df, use_container_width=True)

    st.subheader("Download Report")

    pdf_report = generate_pdf_report(
        total_rows,
        total_columns,
        duplicate_rows,
        missing_values,
        quality_score,
        missing_df
    )

    st.download_button(
        label="Download Data Quality PDF Report",
        data=pdf_report,
        file_name="data_quality_report.pdf",
        mime="application/pdf"
    )

    st.subheader("Missing Values Chart")

    missing_chart = missing_df[missing_df["Missing Values"] > 0]

    if len(missing_chart) > 0:
        fig = px.bar(
            missing_chart,
            x="Column",
            y="Missing Values",
            title="Missing Values by Column"
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.success("No missing values found.")

    numeric_cols = df.select_dtypes(include="number").columns

    if len(numeric_cols) > 0:
        st.subheader("Numeric Column Analysis")

        selected_col = st.selectbox(
            "Select Numeric Column",
            numeric_cols
        )

        fig = px.histogram(
            df,
            x=selected_col,
            title=f"Distribution of {selected_col}"
        )

        st.plotly_chart(fig, use_container_width=True)

    if st.button("Clean Data"):
        cleaned_df = df.copy()

        rows_before = len(cleaned_df)

        cleaned_df = cleaned_df.dropna(how="all")
        cleaned_df = cleaned_df.drop_duplicates()

        cleaned_df.columns = cleaned_df.columns.astype(str).str.strip()

        for col in cleaned_df.select_dtypes(include=["object", "string"]).columns:
            cleaned_df[col] = cleaned_df[col].astype(str).str.strip()

        rows_after = len(cleaned_df)
        rows_removed = rows_before - rows_after

        st.success("Data cleaned successfully!")

        st.subheader("Cleaning Summary")
        st.write(f"Rows before cleaning: {rows_before}")
        st.write(f"Rows after cleaning: {rows_after}")
        st.write(f"Rows removed: {rows_removed}")

        st.subheader("Cleaned Data Preview")
        st.dataframe(cleaned_df.head(20), use_container_width=True)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            cleaned_df.to_excel(writer, index=False, sheet_name="Cleaned_Data")

        st.download_button(
            label="Download Cleaned Excel File",
            data=output.getvalue(),
            file_name="cleaned_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload an Excel file to begin.")