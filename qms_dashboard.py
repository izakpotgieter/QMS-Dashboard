import streamlit as st
import pandas as pd
import plotly.express as px
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import datetime as dt
import os

st.set_page_config(page_title="QMS Dashboard", layout="wide")

st.sidebar.image("aklin_logo.png", width=180)
st.sidebar.title("Aklin Carbide QMS Portal")
st.sidebar.caption("ISO 9001:2015 Analytics Dashboard")

uploaded = st.sidebar.file_uploader("Upload STAGE1_QMS_TEMP.xlsx", type=["xlsx"])

st.title("ISO 9001:2015 QMS Dashboard")
st.caption("Executive overview of risks, NCR/CAPA, audits, inspections, suppliers, training, documents, and process performance.")

if uploaded:
    xls = pd.ExcelFile(uploaded)

    def load(name):
        try:
            df = xls.parse(name)
            df.columns = df.columns.str.strip().str.lower()
            return df
        except:
            return pd.DataFrame()

    df_risk = load("Risk_Register")
    df_ncr = load("NCR_Log")
    df_capa = load("CAPA")
    df_aud = load("Internal_Audits")
    df_insp = load("Inspections")
    df_sup = load("Supplier_Eval")
    df_train = load("Training")
    df_doc = load("Doc_Control")
    df_proc = load("Process_Map")

    st.success("Data loaded successfully.")

    def count_open(df, col):
        if df.empty or col not in df.columns:
            return 0
        return len(df[~df[col].str.lower().str.contains("closed")])

    kpi1 = count_open(df_ncr, "status")
    kpi2 = count_open(df_capa, "status")
    kpi3 = len(df_risk[df_risk["risk score"] >= 15]) if "risk score" in df_risk.columns else 0
    kpi4 = count_open(df_insp, "pass/fail")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Open NCRs", kpi1)
    col2.metric("Open CAPAs", kpi2)
    col3.metric("High Risks (≥15)", kpi3)
    col4.metric("Failed Inspections", kpi4)

    tabs = st.tabs([
        "Risks", "NCR / CAPA", "Audits", "Inspections",
        "Suppliers", "Training", "Documents",
        "Process Map", "Management Review", "Export PDF"
    ])

    with tabs[0]:
        st.subheader("Risk Profile")
        if not df_risk.empty and "risk score" in df_risk.columns:
            bins = [0, 5, 10, 15, 25]
            labels = ["Low", "Medium", "High", "Critical"]
            df_risk["band"] = pd.cut(df_risk["risk score"], bins=bins, labels=labels, include_lowest=True)
            band = df_risk["band"].value_counts().reset_index()
            band.columns = ["Category", "Count"]
            fig = px.bar(
                band,
                x="Category",
                y="Count",
                title="Risk Bands",
                color="Category",
                color_discrete_map={
                    "Low": "#2ecc71",
                    "Medium": "#f1c40f",
                    "High": "#e67e22",
                    "Critical": "#e74c3c"
                }
            )
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Risk data incomplete or missing 'Risk Score' column.")
        st.dataframe(df_risk, use_container_width=True)

    with tabs[1]:
        st.subheader("NCR and CAPA Overview")
        col_ncr, col_capa = st.columns(2)

        with col_ncr:
            st.markdown("**NCR Status**")
            if not df_ncr.empty and "status" in df_ncr.columns:
                status = df_ncr["status"].value_counts().reset_index()
                status.columns = ["Category", "Count"]
                fig = px.bar(
                    status,
                    x="Category",
                    y="Count",
                    title="NCR Status",
                    color="Category",
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("NCR data missing or incomplete.")
            st.dataframe(df_ncr, use_container_width=True)

        with col_capa:
            st.markdown("**CAPA Status**")
            if not df_capa.empty and "status" in df_capa.columns:
                status = df_capa["status"].value_counts().reset_index()
                status.columns = ["Category", "Count"]
                fig = px.bar(
                    status,
                    x="Category",
                    y="Count",
                    title="CAPA Status",
                    color="Category",
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("CAPA data missing or incomplete.")
            st.dataframe(df_capa, use_container_width=True)

    with tabs[2]:
        st.subheader("Internal Audit Results")
        if not df_aud.empty and "severity" in df_aud.columns:
            sev = df_aud["severity"].value_counts().reset_index()
            sev.columns = ["Category", "Count"]
            fig = px.bar(
                sev,
                x="Category",
                y="Count",
                title="Audit Severity",
                color="Category",
                color_discrete_map={
                    "minor": "#f1c40f",
                    "major": "#e67e22",
                    "critical": "#e74c3c"
                }
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Audit data missing 'Severity' column.")
        st.dataframe(df_aud, use_container_width=True)

    with tabs[3]:
        st.subheader("Inspection Performance")
        if not df_insp.empty and "pass/fail" in df_insp.columns:
            pf = df_insp["pass/fail"].value_counts().reset_index()
            pf.columns = ["Category", "Count"]
            fig = px.bar(
                pf,
                x="Category",
                y="Count",
                title="Inspection Results",
                color="Category",
                color_discrete_map={
                    "pass": "#2ecc71",
                    "fail": "#e74c3c"
                }
            )
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Inspection data missing 'Pass/Fail' column.")
        st.dataframe(df_insp, use_container_width=True)

    with tabs[4]:
        st.subheader("Supplier Performance")
        if not df_sup.empty and "band" in df_sup.columns:
            band = df_sup["band"].value_counts().reset_index()
            band.columns = ["Category", "Count"]
            fig = px.bar(
                band,
                x="Category",
                y="Count",
                title="Supplier Bands",
                color="Category",
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Supplier data missing 'Band' column.")
        st.dataframe(df_sup, use_container_width=True)

    with tabs[5]:
        st.subheader("Training Status")
        if not df_train.empty and "status" in df_train.columns:
            status = df_train["status"].value_counts().reset_index()
            status.columns = ["Category", "Count"]
            fig = px.bar(
                status,
                x="Category",
                y="Count",
                title="Training Status",
                color="Category",
                color_discrete_sequence=px.colors.qualitative.Set1
            )
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Training data missing 'Status' column.")
        st.dataframe(df_train, use_container_width=True)

    with tabs[6]:
        st.subheader("Document Control")
        if not df_doc.empty and "status" in df_doc.columns:
            status = df_doc["status"].value_counts().reset_index()
            status.columns = ["Category", "Count"]
            fig = px.bar(
                status,
                x="Category",
                y="Count",
                title="Document Status",
                color="Category",
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Document data missing 'Status' column.")
        st.dataframe(df_doc, use_container_width=True)

    with tabs[7]:
        st.subheader("Process Map Overview")
        st.dataframe(df_proc, use_container_width=True)

    with tabs[8]:
        st.subheader("Management Review – Auto Summary")
        topics = []
        if not df_ncr.empty:
            topics.append(["Customer complaints / NCRs", len(df_ncr)])
        if not df_aud.empty:
            topics.append(["Internal audit results", len(df_aud)])
        if not df_capa.empty:
            topics.append(["CAPA status", len(df_capa)])
        if not df_sup.empty:
            topics.append(["Supplier performance", len(df_sup)])
        if not df_train.empty:
            topics.append(["Training status", len(df_train)])
        if not df_risk.empty:
            topics.append(["Risk profile", len(df_risk)])

        if topics:
            df_mr = pd.DataFrame(topics, columns=["Topic", "Count"])
            st.dataframe(df_mr, use_container_width=True)
        else:
            df_mr = pd.DataFrame(columns=["Topic", "Count"])
            st.info("No management review data available.")

    with tabs[9]:
        st.subheader("Export Management Review to PDF")
        if df_mr.empty:
            st.info("No Management Review data to export.")
        else:
            if st.button("Generate PDF"):
                file_path = "Management_Review.pdf"
                c = canvas.Canvas(file_path, pagesize=A4)
                c.setFont("Helvetica-Bold", 16)
                c.drawString(50, 800, "Aklin Carbide – Management Review Summary")
                c.setFont("Helvetica", 12)
                y = 760
                for _, row in df_mr.iterrows():
                    text = f"{row['Topic']}: {row['Count']}"
                    c.drawString(50, y, text)
                    y -= 20
                    if y < 50:
                        c.showPage()
                        c.setFont("Helvetica", 12)
                        y = 800
                c.save()
                with open(file_path, "rb") as f:
                    st.download_button(
                        label="Download Management Review PDF",
                        data=f,
                        file_name="Management_Review.pdf",
                        mime="application/pdf"
                    )
                st.success("PDF generated successfully.")
else:
    st.info("Upload your Stage 1 QMS Excel file using the sidebar to begin.")