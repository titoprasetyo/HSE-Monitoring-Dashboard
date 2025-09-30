import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns  # <-- tambahan untuk Risk Matrix
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

# -----------------------------
# Fungsi buat PDF rapi
# -----------------------------
def export_pdf(summary_dict, charts):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("<b>Laporan HSE</b>", styles['Title']))
    story.append(Spacer(1, 12))

    for sheet_name, summary in summary_dict.items():
        story.append(Paragraph(f"📌 <b>{sheet_name}</b>", styles['Heading2']))
        for key, val in summary.items():
            if isinstance(val, pd.DataFrame):
                continue
            story.append(Paragraph(f"- {key}: {val}", styles['Normal']))
        story.append(Spacer(1, 6))

        if "Trend" in summary:
            trend_df = summary["Trend"]
            data = [trend_df.columns.tolist()] + trend_df.values.tolist()
            table = Table(data, hAlign="LEFT")
            table.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.grey),
                ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                ("GRID", (0,0), (-1,-1), 0.5, colors.black),
            ]))
            story.append(table)
            story.append(Spacer(1, 6))

        # Masukkan semua grafik untuk sheet ini
        for key, img_path in charts.items():
            if key.startswith(sheet_name):
                story.append(Image(img_path, width=400, height=250))
                story.append(Spacer(1, 12))

    # Footer
    story.append(Spacer(1, 50))
    story.append(Paragraph("<para align='center'>© 2025 Tito Prasetyo Ashiddiq</para>", styles['Normal']))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf


# -----------------------------
# Fungsi export Excel (support pie/column chart)
# -----------------------------
def export_excel(dfs_dict, summary_dict, chart_type_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        for sheet, df in dfs_dict.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

            # Grafik tren
            if "Trend" in summary_dict.get(sheet, {}):
                trend = summary_dict[sheet]["Trend"]
                trend.to_excel(writer, sheet_name=f"{sheet}_Trend", index=False)

                worksheet = writer.sheets[f"{sheet}_Trend"]
                chart_type = chart_type_dict.get(f"{sheet}_Trend", "line")
                chart = workbook.add_chart({"type": chart_type})
                chart.add_series({
                    "name": f"Trend {sheet}",
                    "categories": [f"{sheet}_Trend", 1, 0, len(trend), 0],
                    "values":     [f"{sheet}_Trend", 1, 1, len(trend), 1],
                })
                chart.set_title({"name": f"Trend {sheet}"})
                chart.set_x_axis({"name": "Bulan"})
                chart.set_y_axis({"name": "Jumlah"})
                worksheet.insert_chart("E2", chart)

            # Grafik distribusi
            for col in ["Jenis", "Severity", "Status"]:
                if col in df.columns and not df[col].empty:
                    counts = df[col].value_counts().reset_index()
                    counts.columns = [col, "Jumlah"]
                    counts.to_excel(writer, sheet_name=f"{sheet}_{col}", index=False)

                    worksheet = writer.sheets[f"{sheet}_{col}"]
                    chart_type = chart_type_dict.get(f"{sheet}_{col}", "column")
                    chart = workbook.add_chart({"type": chart_type})

                    if chart_type == "pie":
                        chart.add_series({
                            "name": f"Distribusi {col} - {sheet}",
                            "categories": [f"{sheet}_{col}", 1, 0, len(counts), 0],
                            "values":     [f"{sheet}_{col}", 1, 1, len(counts), 1],
                            "data_labels": {"percentage": True},
                        })
                    else:
                        chart.add_series({
                            "name": f"Distribusi {col} - {sheet}",
                            "categories": [f"{sheet}_{col}", 1, 0, len(counts), 0],
                            "values":     [f"{sheet}_{col}", 1, 1, len(counts), 1],
                            "data_labels": {"value": True},
                        })

                    chart.set_title({"name": f"Distribusi {col}"})
                    worksheet.insert_chart("E2", chart)

    return output.getvalue()


# -----------------------------
# Streamlit Dashboard
# -----------------------------
st.sidebar.title("📊 HSE Dashboard")

menu = ["Home", "Upload File Excel"]
if "xls" in st.session_state:
    menu += st.session_state["xls"].sheet_names
menu.append("Download Laporan")

choice = st.sidebar.radio("Pilih menu:", menu)

if choice == "Home":
    st.title("📊 HSE Monitoring Dashboard")
    st.markdown(""" 
                Selamat datang di **HSE Dashboard** Dashboard ini membantu memonitor data HSE: 
                - 👷 **Incidents & Near Miss** 
                - 📜 **Permit To Work** 
                - 🔒 **LOTO (Lock Out Tag Out)** 
                - 🎓 **Training & Competency** 
                - 📝 **HIRADC (Hazard Identification, Risk Assessment, and Determining Control)**
                -    ** Dan yang lainnya**
                """)
    st.info("Silakan upload file Excel di menu samping.")

elif choice == "Upload File Excel":
    st.title("📂 Upload File Excel")
    uploaded_file = st.file_uploader("Upload File Excel (xlsx)", type=["xlsx"])

    if uploaded_file:
        if st.button("📥 Load File dengan double klik"):
            st.session_state["xls"] = pd.ExcelFile(uploaded_file)
            st.session_state["summary_dict"] = {}
            st.session_state["dfs_dict"] = {}
            st.session_state["charts"] = {}
            st.session_state["chart_type_dict"] = {}
            st.success("✅ File berhasil dimuat!")

elif choice in (st.session_state.get("xls").sheet_names if "xls" in st.session_state else []):
    sheet = choice
    st.title(f"📌 Analisa Sheet: {sheet}")

    df = pd.read_excel(st.session_state["xls"], sheet_name=sheet)

    if df.empty:
        st.warning("⚠️ Data kosong.")
    else:
        # Analisa HIRADC jika ada Likelihood & Severity
        if all(col in df.columns for col in ["Likelihood", "Severity"]):
            st.subheader("📊 Analisa HIRADC")

            df["Risk Rating"] = df["Likelihood"] * df["Severity"]
            st.dataframe(df)

            st.write("### Ringkasan HIRADC")
            st.write(f"- Rata-rata Risk Rating: {df['Risk Rating'].mean():.2f}")
            st.write(f"- Risk Rating tertinggi: {df['Risk Rating'].max()}")

            # Risk Matrix
            st.write("### 📌 Risk Matrix")
            matrix = pd.crosstab(df["Likelihood"], df["Severity"])
            fig, ax = plt.subplots(figsize=(6,5))
            sns.heatmap(matrix, annot=True, fmt="d", cmap="Reds", ax=ax, cbar=False)
            ax.set_title("Risk Matrix (Likelihood vs Severity)")
            ax.set_xlabel("Severity")
            ax.set_ylabel("Likelihood")
            st.pyplot(fig)

            img_path = f"{sheet}_RiskMatrix.png"
            fig.savefig(img_path, bbox_inches="tight")
            st.session_state["charts"][f"{sheet}_RiskMatrix"] = img_path

        # Analisa lain (Trend + Distribusi)
        else:
            # Filter tanggal
            if "Tanggal" in df.columns:
                df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
                min_date, max_date = df["Tanggal"].min(), df["Tanggal"].max()
                start_date, end_date = st.date_input("📅 Filter tanggal", [min_date, max_date])
                df = df[(df["Tanggal"].dt.date >= start_date) & (df["Tanggal"].dt.date <= end_date)]

            # Filter dinamis
            for col in ["Jenis", "Severity", "Status"]:
                if col in df.columns:
                    options = df[col].dropna().unique().tolist()
                    selected = st.multiselect(f"Filter {col}", options, default=options)
                    df = df[df[col].isin(selected)]

            st.dataframe(df)

            # Analisa ringkasan
            summary = {}
            if "Jenis" in df.columns and not df["Jenis"].empty:
                summary["Jenis terbanyak"] = df["Jenis"].mode()[0]
            if "Severity" in df.columns and not df["Severity"].empty:
                summary["Severity dominan"] = df["Severity"].mode()[0]
            if "Status" in df.columns and not df["Status"].empty:
                summary["Status dominan"] = df["Status"].mode()[0]
            if "Tanggal" in df.columns:
                trend = df.groupby(df["Tanggal"].dt.to_period("M")).size().reset_index(name="Jumlah")
                trend["Tanggal"] = trend["Tanggal"].astype(str)
                summary["Trend"] = trend

                chart_type = st.selectbox(f"Jenis grafik tren {sheet}", ["line", "column"], key=f"{sheet}_Trend")
                st.session_state["chart_type_dict"][f"{sheet}_Trend"] = chart_type

                fig, ax = plt.subplots()
                if chart_type == "line":
                    ax.plot(trend["Tanggal"], trend["Jumlah"], marker="o", label="Jumlah Kasus")
                else:
                    ax.bar(trend["Tanggal"], trend["Jumlah"], label="Jumlah Kasus")
                ax.set_title(f"Trend {sheet} Perbulan")
                ax.set_xlabel("Bulan")
                ax.set_ylabel("Jumlah")
                for i, val in enumerate(trend["Jumlah"]):
                    ax.text(i, val, str(val), ha="center", va="bottom", fontsize=8)
                plt.xticks(rotation=45, ha="right")
                ax.legend()
                st.pyplot(fig)

                img_path = f"{sheet}_Trend.png"
                fig.savefig(img_path, bbox_inches="tight")
                st.session_state["charts"][f"{sheet}_Trend"] = img_path

            # Distribusi kategorikal
            for col in ["Jenis", "Severity", "Status"]:
                if col in df.columns:
                    chart_type = st.selectbox(f"Jenis grafik distribusi {col} - {sheet}", ["column", "pie"], key=f"{sheet}_{col}")
                    st.session_state["chart_type_dict"][f"{sheet}_{col}"] = chart_type

                    counts = df[col].value_counts()

                    fig, ax = plt.subplots()
                    if chart_type == "pie":
                        ax.pie(counts.values, labels=counts.index, autopct="%1.1f%%")
                    else:
                        ax.bar(counts.index, counts.values)
                        for i, val in enumerate(counts.values):
                            ax.text(i, val, str(val), ha="center", va="bottom", fontsize=8)
                    ax.set_title(f"Distribusi {col} - {sheet}")
                    st.pyplot(fig)

                    img_path = f"{sheet}_{col}.png"
                    fig.savefig(img_path, bbox_inches="tight")
                    st.session_state["charts"][f"{sheet}_{col}"] = img_path

            st.write("### Ringkasan")
            for k, v in summary.items():
                if isinstance(v, pd.DataFrame):
                    st.table(v)
                else:
                    st.write(f"- {k}: {v}")

            # Simpan hasil
            st.session_state["summary_dict"][sheet] = summary
            st.session_state["dfs_dict"][sheet] = df

elif choice == "Download Laporan":
    st.title("⬇️ Download Laporan")

    if "summary_dict" not in st.session_state or "dfs_dict" not in st.session_state:
        st.warning("⚠️ Belum ada analisa yang tersimpan.")
    else:
        pdf_bytes = export_pdf(st.session_state["summary_dict"], st.session_state["charts"])
        st.download_button("⬇️ Download PDF", data=pdf_bytes, file_name="laporan_hse.pdf", mime="application/pdf")

        excel_bytes = export_excel(st.session_state["dfs_dict"], st.session_state["summary_dict"], st.session_state["chart_type_dict"])
        st.download_button("⬇️ Download Excel", data=excel_bytes, file_name="laporan_hse.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------------------------
# Copyright Footer
# -------------------------
footer = """
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: transparent;
        color: grey;
        text-align: center;
        padding: 10px;
        font-size: 13px;
    }
    </style>
    <div class="footer">
        © 2025 Tito Prasetyo Ashiddiq
    </div>
"""
st.markdown(footer, unsafe_allow_html=True)

