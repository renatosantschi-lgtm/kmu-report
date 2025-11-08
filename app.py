# app.py (kleine stabile Version)
import io, numpy as np, pandas as pd, streamlit as st
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4

st.set_page_config(page_title="KMU Report", page_icon="📄", layout="centered")
st.title("KMU Jahresreport Generator")

# --- Helpers ---
def _normalize(df):
    df = df.iloc[:, :2].copy()
    df.columns = ["account", "amount"]
    df["account"] = df["account"].astype(str).str.strip()
    return df

def _get(df, key):
    row = df.loc[df["account"].str.lower()==key.lower()]
    return float(row["amount"].values[0]) if not row.empty else 0.0

def load_excel_from_bytes(raw_bytes: bytes):
    bio = io.BytesIO(raw_bytes); bio.seek(0)
    xls = pd.ExcelFile(bio)
    b = _normalize(pd.read_excel(xls, "balance_sheet", header=None))
    p = _normalize(pd.read_excel(xls, "profit_loss",   header=None))
    return b, p

def compute_kpis(b, p):
    cash=_get(b,'cash'); rec=_get(b,'receivables'); inv=_get(b,'inventory')
    cl=_get(b,'current_liabilities'); eq=_get(b,'equity')
    rev=_get(p,'revenue'); cogs=_get(p,'cogs'); pers=_get(p,'personnel'); depr=_get(p,'depr'); it=_get(p,'interest')
    ebit = rev - cogs - pers - depr - it
    return dict(revenue=rev, ebit=ebit, ebit_margin= (ebit/rev if rev else np.nan),
                equity_ratio=(eq/(b['amount'].sum()) if b['amount'].sum() else np.nan),
                liquidity_ratio_2=((cash+rec)/cl if cl else np.nan))

def fmt_pct(x): 
    return "-" if (x is None or (isinstance(x,float) and np.isnan(x))) else f"{x*100:.1f}%"

def fmt_num(x): 
    return "-" if (x is None or (isinstance(x,float) and np.isnan(x))) else f"{x:,.0f}".replace(",", "'")

def simple_pdf(k, org, period):
    buf = io.BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(buf, pagesize=A4, title=f"{org} – Jahresreport")
    story=[]
    story.append(Paragraph(f"<b>{org}</b>", styles["Title"]))
    story.append(Paragraph(period, styles["Normal"]))
    story.append(Spacer(1,8))
    data = [
        ["Umsatz", f"{fmt_num(k['revenue'])} CHF"],
        ["EBIT",   f"{fmt_num(k['ebit'])} CHF"],
        ["EBIT-Marge", fmt_pct(k['ebit_margin'])],
        ["EK-Quote",   fmt_pct(k['equity_ratio'])],
        ["Liquidität II", fmt_pct(k['liquidity_ratio_2'])]
    ]
    t = Table(data, colWidths=[140,240])
    t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#E5E7EB"))]))
    story.append(t)
    doc.build(story)
    buf.seek(0)
    return buf

# --- UI ---
org = st.text_input("Firma", "Bäckerei Santschi GmbH")
period = st.text_input("Periode", "Geschäftsjahr 2024")
uploaded = st.file_uploader("Excel auswählen (.xlsx) – benötigt die Sheets balance_sheet und profit_loss", type=["xlsx"])

if uploaded:
    try:
        b, p = load_excel_from_bytes(uploaded.read())
    except Exception as e:
        st.error(f"Lesefehler: {e}")
        st.stop()
    k = compute_kpis(b, p)
    c1,c2,c3 = st.columns(3)
    c1.metric("Umsatz", fmt_num(k["revenue"]))
    c2.metric("EBIT", fmt_num(k["ebit"]))
    c3.metric("EBIT-Marge", fmt_pct(k["ebit_margin"]))
    if st.button("📄 PDF erzeugen"):
        pdf_buf = simple_pdf(k, org, period)
        st.download_button("Download PDF", data=pdf_buf.getvalue(), file_name="report.pdf", mime="application/pdf")
