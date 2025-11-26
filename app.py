import streamlit as st
import pandas as pd
import datetime
import io
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pytz  # لإعداد التوقيت المحلي

# ---------- Arabic helpers ----------
def fix_arabic(text):
    if pd.isna(text):
        return ""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)

def fill_down(series):
    return series.ffill()



    final_cols = [
        'كود الاوردر', 'اسم العميل', 'المنطقة', 'العنوان',
        'المدينة', 'رقم موبايل العميل', 'حالة الاوردر',
        'عدد القطع', 'الملاحظات', 'اسم الصنف',
        'اللون', 'المقاس', 'الكمية',
        'الإجمالي مع الشحن'
    ]
    df = df[[c for c in final_cols if c in df.columns]].copy()

    if 'رقم موبايل العميل' in df.columns:
        df['رقم موبايل العميل'] = df['رقم موبايل العميل'].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.','',1).isdigit()
            else ("" if pd.isna(x) else str(x))
        )

    safe_cols = {'الإجمالي مع الشحن','كود الاوردر','رقم موبايل العميل','اسم العميل',
                 'المنطقة','العنوان','المدينة','حالة الاوردر','الملاحظات','اسم الصنف','اللون','المقاس'}
    for col in df.columns:
        if col not in safe_cols:
            df[col] = df[col].apply(
                lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.','',1).isdigit()
                else ("" if pd.isna(x) else str(x))
            )

    styleN = ParagraphStyle(name='Normal', fontName='Arabic-Bold', fontSize=9,
                            alignment=1, wordWrap='RTL')
    styleBH = ParagraphStyle(name='Header', fontName='Arabic-Bold', fontSize=10,
                             alignment=1, wordWrap='RTL')
    styleTitle = ParagraphStyle(name='Title', fontName='Arabic-Bold', fontSize=14,
                                alignment=1, wordWrap='RTL')

    data = []
    data.append([Paragraph(fix_arabic(col), styleBH) for col in df.columns])
    for _, row in df.iterrows():
        data.append([Paragraph(fix_arabic("" if pd.isna(row[col]) else str(row[col])), styleN)
                     for col in df.columns])

    col_widths_cm = [2, 2, 1.5, 3, 2, 3, 1.5, 1.5, 2.5, 3.5, 1.5, 1.5, 1, 1.5]
    col_widths = [max(c * 28.35, 15) for c in col_widths_cm]

    tz = pytz.timezone('Africa/Cairo')
    today = datetime.datetime.now(tz).strftime("%Y-%m-%d")
    title_text = f"kuwait-mall {title} | kuwait-mall | {today} "

    elements = [
        Paragraph(fix_arabic(title_text), styleTitle),
        Spacer(1, 14)
    ]

    table = Table(data, colWidths=col_widths[:len(df.columns)], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#90EE90")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))

    elements.append(table)
    elements.append(PageBreak())
    return elements

# ---------- Streamlit App ----------
st.set_page_config(page_title="🛍️🛍️TREND Orders Processor", layout="wide")
st.title("🛍️🛍️ TREND Orders Processor")
st.markdown("...ارفعي الملفات لاستلام الشيت")

uploaded_files = st.file_uploader(
    "Upload Excel files (.xlsx)",
    accept_multiple_files=True,
    type=["xlsx"]
)

# ------------ ❗ هنا التعديل اللي بيوقف كل حاجة ❗ ------------
if uploaded_files:
    st.error("⚠️ هذا الكود للسالكين فقط ⚠️")
    st.stop()
# -------------------------------------------------------------


# الكود تحت دا مش هيشتغل أبداً بعد stop()
