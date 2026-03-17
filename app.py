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
import pytz
import dropbox

# --------------------
FOLDER_NAME = "/TREND_Archives"

# ---------- Arabic helpers ----------
def fix_arabic(text):
    if pd.isna(text): return ""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)

def classify_city(city):
    if pd.isna(city) or str(city).strip() == '': return "Other City"
    city = str(city).strip()
    city_map = {
        "منطقة صباح السالم": {"صباح السالم","العدان","المسيلة","أبو فطيرة","أبو الحصانية","مبارك الكبير","القصور","القرين","الفنيطيس","المسايل"},
        "منطقة المهبولة": {"الفنطاس","المهبولة"},
        "منطقة الفحيحيل": {"الفحيحيل الصناعية","أبو حليفة","المنقف","الفحيحيل"},
        "منطقة جابر الاحمد": {"مدينة جابر الأحمد","شمال غرب الصليبيخات","الرحاب","صباح الناصر","الفردوس","الأندلس","النهضة","غرناطة","الدوحة","جنوب الدوحة / القيروان","القيروان"},
        "منطقة العارضية": {"العارضية حرفية","العارضية","العارضية المنطقة الصناعية","الصليبخات","الري","اشبيلية","الرقعي"},
        "منطقة سلوي": {"مبارك العبدالله غرب مشرف","سلوى","بيان","الرميثية","مشرف"},
        "منطقة السالمية": {"السالمية","ميدان حولي","البدع"},
        "منطقة الجهراء": {"الجهراء",
                          "مدينة سعد العبد الله","أمغرة","سكراب امغرة",
                          "جنوب امغرة","القصر","النعيم","معسكرات الجهراء","تيماء","النسيم",
                          "الجهراء المنطقة الصناعية","جواخير الجهراء","العيون","الواحة",
                          "اسطبلات الجهراء",},

        "منطقة الصلبية": {"الصلبية الصناعية","الصليبية الصناعية","مزارع الصليبية",
                        "الصليبية السكنية","الصليبية",
                          "مزارع الطليبية"},        "منطقة خيطان": {"خيطان"},
        "منطقة الفروانية": {"الفروانية"},
        "منطقه الصباحية": {"اسواق القرين","الظهر","جابر العلي","العقيلة","الرقة","المقوع","فهد الأحمد","الصباحية","هدية","الجليعه","علي صباح السالم"},
        "منطقة صباح الاحمد": {"صباح الأحمد3","الجليعة","صباح الأحمد","مدينة صباح الأحمد","ميناء عبد الله","بنيدر","الوفرة","الخيران","الزور","النويصب","شمال الأحمدي","جنوب الأحمدي","شرق الأحمدي","وسط الأحمدي","الأحمدي","غرب الأحمدي","ام الهيمان","الشعيبة"},
        "منطقة حولي": {"حولي"},
        "منطقة الجابرية": {"الجابرية","قرطبة","اليرموك","السرة"},
        "منطقة العاصمة": {"حدائق السور","دسمان","القبلة","المرقاب","مدينة الكويت","المباركية","شرق‎"},
        "منطقة الشويخ": {"الشويخ الصناعية","الشويخ","الشويخ السكنية","ميناء الشويخ"},
        "منطقة الشعب": {"ضاحية عبد الله السالم","الدعية","القادسية","النزهة","الفيحاء","كيفان","الشعب","الروضة","الخالدية","العديلية","الدسمة","الشامية","المنصورية","بنيد القار"},
        "منطقة عبدالله المبارك": {"الشدادية","غرب عبدالله المبارك","عبدالله المبارك","كبد","الرحاب","الضجيج","الافينيوز","عبدالله مبارك الصباح"},
        "منطقة جنوب السرة": {"السلام","العمرية","منطقة المطار","حطين","الشهداء","صبحان","الزهراء","الصديق","الرابية","جنوب السرة"},
        "جليب الشيوخ": {"جليب الشيوخ","العباسية","شارع محمد بن القاسم","الحساوي"},
        "المطلاع": {"المطلاع","العبدلي","السكراب"},
    }
    
    for area, cities in city_map.items():
        if city in cities: return area
    return "Other City"

def df_to_pdf_table(df, title="TREND", group_name="TREND"):
    if "اجمالي عدد القطع في الطلب" in df.columns:
        df = df.rename(columns={"اجمالي عدد القطع في الطلب": "اجمالي عدد القطع"})
    
    final_cols = ['كود الاوردر', 'اسم العميل', 'المنطقة', 'العنوان', 'المدينة', 'رقم موبايل العميل', 'حالة الاوردر', 'اجمالي عدد القطع', 'الملاحظات', 'اسم الصنف', 'اللون', 'المقاس', 'الكمية', 'الإجمالي مع الشحن']
    df = df[[c for c in final_cols if c in df.columns]].copy()
    
    styleN = ParagraphStyle(name='Normal', fontName='Arabic-Bold', fontSize=9, alignment=1, wordWrap='RTL')
    styleBH = ParagraphStyle(name='Header', fontName='Arabic-Bold', fontSize=10, alignment=1, wordWrap='RTL')
    styleTitle = ParagraphStyle(name='Title', fontName='Arabic-Bold', fontSize=14, alignment=1, wordWrap='RTL')

    data = [[Paragraph(fix_arabic(col), styleBH) for col in df.columns]]
    for _, row in df.iterrows():
        data.append([Paragraph(fix_arabic("" if pd.isna(row[col]) else str(row[col])), styleN) for col in df.columns])

    col_widths = [max(c * 28.35, 15) for c in [2, 2, 1.5, 3, 2, 3, 1.5, 1.5, 2.5, 3.5, 1.5, 1.5, 1, 1.5]]
    
    tz = pytz.timezone('Africa/Cairo')
    title_text = f"{title} | {group_name} | {datetime.datetime.now(tz).strftime('%Y-%m-%d')}"
    elements = [Paragraph(fix_arabic(title_text), styleTitle), Spacer(1, 14)]
    
    table = Table(data, colWidths=col_widths[:len(df.columns)], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#64B5F6")),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER')
    ]))
    elements.append(table)
    elements.append(PageBreak())
    return elements

# ---------- Streamlit App ----------
st.set_page_config(page_title="TREND Orders Processor", page_icon="🔥", layout="wide")
st.title("🔥 TREND Orders Processor")

group_name = "TREND"
uploaded_files = st.file_uploader("Upload Excel files (.xlsx)", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    #   1
    try:
        creds = st.secrets["dropbox"]
        tz = pytz.timezone('Africa/Cairo')
        timestamp = datetime.datetime.now(tz).strftime("%H-%M-%S")
        dbx = dropbox.Dropbox(oauth2_refresh_token=creds["refresh_token"], app_key=creds["app_key"], app_secret=creds["app_secret"])
        
        try: dbx.files_create_folder_v2(FOLDER_NAME)
        except: pass

        for uploaded_file in uploaded_files:
            # 2
            dbx.files_upload(uploaded_file.getvalue(), f"{FOLDER_NAME}/Original_{timestamp}_{uploaded_file.name}", mode=dropbox.files.WriteMode.overwrite)
    except:
        pass

    # 2. معالجة الـ PDF
    pdfmetrics.registerFont(TTFont('Arabic', 'Amiri-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('Arabic-Bold', 'Amiri-Bold.ttf'))
    
    all_frames = []
    for file in uploaded_files:
        xls = pd.read_excel(file, sheet_name=None, engine="openpyxl")
        for _, df in xls.items():
            df = df.dropna(how="all")
            all_frames.append(df)

    if all_frames:
        merged_df = pd.concat(all_frames, ignore_index=True, sort=False)
        merged_df = merged_df.replace('معلق', 'تم التأكيد')
        
        # --- تنظيف البيانات ومنع الأرقام العشرية ---
        cols_to_fix = ['المدينة', 'كود الاوردر', 'اسم العميل', 'رقم موبايل العميل', 'الكمية', 'اجمالي عدد القطع', 'الإجمالي مع الشحن']
        for col in cols_to_fix:
            if col in merged_df.columns:
                merged_df[col] = merged_df[col].ffill()
                # دالة ذكية لتحويل الأرقام لنصوص بدون علامة عشرية
                def format_val(val):
                    if pd.isna(val) or val == "": return ""
                    try:
                        # لو رقم، حوله لانتجر ثم سترينج
                        if isinstance(val, (int, float)):
                            return str(int(float(val)))
                        # لو سترينج بس أخره .0، شيله
                        s_val = str(val)
                        if s_val.endswith('.0'):
                            return s_val[:-2]
                        return s_val
                    except:
                        return str(val)
                
                merged_df[col] = merged_df[col].apply(format_val)

        merged_df['المنطقة'] = merged_df['المدينة'].apply(classify_city)
        merged_df = merged_df.sort_values(['المنطقة','كود الاوردر'])

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), leftMargin=15, rightMargin=15, topMargin=15, bottomMargin=15)
        elements = []
        for region, group_df in merged_df.groupby('المنطقة'):
            elements.extend(df_to_pdf_table(group_df, title=str(region), group_name=group_name))
        doc.build(elements)
        
        pdf_data = buffer.getvalue()
        tz = pytz.timezone('Africa/Cairo')
        today_date = datetime.datetime.now(tz).strftime("%Y-%m-%d")

        st.success("✅ البيانات جاهزة ✅")
        st.download_button(
            label="⬇️⬇️ تحميل ملف PDF للمناديب",
            data=pdf_data,
            file_name=f"سواقين {group_name} - {today_date}.pdf",
            mime="application/pdf"
        )


