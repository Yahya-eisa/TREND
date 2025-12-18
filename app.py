import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import pytz

st.set_page_config(page_title="📦 المشتريات المجمعة", layout="wide")
st.title("📦 المشتريات المجمعة")
st.markdown("ارفع الملف وخد المشتريات على طول 🔥")

uploaded_file = st.file_uploader("📤 ارفع ملف Excel", type=["xlsx"])

if uploaded_file:
    # قراءة الملف
    xls = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl", dtype=str)
    
    all_frames = []
    for _, df in xls.items():
        df = df.dropna(how="all")
        all_frames.append(df)
    
    if all_frames:
        merged_df = pd.concat(all_frames, ignore_index=True, sort=False)
        
        # تحديد أسماء الأعمدة
        product_col = None
        color_col = None
        size_col = None
        qty_col = None
        
        # البحث عن الأعمدة
        for col in merged_df.columns:
            if 'منتج' in str(col) or 'صنف' in str(col):
                product_col = col
            elif 'لون' in str(col):
                color_col = col
            elif 'مقاس' in str(col):
                size_col = col
            elif 'كمية' in str(col) or 'الكمية' in str(col):
                qty_col = col
        
        if product_col and qty_col:
            # تحويل الكمية لأرقام
            merged_df[qty_col] = pd.to_numeric(merged_df[qty_col], errors='coerce').fillna(0)
            
            # تجميع المشتريات
            group_cols = [product_col]
            if color_col and color_col in merged_df.columns:
                group_cols.append(color_col)
            if size_col and size_col in merged_df.columns:
                group_cols.append(size_col)
            
            products_df = merged_df.groupby(group_cols)[qty_col].sum().reset_index()
            products_df.columns = group_cols + ['إجمالي الكمية']
            products_df = products_df.sort_values('إجمالي الكمية', ascending=False)
            
            # عرض النتيجة
            st.success(f"✅ تم تجميع {len(products_df)} منتج")
            st.dataframe(products_df, use_container_width=True)
            
            # تحميل الملف
            buffer = BytesIO()
            products_df.to_excel(buffer, sheet_name='المشتريات', index=False, engine='openpyxl')
            buffer.seek(0)
            
            tz = pytz.timezone('Africa/Cairo')
            today = datetime.datetime.now(tz).strftime("%Y-%m-%d")
            file_name = f"المشتريات - {today}.xlsx"
            
            st.download_button(
                label="🛒 تحميل المشتريات",
                data=buffer.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ مش لاقي أعمدة المنتج أو الكمية في الملف!")
