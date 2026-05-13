
"""
Smart Data Analyzer – Dashboard تفاعلي
تشغيل: streamlit run streamlit_dashboard.py
"""
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pathlib import Path
import sys, os

st.set_page_config(
    page_title="Smart Data Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS مخصص ──
st.markdown("""
<style>
.main {background-color: #F0F4F8;}
.stMetric label {font-size: 12px; color: #7F8C8D;}
.stMetric value {font-size: 24px; font-weight: bold; color: #1B4F72;}
h1 {color: #1B4F72; border-bottom: 3px solid #F39C12; padding-bottom: 8px;}
h2 {color: #2E86C1;}
</style>
""", unsafe_allow_html=True)

# ── الشريط الجانبي ──
with st.sidebar:
    st.image("https://via.placeholder.com/200x60/1B4F72/FFFFFF?text=Smart+Analyzer",
             use_column_width=True)
    st.title("📁 رفع الملف")
    uploaded = st.file_uploader("اختر ملف Excel", type=["xlsx","xls"])
    st.markdown("---")
    st.info("📊 Smart Data Analyzer\nنظام تحليل بيانات متكامل")

st.title("📊 Smart Data Analyzer – لوحة المعلومات التفاعلية")

if uploaded is None:
    st.warning("⬆️ يرجى رفع ملف Excel من الشريط الجانبي للبدء")
    st.stop()

# ── تحميل البيانات ──
@st.cache_data
def load_excel(file):
    xl = pd.ExcelFile(file)
    return {name: xl.parse(name) for name in xl.sheet_names}

sheets = load_excel(uploaded)
sheet_name = st.selectbox("اختر الورقة", list(sheets.keys()))
df = sheets[sheet_name].copy()

# ── تنظيف تلقائي ──
df.dropna(how="all", inplace=True)
num_cols = df.select_dtypes(include=[np.number]).columns
df[num_cols] = df[num_cols].fillna(df[num_cols].median())

# ── KPIs ──
st.subheader("📈 المؤشرات الرئيسية")
cols = st.columns(4)
numeric = df.select_dtypes(include=[np.number])
for i, col in enumerate(list(numeric.columns)[:4]):
    with cols[i]:
        total = numeric[col].sum()
        mean  = numeric[col].mean()
        st.metric(col, f"{total:,.0f}", f"متوسط: {mean:,.0f}")

st.markdown("---")

# ── البيانات الخام ──
with st.expander("📋 عرض البيانات المنظّفة"):
    st.dataframe(df, use_container_width=True)
    st.caption(f"{len(df):,} صف × {len(df.columns)} عمود")

# ── الرسوم البيانية ──
st.subheader("📊 الرسوم البيانية")
cat_cols = [c for c in df.columns if df[c].nunique() < 20 and df[c].dtype == object]
val_cols = list(numeric.columns)

if cat_cols and val_cols:
    c1, c2 = st.columns(2)
    with c1:
        cat = st.selectbox("فئة المحور X", cat_cols, key="bar_cat")
        val = st.selectbox("قيمة المحور Y", val_cols, key="bar_val")
        grp = df.groupby(cat)[val].sum().sort_values(ascending=False).head(10)
        fig, ax = plt.subplots(figsize=(8,5))
        ax.bar(range(len(grp)), grp.values,
               color=["#2E86C1","#F39C12","#27AE60","#E74C3C","#8E44AD"] * 3)
        ax.set_xticks(range(len(grp)))
        ax.set_xticklabels([str(x)[:15] for x in grp.index], rotation=30, ha="right")
        ax.set_title(f"{cat} vs {val}", fontweight="bold")
        ax.spines[["top","right"]].set_visible(False)
        st.pyplot(fig)
        plt.close()
    with c2:
        grp2 = df.groupby(cat)[val].sum().sort_values(ascending=False).head(6)
        fig2, ax2 = plt.subplots(figsize=(7,6))
        ax2.pie(grp2.values, labels=[str(x)[:12] for x in grp2.index],
                autopct="%1.1f%%",
                colors=["#2E86C1","#F39C12","#27AE60","#E74C3C","#8E44AD","#17A589"],
                wedgeprops=dict(edgecolor="white", linewidth=1.5))
        ax2.set_title(f"توزيع {cat}", fontweight="bold")
        st.pyplot(fig2)
        plt.close()

# ── مصفوفة الارتباط ──
if len(val_cols) >= 2:
    st.subheader("🌡️ مصفوفة الارتباط")
    corr = numeric.corr()
    fig3, ax3 = plt.subplots(figsize=(10, max(4, len(corr)-1)))
    im = ax3.imshow(corr.values, cmap="RdYlGn", vmin=-1, vmax=1)
    ax3.set_xticks(range(len(corr.columns)))
    ax3.set_yticks(range(len(corr.columns)))
    ax3.set_xticklabels([c[:12] for c in corr.columns], rotation=45, ha="right")
    ax3.set_yticklabels([c[:12] for c in corr.columns])
    for i in range(len(corr)):
        for j in range(len(corr)):
            ax3.text(j, i, f"{corr.values[i,j]:.2f}", ha="center", va="center", fontsize=8)
    plt.colorbar(im, ax=ax3)
    st.pyplot(fig3)
    plt.close()

# ── الإحصاء الوصفي ──
st.subheader("📐 الإحصاء الوصفي")
st.dataframe(df.describe().round(2), use_container_width=True)

st.markdown("---")
st.caption("🤖 Smart Data Analyzer | تحليل ذكي آلي | 2026")
