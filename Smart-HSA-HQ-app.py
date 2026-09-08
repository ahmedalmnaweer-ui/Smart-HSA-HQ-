import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import json

# ============= إعدادات الصفحة =============
st.set_page_config(
    page_title="HSA Smart Cost Analyzer",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============= الأنماط المخصصة =============
st.markdown("""
<style>
    :root {
        --blue: #1565C0;
        --blue-d: #0D47A1;
        --bg: #F8F9FB;
        --card: #FFFFFF;
        --ink: #263238;
        --soft: #607D8B;
        --line: #E3E8EE;
        --green: #2E7D32;
        --red: #C62828;
        --amber: #F57F17;
    }
    
    body {
        font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
        direction: rtl;
        background-color: var(--bg);
    }
    
    .header-container {
        background: linear-gradient(135deg, var(--blue), var(--blue-d));
        color: white;
        padding: 20px;
        border-radius: 12px;
        margin-bottom: 20px;
        box-shadow: 0 2px 8px rgba(13, 71, 161, 0.25);
    }
    
    .stat-card {
        background: white;
        padding: 15px;
        border-radius: 12px;
        border-left: 4px solid var(--blue);
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
    }
    
    .price-highlight {
        background: #E8F5E9;
        color: var(--green);
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        font-size: 20px;
        font-weight: bold;
        margin: 15px 0;
    }
</style>
""", unsafe_allow_html=True)

# ============= إدارة الجلسة =============
if 'hsa_items' not in st.session_state:
    st.session_state.hsa_items = []

if 'hsa_materials' not in st.session_state:
    st.session_state.hsa_materials = {
        "مرحاض عربي نوع خزف": {"price": 12000},
        "مواسير بلاستيك ضغط عالي (UPVC)": {"price": 5000},
        "مفاتيح صنابير نحاس": {"price": 800}
    }

if 'hsa_resources' not in st.session_state:
    st.session_state.hsa_resources = {
        "عمل يد سباكه شبكه صرف صحي": {"daily_rate": 15000},
        "عمل يد تكسير وحفر": {"daily_rate": 10000},
        "عمل يد لحام وتركيب": {"daily_rate": 12000}
    }

# ============= العنوان الرئيسي =============
st.markdown("""
<div class="header-container">
    <h1>🏗️ نظام التكاليف الذكي - Smart HSA Cost Analyzer</h1>
    <p>نسخة Streamlit المتطورة | تحليل شامل لتكاليف المشاريع</p>
</div>
""", unsafe_allow_html=True)

# ============= الشريط الجانبي =============
with st.sidebar:
    st.title("⚙️ الإعدادات والخيارات")
    page = st.radio(
        "اختر القسم:",
        ["🏠 الرئيسية", "📊 التقديرية", "🏭 المواد والمواصفات", "⚖️ عروض الأسعار", "📈 التحليلات"]
    )

# ============= الصفحة الرئيسية =============
if page == "🏠 الرئيسية":
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <h3>📦 البنود المحفوظة</h3>
            <p style="font-size: 28px; color: var(--blue); font-weight: bold; margin: 10px 0;">
                {len(st.session_state.hsa_items)}
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="stat-card">
            <h3>🏢 المواد</h3>
            <p style="font-size: 28px; color: var(--blue); font-weight: bold; margin: 10px 0;">
                {len(st.session_state.hsa_materials)}
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="stat-card">
            <h3>👷 الموارد البشرية</h3>
            <p style="font-size: 28px; color: var(--blue); font-weight: bold; margin: 10px 0;">
                {len(st.session_state.hsa_resources)}
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.subheader("📂 استيراد البيانات من ملف إ��سل")
    uploaded_file = st.file_uploader("اختر ملف Excel (BOQ أو تحليل تفصيلي)", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ تم تحميل الملف بنجاح! {len(df)} صف")
            st.dataframe(df.head(10), use_container_width=True)
            
            # إضافة البنود
            for idx, row in df.iterrows():
                item = {
                    "id": f"BOQ_{datetime.now().timestamp()}_{idx}",
                    "itemName": str(row.iloc[0] if len(row) > 0 else ""),
                    "unit": str(row.iloc[1] if len(row) > 1 else ""),
                    "qty": float(row.iloc[2]) if len(row) > 2 else 0,
                    "directPrice": float(row.iloc[3]) if len(row) > 3 else 0,
                    "type": "مقطوعية_مباشرة"
                }
                if item["itemName"] and item["qty"] > 0:
                    st.session_state.hsa_items.append(item)
            
            st.success(f"✅ تم إضافة {len(st.session_state.hsa_items)} بند للنظام")
        except Exception as e:
            st.error(f"❌ خطأ في قراءة الملف: {e}")

# ============= صفحة التكلفة التقديرية =============
elif page == "📊 التقديرية":
    st.subheader("📍 البنود المسحوبة لتسعير الموقع")
    
    if len(st.session_state.hsa_items) == 0:
        st.info("⚠️ لم يتم استيراد أي بنود بعد. قم باستيراد ملف إكسل من القسم الرئيسي.")
    else:
        selected_items = []
        for item in st.session_state.hsa_items:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"**{item['itemName']}** - {item['qty']} {item.get('unit', 'وحدة')}")
            with col2:
                if st.checkbox("✓", value=True, key=item['id']):
                    selected_items.append(item)
        
        if st.button("🔄 احسب التكلفة التقديرية", use_container_width=True):
            total_cost = 0
            material_cost = 0
            labor_cost = 0
            
            for item in selected_items:
                item_total = item['qty'] * item['directPrice']
                material_cost += item_total * 0.70
                labor_cost += item_total * 0.30
                total_cost += item_total
            
            profit = (material_cost + labor_cost) * 0.15
            grand_total = material_cost + labor_cost + profit
            
            # عرض النتائج
            st.markdown(f"""
            <div class="price-highlight">
                💰 إجمالي التكلفة التقديرية: {grand_total:,.0f} ريال
            </div>
            """, unsafe_allow_html=True)
            
            # تفصيل التكاليف
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("المواد", f"{material_cost:,.0f}", delta=None)
            with col2:
                st.metric("العمالة", f"{labor_cost:,.0f}", delta=None)
            with col3:
                st.metric("الربح (15%)", f"{profit:,.0f}", delta=None)
            with col4:
                st.metric("الإجمالي", f"{grand_total:,.0f}", delta=None)
            
            # رسم بياني توزيع التكاليف
            fig = go.Figure(data=[go.Pie(
                labels=['المواد', 'العمالة', 'الربح'],
                values=[material_cost, labor_cost, profit],
                marker=dict(colors=['#1565C0', '#F57F17', '#2E7D32']),
                textposition='inside',
                textinfo='label+percent'
            )])
            fig.update_layout(
                title="توزيع التكاليف",
                font=dict(size=12)
            )
            st.plotly_chart(fig, use_container_width=True)

# ============= صفحة المواد والمواصفات =============
elif page == "🏭 المواد والمواصفات":
    st.subheader("🧱 إدارة المواد والمواصفات")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### المواد الحالية")
        for mat_name, mat_data in st.session_state.hsa_materials.items():
            st.write(f"**{mat_name}**: {mat_data['price']:,} ريال")
    
    with col2:
        st.markdown("### إضافة مادة جديدة")
        new_material = st.text_input("اسم المادة")
        new_price = st.number_input("السعر (ريال)", min_value=0, value=0)
        
        if st.button("➕ إضافة", use_container_width=True):
            if new_material and new_price > 0:
                st.session_state.hsa_materials[new_material] = {"price": new_price}
                st.success(f"✅ تمت إضافة {new_material}")
            else:
                st.error("❌ أكمل جميع الحقول")
    
    st.markdown("---")
    
    # تحديث الأسعار
    st.markdown("### تحديث أسعار المواد")
    material_choice = st.selectbox("اختر المادة", list(st.session_state.hsa_materials.keys()))
    new_price = st.number_input(f"السعر الجديد ل {material_choice}", min_value=0, value=int(st.session_state.hsa_materials[material_choice]['price']))
    
    if st.button("🔄 تحديث السعر", use_container_width=True):
        st.session_state.hsa_materials[material_choice]['price'] = new_price
        st.success("✅ تم تحديث السعر")

# ============= صفحة عروض الأسعار =============
elif page == "⚖️ عروض الأسعار":
    st.subheader("📊 مقارنة عطاءات المقاولين")
    
    # بيانات تجريبية للمقارنة
    quotations = {
        "التقديرية": 145000,
        "مؤسسة البناء": 160000,
        "الثناء للمقاولات": 138000,
        "النهضة والتطوير": 155000
    }
    
    fig = go.Figure(data=[
        go.Bar(
            x=list(quotations.keys()),
            y=list(quotations.values()),
            marker=dict(color=['#1565C0', '#F57F17', '#2E7D32', '#C62828'])
        )
    ])
    
    fig.update_layout(
        title="مقارنة العطاءات",
        yaxis_title="السعر (ريال)",
        xaxis_title="الجهة",
        font=dict(size=12),
        showlegend=False
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    col1, col2, col3, col4 = st.columns(4)
    for idx, (name, price) in enumerate(quotations.items()):
        with st.columns(4)[idx]:
            st.metric(name, f"{price:,.0f} ريال")

# ============= صفحة التحليلات =============
elif page == "📈 التحليلات":
    st.subheader("📊 التحليلات والتقارير")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### تطور أسعار المواسير (UPVC)")
        months = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو']
        prices = [4200, 4250, 4400, 4800, 5000, 5200]
        
        fig_line = go.Figure(data=[
            go.Scatter(x=months, y=prices, mode='lines+markers', 
                      marker=dict(color='#1565C0', size=8),
                      line=dict(color='#1565C0', width=3))
        ])
        fig_line.update_layout(title="مؤشر تغير الأسعار", xaxis_title="الشهر", yaxis_title="السعر")
        st.plotly_chart(fig_line, use_container_width=True)
    
    with col2:
        st.markdown("### توزيع المقاولين حسب السعر")
        fig_hist = px.histogram(
            x=['145000', '160000', '138000', '155000'],
            nbins=4,
            labels={'x': 'نطاق السعر'},
            title='توزيع العطاءات'
        )
        st.plotly_chart(fig_hist, use_container_width=True)

# ============= التذييل =============
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #607D8B; padding: 20px;">
    <p>🔧 <strong>نظام التكاليف الذكي</strong> - Smart HSA Cost Analyzer</p>
    <p>تم التطوير بواسطة: أحمد المنويّر | نسخة Streamlit المحسّنة</p>
</div>
""", unsafe_allow_html=True)
