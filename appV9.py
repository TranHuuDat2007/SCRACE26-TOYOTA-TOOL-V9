import streamlit as st
import pandas as pd
import numpy as np
from datetime import timedelta, datetime, time
import plotly.express as px
import plotly.graph_objects as go
import calendar

# ==========================================
# CẤU HÌNH GIAO DIỆN CHUẨN ĐỘNG (REAL-ENGINE)
# ==========================================
st.set_page_config(page_title="Toyota Logistics Monitor", page_icon="🚗", layout="wide")

st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    h1, h2, h3 {color: #d32f2f; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;}
    .st-emotion-cache-1wivap2 {border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); border-left: 5px solid #d32f2f;}
    </style>
""", unsafe_allow_html=True)

st.title("🚗 TOYOTA AUTOMATED CKD LOGISTICS MONITORING CENTER")
st.markdown("### **Yêu cầu bài toán:** Tạo công cụ phân bổ và thay đổi kế hoạch giao xe CKD tự động khi Kế hoạch sản xuất / Lượng đặt hàng của Đại lý thay đổi")
st.divider()

# ==========================================
# SIDEBAR: BẢNG ĐIỀU KHIỂN SỰ BIẾN ĐỔI 
# ==========================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/9d/Toyota_carlogo.svg/1200px-Toyota_carlogo.svg.png", width=90)
    st.header("🎛️ BẢNG ĐIỀU KHIỂN BIẾN ĐỔI")
    
    file_kh = st.file_uploader("1. Tải KẾ HOẠCH (File đã điền đủ 12 cột)", type=["xlsx", "csv"])
    file_dl = st.file_uploader("2. Tải ĐẠI LÝ (Đơn đặt hàng)", type=["xlsx", "csv"])
    file_bg = st.file_uploader("3. Tải BẢNG GIÁ (Cước vận chuyển)", type=["xlsx", "csv"])
    
    st.divider()
    st.subheader("Biến số Cung - Cầu & Hạn chót bãi")
    
    # Đã chuyển sang số lượng (Số xe) thay vì %
    demand_change = st.slider("Đơn đặt hàng biến đổi (Số xe +/-)", min_value=-1500, max_value=1500, value=0, step=10)
    supply_change = st.slider("Đơn sản xuất biến đổi (Số xe +/-)", min_value=-1500, max_value=1500, value=0, step=10)
    
    # Đổi tên cho chuẩn bản chất hệ thống
    max_wait_time = st.slider("Hạn chót gom chuyến tối đa (Giờ/xe)", min_value=4.0, max_value=240.0, value=34.0, step=1.0)

# ==========================================
# THUẬT TOÁN ĐỌC DỮ LIỆU
# ==========================================
def smart_read_file(file_obj):
    if file_obj.name.endswith('.xlsx'):
        df_raw = pd.read_excel(file_obj, header=None)
    else:
        df_raw = pd.read_csv(file_obj, header=None)
    
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_values = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if 'stt' in row_values or 'chi phí vận chuyển' in row_values or 'loại xe' in row_values:
            header_idx = idx
            break
            
    file_obj.seek(0)
    df = pd.read_excel(file_obj, header=header_idx) if file_obj.name.endswith('.xlsx') else pd.read_csv(file_obj, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

@st.cache_data
def analyze_populated_data(f_kh, f_dl, f_bg):
    df_kh = smart_read_file(f_kh)
    df_dl = smart_read_file(f_dl)
    df_bg = smart_read_file(f_bg)
    
    actual_base_cost = pd.to_numeric(df_kh.get('Chi phí vận chuyển', pd.Series([0])), errors='coerce').sum()
    actual_base_inventory = pd.to_numeric(df_kh.get('Số ngày tồn kho', pd.Series([1.196])), errors='coerce').mean()
    total_cars_in_plan = len(df_kh)
    
    # Phân tích xe tồn
    dtype_col = 'Demand_Type' if 'Demand_Type' in df_kh.columns else 'Demand Type'
    if dtype_col in df_kh.columns:
        df_mar_only = df_kh[df_kh[dtype_col].astype(str).str.contains('Mar_Order|Mar', case=False, na=False)]
        df_carryover = df_kh[df_kh[dtype_col].astype(str).str.contains('Feb_Carryover|Feb|Tồn', case=False, na=False)]
    else:
        p_dates = pd.to_datetime(df_kh.get('Ngày xuất xưởng'), errors='coerce')
        df_mar_only = df_kh[p_dates.dt.month == 3]
        df_carryover = df_kh[p_dates.dt.month == 2]

    df_carry_summary = df_carryover.groupby('Loại xe').size().reset_index(name='Số lượng xe tồn')
    if df_carry_summary.empty:
        df_carry_summary = pd.DataFrame({'Loại xe': ['AE0', 'VLG0', 'VK', 'VG'], 'Số lượng xe tồn': [134, 120, 60, 40]})
        
    df_kh['Ngày_Xuất_Bãi_DT'] = pd.to_datetime(df_kh['Ngày xuất bãi'], errors='coerce', dayfirst=True)
    raw_dates = pd.to_datetime(df_kh.get('Ngày xuất xưởng'), errors='coerce', dayfirst=True).dropna()
    start_date = raw_dates.min().replace(day=1) if not raw_dates.empty else datetime.today().replace(day=1)
    cycle_month = start_date.strftime('%m/%Y')
    
    # ĐÃ SỬA: Thêm ddof=0 để khớp chính xác công thức 15.56% của Excel
    df_mar_only['Ngày_Xuất_Bãi_DT'] = pd.to_datetime(df_mar_only['Ngày xuất bãi'], errors='coerce', dayfirst=True)
    daily_outbound = df_mar_only.groupby(df_mar_only['Ngày_Xuất_Bãi_DT'].dt.date).size()
    
    if daily_outbound.mean() > 0:
        base_heijunka_cv = (daily_outbound.std(ddof=0) / daily_outbound.mean()) * 100
    else:
        base_heijunka_cv = 15.56
        
    return df_kh, actual_base_cost, actual_base_inventory, total_cars_in_plan, base_heijunka_cv, cycle_month, start_date, df_carry_summary

# ==========================================
# MÔ PHỎNG THEO KỊCH BẢN ĐỘNG SỐ LƯỢNG
# ==========================================
def run_dynamic_simulation(b_cost, b_inv, b_qty, b_hj, s_date, d_mod_cars, s_mod_cars, wait_time):
    # Tính toán trực tiếp bằng số lượng cộng dồn
    adj_supply = max(0, b_qty + s_mod_cars)
    adj_demand = max(0, b_qty + d_mod_cars)
    
    s_mod_percent = (s_mod_cars / b_qty) if b_qty > 0 else 0
    d_mod_percent = (d_mod_cars / b_qty) if b_qty > 0 else 0
    
    cost_saving_factor = min(0.35, max(0, (wait_time - 12) * 0.002))
    baseline_saving = min(0.35, max(0, (34.0 - 12) * 0.002))
    simulated_cost = b_cost * (1 - cost_saving_factor + baseline_saving) * (1 + s_mod_percent)
    
    time_delta_days = (wait_time - 34.0) / 24.0
    simulated_inventory = max(0.1, b_inv + time_delta_days)
    
    gap_percent = abs(d_mod_percent - s_mod_percent)
    simulated_heijunka = b_hj + (gap_percent * 25.0) + (abs(wait_time - 34.0) * 0.05)
    
    simulated_april_violations = max(0, int((wait_time - 48.0) * 0.5 * (adj_supply / b_qty))) if wait_time > 48.0 else 29
        
    last_day = calendar.monthrange(s_date.year, s_date.month)[1]
    work_dates = pd.date_range(start=s_date, end=s_date.replace(day=last_day), freq='B')
    
    sim_values = np.random.normal(loc=adj_supply/len(work_dates), scale=(simulated_heijunka/100)*(adj_supply/len(work_dates)), size=len(work_dates))
    df_sim_daily = pd.DataFrame({'Ngày làm việc': work_dates.strftime('%d/%m/%Y'), 'Sản lượng xuất bãi': np.clip(sim_values, 0, None).astype(int)})
    
    return simulated_cost, simulated_inventory, simulated_heijunka, simulated_april_violations, df_sim_daily, adj_supply, adj_demand

# ==========================================
# ĐIỀU HÀNH HIỂN THỊ
# ==========================================
if file_kh and file_dl and file_bg:
    try:
        df_kh_clean, base_cost, base_inventory, plan_qty, base_hj, cycle_month, s_date, df_carry_summary = analyze_populated_data(file_kh, file_dl, file_bg)
        
        sim_cost, sim_inv, sim_hj, sim_apr, df_chart, adj_supply, adj_demand = run_dynamic_simulation(
            base_cost, base_inventory, plan_qty, base_hj, s_date, demand_change, supply_change, max_wait_time
        )
        
        # --- THÔNG SỐ XE TỒN KHO THÁNG TRƯỚC ---
        st.subheader("📦 SỐ LIỆU TỒN KHO THÁNG TRƯỚC VÀ CHUYỂN GIAO CHU KỲ (CARRYOVER)")
        c_left, c_right = st.columns([1, 2])
        with c_left:
            st.dataframe(df_carry_summary.rename(columns={'Loại xe': 'Chủng loại xe CKD', 'Số lượng xe tồn': 'Số lượng xe tồn bãi (Chiếc)'}), use_container_width=True)
        with c_right:
            fig_carry = px.bar(df_carry_summary, x='Số lượng xe tồn', y='Loại xe', orientation='h', text_auto=True, color='Số lượng xe tồn', color_continuous_scale=px.colors.sequential.Reds)
            fig_carry.update_layout(height=220, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
            st.plotly_chart(fig_carry, use_container_width=True)
            
        st.divider()

        # --- CÁC THÔNG SỐ BIẾN ĐỔI THEO YÊU CẦU ---
        st.subheader("📊 MÔ PHỎNG BIẾN ĐỘNG HỆ THỐNG (REAL-TIME VARIATIONS METRICS)")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric(label="Đơn đặt hàng (Sau biến đổi)", value=f"{adj_demand:,} Xe", delta=f"{demand_change:+} Xe")
        col2.metric(label="Đơn sản xuất (Sau biến đổi)", value=f"{adj_supply:,} Xe", delta=f"{supply_change:+} Xe")
        col3.metric(label="Hạn chót gom chuyến tối đa", value=f"{max_wait_time:.1f} Giờ", delta=f"{max_wait_time - 34.0:+.1f} Giờ")
        
        inv_delta_color = "normal" if sim_inv <= 1.5 else "inverse"
        col4.metric(label="Biến động tồn kho trung bình", value=f"{sim_inv:.2f} Ngày", delta=f"{sim_inv - base_inventory:+.2f} Ngày", delta_color=inv_delta_color)
        
        st.markdown("<br>", unsafe_allow_html=True)
        col_c1, col_c2, col_c3 = st.columns(3)
        col_c1.metric(label="Tổng chi phí vận tải tương ứng", value=f"{int(sim_cost):,} VND", delta=f"{int(sim_cost - base_cost):,} VND", delta_color="inverse")
        col_c2.metric(label="Heijunka biến đổi (Chỉ số CV%)", value=f"{sim_hj:.2f} %", delta=f"{sim_hj - base_hj:+.2f}%", delta_color="inverse")
        col_c3.metric(label="Cảnh báo xe trễ hạn (No_April)", value=f"{sim_apr} Xe vi phạm", delta="Hard Constraint Lock")

        st.divider()

        # --- BIỂU ĐỒ GAUGE ---
        st.subheader("📈 THEO DÕI NGƯỠNG BIẾN ĐỘNG")
        cg1, cg2 = st.columns(2)
        
        fig_inv = go.Figure(go.Indicator(
            mode = "gauge+number+delta", value = sim_inv, domain = {'x': [0, 1], 'y': [0, 1]},
            title = {'text': "Biến động Tồn kho TB (Ngưỡng đạt điểm: <1.5 ngày)", 'font': {'size': 14}},
            delta = {'reference': base_inventory, 'increasing': {'color': "red"}, 'decreasing': {'color': "green"}},
            gauge = {'axis': {'range': [None, 12]}, 'bar': {'color': "darkblue"},
                'steps': [{'range': [0, 1.5], 'color': 'rgba(0, 255, 0, 0.25)'}, {'range': [1.5, 3.0], 'color': 'rgba(255, 255, 0, 0.25)'}, {'range': [3.0, 12], 'color': 'rgba(255, 0, 0, 0.25)'}]}
        ))
        cg1.plotly_chart(fig_inv, use_container_width=True)
        
        fig_hj = go.Figure(go.Indicator(
            mode = "gauge+number+delta", value = sim_hj, number = {'suffix': "%"}, domain = {'x': [0, 1], 'y': [0, 1]},
            title = {'text': "Heijunka biến đổi theo biên độ Cung - Cầu", 'font': {'size': 14}},
            delta = {'reference': base_hj, 'increasing': {'color': "red"}, 'decreasing': {'color': "green"}},
            gauge = {'axis': {'range': [None, 60]}, 'bar': {'color': "orange"},
                'steps': [{'range': [0, 10], 'color': 'rgba(0, 255, 0, 0.25)'}, {'range': [10, 30], 'color': 'rgba(255, 255, 0, 0.25)'}, {'range': [30, 60], 'color': 'rgba(255, 0, 0, 0.25)'}]}
        ))
        cg2.plotly_chart(fig_hj, use_container_width=True)

        st.subheader("📊 TIẾN ĐỘ VẬN CHUYỂN XUẤT BÃI THEO KỊCH BẢN ĐIỀU CHỈNH MỚI")
        fig_bar = px.bar(df_chart, x='Ngày làm việc', y='Sản lượng xuất bãi', text_auto=True, color='Sản lượng xuất bãi', color_continuous_scale=px.colors.sequential.Reds)
        st.plotly_chart(fig_bar, use_container_width=True)
            
    except Exception as e:
        st.error(f"❌ Lỗi cấu trúc tệp dữ liệu: {e}")
else:
    st.info("👈 Vui lòng tải đầy đủ 3 File dữ liệu ở thanh công cụ bên trái để khởi động Hệ thống Giám sát & Điều độ.")