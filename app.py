import streamlit as st
import pandas as pd
import datetime

# ==============================================================================
# 0. 页面全局配置
# ==============================================================================
st.set_page_config(
    page_title="酒店AI运营报告自动化统计系统",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义极简样式
st.markdown("""
    <style>
    .block-container {padding-top: 2rem; padding-bottom: 2rem;}
    h1, h2, h3 {margin-top: 0rem; font-weight: 600;}
    div[data-testid="stMetricValue"] {font-size: 2.2rem; font-weight: bold; color: #1E1E1E;}
    </style>
""", unsafe_allow_html=True)

# 纯净大标题
st.title("🏨 酒店 AI 运营报告数据自动化统计系统")
st.markdown("---")

# ==============================================================================
# 1. PART 1 & PART 2：数据源上传区
# ==============================================================================
st.subheader("📦 PART 1 & PART 2：全板块一体化账目")

col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    uploaded_file_call = st.file_uploader(
        "1. 上传【云总机通话详单】", 
        type=["xlsx", "xls"],
        key="call_uploader"
    )

with col_upload2:
    uploaded_file_ext = st.file_uploader(
        "2. 上传【分机号表】", 
        type=["xlsx", "xls"],
        key="ext_uploader"
    )

# 动态日期捕获提示条
if uploaded_file_call or uploaded_file_ext:
    st.info("📅 成功从文件名中捕获观测日期周期：0612-0618")
else:
    st.info("💡 请上传上方基础表格以激活数据平账模型")

st.markdown("---")

# ==============================================================================
# 2. PART 3：工单联动核算区 (加入状态锁，防止未确认前剧透数据)
# ==============================================================================
st.subheader("⚙️ PART 3：工单大盘自动核算")

# 配置选择框
col_ctrl1, col_ctrl2 = st.columns(2)

with col_ctrl1:
    hotel_name = st.selectbox(
        "酒店名称",
        options=["请选择酒店", "长沙延年檀香山酒店", "其他备选酒店1"],
        index=0  # 默认让用户主动选择，或者改成 1
    )

with col_ctrl2:
    date_range = st.date_input(
        "时间周期",
        value=(datetime.date(2026, 6, 12), datetime.date(2026, 6, 18)),
        min_value=datetime.date(2025, 1, 1),
        max_value=datetime.date(2027, 1, 1)
    )

# 🔔 新增：确认核算触发按钮（将主动权交还给用户）
run_calculation = st.button("🔍 确认酒店及日期，开始自动核算大盘", type="secondary")

# 创建大盘数字占位槽
col_m1, col_m2, col_m3 = st.columns(3)

# 逻辑分流：只有用户点击了按钮，并且选了正确的酒店，才吐出数据
if run_calculation:
    if hotel_name == "请选择酒店":
        st.warning("⚠️ 请先在上方选择具体的酒店名称后再进行核算。")
        with col_m1: st.metric(label="AI 服务工单总数", value="-")
        with col_m2: st.metric(label="超时处理工单数", value="-")
        with col_m3: st.metric(label="服务工单超时率", value="-")
    else:
        # 🔥 这里是未来对接后端真实计算逻辑的地方，目前先动态呈现核算结果
        with col_m1:
            st.metric(label="AI 服务工单总数", value="57 个", delta="↑")
        with col_m2:
            st.metric(label="超时处理工单数 (含'已超时'状态)", value="14 个", delta="↑")
        with col_m3:
            st.metric(label="服务工单超时率", value="24.56 %", delta=None)
else:
    # 💤 初始冷启动状态：显示横杠，保持静默
    with col_m1:
        st.metric(label="AI 服务工单总数", value="-")
    with col_m2:
        st.metric(label="超时处理工单数", value="-")
    with col_m3:
        st.metric(label="服务工单超时率", value="-")

st.markdown("---")

# ==============================================================================
# 3. 底栏状态与一键融合导出区
# ==============================================================================
# 只有在点击核算成功后，才亮起最终导出的绿灯
if run_calculation and hotel_name != "请选择酒店":
    st.success("🟩 数据平账模型匹配完毕，已就绪一键三表联动导出！")
    st.button(
        "📥 导出【0612-0618】三大板块融合版运营报告", 
        type="primary", 
        use_container_width=True
    )
else:
    st.error("🛑 请先完成 PART 3 的酒店和日期核算，以锁定制表大盘数据。")
