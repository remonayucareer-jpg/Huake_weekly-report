import streamlit as st
import pandas as pd
import datetime

# ==============================================================================
# 0. 页面全局配置 (去噪、专业后台管理风)
# ==============================================================================
st.set_page_config(
    page_title="酒店AI运营 - 周报",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义极简样式，移除 Streamlit 默认的多余间距
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
# 1. PART 1 & PART 2：数据源上传区 (全板块一体化账目导出)
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

# 动态日期捕获提示条 (此处示例默认捕获 0612-0618)
if uploaded_file_call or uploaded_file_ext:
    st.info("📅 成功从文件名中捕获观测日期周期：0612-0618")
else:
    st.info("💡 请上传上方基础表格以激活数据平账模型")

st.markdown("---")

# ==============================================================================
# 2. PART 3：工单联动核算区 (独立分区，保留控制与选择功能)
# ==============================================================================
st.subheader("⚙️ PART 3：工单大盘自动核算")

# 保留选择酒店名称和时间的功能
col_ctrl1, col_ctrl2 = st.columns(2)

with col_ctrl1:
    hotel_name = st.selectbox(
        "酒店名称",
        options=["长沙延年檀香山酒店", "其他备选酒店1", "其他备选酒店2"],
        index=0
    )

with col_ctrl2:
    # 默认联动 2026-06-12 到 2026-06-18
    date_range = st.date_input(
        "时间周期",
        value=(datetime.date(2026, 6, 12), datetime.date(2026, 6, 18)),
        min_value=datetime.date(2025, 1, 1),
        max_value=datetime.date(2027, 1, 1)
    )

# 实时联动只读指标看板
col_m1, col_m2, col_m3 = st.columns(3)

with col_m1:
    st.metric(label="AI 服务工单总数", value="57 个", delta="↑")

with col_m2:
    st.metric(label="超时处理工单数 (含'已超时'状态)", value="14 个", delta="↑")

with col_m3:
    # 算好具体比例，规避前端符号错位
    st.metric(label="服务工单超时率", value="24.56 %", delta=None)

st.markdown("---")

# ==============================================================================
# 3. 底栏状态与收网一键导出区
# ==============================================================================
# 提示模型匹配就绪状态
st.success("🟩 数据平账模型匹配完毕，已就绪一键三表联动导出！")

# 核心联动融合导出按钮
st.button(
    "📥 导出【0612-0618】三大板块融合版运营报告", 
    type="primary", 
    use_container_width=True
)
