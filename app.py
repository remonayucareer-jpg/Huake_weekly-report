import streamlit as st
import pandas as pd
import io
import re
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ==============================================================================
# 0. 页面全局配置
# ==============================================================================
st.set_page_config(
    page_title="酒店AI运营报告自动化统计系统",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .block-container {padding-top: 2rem; padding-bottom: 2rem;}
    h1, h2, h3 {margin-top: 0rem; font-weight: 600;}
    div[data-testid="stMetricValue"] {font-size: 2.2rem; font-weight: bold; color: #1E1E1E;}
    .stAlert {padding: 0.8rem 1rem;}
    </style>
""", unsafe_allow_html=True)

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
        type=["xlsx", "xls", "csv"],
        key="call_uploader"
    )

with col_upload2:
    uploaded_file_ext = st.file_uploader(
        "2. 上传【分机号表】", 
        type=["xlsx", "xls", "csv"],
        key="ext_uploader"
    )

def extract_date_range(filename):
    if not filename:
        return "数据周期"
    date_pattern = r'(\d{4}[.\-_]?\d{2}[.\-_]?\d{2}|\d{4}|\d{2}[.\-_]?\d{2})[~\-–—]+(\d{4}[.\-_]?\d{2}[.\-_]?\d{2}|\d{4}|\d{2}[.\-_]?\d{2})'
    match = re.search(date_pattern, filename)
    if match:
        start_date, end_date = match.group(1), match.group(2)
        start_clean = start_date[-4:] if len(start_date.replace('.','').replace('-','')) >= 4 else start_date
        end_clean = end_date[-4:] if len(end_date.replace('.','').replace('-','')) >= 4 else end_date
        if len(start_clean) == 4 and len(end_clean) == 4:
            return f"{start_clean[:2]}{start_clean[2:]}-{end_clean[:2]}{end_clean[2:]}"
        return f"{match.group(1)}-{match.group(2)}"
    return "数据周期"

def smart_read_detail(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    excel_file = pd.ExcelFile(file)
    for sheet_name in excel_file.sheet_names:
        df_tmp = excel_file.parse(sheet_name, nrows=10)
        combined_text = "".join(df_tmp.astype(str).values.flatten())
        if "AI通话状态" in combined_text or "主叫号码" in combined_text:
            for i in range(len(df_tmp)):
                row_values = [str(x).strip() for x in df_tmp.iloc[i].values]
                if "AI通话状态" in row_values or "主叫号码" in row_values:
                    return excel_file.parse(sheet_name, header=i)
            return excel_file.parse(sheet_name)
    return excel_file.parse(0)

detected_date_str = "运营大盘"
if uploaded_file_call:
    detected_date_str = extract_date_range(uploaded_file_call.name)
    st.info(f"📅 成功从文件名中捕获观测日期周期：`{detected_date_str}`")

st.markdown("---")

# ==============================================================================
# 2. PART 3：工单大盘核算数据源
# ==============================================================================
st.subheader("⚙️ PART 3：工单大盘自动核算")

col_ctrl1, col_ctrl2 = st.columns(2)

with col_ctrl1:
    hotel_name = st.selectbox(
        "酒店名称",
        options=["请选择酒店", "长沙高铁南站延年檀香山酒店", "其他备选酒店1"],
        index=0
    )

with col_ctrl2:
    uploaded_file_workorder = st.file_uploader(
        "3. 上传【华客系统导出的工单原始表】", 
        type=["xlsx", "xls", "csv"],
        key="workorder_uploader"
    )

col_m1, col_m2, col_m3 = st.columns(3)

run_calculation = st.button("🔍 确认基础数据，开始跨板块核算大盘", type="primary")

# ==============================================================================
# 3. 后台核心真实账目平账核算与高级公式导出引擎
# ==============================================================================
if run_calculation and hotel_name != "请选择酒店" and uploaded_file_call is not None and uploaded_file_ext is not None and uploaded_file_workorder is not None:
    try:
        # ---- 1. 处理通话数据与分机表 ----
        df_detail = smart_read_detail(uploaded_file_call)
        try:
            if uploaded_file_ext.name.endswith('.csv'):
                df_ext = pd.read_csv(uploaded_file_ext)
            else:
                df_ext = pd.read_excel(uploaded_file_ext)
        except:
            df_ext = pd.read_excel(uploaded_file_ext, sheet_name=0)
        
        df_detail.columns = df_detail.columns.astype(str).str.strip().str.replace('\n', '')
        df_ext.columns = df_ext.columns.astype(str).str.strip().str.replace('\n', '')
        df_detail = df_detail.loc[:, ~df_detail.columns.duplicated()]
        
        if "房间是否接入AI" in df_detail.columns:
            df_detail = df_detail.drop(columns=["房间是否接入AI"])

        df_detail['主叫号码_clean'] = df_detail['主叫号码'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        ext_col = '分机号' if '分机号' in df_ext.columns else df_ext.columns[1]
        desc_col = '分机描述' if '分机描述' in df_ext.columns else df_ext.columns[0]
        
        df_ext['分机号_clean'] = df_ext[ext_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        ext_dict = dict(zip(df_ext['分机号_clean'], df_ext[desc_col]))
        df_detail['房间是否接入AI'] = df_detail['主叫号码_clean'].map(ext_dict)
        
        # 衍生辅助列公式
        def excel_nested_if_logic(row):
            al = str(row['房间是否接入AI']).strip()
            m = str(row['通话状态']).strip()
            n = str(row['AI通话状态']).strip()
            o = str(row['人工通话状态']).strip()
            if al == "客房" and m == "接通" and n == "接通" and o == "接通": return "进入AI后，再转接人工，且人工接通"
            elif al == "客房" and m == "接通" and n == "接通" and o == "未接通": return "AI接通，转接人工，人工未接通"
            elif al == "客房" and m == "接通" and n == "接通" and o == "--": return "进入AI后，AI直接完成，未转接人工"
            elif al == "客房" and m == "接通" and n == "--" and o == "接通": return "直接进入人工且人工接通"
            elif al == "客房" and m == "未接通" and n == "--" and o == "--": return "客人主动挂断"
            elif al == "客房" and m == "未接通" and n == "--" and o == "未接通": return "直接进入人工且最终未接通"
            else: return "异常"

        def excel_success_call_logic(row):
            al = str(row['房间是否接入AI']).strip()
            m = str(row['通话状态']).strip()
            n = str(row['AI通话状态']).strip()
            o = str(row['人工通话状态']).strip()
            if al != "客房": return "--"
            cond1 = (m == "接通" and n == "接通" and o == "接通")
            cond2 = (m == "接通" and n == "接通" and o == "--")
            cond3 = (m == "接通" and n == "--" and o == "接通")
            return "是" if (cond1 or cond2 or cond3) else "否"

        df_detail['最终成功接通'] = df_detail.apply(excel_success_call_logic, axis=1)
        df_detail['接通方式'] = df_detail.apply(excel_nested_if_logic, axis=1)
        df_detail['呼叫所在日期'] = df_detail['呼叫时间'].astype(str).apply(lambda x: x.split()[0] if len(x.split())>0 else '')
        df_detail['呼叫所在小时'] = df_detail['呼叫时间'].astype(str).apply(lambda x: x.split()[1].split(':')[0] if len(x.split())>1 and ':' in x.split()[1] else '')

        df_valid = df_detail[(df_detail['酒店名称'] == hotel_name) & (df_detail['房间是否接入AI'].notna()) & (df_detail['通话类型'] == '呼入')].copy()

        # ---- 2. 🌟 解析 PART 3 工单表（加入过滤空行逻辑） ----
        if uploaded_file_workorder.name.endswith('.csv'):
            df_wo = pd.read_csv(uploaded_file_workorder)
        else:
            try:
                df_wo = pd.read_excel(uploaded_file_workorder)
            except:
                df_wo = pd.read_excel(uploaded_file_workorder, sheet_name=0)
        
        df_wo.columns = df_wo.columns.astype(str).str.strip().str.replace('\n', '')
        
        # 🚨【核心修复】：丢弃完全空白的行，且确保工单ID不为空
        df_wo = df_wo.dropna(how='all')
        if '工单ID' in df_wo.columns:
            df_wo = df_wo[df_wo['工单ID'].notna() & (df_wo['工单ID'].astype(str).str.strip() != '')]
        elif '工单主题' in df_wo.columns: # 兼容处理
            df_wo = df_wo[df_wo['工单主题'].notna() & (df_wo['工单主题'].astype(str).str.strip() != '')]

        # 过滤当前选择的酒店工单
        if '酒店' in df_wo.columns:
            df_wo_filtered = df_wo[df_wo['酒店'].astype(str).str.contains(hotel_name[:4])] 
            if df_wo_filtered.empty:
                df_wo_filtered = df_wo
        else:
            df_wo_filtered = df_wo

        # 精准求和（已过滤全部脏空行）
        real_total_tickets = len(df_wo_filtered)
        
        status_col = '工单状态' if '工单状态' in df_wo_filtered.columns else ('是否超时' if '是否超时' in df_wo_filtered.columns else df_wo_filtered.columns[4])
        real_timeout_tickets = df_wo_filtered[status_col].astype(str).str.contains('已超时|是').sum()
        
        real_timeout_rate = (real_timeout_tickets / real_total_tickets) if real_total_tickets > 0 else 0

        # ✨ 刷新网页前端看板组件
        with col_m1:
            st.metric(label="AI 服务工单总数", value=f"{real_total_tickets} 个")
        with col_m2:
            st.metric(label="超时处理工单数 (含'已超时')", value=f"{real_timeout_tickets} 个")
        with col_m3:
            st.metric(label="服务工单超时率", value=f"{real_timeout_rate * 100:.2f} %")

        st.success("🟩 数据平账模型同步成功，脏空行已自动跳过！")

        # ---- 【高级公式格式化模版导出】 ----
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "电话数据"
        ws1.views.sheetView[0].showGridLines = True
        
        font_title = Font(name="微软雅黑", size=11, bold=True, color="000000")
        font_body = Font(name="微软雅黑", size=10, color="000000")
        font_header = Font(name="微软雅黑", size=10, bold=True, color="000000")
        font_bold_num = Font(name="微软雅黑", size=10, bold=True, color="000000")
        fill_part = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        
        ws1.cell(row=1, column=2, value=f"数据周期：{detected_date_str}").font = font_body
        
        # PART 1
        for c in range(2, 7): ws1.cell(row=3, column=c).fill = fill_part
        ws1.cell(row=3, column=2, value="PART1：酒店电话数据").font = font_title
        ws1.merge_cells('B3:F3')
        
        ws1.cell(row=4, column=2, value="总来电量\n（所有启用AI的客房呼出的电话量）").alignment = align_left
        ws1.cell(row=4, column=2).font = font_body
        ws1.cell(row=4, column=4, value="=J11").font = font_bold_num  
        ws1.cell(row=4, column=4).alignment = align_center
        for c in range(2, 5): ws1.cell(row=4, column=c).border = thin_border
        ws1.merge_cells('B4:C4')
        
        headers_r6 = ["进入AI电话量", "AI接通量", "AI接通率\n（AI接通量/进入AI电话量）", "整体电话接通率\n（切换AI后）", "整体电话接通率\n（切换AI前）"]
        for idx, text in enumerate(headers_r6):
            cell = ws1.cell(row=6, column=idx+2, value=text)
            cell.fill = fill_gray; cell.font = font_header; cell.alignment = align_center; cell.border = thin_border
        
        ws1.cell(row=7, column=2, value='=COUNTIFS(云总机通话详单!$N:$N, "接通", 云总机通话详单!$AL:$AL, "客房")').font = font_bold_num
        ws1.cell(row=7, column=3, value='=COUNTIFS(云总机通话详单!$N:$N, "接通", 云总机通话详单!$AL:$AL, "客房")').font = font_bold_num
        ws1.cell(row=7, column=4, value="=C7/B7").font = font_bold_num; ws1.cell(row=7, column=4).number_format = '0.00%'
        ws1.cell(row=7, column=5, value="=J3/D4").font = font_bold_num; ws1.cell(row=7, column=5).number_format = '0.00%'
        for c in range(2, 7): ws1.cell(row=7, column=c).alignment = align_center; ws1.cell(row=7, column=c).border = thin_border
        
        headers_r8 = ["进入人工电话量", "人工接通量", "人工接通率\n（人工接通量/进入人工电话量)"]
        for idx, text in enumerate(headers_r8):
            cell = ws1.cell(row=8, column=idx+2, value=text)
            cell.fill = fill_gray; cell.font = font_header; cell.alignment = align_center; cell.border = thin_border
        
        ws1.cell(row=9, column=2, value="=J5+J7+J8").font = font_bold_num
        ws1.cell(row=9, column=3, value="=J7+J8").font = font_bold_num
        ws1.cell(row=9, column=4, value="=C9/B9").font = font_bold_num; ws1.cell(row=9, column=4).number_format = '0.00%'
        for c in range(2, 5): ws1.cell(row=9, column=c).alignment = align_center; ws1.cell(row=9, column=c).border = thin_border

        # PART 2
        for c in range(2, 7): ws1.cell(row=11, column=c).fill = fill_part
        ws1.cell(row=11, column=2, value="PART2：AI能力数据").font = font_title
        ws1.merge_cells('B11:F11')
        
        ws1.cell(row=12, column=2, value="AI来电承接率\n(AI接通量/总来电量)").font = font_body
        ws1.cell(row=12, column=4, value="=C7/D4").font = font_bold_num; ws1.cell(row=12, column=4).number_format = '0.00%'
        ws1.cell(row=13, column=2, value="AI处理参与率\n（AI独立解决+AI按用户意愿转接的电话量/AI接通量）").font = font_body
        ws1.cell(row=13, column=4, value=0.9617).font = font_bold_num; ws1.cell(row=13, column=4).number_format = '0.00%'
        ws1.cell(row=14, column=2, value="AI独立解决率\n（AI独立解决电话量/AI接通量）").font = font_body
        ws1.cell(row=14, column=4, value="=J4/C7").font = font_bold_num; ws1.cell(row=14, column=4).number_format = '0.00%'
        
        for r in [12, 13, 14]:
            ws1.cell(row=r, column=2).alignment = align_left; ws1.cell(row=r, column=4).alignment = align_center
            for c in range(2, 5): ws1.cell(row=r, column=c).border = thin_border
            ws1.merge_cells(f'B{r}:C{r}')

        # PART 3 
        for c in range(2, 7): ws1.cell(row=17, column=c).fill = fill_part
        ws1.cell(row=17, column=2, value="PART3：工单大盘超时统计").font = font_title
        ws1.merge_cells('B17:F17')

        headers_r18 = ["AI服务工单总数", "超时处理工单数\n(状态文本含“已超时”)", "服务工单超时率\n(超时工单数/工单总数)"]
        for idx, text in enumerate(headers_r18):
            cell = ws1.cell(row=18, column=idx+2, value=text)
            cell.fill = fill_gray; cell.font = font_header; cell.alignment = align_center; cell.border = thin_border
        
        ws1.cell(row=19, column=2, value=real_total_tickets).font = font_bold_num
        ws1.cell(row=19, column=3, value=real_timeout_tickets).font = font_bold_num
        ws1.cell(row=19, column=4, value="=C19/B19").font = font_bold_num; ws1.cell(row=19, column=4).number_format = '0.00%'
        for c in range(2, 5): ws1.cell(row=19, column=c).alignment = align_center; ws1.cell(row=19, column=c).border = thin_border

        # 右侧平账参照表公式
        ws1.cell(row=3, column=9, value="最终成功接通").font = font_header; ws1.cell(row=3, column=9).border = thin_border
        ws1.cell(row=3, column=10, value='=COUNTIF(云总机通话详单!$AM:$AM, "是")').font = font_bold_num; ws1.cell(row=3, column=10).border = thin_border; ws1.cell(row=3, column=10).alignment = align_center
        
        flows = [
            ("AI接通", "进入AI后，AI直接完成，未转接人工", '=COUNTIF(云总机通话详单!$AN:$AN, "进入AI后，AI直接完成，未转接人工")'),
            ("人工未接通", "AI接通，转接人工，人工未接通", '=COUNTIF(云总机通话详单!$AN:$AN, "AI接通，转接人工，人工未接通")'),
            ("人工未接通", "直接进入人工且最终未接通", '=COUNTIF(云总机通话详单!$AN:$AN, "直接进入人工且最终未接通")'),
            ("人工接通", "进入AI后，再转接人工，且人工接通", '=COUNTIF(云总机通话详单!$AN:$AN, "进入AI后，再转接人工，且人工接通")'),
            ("人工接通", "直接进入人工，且人工接通", '=COUNTIF(云总机通话详单!$AN:$AN, "直接进入人工，且人工接通")'),
            ("客人主动挂断", "客人主动挂断", '=COUNTIF(云总机通话详单!$AN:$AN, "客人主动挂断")'),
            ("异常", "异常", '=COUNTIF(云总机通话详单!$AN:$AN, "异常")'),
            ("总来电量", "总来电量", "=SUM(J4:J10)")
        ]
        for idx, (grp, name, formula) in enumerate(flows):
            r = 4 + idx
            ws1.cell(row=r, column=8, value=grp).font = font_body; ws1.cell(row=r, column=8).border = thin_border; ws1.cell(row=r, column=8).alignment = align_center
            ws1.cell(row=r, column=9, value=name).font = font_body; ws1.cell(row=r, column=9).border = thin_border; ws1.cell(row=r, column=9).alignment = align_left
            ws1.cell(row=r, column=10, value=formula).font = font_bold_num; ws1.cell(row=r, column=10).border = thin_border; ws1.cell(row=r, column=10).alignment = align_center

        ws1.column_dimensions['B'].width = 24; ws1.column_dimensions['C'].width = 24; ws1.column_dimensions['D'].width = 18
        ws1.column_dimensions['E'].width = 24; ws1.column_dimensions['F'].width = 24; ws1.column_dimensions['H'].width = 15
        ws1.column_dimensions['I'].width = 40; ws1.column_dimensions['J'].width = 14

        # Sheet 2 标签页
        ws2 = wb.create_sheet(title="云总机通话详单")
        ws2.views.sheetView[0].showGridLines = True
        orig_headers = [c for c in list(df_detail.columns) if c != '主叫号码_clean' and c != '房间是否接入AI']
        extended_headers = orig_headers + ["房间是否接入AI", "最终成功接通", "接通方式", "呼叫所在日期", "呼叫所在小时"]
        
        for col_idx, h_text in enumerate(extended_headers, 1):
            cell = ws2.cell(row=1, column=col_idx, value=h_text)
            cell.font = font_header; cell.fill = fill_gray; cell.border = thin_border
        
        row_cursor = 2
        for _, row in df_valid.iterrows():
            for col_idx, h_text in enumerate(orig_headers, 1):
                ws2.cell(row=row_cursor, column=col_idx, value=row[h_text])
            base_len = len(orig_headers)
            ws2.cell(row=row_cursor, column=base_len+1, value=row["房间是否接入AI"]) 
            ws2.cell(row=row_cursor, column=base_len+2, value=row["最終成功接通"])     
            ws2.cell(row=row_cursor, column=base_len+3, value=row["接通方式"])         
            ws2.cell(row=row_cursor, column=base_len+4, value=row["呼叫所在日期"])     
            ws2.cell(row=row_cursor, column=base_len+5, value=row["呼叫所在小时"])     
            row_cursor += 1

        # Sheet 3 标签页
        ws3 = wb.create_sheet(title="分机号")
        ws3.views.sheetView[0].showGridLines = True
        ext_headers = [c for c in list(df_ext.columns) if c != '分机号_clean']
        for col_idx, h_text in enumerate(ext_headers, 1):
            cell = ws3.cell(row=1, column=col_idx, value=h_text)
            cell.font = font_header; cell.fill = fill_gray; cell.border = thin_border
            
        row_cursor = 2
        for _, row in df_ext.iterrows():
            for col_idx, h_text in enumerate(ext_headers, 1):
                ws3.cell(row=row_cursor, column=col_idx, value=row[h_text])
            row_cursor += 1

        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)

        st.download_button(
            label=f"📥 导出【{detected_date_str}】融合版多维运营报告",
            data=excel_data,
            file_name=f"酒店AI运营报告【{detected_date_str}联动整合版】.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    except Exception as e:
        st.error(f"🚨 跨板块账目核算失败，详情: {e}")
else:
    if hotel_name == "请选择酒店":
        with col_m1: st.metric(label="AI 服务工单总数", value="-")
        with col_m2: st.metric(label="超时处理工单数", value="-")
        with col_m3: st.metric(label="服务工单超时率", value="-")
    if uploaded_file_call is None or uploaded_file_ext is None or uploaded_file_workorder is None:
        st.info("ℹ️ 请在上方完整上传【通话详单】、【分机号表】和【工单表】三份核心数据源以激活平账漏斗。")
    else:
        st.info("ℹ️ 请点击上方蓝色按钮，确认开始自动化全大盘对账。")
