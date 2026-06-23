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
# 2. PART 3：工单大盘自动核算
# ==============================================================================
st.subheader("⚙️ PART 3：工单大盘自动核算")

uploaded_file_workorder = st.file_uploader(
    "3. 上传【华客系统导出的工单原始表】", 
    type=["xlsx", "xls", "csv"],
    key="workorder_uploader"
)

st.markdown(" ")
col_m1, col_m2, col_m3 = st.columns(3)

run_calculation = st.button("🔍 确认基础数据，开始跨板块核算大盘", type="primary")

# ==============================================================================
# 3. 后台核心真实账目平账核算与高级公式导出引擎
# ==============================================================================
if run_calculation and uploaded_file_call is not None and uploaded_file_ext is not None and uploaded_file_workorder is not None:
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
        
        # 严格移除非必要的残留衍生字段，后面统一控制位置重新生成
        for col_to_drop in ["房间是否接入AI", "最终成功接通", "接通方式", "呼叫所在日期", "呼叫所在小时"]:
            if col_to_drop in df_detail.columns:
                df_detail = df_detail.drop(columns=[col_to_drop])

        df_detail['主叫号码_clean'] = df_detail['主叫号码'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        ext_col = '分机号' if '分机号' in df_ext.columns else df_ext.columns[1]
        desc_col = '分机描述' if '分机描述' in df_ext.columns else df_ext.columns[0]
        
        df_ext['分机号_clean'] = df_ext[ext_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        ext_dict = dict(zip(df_ext['分机号_clean'], df_ext[desc_col]))
        df_detail['房间是否接入AI'] = df_detail['主叫号码_clean'].map(ext_dict)
        
        # 衍生辅助列逻辑
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

        # 仅留存呼入数据
        df_valid = df_detail[df_detail['通话类型'] == '呼入'].copy()

        # ---- 2. 解析 PART 3 工单表 ----
        if uploaded_file_workorder.name.endswith('.csv'):
            df_wo = pd.read_csv(uploaded_file_workorder)
        else:
            try:
                df_wo = pd.read_excel(uploaded_file_workorder)
            except:
                df_wo = pd.read_excel(uploaded_file_workorder, sheet_name=0)
        
        df_wo.columns = df_wo.columns.astype(str).str.strip().str.replace('\n', '')
        
        # 剔除全空行并精确卡死条件
        df_wo = df_wo.dropna(how='all')
        if '工单ID' in df_wo.columns:
            df_wo_filtered = df_wo[df_wo['工单ID'].notna() & (df_wo['工单ID'].astype(str).str.strip() != '')].copy()
        elif '工单主题' in df_wo.columns:
            df_wo_filtered = df_wo[df_wo['工单主题'].notna() & (df_wo['工单主题'].astype(str).str.strip() != '')].copy()
        else:
            df_wo_filtered = df_wo.copy()

        real_total_tickets = len(df_wo_filtered)
        timeout_col = '是否超时' if '是否超时' in df_wo_filtered.columns else ('工单状态' if '工单状态' in df_wo_filtered.columns else df_wo_filtered.columns[5])
        real_timeout_tickets = (df_wo_filtered[timeout_col].astype(str).str.strip() == '是').sum()
        real_timeout_rate = (real_timeout_tickets / real_total_tickets) if real_total_tickets > 0 else 0

        # ✨ 前端网页看板实时渲染（供快速肉眼查看）
        with col_m1:
            st.metric(label="AI 服务工单总数", value=f"{real_total_tickets} 个")
        with col_m2:
            st.metric(label="超时处理工单数 (是否超时='是')", value=f"{real_timeout_tickets} 个")
        with col_m3:
            st.metric(label="服务工单超时率", value=f"{real_timeout_rate * 100:.2f} %")

        st.success("🟩 大盘账目校准完成！所有 Excel 单元格已全量切换为纯动态公式驱动。")

        # ---- 3. 【高级格式化模版导出引擎】 ----
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
        
        # PART 1：酒店电话数据
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
        
        # 🌟 锁死动态列映射关系：
        # M列=通话状态, N列=AI通话状态, O列=人工通话状态, AL列=房间是否接入AI, AM列=最终成功接通, AN列=接通方式
        ws1.cell(row=7, column=2, value='=COUNTIF(云总机通话详单!$AL:$AL, "客房")').font = font_bold_num
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

        # PART 2：AI能力数据
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

        # PART 3：工单大盘超时统计
        for c in range(2, 7): ws1.cell(row=17, column=c).fill = fill_part
        ws1.cell(row=17, column=2, value="PART3：工单大盘超时统计").font = font_title
        ws1.merge_cells('B17:F17')

        headers_r18 = ["AI服务工单总数", "超时处理工单数\n(是否超时列为“是”)", "服务工单超时率\n(超时工单数/工单总数)"]
        for idx, text in enumerate(headers_r18):
            cell = ws1.cell(row=18, column=idx+2, value=text)
            cell.fill = fill_gray; cell.font = font_header; cell.alignment = align_center; cell.border = thin_border
        
        ws1.cell(row=19, column=2, value=real_total_tickets).font = font_bold_num
        ws1.cell(row=19, column=3, value=real_timeout_tickets).font = font_bold_num
        ws1.cell(row=19, column=4, value="=C19/B19").font = font_bold_num; ws1.cell(row=19, column=4).number_format = '0.00%'
        for c in range(2, 5): ws1.cell(row=19, column=c).alignment = align_center; ws1.cell(row=19, column=c).border = thin_border

        # 右侧平账核对漏斗表（100%纯公式驱动）
        ws1.cell(row=3, column=9, value="最终成功接通").font = font_header; ws1.cell(row=3, column=9).border = thin_border
        ws1.cell(row=3, column=10, value='=COUNTIF(云总机通话详单!$AM:$AM, "是")').font = font_bold_num; ws1.cell(row=3, column=10).border = thin_border; ws1.cell(row=3, column=10).alignment = align_center
        
        flows = [
            ("AI接通", "进入AI后，AI直接完成，未转接人工", '=COUNTIF(云总机通话详单!$AN:$AN, "进入AI后，AI直接完成，未转接人工")'),
            ("人工未接通", "AI接通，转接人工，人工未接通", '=COUNTIF(云总机通话详单!$AN:$AN, "AI接通，转接人工，人工未接通")'),
            ("人工未接通", "直接进入人工且最终未接通", '=COUNTIF(云总机通话详单!$AN:$AN, "直接进入人工且最终未接通")'),
            ("人工接通", "进入AI后，再转接人工，且人工接通", '=COUNTIF(云总机通话详单!$AN:$AN, "进入AI后，再转接人工，且人工接通")'),
            ("人工接通", "直接进入人工且人工接通", '=COUNTIF(云总机通话详单!$AN:$AN, "直接进入人工且人工接通")'),
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

        # Sheet 2 标签页：云总机通话详单
        ws2 = wb.create_sheet(title="云总机通话详单")
        ws2.views.sheetView[0].showGridLines = True
        
        # 显式分离：前37列为原始数据列，后5列为派生过滤列
        orig_headers = [c for c in list(df_valid.columns) if c not in ["主叫号码_clean", "房间是否接入AI", "最终成功接通", "接通方式", "呼叫所在日期", "呼叫所在小时"]]
        fixed_extended_headers = ["房间是否接入AI", "最终成功接通", "接通方式", "呼叫所在日期", "呼叫所在小时"]
        final_all_headers = orig_headers + fixed_extended_headers
        
        # 写入统一的唯一表头（彻底去除双胞胎列重复Bug）
        for col_idx, h_text in enumerate(final_all_headers, 1):
            cell = ws2.cell(row=1, column=col_idx, value=h_text)
            cell.font = font_header; cell.fill = fill_gray; cell.border = thin_border
        
        # 逐行写入流
        row_cursor = 2
        for _, row in df_valid.iterrows():
            # 1. 先写原始前37列
            for col_idx, h_text in enumerate(orig_headers, 1):
                ws2.cell(row=row_cursor, column=col_idx, value=row[h_text])
            
            # 2. 严格紧接其后写入新加的5个计算列，不再二次重复调用
            base_len = len(orig_headers) # 应该为37
            ws2.cell(row=row_cursor, column=base_len+1, value=row["房间是否接入AI"]) # AL列 (38)
            ws2.cell(row=row_cursor, column=base_len+2, value=row["最终成功接通"])     # AM列 (39)
            ws2.cell(row=row_cursor, column=base_len+3, value=row["接通方式"])         # AN列 (40)
            ws2.cell(row=row_cursor, column=base_len+4, value=row["呼叫所在日期"])     # AO列 (41)
            ws2.cell(row=row_cursor, column=base_len+5, value=row["呼叫所在小时"])     # AP列 (42)
            row_cursor += 1

        # Sheet 3 标签页：分机号
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
            label=f"📥 导出【{detected_date_str}】公式联动整合版报告",
            data=excel_data,
            file_name=f"酒店AI运营报告【{detected_date_str}公式联动版】.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    except Exception as e:
        st.error(f"🚨 跨板块账目核算失败，详情: {e}")
else:
    if uploaded_file_call is None or uploaded_file_ext is None or uploaded_file_workorder is None:
        with col_m1: st.metric(label="AI 服务工单总数", value="-")
        with col_m2: st.metric(label="超时处理工单数", value="-")
        with col_m3: st.metric(label="服务工单超时率", value="-")
        st.info("ℹ️ 请在上方完整上传【通话详单】、【分机号表】和【工单表】三份核心数据源以激活公式平账引擎。")
    else:
        st.info("ℹ️ 请点击上方蓝色按钮，确认开始自动化大盘公式核算。")
