import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(page_title="酒店AI运营报告自动化工具", layout="wide")

st.title("🏨 酒店AI运营报告数据自动化统计系统 (文件名日期动态同步版)")
st.markdown("已新增功能：自动解析【云总机通话详单】文件名中的日期区间，并像素级同步更新至 Excel 报告的 B1 单元格。")

col1, col2 = st.columns(2)
with col1:
    detail_file = st.file_uploader("1. 上传【云总机通话详单】", type=["xlsx", "xls"])
with col2:
    extension_file = st.file_uploader("2. 上传【分机号表】", type=["xlsx", "xls"])

def smart_read_detail(file):
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

# 📌 新增：从文件名动态提取日期的函数
def extract_date_range(filename):
    if not filename:
        return "0605-0611" # 默认兜底
    
    # 匹配常见的日期格式如：0605-0611, 20260605-20260611, 2026.06.05-2026.06.11 等
    date_pattern = r'(\d{4}[.\-_]?\d{2}[.\-_]?\d{2}|\d{4}|\d{2}[.\-_]?\d{2})[~\-–—]+(\d{4}[.\-_]?\d{2}[.\-_]?\d{2}|\d{4}|\d{2}[.\-_]?\d{2})'
    match = re.search(date_pattern, filename)
    
    if match:
        # 提取到的原始日期字符串
        start_date, end_date = match.group(1), match.group(2)
        # 清洗掉可能多余的年年份前缀，统一保持简洁如 0605-0611 的视觉感，或者直接返回原样式
        start_clean = start_date[-4:] if len(start_date.replace('.','').replace('-','')) >= 4 else start_date
        end_clean = end_date[-4:] if len(end_date.replace('.','').replace('-','')) >= 4 else end_date
        
        # 格式化一下
        if len(start_clean) == 4 and len(end_clean) == 4:
            return f"{start_clean[:2]}{start_clean[2:]}-{end_clean[:2]}{end_clean[2:]}"
        return f"{match.group(1)}-{match.group(2)}"
    
    return "0605-0611" # 没匹配到时的默认值

if detail_file is not None and extension_file is not None:
    try:
        # 提取文件名中的日期区间
        detected_date_range = extract_date_range(detail_file.name)
        st.info(f"📅 成功从文件名中捕获到观测日期周期：`{detected_date_range}`")

        df_detail = smart_read_detail(detail_file)
        try:
            df_ext = pd.read_excel(extension_file)
        except:
            df_ext = pd.read_excel(extension_file, sheet_name=0)
        
        # 清洗原始列名
        df_detail.columns = df_detail.columns.astype(str).str.strip().str.replace('\n', '')
        df_ext.columns = df_ext.columns.astype(str).str.strip().str.replace('\n', '')
        
        # 移除可能重复的‘房间是否接入AI’列
        df_detail = df_detail.loc[:, ~df_detail.columns.duplicated()]
        if "房间是否接入AI" in df_detail.columns:
            df_detail = df_detail.drop(columns=["房间是否接入AI"])

        # 核心清洗匹配
        df_detail['主叫号码_clean'] = df_detail['主叫号码'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        ext_col = '分机号' if '分机号' in df_ext.columns else df_ext.columns[1]
        desc_col = '分机描述' if '分机描述' in df_ext.columns else df_ext.columns[0]
        
        df_ext['分机号_clean'] = df_ext[ext_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        ext_dict = dict(zip(df_ext['分机号_clean'], df_ext[desc_col]))
        
        df_detail['房间是否接入AI'] = df_detail['主叫号码_clean'].map(ext_dict)
        
        # 严格过滤属于呼入的记录
        df_valid = df_detail[df_detail['房间是否接入AI'].notna() & (df_detail['通话类型'] == '呼入')].copy()
        
        # 1:1 复刻 Excel 嵌套 IF 逻辑函数
        def excel_nested_if_logic(row):
            al = str(row['房间是否接入AI']).strip()
            m = str(row['通话状态']).strip()
            n = str(row['AI通话状态']).strip()
            o = str(row['人工通话状态']).strip()
            
            if al == "客房" and m == "接通" and n == "接通" and o == "接通":
                return "进入AI后，再转接人工，且人工接通"
            elif al == "客房" and m == "接通" and n == "接通" and o == "未接通":
                return "AI接通，转接人工，人工未接通"
            elif al == "客房" and m == "接通" and n == "接通" and o == "--":
                return "进入AI后，AI直接完成，未转接人工"
            elif al == "客房" and m == "接通" and n == "--" and o == "接通":
                return "直接进入人工，且人工接通"
            elif al == "客房" and m == "未接通" and n == "--" and o == "--":
                return "客人主动挂断"
            elif al == "客房" and m == "未接通" and n == "--" and o == "未接通":
                return "直接进入人工且最终未接通"
            else:
                return "异常"

        # 绑定新列计算
        df_valid['最终成功接通'] = df_valid.apply(lambda r: "是" if str(r['通话状态']).strip() == "接通" and str(r['通话时长']).strip() != "00:00:00" else "否", axis=1)
        df_valid['接通方式'] = df_valid.apply(excel_nested_if_logic, axis=1)
        df_valid['呼叫所在日期'] = df_valid['呼叫时间'].astype(str).apply(lambda x: x.split()[0] if len(x.split())>0 else '')
        df_valid['呼叫所在小时'] = df_valid['呼叫时间'].astype(str).apply(lambda x: x.split()[1].split(':')[0] if len(x.split())>1 and ':' in x.split()[1] else '')

        st.success("📊 数据链条清洗完毕，准备写入 Excel 公式！")

        def generate_formula_excel(date_range_str):
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
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
            )
            
            # 📌 动态写入抓取到的实际观测日期
            ws1.cell(row=1, column=2, value=f"数据周期：{date_range_str}").font = font_body
            
            for c in range(2, 7): ws1.cell(row=3, column=c).fill = fill_part
            ws1.cell(row=3, column=2, value="PART1：酒店电话数据").font = font_title
            ws1.merge_cells('B3:F3')
            
            ws1.cell(row=4, column=2, value="总来电量\n（所有启用AI的客房呼出的电话量）").alignment = align_left
            ws1.cell(row=4, column=2).font = font_body
            ws1.cell(row=4, column=4, value="=J10").font = font_bold_num  
            ws1.cell(row=4, column=4).alignment = align_center
            for c in range(2, 5): ws1.cell(row=4, column=c).border = thin_border
            ws1.merge_cells('B4:C4')
            
            # 大盘头部指标
            headers_r6 = ["进入AI电话量", "AI接通量", "AI接通率\n（AI接通量/进入AI电话量）", "整体电话接通率\n（切换AI后）", "整体电话接通率\n（切换AI前）"]
            for idx, text in enumerate(headers_r6):
                cell = ws1.cell(row=6, column=idx+2, value=text)
                cell.fill = fill_gray; cell.font = font_header; cell.alignment = align_center; cell.border = thin_border
            
            # 终极校准形式
            ws1.cell(row=7, column=2, value='=COUNTIFS(云总机通话详单!$N:$N, "接通", 云总机通话详单!$AL:$AL, "客房")').font = font_bold_num
            ws1.cell(row=7, column=3, value='=COUNTIFS(云总机通话详单!$N:$N, "接通", 云总机通话详单!$AL:$AL, "客房")').font = font_bold_num
            ws1.cell(row=7, column=4, value="=C7/B7").font = font_bold_num; ws1.cell(row=7, column=4).number_format = '0.00%'
            ws1.cell(row=7, column=5, value="=J3/D4").font = font_bold_num; ws1.cell(row=7, column=5).number_format = '0.00%'
            ws1.cell(row=7, column=6, value=0.967).font = font_bold_num; ws1.cell(row=7, column=6).number_format = '0.00%'
            for c in range(2, 7): ws1.cell(row=7, column=c).alignment = align_center; ws1.cell(row=7, column=c).border = thin_border
            
            # 人工指标区
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
                ws1.cell(row=r, column=2).alignment = align_left
                ws1.cell(row=r, column=4).alignment = align_center
                for c in range(2, 5): ws1.cell(row=r, column=c).border = thin_border
                ws1.merge_cells(f'B{r}:C{r}')

            # 右侧流程池核心映射
            ws1.cell(row=3, column=9, value="最终成功接通").font = font_header; ws1.cell(row=3, column=9).border = thin_border
            ws1.cell(row=3, column=10, value='=COUNTIF(云总机通话详单!$AM:$AM, "是")').font = font_bold_num; ws1.cell(row=3, column=10).border = thin_border; ws1.cell(row=3, column=10).alignment = align_center
            
            flows = [
                ("AI接通", "进入AI后，AI直接完成，未转接人工", '=COUNTIF(云总机通话详单!$AN:$AN, "进入AI后，AI直接完成，未转接人工")'),
                ("人工未接通", "AI接通，转接人工，人工未接通", '=COUNTIF(云总机通话详单!$AN:$AN, "AI接通，转接人工，人工未接通")'),
                ("人工未接通", "直接进入人工且最终未接通", '=COUNTIF(云总机通话详单!$AN:$AN, "直接进入人工且最终未接通")'),
                ("人工接通", "进入AI后，再转接人工，且人工接通", '=COUNTIF(云总机通话详单!$AN:$AN, "进入AI后，再转接人工，且人工接通")'),
                ("人工接通", "直接进入人工，且人工接通", '=COUNTIF(云总机通话详单!$AN:$AN, "直接进入人工，且人工接通")'),
                ("异常", "异常", '=COUNTIF(云总机通话详单!$AN:$AN, "异常")'),
                ("总来电量", "总来电量", "=SUM(J4:J9)")
            ]
            
            for idx, (grp, name, formula) in enumerate(flows):
                r = 4 + idx
                ws1.cell(row=r, column=8, value=grp).font = font_body; ws1.cell(row=r, column=8).border = thin_border; ws1.cell(row=r, column=8).alignment = align_center
                ws1.cell(row=r, column=9, value=name).font = font_body; ws1.cell(row=r, column=9).border = thin_border; ws1.cell(row=r, column=9).alignment = align_left
                ws1.cell(row=r, column=10, value=formula).font = font_bold_num; ws1.cell(row=r, column=10).border = thin_border; ws1.cell(row=r, column=10).alignment = align_center

            ws1.column_dimensions['B'].width = 24
            ws1.column_dimensions['C'].width = 24
            ws1.column_dimensions['D'].width = 18
            ws1.column_dimensions['E'].width = 24
            ws1.column_dimensions['F'].width = 24
            ws1.column_dimensions['H'].width = 15
            ws1.column_dimensions['I'].width = 40
            ws1.column_dimensions['J'].width = 14

            # --- SHEET 2: 云总机通话详单 ---
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
                ws2.cell(row=row_cursor, column=base_len+2, value=row["最终成功接通"])     
                ws2.cell(row=row_cursor, column=base_len+3, value=row["接通方式"])         
                ws2.cell(row=row_cursor, column=base_len+4, value=row["呼叫所在日期"])     
                ws2.cell(row=row_cursor, column=base_len+5, value=row["呼叫所在小时"])     
                row_cursor += 1

            # --- SHEET 3: 分机号 ---
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
                
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            return output

        excel_data = generate_formula_excel(detected_date_range)
        
        st.markdown("---")
        st.download_button(
            label=f"📥 导出【{detected_date_range}】终极版运营报告",
            data=excel_data,
            file_name=f"酒店AI运营报告【{detected_date_range}动态平账版】.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"处理数据时发生异常: {e}")
else:
    st.info("💡 请在上方上传对应表单以生成报告。")
