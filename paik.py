import streamlit as st
import pandas as pd
import random
import copy
import math
import io
import re
import hashlib
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

# ================= 页面全局设置 =================
st.set_page_config(page_title="英华学校教务综合平台", layout="wide", page_icon="🏫")

# 初始化缓存
if 'schedule_result' not in st.session_state: st.session_state['schedule_result'] = None
if 'schedule_classes' not in st.session_state: st.session_state['schedule_classes'] = None
if 'exam_result' not in st.session_state: st.session_state['exam_result'] = None
if 'exam_slips' not in st.session_state: st.session_state['exam_slips'] = None
if 'slip_export_data' not in st.session_state: st.session_state['slip_export_data'] = None

# ================= 核心工具函数 =================

# 严格新高考学科顺位，囊括所有小语种
EDITOR_ORDER = [
    "语文", "数学", "物理", "历史", 
    "英语", "外语", "日语", "俄语", "西班牙语", "法语", "德语", 
    "化学", "地理", "政治", "思想政治", "生物", "生物学", "技术"
]

# 选考科目池（将重新打乱排考场）
DYNAMIC_POOL = ["化学", "地理", "政治", "思想政治", "生物", "生物学", "技术"]

def get_exam_sort_key(ex):
    sub = str(ex.get('科目', '')).strip()
    try:
        return EDITOR_ORDER.index(sub)
    except ValueError:
        return 999  

def format_time_display(time_str, is_html=False):
    br = '<br>' if is_html else '\n'
    time_str = str(time_str).strip()
    if ' ' in time_str:
        return time_str.replace(' ', br, 1)
    for kw in ['日', '号']:
        if kw in time_str:
            parts = time_str.split(kw, 1)
            return f"{parts[0]}{kw}{br}{parts[1]}"
    return time_str

DEFAULT_EXAM_TIMES = {
    "语文": "4月22日 09:00-11:30", 
    "数学": "4月22日 15:00-17:00",
    "物理": "4月23日 09:00-10:15", 
    "历史": "4月23日 09:00-10:15",
    "英语": "4月23日 15:00-17:00", 
    "外语": "4月23日 15:00-17:00",
    "日语": "4月23日 15:00-17:00",
    "俄语": "4月23日 15:00-17:00",
    "西班牙语": "4月23日 15:00-17:00",
    "法语": "4月23日 15:00-17:00",
    "德语": "4月23日 15:00-17:00",
    "化学": "4月24日 08:30-09:45", 
    "地理": "4月24日 11:00-12:15", 
    "政治": "4月24日 14:30-15:45", 
    "思想政治": "4月24日 14:30-15:45",
    "生物": "4月24日 17:00-18:15",
    "生物学": "4月24日 17:00-18:15"
}

# ================= 侧边栏导航 =================
st.sidebar.title("🛠️ 英华教务工作台")
app_mode = st.sidebar.radio("请选择功能模块：", ["📅 智能排课系统", "📝 全科考场编排"])
st.sidebar.markdown("---")

# =====================================================================
#                          模块一：智能排课系统 (占位，防报错)
# =====================================================================
if app_mode == "📅 智能排课系统":
    st.title("📅 全校智能排课系统")
    st.info("💡 排课系统模块已暂时安全折叠。请切换至【全科考场编排】进行考务操作。")

# =====================================================================
#                          模块二：全科考场编排 (核心业务)
# =====================================================================
elif app_mode == "📝 全科考场编排":
    with st.sidebar:
        st.header("⚙️ 考务参数设置")
        room_capacity = st.number_input("选考标准考场容量 (人数)", min_value=10, max_value=60, value=30)
        enable_balance = st.checkbox("✅ 开启选考科目考场人数均衡\n(防止尾考场人数过少)", value=True)
        st.markdown("---")
        st.header("🖨️ 打印排版设置")
        exam_name = st.text_input("考试名称 (如：一模/月考)", value="二模考试")
        slips_per_page = st.selectbox("每页打印学生数", [4, 6, 8], index=1)

    class ExamScheduler:
        def __init__(self, student_df, capacity, time_map, balance):
            self.student_data = student_df.fillna("")
            self.capacity = int(capacity)
            self.time_map = time_map
            self.balance = balance  
            self.fixed_subjects, self.dynamic_subjects = {}, {}

        def parse_data(self):
            for _, row in self.student_data.iterrows():
                student_info = {
                    '准考证号': str(row.get('准考证号', '')), 
                    '姓名': str(row.get('姓名', '')),
                    '行政班': str(row.get('行政班', '')),
                    '原考场': str(row.get('考场', '')).zfill(2) if str(row.get('考场', '')) else "00",
                    '原座位': str(row.get('座位号', '')).zfill(2) if str(row.get('座位号', '')) else "00"
                }
                
                subs_for_student = ['语文', '数学']
                for col in ['语种', '科类', '选考1', '选考2']:
                    if col in self.student_data.columns:
                        sub = str(row.get(col, '')).strip()
                        if sub and sub != 'nan':
                            subs_for_student.append(sub)
                
                for sub in set(subs_for_student): 
                    if sub in DYNAMIC_POOL:
                        if sub not in self.dynamic_subjects: self.dynamic_subjects[sub] = []
                        self.dynamic_subjects[sub].append(student_info)
                    else:
                        if sub not in self.fixed_subjects: self.fixed_subjects[sub] = []
                        self.fixed_subjects[sub].append(student_info)

        def arrange(self):
            self.parse_data()
            exam_results, student_slips = {}, {}
            def add_to_slip(zkz, name, cls, subject, room_name, seat_name):
                if zkz not in student_slips: 
                    student_slips[zkz] = {'姓名': name, '行政班': cls, '准考证号': zkz, 'exams': []}
                time_val = self.time_map.get(subject, "时间待定")
                student_slips[zkz]['exams'].append({'科目': subject, '时间': time_val, '考场': room_name, '座位': seat_name})

            # 固定主科（沿用原考场）
            for subject, students in self.fixed_subjects.items():
                if not students: continue 
                data_list = []
                for stu in students:
                    room_name = f"第{stu['原考场']}考场" if stu['原考场'] != "00" else "未分配"
                    add_to_slip(stu['准考证号'], stu['姓名'], stu['行政班'], subject, room_name, stu['原座位'])
                    data_list.append({'考场号': room_name, '座位号': stu['原座位'], '准考证号': stu['准考证号'], '姓名': stu['姓名'], '行政班': stu['行政班']})
                if data_list:
                    exam_results[subject] = pd.DataFrame(data_list, columns=['考场号', '座位号', '准考证号', '姓名', '行政班']).sort_values(by=['考场号', '座位号'])
                
            # 选考科目（动态智能排考，蛇形+均分）
            for subject, students in self.dynamic_subjects.items():
                if not students: continue 
                students_sorted = sorted(students, key=lambda x: (x['原考场'], x['原座位']))
                data_list = []
                
                total_stu = len(students_sorted)
                num_rooms = math.ceil(total_stu / self.capacity)
                
                if num_rooms > 0:
                    if self.balance:
                        base_cap = total_stu // num_rooms
                        remainder = total_stu % num_rooms
                        room_caps = [base_cap + (1 if i < remainder else 0) for i in range(num_rooms)]
                    else:
                        room_caps = [self.capacity for _ in range(num_rooms - 1)]
                        last_room = total_stu % self.capacity
                        room_caps.append(last_room if last_room != 0 else self.capacity)
                else:
                    room_caps = []

                current_room_idx = 0
                current_seat_idx = 0

                for stu in students_sorted:
                    if current_seat_idx >= room_caps[current_room_idx]:
                        current_room_idx += 1
                        current_seat_idx = 0

                    rn = current_room_idx + 1
                    sn = current_seat_idx + 1
                    room_name = f"第{rn:02d}考场"
                    seat_name = f"{sn:02d}"
                    
                    add_to_slip(stu['准考证号'], stu['姓名'], stu['行政班'], subject, room_name, seat_name)
                    data_list.append({'考场号': room_name, '座位号': seat_name, '准考证号': stu['准考证号'], '姓名': stu['姓名'], '行政班': stu['行政班']})
                    current_seat_idx += 1

                if data_list:
                    exam_results[subject] = pd.DataFrame(data_list, columns=['考场号', '座位号', '准考证号', '姓名', '行政班']).sort_values(by=['考场号', '座位号'])
            
            slip_export_data = []
            for zkz, info in student_slips.items():
                row = {'准考证号': zkz, '姓名': info['姓名'], '行政班': info['行政班']}
                sorted_exams = sorted(info['exams'], key=get_exam_sort_key)
                for idx, ex in enumerate(sorted_exams):
                    row[f'科目{idx+1}'] = ex['科目']; row[f'时间{idx+1}'] = ex['时间']
                    row[f'考场{idx+1}'] = ex['考场']; row[f'座位{idx+1}'] = ex['座位']
                slip_export_data.append(row)
            return exam_results, student_slips, slip_export_data

    # ================= 完美 Word 导出引擎 =================
    def export_slips_to_word(slips_dict, per_page, exam_title):
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Inches(0.25)
        section.bottom_margin = Inches(0.25)
        section.left_margin = Inches(0.35)
        section.right_margin = Inches(0.35)

        def set_font(run, font_name, size, bold=False, color=None):
            run.font.name = font_name
            run.font.size = Pt(size)
            run.bold = bold
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            if color: 
                run.font.color.rgb = RGBColor(*color)

        student_list = list(slips_dict.values())
        num_cols, num_rows = 2, per_page // 2
        for i in range(0, len(student_list), per_page):
            grid_table = doc.add_table(rows=num_rows, cols=num_cols)
            tblPr = grid_table._tbl.tblPr
            tblBorders = OxmlElement('w:tblBorders')
            for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                edge = OxmlElement(f'w:{b}')
                edge.set(qn('w:val'), 'dashed')
                edge.set(qn('w:sz'), '4')
                edge.set(qn('w:color'), 'AAAAAA')
                tblBorders.append(edge)
            tblPr.append(tblBorders)

            for row in grid_table.rows:
                row.height = Inches(10.0 / num_rows)
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

            chunk = student_list[i : i + per_page]
            for idx, info in enumerate(chunk):
                cell = grid_table.cell(idx // num_cols, idx % num_cols)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                # 头部：学校名称居左，标题居中
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_after = Pt(8)
                tab_stops = p.paragraph_format.tab_stops
                tab_stops.add_tab_stop(Inches(1.75), WD_TAB_ALIGNMENT.CENTER)
                
                run_s = p.add_run("🏫 英华学校\t")
                set_font(run_s, '微软雅黑', 9, color=(130, 130, 130))
                
                run_t = p.add_run(f"{exam_title}准考证")
                set_font(run_t, '黑体', 15, True)
                
                # 个人信息：单行精简版
                p2 = cell.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.paragraph_format.space_after = Pt(8)
                
                run_l1 = p2.add_run("姓名: ")
                set_font(run_l1, '宋体', 9, False)
                run_v1 = p2.add_run(f"{info['姓名']} ")
                set_font(run_v1, '黑体', 10, True)
                
                run_l2 = p2.add_run("班级: ")
                set_font(run_l2, '宋体', 9, False)
                run_v2 = p2.add_run(f"{info['行政班']} ")
                set_font(run_v2, '黑体', 10, True)
                
                run_l3 = p2.add_run("考号: ")
                set_font(run_l3, '宋体', 9, False)
                run_v3 = p2.add_run(f"{info['准考证号']}")
                set_font(run_v3, 'Arial', 10, True, color=(200, 60, 0))
                
                # 内部表格配置
                inner = cell.add_table(rows=1, cols=4)
                inner.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                in_tblPr = inner._tbl.tblPr
                in_tblBorders = OxmlElement('w:tblBorders')
                for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                    edge = OxmlElement(f'w:{b}')
                    edge.set(qn('w:val'), 'single')
                    edge.set(qn('w:sz'), '4')
                    edge.set(qn('w:color'), 'E0E0E0')
                    in_tblBorders.append(edge)
                in_tblPr.append(in_tblBorders)

                widths = [Inches(0.6), Inches(1.4), Inches(1.0), Inches(0.5)] 
                for c_idx, w in enumerate(widths): 
                    inner.columns[c_idx].width = w

                # 表头渲染
                hdr = inner.rows[0].cells
                inner.rows[0].height = Inches(0.28) 
                inner.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                
                for j, txt in enumerate(['科目', '考试时间', '考场名称', '座号']):
                    hdr[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER 
                    cp = hdr[j].paragraphs[0]
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER 
                    cp.paragraph_format.space_before = Pt(0) 
                    cp.paragraph_format.space_after = Pt(0)
                    set_font(cp.add_run(txt), '黑体', 9, True, color=(40, 40, 40))
                    shd = parse_xml(r'<w:shd {} w:fill="F5F5F5"/>'.format(nsdecls('w'))) 
                    hdr[j]._tc.get_or_add_tcPr().append(shd)

                # 数据行渲染
                for ex in sorted(info['exams'], key=get_exam_sort_key):
                    row_c = inner.add_row().cells
                    inner.rows[-1].height = Inches(0.28)
                    inner.rows[-1].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                    
                    fmt_time = format_time_display(ex['时间'], is_html=False)
                    vals = [ex['科目'], fmt_time, ex['考场'], ex['座位']]
                    for j, v in enumerate(vals):
                        row_c[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER 
                        cp = row_c[j].paragraphs[0]
                        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER 
                        cp.paragraph_format.space_before = Pt(0)
                        cp.paragraph_format.space_after = Pt(0)
                        
                        if j == 0: 
                            set_font(cp.add_run(v), '宋体', 9, color=(80, 80, 80))
                        elif j == 1: 
                            set_font(cp.add_run(v), 'Arial', 8, color=(80, 80, 80))
                        elif j == 2: 
                            set_font(cp.add_run(v), '黑体', 10, True)
                        elif j == 3: 
                            set_font(cp.add_run(v), 'Arial', 9)

            if i + per_page < len(student_list): 
                doc.add_page_break()

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    st.title("📝 全科考场编排 (教务智能终极版)")
    
    st.info("💡 **分流引擎已激活**：【语文/数学/英语/物理/历史】将自动提取原考场；【化学/生物/政治/地理】将重新分配考场并自动均衡人数。")
    up_exam = st.file_uploader("📂 第一步：上传【考生信息表】", type=['xlsx', 'xls', 'csv'])

    if up_exam:
        df_stu = pd.read_csv(up_exam, dtype=str) if up_exam.name.endswith('.csv') else pd.read_excel(up_exam, dtype=str)
        st.success(f"✅ 成功读取 {len(df_stu)} 名考生")

        st.markdown("### 📅 第二步：确认考试时间")
        
        all_subs = set(['语文', '数学'])
        for col in ['语种', '选考1', '选考2', '科类']:
            if col in df_stu.columns:
                for val in df_stu[col].dropna().unique():
                    v_str = str(val).strip()
                    if v_str and v_str != 'nan':
                        all_subs.add(v_str)
        
        time_data = []
        for s in sorted(list(all_subs), key=lambda x: EDITOR_ORDER.index(x) if x in EDITOR_ORDER else 999):
            time_data.append({"科目": s, "考试时间": DEFAULT_EXAM_TIMES.get(s, "时间待定")})
        
        dynamic_key = "exam_time_editor_" + hashlib.md5("".join(sorted(list(all_subs))).encode()).hexdigest()[:8]
        
        edited_time_df = st.data_editor(
            pd.DataFrame(time_data), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "科目": st.column_config.TextColumn("考试科目 (已锁定)", disabled=True),
                "考试时间": st.column_config.TextColumn("考试时间 (点击可修改)")
            },
            key=dynamic_key 
        )
        
        if '科目' in edited_time_df.columns and '考试时间' in edited_time_df.columns:
            final_time_map = dict(zip(edited_time_df['科目'], edited_time_df['考试时间']))
        else:
            final_time_map = {row['科目']: row['考试时间'] for row in time_data}

        if st.button("🚀 第三步：生成准考证与考场表", type="primary"):
            with st.spinner("正在安全排位并生成教务数据..."):
                sch = ExamScheduler(df_stu, room_capacity, final_time_map, enable_balance)
                res, slips, export = sch.arrange()
                st.session_state['exam_result'] = res
                st.session_state['exam_slips'] = slips
                st.session_state['slip_export_data'] = export

    if st.session_state['exam_result'] is not None:
        st.success("✅ 完美处理完毕！Excel已新增【考场分卷清单】。")
        col1, col2 = st.columns(2)
        with col1:
            output_ex = io.BytesIO()
            with pd.ExcelWriter(output_ex, engine='openpyxl') as writer:
                # 1. 各科目明细
                for sub in sorted(list(st.session_state['exam_result'].keys()), key=lambda x: EDITOR_ORDER.index(x) if x in EDITOR_ORDER else 999):
                    st.session_state['exam_result'][sub].to_excel(writer, sheet_name=f"{sub}考场", index=False)
                
                # 2. 考场分卷统计 Sheet (超稳定重写版)
                all_data_list = []
                for sub, df in st.session_state['exam_result'].items():
                    if not df.empty:
                        temp_df = df.copy()
                        temp_df['科目'] = sub
                        all_data_list.append(temp_df)
                
                if all_data_list:
                    full_df = pd.concat(all_data_list, ignore_index=True)
                    # 按照 考场号 和 科目 统计人数
                    stats_df = full_df.groupby(['考场号', '科目']).size().unstack(fill_value=0)
                    
                    # 按照规定的考试顺序重新排列列名
                    ordered_cols = [c for c in EDITOR_ORDER if c in stats_df.columns]
                    stats_df = stats_df[ordered_cols]
                    
                    # 增加总计列
                    stats_df['考场总卷数'] = stats_df.sum(axis=1)
                    
                    # 重置索引，让“考场号”变成普通列输出到 Excel
                    stats_df.reset_index(inplace=True)
                    stats_df.to_excel(writer, sheet_name="考场分卷清单", index=False)
                
                # 3. 全科汇总
                pd.DataFrame(st.session_state['slip_export_data']).to_excel(writer, sheet_name="学生个人全科汇总", index=False)
            
            st.download_button("📊 下载考务 Excel 总表 (含分卷统计)", data=output_ex.getvalue(), file_name=f"英华_{exam_name}_考务总表.xlsx")
        
        with col2:
            word_f = export_slips_to_word(st.session_state['exam_slips'], slips_per_page, exam_name)
            st.download_button(f"📄 下载 {exam_name} 准考证 (Word)", data=word_f, file_name=f"英华_{exam_name}_准考证_{slips_per_page}人版.docx")

        view = st.radio("👀 预览模式", ["🎫 准考条 UI 预览", "📋 各科目考场门贴"], horizontal=True)
        if "预览" in view:
            html = '<div style="display:flex; flex-wrap:wrap; justify-content:center; gap:15px; padding:10px; background:#f0f2f6;">'
            for zkz, info in list(st.session_state['exam_slips'].items())[:20]: 
                html += f"""
                <div style="width:340px; border:1px dashed #999; border-radius:8px; background:#fff; padding:12px; margin-bottom:10px;">
                    <div style="border-bottom:1px solid #eee; padding-bottom:6px; margin-bottom:10px; position:relative; text-align:center;">
                        <span style="font-size:10px; color:#888; position:absolute; left:0; bottom:3px;">🏫 英华学校</span>
                        <b style="font-size:17px; letter-spacing:1px;">{exam_name}准考证</b>
                    </div>
                    <div style="font-size:11px; margin-bottom:10px; text-align:center;">
                        姓名: <b>{info['姓名']}</b> &nbsp; 班级: <b>{info['行政班']}</b> &nbsp; 考号: <b style="color:#cc4400; font-family:Arial;">{info['准考证号']}</b>
                    </div>
                    <table style="width:100%; border-collapse:collapse; font-size:11px; text-align:center; border:1px solid #e0e0e0;">
                        <tr style="background:#f5f5f5; color:#333;">
                            <th style="padding:6px; border:1px solid #e0e0e0;">科目</th>
                            <th style="padding:6px; border:1px solid #e0e0e0;">考试时间</th>
                            <th style="padding:6px; border:1px solid #e0e0e0;">考场名称</th>
                            <th style="padding:6px; border:1px solid #e0e0e0;">座号</th>
                        </tr>
                """
                for ex in sorted(info['exams'], key=get_exam_sort_key):
                    fmt_time_html = format_time_display(ex['时间'], is_html=True)
                    html += f"<tr><td style='padding:5px; border:1px solid #e0e0e0; color:#555;'>{ex['科目']}</td><td style='line-height:1.3; padding:5px 0; border:1px solid #e0e0e0; color:#555;'>{fmt_time_html}</td><td style='font-weight:bold; border:1px solid #e0e0e0;'>{ex['考场']}</td><td style='border:1px solid #e0e0e0;'>{ex['座位']}</td></tr>"
                html += "</table></div>"
            st.markdown(html + '</div>', unsafe_allow_html=True)
            if len(st.session_state['exam_slips']) > 20: st.warning("💡 网页仅显示前 20 位考生进行预览，下载 Word 查看全校数据。")
        else:
            tabs = st.tabs(sorted(list(st.session_state['exam_result'].keys()), key=lambda x: EDITOR_ORDER.index(x) if x in EDITOR_ORDER else 999))
            for idx, sub in enumerate(sorted(list(st.session_state['exam_result'].keys()), key=lambda x: EDITOR_ORDER.index(x) if x in EDITOR_ORDER else 999)):
                with tabs[idx]: st.dataframe(st.session_state['exam_result'][sub], use_container_width=True, hide_index=True)