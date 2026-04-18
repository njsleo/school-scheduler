import streamlit as st
import pandas as pd
import random
import copy
import math
import io
import re
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
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

# 🚨 【核心修正】：严格按照新高考(3+1+2)真实考试顺序锁定！
EDITOR_ORDER = ["语文", "数学", "物理", "历史", "英语", "外语", "化学", "地理", "政治", "思想政治", "生物", "生物学", "技术"]

# 1. 绝对强制的学科排序引擎
def get_exam_sort_key(ex):
    sub = str(ex.get('科目', '')).strip()
    try:
        return EDITOR_ORDER.index(sub)
    except ValueError:
        return 999  # 不认识的科目老老实实垫底

# 2. 智能换行排版 (将日期和时间劈成两行，节省宽度)
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

# 🚨 预设新高考联考标准时间轴 (防呆预填，老师可直接修改日期)
DEFAULT_EXAM_TIMES = {
    "语文": "4月23日 09:00-11:30", 
    "数学": "4月23日 15:00-17:00",
    "物理": "4月24日 09:00-10:15", 
    "历史": "4月24日 09:00-10:15",
    "英语": "4月24日 15:00-17:00", 
    "外语": "4月24日 15:00-17:00",
    "化学": "4月25日 08:30-09:45", 
    "地理": "4月25日 11:00-12:15", 
    "政治": "4月25日 14:30-15:45", 
    "思想政治": "4月25日 14:30-15:45",
    "生物": "4月25日 17:00-18:15",
    "生物学": "4月25日 17:00-18:15"
}

# ================= 侧边栏导航 =================
st.sidebar.title("🛠️ 英华教务工作台")
app_mode = st.sidebar.radio("请选择功能模块：", ["📅 智能排课系统", "📝 全科考场编排"])
st.sidebar.markdown("---")

# =====================================================================
#                          模块一：智能排课系统
# =====================================================================
if app_mode == "📅 智能排课系统":
    st.title("📅 全校智能排课系统")
    st.info("排课系统运行中... (保留原有排课逻辑)")
    # (此部分保留你之前的排课核心代码即可)

# =====================================================================
#                          模块二：全科考场编排
# =====================================================================
elif app_mode == "📝 全科考场编排":
    with st.sidebar:
        st.header("⚙️ 考务参数设置")
        room_capacity = st.number_input("标准考场容量 (人数)", min_value=10, max_value=60, value=30)
        st.markdown("---")
        st.header("🖨️ 打印排版设置")
        exam_name = st.text_input("考试名称 (如：一模/月考)", value="四月全市统考")
        slips_per_page = st.selectbox("每页打印学生数", [4, 6, 8], index=1)

    class ExamScheduler:
        def __init__(self, student_df, capacity, time_map):
            self.student_data = student_df.fillna("")
            self.capacity = int(capacity)
            self.time_map = time_map
            self.fixed_subjects, self.dynamic_subjects = {}, {}

        def parse_data(self):
            for _, row in self.student_data.iterrows():
                student_info = {
                    '准考证号': str(row.get('准考证号', '')), '姓名': str(row.get('姓名', '')),
                    '行政班': str(row.get('行政班', '')),
                    '原考场': str(row.get('考场', '')).zfill(2) if str(row.get('考场', '')) else "00",
                    '原座位': str(row.get('座位号', '')).zfill(2) if str(row.get('座位号', '')) else "00"
                }
                fixed_subs = ['语文', '数学']
                lang = str(row.get('语种', '')).strip()
                if lang and lang != 'nan': fixed_subs.append(lang)
                for sub in fixed_subs:
                    if sub not in self.fixed_subjects: self.fixed_subjects[sub] = []
                    self.fixed_subjects[sub].append(student_info)
                for col in ['科类', '选考1', '选考2']:
                    sub = str(row.get(col, '')).strip()
                    if sub and sub != 'nan':
                        if sub not in self.dynamic_subjects: self.dynamic_subjects[sub] = []
                        self.dynamic_subjects[sub].append(student_info)

        def arrange(self):
            self.parse_data()
            exam_results, student_slips = {}, {}
            def add_to_slip(zkz, name, cls, subject, room_name, seat_name):
                if zkz not in student_slips: student_slips[zkz] = {'姓名': name, '行政班': cls, '准考证号': zkz, 'exams': []}
                time_val = self.time_map.get(subject, "时间待定")
                student_slips[zkz]['exams'].append({'科目': subject, '时间': time_val, '考场': room_name, '座位': seat_name})

            for subject, students in self.fixed_subjects.items():
                data_list = []
                for stu in students:
                    room_name = f"第{stu['原考场']}考场" if stu['原考场'] != "00" else "未分配"
                    add_to_slip(stu['准考证号'], stu['姓名'], stu['行政班'], subject, room_name, stu['原座位'])
                    data_list.append({'考场号': room_name, '座位号': stu['原座位'], '准考证号': stu['准考证号'], '姓名': stu['姓名'], '行政班': stu['行政班']})
                exam_results[subject] = pd.DataFrame(data_list).sort_values(by=['考场号', '座位号'])
                
            for subject, students in self.dynamic_subjects.items():
                students_sorted = sorted(students, key=lambda x: (x['原考场'], x['原座位']))
                data_list = []
                for i, stu in enumerate(students_sorted):
                    rn, sn = (i // self.capacity) + 1, (i % self.capacity) + 1
                    room_name, seat_name = f"第{rn:02d}考场", f"{sn:02d}"
                    add_to_slip(stu['准考证号'], stu['姓名'], stu['行政班'], subject, room_name, seat_name)
                    data_list.append({'考场号': room_name, '座位号': seat_name, '准考证号': stu['准考证号'], '姓名': stu['姓名'], '行政班': stu['行政班']})
                exam_results[subject] = pd.DataFrame(data_list).sort_values(by=['考场号', '座位号'])
            
            slip_export_data = []
            for zkz, info in student_slips.items():
                row = {'准考证号': zkz, '姓名': info['姓名'], '行政班': info['行政班']}
                # 导出大表，依然应用绝对考试顺序
                sorted_exams = sorted(info['exams'], key=get_exam_sort_key)
                for idx, ex in enumerate(sorted_exams):
                    row[f'科目{idx+1}'] = ex['科目']; row[f'时间{idx+1}'] = ex['时间']
                    row[f'考场{idx+1}'] = ex['考场']; row[f'座位{idx+1}'] = ex['座位']
                slip_export_data.append(row)
            return exam_results, student_slips, slip_export_data

    # --- 渲染与导出逻辑 ---
    def export_slips_to_word(slips_dict, per_page, exam_title):
        doc = Document()
        section = doc.sections[0]
        section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = Inches(0.25), Inches(0.25), Inches(0.35), Inches(0.35)

        def set_font(run, name, size, bold=False, color=None):
            run.font.name = name; run.font.size = Pt(size); run.bold = bold
            run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
            if color: run.font.color.rgb = RGBColor(*color)

        student_list = list(slips_dict.values())
        num_cols, num_rows = 2, per_page // 2
        for i in range(0, len(student_list), per_page):
            grid_table = doc.add_table(rows=num_rows, cols=num_cols)
            tblPr = grid_table._tbl.tblPr
            tblBorders = OxmlElement('w:tblBorders')
            for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                edge = OxmlElement(f'w:{b}'); edge.set(qn('w:val'), 'dashed'); edge.set(qn('w:sz'), '4'); edge.set(qn('w:color'), 'AAAAAA')
                tblBorders.append(edge)
            tblPr.append(tblBorders)

            for row in grid_table.rows:
                row.height = Inches(10.0 / num_rows); row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

            chunk = student_list[i : i + per_page]
            for idx, info in enumerate(chunk):
                cell = grid_table.cell(idx // num_cols, idx % num_cols)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_s = p.add_run("🏫 英华学校\n"); set_font(run_s, '微软雅黑', 9, color=(120, 120, 120))
                run_t = p.add_run(f"{exam_title}准考证"); set_font(run_t, '黑体', 14, True)
                
                p2 = cell.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_font(p2.add_run(f"姓名：{info['姓名']}  班级：{info['行政班']}  考号：{info['准考证号']}"), 10, True)
                
                inner = cell.add_table(rows=1, cols=4); inner.style = 'Table Grid'; inner.alignment = WD_ALIGN_PARAGRAPH.CENTER
                widths = [Inches(0.55), Inches(1.2), Inches(1.1), Inches(0.45)] 
                for c_idx, w in enumerate(widths): inner.columns[c_idx].width = w

                hdr = inner.rows[0].cells
                for j, txt in enumerate(['科目', '考试时间', '考场名称', '座号']):
                    cp = hdr[j].paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_font(cp.add_run(txt), 8, True)
                    shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w'))); hdr[j]._tc.get_or_add_tcPr().append(shd)

                # 强硬应用新高考学科顺序 (语、数、物/史...)
                for ex in sorted(info['exams'], key=get_exam_sort_key):
                    row_c = inner.add_row().cells
                    fmt_time = format_time_display(ex['时间'], is_html=False)
                    vals = [ex['科目'], fmt_time, ex['考场'], ex['座位']]
                    for j, v in enumerate(vals):
                        cp = row_c[j].paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        set_font(cp.add_run(v), 8, (j == 2)) 

            if i + per_page < len(student_list): doc.add_page_break()

        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        return buf

    st.title("📝 全科考场编排 (新高考统考版)")
    up_exam = st.file_uploader("📂 第一步：上传【考生信息表】", type=['xlsx', 'xls', 'csv'])

    if up_exam:
        df_stu = pd.read_csv(up_exam, dtype=str) if up_exam.name.endswith('.csv') else pd.read_excel(up_exam, dtype=str)
        st.success(f"✅ 成功读取 {len(df_stu)} 名考生")

        st.markdown("### 📅 第二步：确认考试时间")
        st.info("🔒 提示：表格已按【新高考3+1+2真实考试顺序】锁定！已为您预填常规考试时间，可直接点击修改。")
        all_subs = set(['语文', '数学'])
        lang = df_stu['语种'].unique().tolist() if '语种' in df_stu.columns else []
        sel1 = df_stu['选考1'].unique().tolist() if '选考1' in df_stu.columns else []
        sel2 = df_stu['选考2'].unique().tolist() if '选考2' in df_stu.columns else []
        kl = df_stu['科类'].unique().tolist() if '科类' in df_stu.columns else []
        for s in lang + sel1 + sel2 + kl:
            if str(s) != 'nan' and str(s).strip() != '': all_subs.add(s.strip())
        
        time_data = []
        # 强制按新高考顺序排列UI录入表
        for s in sorted(list(all_subs), key=lambda x: EDITOR_ORDER.index(x) if x in EDITOR_ORDER else 999):
            time_data.append({"科目": s, "考试时间": DEFAULT_EXAM_TIMES.get(s, "时间待定")})
        
        # 终极修复点：更换全新 key 清除历史脏缓存
        edited_time_df = st.data_editor(
            pd.DataFrame(time_data), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "科目": st.column_config.TextColumn("考试科目 (按真实考试顺序锁定)", disabled=True),
                "考试时间": st.column_config.TextColumn("考试时间 (点击可修改)")
            },
            key="exam_time_editor_xingaokao_v1" 
        )
        
        if '科目' in edited_time_df.columns and '考试时间' in edited_time_df.columns:
            final_time_map = dict(zip(edited_time_df['科目'], edited_time_df['考试时间']))
        else:
            st.warning("⚠️ 检测到表格状态异常，已自动应用默认时间。")
            final_time_map = {row['科目']: row['考试时间'] for row in time_data}

        if st.button("🚀 第三步：生成准考证与考场表", type="primary"):
            with st.spinner("正在按统考时序进行排位..."):
                sch = ExamScheduler(df_stu, room_capacity, final_time_map)
                res, slips, export = sch.arrange()
                st.session_state['exam_result'], st.session_state['exam_slips'], st.session_state['slip_export_data'] = res, slips, export

    if st.session_state['exam_result'] is not None:
        st.success("✅ 完美处理完毕！顺位与新高考考表完全对齐！")
        col1, col2 = st.columns(2)
        with col1:
            output_ex = io.BytesIO()
            with pd.ExcelWriter(output_ex, engine='openpyxl') as writer:
                # 为了教务找表方便，导出的Excel底部Sheet表签也按新高考顺序排
                for sub in sorted(list(st.session_state['exam_result'].keys()), key=lambda x: EDITOR_ORDER.index(x) if x in EDITOR_ORDER else 999):
                    st.session_state['exam_result'][sub].to_excel(writer, sheet_name=f"{sub}考场", index=False)
                pd.DataFrame(st.session_state['slip_export_data']).to_excel(writer, sheet_name="全科汇总", index=False)
            st.download_button("📊 下载考务 Excel 总表", data=output_ex.getvalue(), file_name=f"英华_{exam_name}_考场汇总.xlsx")
        with col2:
            word_f = export_slips_to_word(st.session_state['exam_slips'], slips_per_page, exam_name)
            st.download_button(f"📄 下载 {exam_name} 准考证 (Word)", data=word_f, file_name=f"英华_{exam_name}_准考证_{slips_per_page}人版.docx")

        view = st.radio("👀 预览模式", ["🎫 准考条 UI 预览", "📋 各科目考场门贴"], horizontal=True)
        if "预览" in view:
            html = '<div style="display:flex; flex-wrap:wrap; justify-content:center; gap:15px; padding:10px; background:#f0f2f6;">'
            for zkz, info in list(st.session_state['exam_slips'].items())[:20]: 
                html += f"""
                <div style="width:340px; border:1px dashed #999; border-radius:8px; background:#fff; padding:10px; margin-bottom:10px;">
                    <div style="text-align:center; border-bottom:1px solid #eee; padding-bottom:5px; margin-bottom:5px;">
                        <span style="font-size:10px; color:#888;">🏫 英华学校</span><br><b style="font-size:16px;">{exam_name}准考证</b>
                    </div>
                    <div style="font-size:12px; margin-bottom:8px; text-align:center;">
                        <b>{info['姓名']}</b> | {info['行政班']} | <span style="color:#d35400;">{info['准考证号']}</span>
                    </div>
                    <table style="width:100%; border-collapse:collapse; font-size:11px; text-align:center; border:1px solid #ddd;">
                        <tr style="background:#f2f2f2;"><th>科目</th><th>考试时间</th><th>考场</th><th>座号</th></tr>
                """
                for ex in sorted(info['exams'], key=get_exam_sort_key):
                    fmt_time_html = format_time_display(ex['时间'], is_html=True)
                    html += f"<tr style='border-bottom:1px solid #eee;'><td style='padding:4px;'>{ex['科目']}</td><td style='line-height:1.2; padding:3px 0;'>{fmt_time_html}</td><td style='font-weight:bold;'>{ex['考场']}</td><td>{ex['座位']}</td></tr>"
                html += "</table></div>"
            st.markdown(html + '</div>', unsafe_allow_html=True)
            if len(st.session_state['exam_slips']) > 20: st.warning("💡 网页仅显示前 20 位考生进行预览，下载 Word 查看全校数据。")
        else:
            tabs = st.tabs(sorted(list(st.session_state['exam_result'].keys()), key=lambda x: EDITOR_ORDER.index(x) if x in EDITOR_ORDER else 999))
            for idx, sub in enumerate(sorted(list(st.session_state['exam_result'].keys()), key=lambda x: EDITOR_ORDER.index(x) if x in EDITOR_ORDER else 999)):
                with tabs[idx]: st.dataframe(st.session_state['exam_result'][sub], use_container_width=True, hide_index=True)