import streamlit as st
import pandas as pd
import random
import re
import copy
import math
import io
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ================= 页面全局设置 =================
st.set_page_config(page_title="英华学校教务综合平台", layout="wide", page_icon="🏫")

# 初始化缓存
if 'schedule_result' not in st.session_state: st.session_state['schedule_result'] = None
if 'schedule_classes' not in st.session_state: st.session_state['schedule_classes'] = None
if 'exam_result' not in st.session_state: st.session_state['exam_result'] = None
if 'exam_slips' not in st.session_state: st.session_state['exam_slips'] = None
if 'slip_export_data' not in st.session_state: st.session_state['slip_export_data'] = None

# ================= 侧边栏导航 =================
st.sidebar.title("🛠️ 英华教务工作台")
app_mode = st.sidebar.radio("请选择功能模块：", ["📅 智能排课系统", "📝 全科考场编排"])
st.sidebar.markdown("---")

# =====================================================================
#                          功能模块：排课系统
# =====================================================================
if app_mode == "📅 智能排课系统":
    # --- 1. 排课专属配置 & 变量 ---
    DAYS = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
    SLOTS = [
        "早自习\n7:00-7:40", "第一节\n8:10-8:55", "第二节\n9:05-9:50", "第三节\n10:15-11:00", 
        "第四节\n11:00-11:45", "第五节\n14:00-14:45", "第六节\n14:55-15:40", "第七节\n15:50-16:35", 
        "第八节\n16:45-18:15", "第九节\n19:00-20:30", "第十节\n20:50-22:30"
    ]
    SUBJECT_COLORS = {
        "走班": "#FFFF99", "语文": "#FFE4E1", "数学": "#E4F1FF", "英语": "#E0FFFF", 
        "物理": "#E8F5E9", "化学": "#FFFACD", "生物": "#FCE4EC", "政治": "#FFDAB9", 
        "历史": "#FFE4B5", "地理": "#E1BEE7", "体育": "#C8E6C9", "音乐": "#E1BEE7", 
        "美术": "#FFCCBC", "班会": "#B3E5FC", "晚自习": "#F8F9FA", "自习": "#F8F9FA",
        "信息": "#E0F7FA", "通用": "#F5F5F5"
    }

    with st.sidebar:
        st.header("⚙️ 排课参数设置")
        class_prefix = st.text_input("班级前缀 (如: 高一 / 高三理科)", value="高三理科")
        class_suffix = st.text_input("班级后缀", value="班")
        include_sunday_evening = st.checkbox("✅ 开启周日晚修 (高三模式)", value=True)

    class CourseScheduler:
        def __init__(self, teacher_df, prefix, suffix, include_sunday):
            self.teacher_data = teacher_df.fillna("") 
            self.classes = set()
            self.teacher_busy = {} 
            self.prefix = prefix
            self.suffix = suffix
            self.evening_days = ['周一', '周二', '周三', '周四', '周五', '周日'] if include_sunday else ['周一', '周二', '周三', '周四', '周五']

        def parse_tasks(self):
            tasks, zouban_tasks = [], []
            self.teacher_data.columns = [str(c).strip() for c in self.teacher_data.columns]
            for _, row in self.teacher_data.iterrows():
                row_dict = {str(k): str(v).strip() for k, v in row.items()}
                t_name, sub, constraint, raw_classes, hours = "", "", "", "", 0
                for k, v in row_dict.items():
                    if v == "": continue
                    if '教师' in k or '老师' in k or '姓名' in k: t_name = v
                    elif '课' in k and '称' in k or '科目' in k: sub = v
                    elif '限制' in k or '连排' in k or '要求' in k: constraint = v
                    elif '班' in k: raw_classes = v
                    elif '课时' in k or '节数' in k:
                        try: hours = int(float(re.findall(r'\d+', v)[0]))
                        except: pass
                
                if not t_name and len(row_dict) >= 1: t_name = list(row_dict.values())[0]
                if not sub and len(row_dict) >= 2: sub = list(row_dict.values())[1]
                if not raw_classes and len(row_dict) >= 3: raw_classes = list(row_dict.values())[2]
                if hours <= 0 or not t_name or not raw_classes: continue

                class_nums = re.findall(r'\d+', raw_classes)
                target_classes = [f"{self.prefix}{n}{self.suffix}" for n in class_nums] 
                for c in target_classes: self.classes.add(c)

                if "Block_ZouBan" in constraint or "走班" in sub or "走班" in t_name:
                    zouban_tasks.append({'classes': target_classes, 'teacher': t_name, 'subject': sub, 'hours': hours})
                else:
                    for c in target_classes: tasks.append({'class': c, 'teacher': t_name, 'subject': sub, 'hours': hours, 'constraint': constraint})

            def sort_key(c_str):
                nums = re.findall(r'\d+', c_str)
                return int(nums[0]) if nums else 0
            self.classes = sorted(list(self.classes), key=sort_key) 
            if not self.classes: self.classes = [f'{self.prefix}1{self.suffix}'] 
            return tasks, zouban_tasks

        def is_free_for(self, c_name, teacher, subject, day, slot_list):
            for slot in slot_list:
                if self.schedules[c_name][day][slot] != "": return False 
                if (teacher, day, slot) in self.teacher_busy: return False 
            
            count = sum(1 for s in SLOTS if self.schedules[c_name][day][s] != "" and ("走班" in subject and "走班" in self.schedules[c_name][day][s] or self.schedules[c_name][day][s].startswith(f"{subject}<br>")))
            if count + len(slot_list) > 2: return False
            return True

        def book(self, c_name, teacher, subject, day, slot):
            content = "走班" if ("走班" in subject or "走班" in teacher) else f"{subject}<br>{teacher}"
            self.schedules[c_name][day][slot] = content
            self.teacher_busy[(teacher, day, slot)] = c_name 

        def run(self):
            self.original_tasks, self.original_zouban = self.parse_tasks()
            if not self.original_tasks and not self.original_zouban: 
                return {c: {day: {slot: "" for slot in SLOTS} for day in DAYS} for c in self.classes}, 0

            PAIRS = [(SLOTS[1], SLOTS[2]), (SLOTS[3], SLOTS[4]), (SLOTS[7], SLOTS[8])]
            best_schedule, best_score = None, 99999
            
            for attempt in range(100):
                tasks, zouban_tasks = copy.deepcopy(self.original_tasks), copy.deepcopy(self.original_zouban)
                self.schedules = {c: {day: {slot: "" for slot in SLOTS} for day in DAYS} for c in self.classes}
                self.teacher_busy = {} 

                zouban_eve_ok = True
                for z_task in zouban_tasks:
                    if z_task['hours'] < 2: continue
                    placed = False
                    days = list(self.evening_days)
                    random.shuffle(days)
                    for day in days:
                        if all(self.is_free_for(c, z_task['teacher'], z_task['subject'], day, [SLOTS[9], SLOTS[10]]) for c in z_task['classes']):
                            for c in z_task['classes']:
                                self.book(c, z_task['teacher'], z_task['subject'], day, SLOTS[9])
                                self.book(c, z_task['teacher'], z_task['subject'], day, SLOTS[10])
                            z_task['hours'] -= 2; placed = True; break
                    if not placed: zouban_eve_ok = False; break
                if not zouban_eve_ok: continue

                eve_placed_ok = True
                eve_assigned = {c: set() for c in self.classes}
                for day in self.evening_days:
                    cls_list = list(self.classes)
                    random.shuffle(cls_list)
                    for c in cls_list:
                        if self.schedules[c][day][SLOTS[9]] != "": continue 
                        c_tasks = [t for t in tasks if t['class'] == c and t['hours'] >= 2]
                        random.shuffle(c_tasks)
                        placed = False
                        for t in [t for t in c_tasks if t['subject'] not in eve_assigned[c]] + [t for t in c_tasks if t['subject'] in eve_assigned[c]]:
                            if self.is_free_for(c, t['teacher'], t['subject'], day, [SLOTS[9], SLOTS[10]]):
                                self.book(c, t['teacher'], t['subject'], day, SLOTS[9])
                                self.book(c, t['teacher'], t['subject'], day, SLOTS[10])
                                t['hours'] -= 2; eve_assigned[c].add(t['subject']); placed = True; break
                        if not placed: eve_placed_ok = False; break
                    if not eve_placed_ok: break
                if not eve_placed_ok: continue 

                for z_task in zouban_tasks:
                    while z_task['hours'] >= 2:
                        placed = False
                        all_days = DAYS[:5]
                        random.shuffle(all_days)
                        for day in all_days:
                            shuffled_pairs = PAIRS.copy()
                            random.shuffle(shuffled_pairs)
                            for s1, s2 in shuffled_pairs:
                                if all(self.is_free_for(c, z_task['teacher'], z_task['subject'], day, [s1, s2]) for c in z_task['classes']):
                                    for c in z_task['classes']:
                                        self.book(c, z_task['teacher'], z_task['subject'], day, s1)
                                        self.book(c, z_task['teacher'], z_task['subject'], day, s2)
                                    z_task['hours'] -= 2; placed = True; break
                            if placed: break
                        if not placed: break
                    
                    while z_task['hours'] > 0:
                        placed = False
                        all_times = [(d, s) for d in DAYS[:5] for s in [SLOTS[5], SLOTS[6]]]
                        random.shuffle(all_times)
                        for day, slot in all_times:
                            if all(self.is_free_for(c, z_task['teacher'], z_task['subject'], day, [slot]) for c in z_task['classes']):
                                for c in z_task['classes']: self.book(c, z_task['teacher'], z_task['subject'], day, slot)
                                z_task['hours'] -= 1; placed = True; break
                        if not placed: break

                for task in tasks:
                    while task['hours'] >= 2:
                        placed = False
                        all_days = DAYS[:5]
                        random.shuffle(all_days)
                        for day in all_days:
                            shuffled_pairs = PAIRS.copy()
                            random.shuffle(shuffled_pairs)
                            for s1, s2 in shuffled_pairs:
                                if self.is_free_for(task['class'], task['teacher'], task['subject'], day, [s1, s2]):
                                    self.book(task['class'], task['teacher'], task['subject'], day, s1)
                                    self.book(task['class'], task['teacher'], task['subject'], day, s2)
                                    task['hours'] -= 2; placed = True; break 
                            if placed: break 
                        if not placed: break 

                for task in tasks:
                    while task['hours'] > 0:
                        placed = False
                        all_days = DAYS[:5]
                        random.shuffle(all_days)
                        for day in all_days:
                            if self.is_free_for(task['class'], task['teacher'], task['subject'], day, [SLOTS[5]]):
                                self.book(task['class'], task['teacher'], task['subject'], day, SLOTS[5]); task['hours'] -= 1; placed = True; break
                        if not placed:
                            all_times = [(d, s) for d in DAYS[:5] for s in [SLOTS[1], SLOTS[2], SLOTS[3], SLOTS[4], SLOTS[7], SLOTS[8]]]
                            random.shuffle(all_times)
                            for day, slot in all_times:
                                if self.is_free_for(task['class'], task['teacher'], task['subject'], day, [slot]):
                                    self.book(task['class'], task['teacher'], task['subject'], day, slot); task['hours'] -= 1; placed = True; break
                        if not placed: break
                        
                unplaced = sum(t['hours'] for t in tasks) + sum(t['hours'] for t in zouban_tasks)
                empty_eve = sum(1 for c in self.classes for day in self.evening_days for s in [SLOTS[9], SLOTS[10]] if self.schedules[c][day][s] == "")
                score = unplaced * 10 + empty_eve
                if score < best_score:
                    best_score = score
                    best_schedule = copy.deepcopy(self.schedules)
                if score == 0: break 
                    
            self.schedules = best_schedule
            for c in self.classes:
                for day in DAYS[:5]:
                    if self.schedules[c][day][SLOTS[6]] == "": self.schedules[c][day][SLOTS[6]] = "自习"
                for day in self.evening_days:
                    for s in [SLOTS[9], SLOTS[10]]:
                        if self.schedules[c][day][s] == "": self.schedules[c][day][s] = "晚自习"

            return self.schedules, len(self.original_tasks) + len(self.original_zouban)

    def render_class_table(class_name, schedule):
        html = f"""
        <div id="table-{class_name}" style="font-family: 'SimSun', '宋体', sans-serif; max-width: 1050px; margin: auto; background-color: white; padding: 20px;">
            <h2 style="text-align: center; color: red; letter-spacing: 2px;">呼和浩特市英华学校 {class_name} 课程表</h2>
            <table border="1" style="width: 100%; text-align: center; border-collapse: collapse; border-color: #333; font-size: 16px;">
                <tr style="background-color: #f8f9fa;"><th style="padding: 12px;">时间 \\ 星期</th><th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th>六</th><th>日</th></tr>
        """
        def get_cell(day, slot):
            content = schedule[day][slot]
            bg_color = "#FFFFFF" 
            for sub, color in SUBJECT_COLORS.items():
                if sub in content: bg_color = color; break
            if "走班" in content: return f'<td style="background-color: {bg_color}; font-weight: bold; padding: 12px; box-shadow: inset 0 0 5px rgba(0,0,0,0.1);">{content}</td>'
            return f'<td style="background-color: {bg_color}; padding: 12px;">{content}</td>'

        html += "<tr><td>早自习(7:00-7:40)</td>"
        for day in DAYS[:5]: html += "<td style='padding: 12px; background-color: #F8F9FA;'>早自习</td>"
        html += "<td rowspan='12' style='width:45px; background-color: #F8F9FA; color: #555;'>考<br><br>试</td><td rowspan='12' style='width:45px; background-color: #F8F9FA; color: #555;'>休<br><br>息</td></tr>" 
        for slot in SLOTS[1:3]:
            html += f"<tr><td>{slot.replace(chr(10), '<br>')}</td>"
            for day in DAYS[:5]: html += get_cell(day, slot)
            html += "</tr>"
        html += "<tr><td colspan='6' style='letter-spacing: 25px; font-weight:bold; padding: 6px; background-color: #EFEFEF; color: #666;'>课间操</td></tr>"
        for slot in SLOTS[3:5]:
            html += f"<tr><td>{slot.replace(chr(10), '<br>')}</td>"
            for day in DAYS[:5]: html += get_cell(day, slot)
            html += "</tr>"
        html += "<tr><td colspan='6' style='letter-spacing: 25px; font-weight:bold; padding: 6px; background-color: #EFEFEF; color: #666;'>午餐</td></tr>"
        for slot in SLOTS[5:9]:
            html += f"<tr><td>{slot.replace(chr(10), '<br>')}</td>"
            for day in DAYS[:5]: html += get_cell(day, slot)
            html += "</tr>"
        html += "<tr><td colspan='8' style='letter-spacing: 25px; font-weight:bold; padding: 6px; background-color: #EFEFEF; color: #666;'>晚餐</td></tr>"
        for slot in SLOTS[9:]:
            html += f"<tr><td>{slot.replace(chr(10), '<br>')}</td>"
            for day in DAYS[:5]: html += get_cell(day, slot)
            html += "<td style='padding: 12px; background-color: #F8F9FA;'>休息</td>" + get_cell('周日', slot) + "</tr>"
        return html + "</table></div><br>"

    def export_course_to_excel(schedules, classes):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for c_name in classes:
                data = []
                for slot in SLOTS:
                    row = [slot.replace('\n', ' ')] 
                    for day in DAYS:
                        row.append(schedules[c_name][day][slot].replace("<br>", "\n"))
                    data.append(row)
                pd.DataFrame(data, columns=['时间段'] + DAYS).to_excel(writer, sheet_name=c_name, index=False)
        return output.getvalue()

    st.title("📅 全校智能排课系统")
    uploaded_course = st.file_uploader("📂 请上传【排课数据】 Excel", type=['xlsx', 'xls'], key="course_upload")

    if uploaded_course:
        all_sheets = pd.read_excel(uploaded_course, sheet_name=None)
        teacher_df = all_sheets[list(all_sheets.keys())[0]] if len(all_sheets) == 1 else [df for df in all_sheets.values() if len(df) > 5 and len(df.columns) >= 3][0]

        if st.button("🚀 开始生成教务课表", type="primary"):
            with st.spinner("正在执行算法排课，寻找完美解..."):
                scheduler = CourseScheduler(teacher_df, class_prefix, class_suffix, include_sunday_evening)
                result, total_tasks = scheduler.run()
                if total_tasks == 0: st.error("❌ 解析失败，请检查 Excel 表头。")
                else:
                    st.session_state['schedule_result'] = result
                    st.session_state['schedule_classes'] = scheduler.classes

    if st.session_state['schedule_result'] is not None:
        st.success("✅ 完美排课成功！")
        result, classes = st.session_state['schedule_result'], st.session_state['schedule_classes']
        
        col1, col2 = st.columns([1, 3])
        with col1:
            st.download_button(
                label="📊 导出课表到 Excel",
                data=export_course_to_excel(result, classes),
                file_name=f"排课结果_{class_prefix}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            components.html("""<button onclick="window.parent.print()" style="padding: 8px 15px; background-color: #FF4B4B; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 14px;">🖨️ 打印课表 / 存为彩色 PDF</button>""", height=50)

        tabs = st.tabs(classes)
        for idx, c_name in enumerate(classes):
            with tabs[idx]:
                st.markdown(render_class_table(c_name, result[c_name]), unsafe_allow_html=True)

# =====================================================================
#                          功能模块：全科考场编排
# =====================================================================
elif app_mode == "📝 全科考场编排":
    with st.sidebar:
        st.header("⚙️ 考务参数设置")
        room_capacity = st.number_input("标准考场容量 (人数)", min_value=10, max_value=60, value=30, step=1)
        st.markdown("---")
        st.header("🖨️ 打印排版设置")
        slips_per_page = st.selectbox("每页打印学生数 (Word导出)", [4, 6, 8], index=1, help="推荐选 6 人，排版紧凑且自带边框裁切线")

    class ExamScheduler:
        def __init__(self, student_df, capacity):
            self.student_data = student_df.fillna("")
            self.capacity = int(capacity)
            self.fixed_subjects = {}   
            self.dynamic_subjects = {} 

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
                if zkz not in student_slips:
                    student_slips[zkz] = {'姓名': name, '行政班': cls, '准考证号': zkz, 'exams': []}
                student_slips[zkz]['exams'].append({'科目': subject, '考场': room_name, '座位': seat_name})

            for subject, students in self.fixed_subjects.items():
                arranged_list = []
                for stu in students:
                    room_name = f"第{stu['原考场']}考场" if stu['原考场'] != "00" else "未分配考场"
                    seat_name = stu['原座位']
                    arranged_list.append({'考场号': room_name, '座位号': seat_name, '准考证号': stu['准考证号'], '姓名': stu['姓名'], '行政班': stu['行政班']})
                    add_to_slip(stu['准考证号'], stu['姓名'], stu['行政班'], subject, room_name, seat_name)
                df = pd.DataFrame(arranged_list).sort_values(by=['考场号', '座位号'])
                exam_results[subject] = df

            for subject, students in self.dynamic_subjects.items():
                students_sorted = sorted(students, key=lambda x: (x['原考场'], x['原座位']))
                arranged_list = []
                for i, stu in enumerate(students_sorted):
                    room_idx, seat_idx = (i // self.capacity) + 1, (i % self.capacity) + 1
                    room_name, seat_name = f"第{room_idx:02d}考场", f"{seat_idx:02d}"
                    arranged_list.append({'考场号': room_name, '座位号': seat_name, '准考证号': stu['准考证号'], '姓名': stu['姓名'], '行政班': stu['行政班']})
                    add_to_slip(stu['准考证号'], stu['姓名'], stu['行政班'], subject, room_name, seat_name)
                df = pd.DataFrame(arranged_list).sort_values(by=['考场号', '座位号'])
                exam_results[subject] = df
                
            slip_export_data = []
            for zkz, info in student_slips.items():
                row = {'准考证号': zkz, '姓名': info['姓名'], '行政班': info['行政班']}
                for idx, ex in enumerate(info['exams']):
                    row[f'科目{idx+1}'] = ex['科目']; row[f'考场{idx+1}'] = ex['考场']; row[f'座位{idx+1}'] = ex['座位']
                slip_export_data.append(row)
            return exam_results, student_slips, slip_export_data

    # --- 2. Word 栅格化导出逻辑 (核心渲染对齐版) ---
    def export_slips_to_word(slips_dict, per_page):
        doc = Document()
        
        # 设置窄边距
        section = doc.sections[0]
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)

        def set_font(run, size, bold=False):
            run.font.name = '宋体'
            run.font.size = Pt(size)
            run.bold = bold
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        student_list = list(slips_dict.values())
        num_cols = 2
        num_rows = per_page // 2
        
        for i in range(0, len(student_list), per_page):
            grid_table = doc.add_table(rows=num_rows, cols=num_cols)
            grid_table.style = 'Table Grid' # <--- 添加外边框，作为完美的裁切线
            grid_table.autofit = False
            
            # 固定行高
            row_height = 10.0 / num_rows 
            for row in grid_table.rows:
                row.height = Inches(row_height)

            chunk = student_list[i : i + per_page]
            for idx, info in enumerate(chunk):
                r, c = idx // num_cols, idx % num_cols
                cell = grid_table.cell(r, c)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                # 标题
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_font(p.add_run("英华学校选考条"), 14, True)
                
                # 基本信息
                p2 = cell.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_font(p2.add_run(f"姓名：{info['姓名']}   班级：{info['行政班']}"), 11)
                
                p3 = cell.add_paragraph()
                p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_font(p3.add_run(f"考号：{info['准考证号']}"), 11)
                p3.paragraph_format.space_after = Pt(6) # 增加与表格的间隙

                # 内嵌考试信息表格
                inner_table = cell.add_table(rows=1, cols=3)
                inner_table.style = 'Table Grid'
                inner_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                inner_table.columns[0].width = Inches(0.9)
                inner_table.columns[1].width = Inches(1.2)
                inner_table.columns[2].width = Inches(0.7)

                hdr_cells = inner_table.rows[0].cells
                for j, txt in enumerate(['科目', '考场', '座号']):
                    cp = hdr_cells[j].paragraphs[0]
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_font(cp.add_run(txt), 10, True)
                    # 为表头注入浅灰色背景，对齐网页UI效果
                    shading_elm = parse_xml(r'<w:shd {} w:fill="F0F0F0"/>'.format(nsdecls('w')))
                    hdr_cells[j]._tc.get_or_add_tcPr().append(shading_elm)

                for ex in info['exams']:
                    row_cells = inner_table.add_row().cells
                    for j, val in enumerate([ex['科目'], ex['考场'], ex['座位']]):
                        cp = row_cells[j].paragraphs[0]
                        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        set_font(cp.add_run(val), 10)

            if i + per_page < len(student_list):
                doc.add_page_break()

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    def export_exam_to_excel(results_dict, slip_export_data):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for subject, df in results_dict.items(): df.to_excel(writer, sheet_name=f"{subject}考场", index=False)
            pd.DataFrame(slip_export_data).to_excel(writer, sheet_name="学生全科汇总", index=False)
        return output.getvalue()

    # --- 3. 排考 UI ---
    st.title("📝 全科考场编排与 Word 准考条生成")
    uploaded_exam = st.file_uploader("📂 上传【考生信息表】", type=['xlsx', 'xls', 'csv'], key="exam_upload")

    if uploaded_exam:
        student_df = pd.read_csv(uploaded_exam, dtype=str) if uploaded_exam.name.endswith('.csv') else pd.read_excel(uploaded_exam, dtype=str)
        
        if st.button("🚀 生成全科考场数据", type="primary"):
            with st.spinner("正在进行排位计算..."):
                scheduler = ExamScheduler(student_df, room_capacity)
                results, slips, slip_export = scheduler.arrange()
                st.session_state['exam_result'], st.session_state['exam_slips'], st.session_state['slip_export_data'] = results, slips, slip_export

    if st.session_state['exam_result'] is not None:
        results = st.session_state['exam_result']
        slips = st.session_state['exam_slips']
        subjects = list(results.keys())
        st.success("✅ 考场编排完毕！")
        
        # 导出按钮区
        col1, col2 = st.columns([1, 1])
        with col1:
            st.download_button(
                label="📊 导出全科考务大表 (Excel)",
                data=export_exam_to_excel(results, st.session_state['slip_export_data']),
                file_name="英华学校_期末全科考场汇总.xlsx"
            )
        with col2:
            word_file = export_slips_to_word(slips, slips_per_page)
            st.download_button(
                label=f"📄 导出考生小票 (Word版 · 自带裁切线)",
                data=word_file,
                file_name=f"英华个人考场条_{slips_per_page}人版.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        view_mode = st.radio("👀 切换查看模式：", ["🎫 考生小票排版预览", "📋 各科目考场门贴大表"], horizontal=True)
        if "预览" in view_mode:
            def render_html(slips_dict):
                html = '<div style="font-family:SimSun; display:flex; flex-wrap:wrap; justify-content:center; gap:20px;">'
                for zkz, info in slips_dict.items():
                    html += f'<div style="border:1px dashed #666; padding:10px; width:280px; margin-bottom:10px;">'
                    html += f'<h5 style="text-align:center; margin:0 0 8px 0; font-size:16px;">英华学校选考条</h5>'
                    html += f'<p style="font-size:13px; text-align:center; margin:0 0 8px 0; line-height: 1.5;">姓名：{info["姓名"]} &nbsp;&nbsp; 班级：{info["行政班"]}<br>考号：{info["准考证号"]}</p>'
                    html += '<table border="1" style="width:100%; font-size:12px; border-collapse:collapse; text-align:center; border-color:#333;">'
                    html += '<tr style="background-color:#f0f0f0;"><th>科目</th><th>考场</th><th>座号</th></tr>'
                    for ex in info['exams']: html += f"<tr><td>{ex['科目']}</td><td>{ex['考场']}</td><td>{ex['座位']}</td></tr>"
                    html += '</table></div>'
                return html + '</div>'
            st.markdown(render_html(slips), unsafe_allow_html=True)
        else:
            tabs = st.tabs(subjects)
            for idx, subject in enumerate(subjects):
                with tabs[idx]:
                    st.dataframe(results[subject], use_container_width=True, hide_index=True)