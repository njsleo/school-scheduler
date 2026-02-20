import streamlit as st
import pandas as pd
import random
import re
import copy
import io
import streamlit.components.v1 as components

# ================= 1. 标准时间网格 =================
DAYS = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
SLOTS = [
    "早自习\n7:00-7:40",   # 0
    "第一节\n8:10-8:55",   # 1
    "第二节\n9:05-9:50",   # 2
    "第三节\n10:15-11:00", # 3
    "第四节\n11:00-11:45", # 4
    "第五节\n14:00-14:45", # 5 
    "第六节\n14:55-15:40", # 6 
    "第七节\n15:50-16:35", # 7
    "第八节\n16:45-18:15", # 8
    "第九节\n19:00-20:30", # 9  (晚修1)
    "第十节\n20:50-22:30"  # 10 (晚修2)
]
DAYTIME_SLOTS = SLOTS[1:9]

# ================= 专属多巴胺配色卡 =================
SUBJECT_COLORS = {
    "走班": "#FFFF99", "语文": "#FFE4E1", "数学": "#E4F1FF", "英语": "#E0FFFF", 
    "物理": "#E8F5E9", "化学": "#FFFACD", "生物": "#FCE4EC", "政治": "#FFDAB9", 
    "历史": "#FFE4B5", "地理": "#E1BEE7", "体育": "#C8E6C9", "音乐": "#E1BEE7", 
    "美术": "#FFCCBC", "班会": "#B3E5FC", "晚自习": "#F8F9FA", "自习": "#F8F9FA",
    "信息": "#E0F7FA", "通用": "#F5F5F5"
}

st.set_page_config(page_title="全校通用排课系统", layout="wide")

# ================= 初始化缓存记忆 =================
if 'schedule_result' not in st.session_state: st.session_state['schedule_result'] = None
if 'schedule_classes' not in st.session_state: st.session_state['schedule_classes'] = None

# ================= 侧边栏：全局规则配置中心 =================
with st.sidebar:
    st.header("⚙️ 引擎参数设置")
    
    st.markdown("### 1. 班级名称配置")
    class_prefix = st.text_input("班级前缀 (如: 高一 / 高三理科)", value="高一")
    class_suffix = st.text_input("班级后缀", value="班")
    st.caption("💡 提示：若 Excel 填了 1,2,3，系统将自动生成：前缀+1+后缀")
    
    st.markdown("### 2. 年级规则模式")
    grade_mode = st.radio(
        "选择作息规则", 
        ["高一 / 高二 (仅周一至周五晚修)", "高三 (包含周日晚修)"]
    )
    st.caption("💡 提示：高一高二模式下，系统会自动避开周末排课。")

# ================= 2. 核心排课引擎 =================
class Scheduler:
    def __init__(self, teacher_df, prefix, suffix, grade_mode):
        self.teacher_data = teacher_df.fillna("") 
        self.classes = set()
        self.teacher_busy = {} 
        self.prefix = prefix
        self.suffix = suffix
        
        # 动态设置晚修天数 (如果是高一高二，周末彻底休息)
        if "高三" in grade_mode:
            self.evening_days = ['周一', '周二', '周三', '周四', '周五', '周日']
        else:
            self.evening_days = ['周一', '周二', '周三', '周四', '周五']

    def parse_tasks(self):
        tasks = []
        zouban_tasks = []
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
            if hours == 0 and len(row_dict) >= 4:
                try: hours = int(float(re.findall(r'\d+', list(row_dict.values())[3])[0]))
                except: pass

            if hours <= 0 or not t_name or not raw_classes: continue

            # === 动态扩展班级：自动拼接前缀和后缀 ===
            class_nums = re.findall(r'\d+', raw_classes)
            target_classes = [f"{self.prefix}{n}{self.suffix}" for n in class_nums] 
            for c in target_classes: self.classes.add(c)

            if "Block_ZouBan" in constraint or "走班" in sub or "走班" in t_name:
                zouban_tasks.append({'classes': target_classes, 'teacher': t_name, 'subject': sub, 'hours': hours})
            else:
                for c in target_classes:
                    tasks.append({'class': c, 'teacher': t_name, 'subject': sub, 'hours': hours, 'constraint': constraint})

        # === 动态排序，支持无数个班 ===
        # 使用自定义排序确保 "高一1班", "高一2班", ..., "高一10班" 的顺序正确
        def sort_key(c_str):
            nums = re.findall(r'\d+', c_str)
            return int(nums[0]) if nums else 0
            
        self.classes = sorted(list(self.classes), key=sort_key) 
        if not self.classes: self.classes = [f'{self.prefix}1{self.suffix}'] 
        return tasks, zouban_tasks

    def get_booked_count(self, c_name, subject, day):
        count = 0
        for s in SLOTS:
            content = self.schedules[c_name][day][s]
            if content == "": continue
            if "走班" in subject and "走班" in content: count += 1
            elif content.startswith(f"{subject}<br>"): count += 1
        return count

    def is_free_for(self, c_name, teacher, subject, day, slot_list):
        for slot in slot_list:
            if self.schedules[c_name][day][slot] != "": return False 
            if (teacher, day, slot) in self.teacher_busy: return False 
        
        if self.get_booked_count(c_name, subject, day) + len(slot_list) > 2: return False
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
        
        best_schedule = None
        best_score = 99999
        
        for attempt in range(100):
            tasks = copy.deepcopy(self.original_tasks)
            zouban_tasks = copy.deepcopy(self.original_zouban)
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
                        z_task['hours'] -= 2
                        placed = True
                        break
                if not placed:
                    zouban_eve_ok = False; break
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

                    unassigned_tasks = [t for t in c_tasks if t['subject'] not in eve_assigned[c]]
                    assigned_tasks = [t for t in c_tasks if t['subject'] in eve_assigned[c]]

                    placed = False
                    for t in unassigned_tasks + assigned_tasks:
                        if self.is_free_for(c, t['teacher'], t['subject'], day, [SLOTS[9], SLOTS[10]]):
                            self.book(c, t['teacher'], t['subject'], day, SLOTS[9])
                            self.book(c, t['teacher'], t['subject'], day, SLOTS[10])
                            t['hours'] -= 2
                            eve_assigned[c].add(t['subject'])
                            placed = True
                            break
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
                                z_task['hours'] -= 2
                                placed = True
                                break
                        if placed: break
                    if not placed: break
                
                while z_task['hours'] > 0:
                    placed = False
                    all_times = [(d, s) for d in DAYS[:5] for s in [SLOTS[5], SLOTS[6]]]
                    random.shuffle(all_times)
                    for day, slot in all_times:
                        if all(self.is_free_for(c, z_task['teacher'], z_task['subject'], day, [slot]) for c in z_task['classes']):
                            for c in z_task['classes']:
                                self.book(c, z_task['teacher'], z_task['subject'], day, slot)
                            z_task['hours'] -= 1
                            placed = True
                            break
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
                                task['hours'] -= 2
                                placed = True
                                break 
                        if placed: break 
                    if not placed: break 

            for task in tasks:
                while task['hours'] > 0:
                    placed = False
                    all_days = DAYS[:5]
                    random.shuffle(all_days)
                    for day in all_days:
                        if self.is_free_for(task['class'], task['teacher'], task['subject'], day, [SLOTS[5]]):
                            self.book(task['class'], task['teacher'], task['subject'], day, SLOTS[5])
                            task['hours'] -= 1
                            placed = True
                            break
                    
                    if not placed:
                        all_times = [(d, s) for d in DAYS[:5] for s in [SLOTS[1], SLOTS[2], SLOTS[3], SLOTS[4], SLOTS[7], SLOTS[8]]]
                        random.shuffle(all_times)
                        for day, slot in all_times:
                            if self.is_free_for(task['class'], task['teacher'], task['subject'], day, [slot]):
                                self.book(task['class'], task['teacher'], task['subject'], day, slot)
                                task['hours'] -= 1
                                placed = True
                                break
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

# ================= 3. HTML 视觉渲染引擎 =================
def render_class_table(class_name, schedule):
    html = f"""
    <div id="table-{class_name}" style="font-family: 'SimSun', '宋体', sans-serif; max-width: 1050px; margin: auto; background-color: white; padding: 20px;">
        <h2 style="text-align: center; color: red; letter-spacing: 2px;">呼和浩特市英华学校 {class_name} 课程表</h2>
        <table border="1" style="width: 100%; text-align: center; border-collapse: collapse; border-color: #333; font-size: 16px;">
            <tr style="background-color: #f8f9fa;">
                <th style="padding: 12px;">时间 \\ 星期</th>
                <th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th>六</th><th>日</th>
            </tr>
    """

    def get_cell(day, slot):
        content = schedule[day][slot]
        bg_color = "#FFFFFF" 
        for sub, color in SUBJECT_COLORS.items():
            if sub in content: bg_color = color; break
        if "走班" in content:
            return f'<td style="background-color: {bg_color}; font-weight: bold; padding: 12px; box-shadow: inset 0 0 5px rgba(0,0,0,0.1);">{content}</td>'
        return f'<td style="background-color: {bg_color}; padding: 12px;">{content}</td>'

    html += "<tr><td>早自习(7:00-7:40)</td>"
    for day in DAYS[:5]: html += "<td style='padding: 12px; background-color: #F8F9FA;'>早自习</td>"
    html += "<td rowspan='12' style='width:45px; background-color: #F8F9FA; color: #555;'>考<br><br>试</td>"
    html += "<td rowspan='12' style='width:45px; background-color: #F8F9FA; color: #555;'>休<br><br>息</td></tr>" 

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
        html += "<td style='padding: 12px; background-color: #F8F9FA;'>休息</td>"
        html += get_cell('周日', slot)
        html += "</tr>"

    html += "</table></div><br>"
    return html

# ================= 4. 导出工具集 =================
def export_to_excel(schedules, classes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for c_name in classes:
            data = []
            for slot in SLOTS:
                row = [slot.replace('\n', ' ')] 
                for day in DAYS:
                    cell_content = schedules[c_name][day][slot].replace("<br>", "\n")
                    row.append(cell_content)
                data.append(row)
            df = pd.DataFrame(data, columns=['时间段'] + DAYS)
            df.to_excel(writer, sheet_name=c_name, index=False)
    return output.getvalue()

# ================= 5. 前端交互 =================
st.title("🏫 英华学校智能排课系统 (全校通用版)")
uploaded_file = st.file_uploader("请先在左侧侧边栏设置年级和班级前缀，然后上传排课数据 Excel", type=['xlsx', 'xls'])

if uploaded_file:
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    teacher_df = pd.DataFrame()
    
    sheet_names = list(all_sheets.keys())
    if len(sheet_names) == 1: teacher_df = all_sheets[sheet_names[0]]
    else:
        for name, df in all_sheets.items():
            if len(df) > 5 and len(df.columns) >= 3: teacher_df = df

    if st.button("🚀 根据当前设置生成课表", type="primary"):
        with st.spinner(f"正在以【{grade_mode}】规则进行排课计算..."):
            scheduler = Scheduler(teacher_df, class_prefix, class_suffix, grade_mode)
            result, total_tasks = scheduler.run()
            
            if total_tasks == 0:
                st.error("❌ 未抓取到数据，请检查 Excel 表头。")
            else:
                st.session_state['schedule_result'] = result
                st.session_state['schedule_classes'] = scheduler.classes

if st.session_state['schedule_result'] is not None:
    st.success("✅ 完美排课成功！请在下方查看结果并导出。")
    
    result = st.session_state['schedule_result']
    classes = st.session_state['schedule_classes']
    
    col1, col2 = st.columns([1, 3])
    with col1:
        excel_data = export_to_excel(result, classes)
        st.download_button(
            label="📊 导出所有班级到 Excel",
            data=excel_data,
            file_name=f"排课结果_{class_prefix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        components.html("""
            <button onclick="window.parent.print()" style="padding: 8px 15px; background-color: #FF4B4B; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; font-family: sans-serif;">
                🖨️ 打印课表 / 保存为彩色 PDF
            </button>
        """, height=50)

    tabs = st.tabs(classes)
    for idx, c_name in enumerate(classes):
        with tabs[idx]:
            final_html = render_class_table(c_name, result[c_name])
            st.markdown(final_html, unsafe_allow_html=True)