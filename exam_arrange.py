import streamlit as st
import pandas as pd
import math
import io
import streamlit.components.v1 as components

st.set_page_config(page_title="智能考场编排系统", layout="wide")

# ================= 初始化缓存记忆 =================
if 'exam_result' not in st.session_state: st.session_state['exam_result'] = None
if 'exam_subjects' not in st.session_state: st.session_state['exam_subjects'] = None

# ================= 核心排考引擎 =================
class ExamScheduler:
    def __init__(self, student_df, capacity):
        self.student_data = student_df.fillna("")
        self.capacity = int(capacity)
        self.subject_students = {}  # {科目名: [学生数据列表]}

    def parse_data(self):
        """解析学生数据，归类每个科目的参考名单"""
        for _, row in self.student_data.iterrows():
            # 提取基础信息
            student_info = {
                '准考证号': row.get('准考证号', ''),
                '姓名': row.get('姓名', ''),
                '行政班': row.get('行政班', ''),
                '身份证号': row.get('身份证号', '')
            }
            
            # 提取该学生需要考的所有科目
            # 默认公共科
            subjects = ['语文', '数学']
            
            # 提取表格中的动态科目
            for col in ['语种', '科类', '选考1', '选考2']:
                sub = str(row.get(col, '')).strip()
                if sub and sub != 'nan':
                    subjects.append(sub)
                    
            # 归入各个科目的参考池
            for sub in set(subjects):
                if sub not in self.subject_students:
                    self.subject_students[sub] = []
                self.subject_students[sub].append(student_info)

    def arrange(self):
        """执行蛇形防作弊排考算法"""
        self.parse_data()
        exam_results = {}
        
        for subject, students in self.subject_students.items():
            # 1. 按照行政班排序（确保同班同学聚在一起，方便后续打散）
            students_sorted = sorted(students, key=lambda x: x['行政班'])
            
            total_students = len(students_sorted)
            if total_students == 0: continue
            
            # 2. 计算需要多少个考场
            num_rooms = math.ceil(total_students / self.capacity)
            
            # 3. 核心：生成“蛇形打散”的考场座位坑位表
            # 逻辑：先遍历座位号，再遍历考场号，像发牌一样把同班同学发到不同考场
            available_spots = []
            for seat in range(1, self.capacity + 1):
                for room in range(1, num_rooms + 1):
                    available_spots.append((room, seat))
            
            # 4. 把学生填入坑位
            arranged_list = []
            for i, student in enumerate(students_sorted):
                room_no, seat_no = available_spots[i]
                arranged_list.append({
                    '考场号': f"{room_no:02d}", # 格式化为 01, 02
                    '座位号': f"{seat_no:02d}",
                    '准考证号': student['准考证号'],
                    '姓名': student['姓名'],
                    '行政班': student['行政班'],
                    '考试科目': subject
                })
            
            # 5. 按照 考场号 -> 座位号 重新排序，生成最终该科目的考场表
            arranged_list = sorted(arranged_list, key=lambda x: (x['考场号'], x['座位号']))
            exam_results[subject] = pd.DataFrame(arranged_list)
            
        return exam_results

# ================= 导出工具 =================
def export_to_excel(results_dict):
    """导出为包含多个科目Sheet的Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for subject, df in results_dict.items():
            df.to_excel(writer, sheet_name=subject, index=False)
    return output.getvalue()

# ================= 前端交互 =================
st.title("🏫 英华学校智能考场编排系统 (防作弊发牌版)")

with st.sidebar:
    st.header("⚙️ 考场规则设置")
    room_capacity = st.number_input("每个考场最大人数", min_value=10, max_value=60, value=30, step=1)
    st.caption("💡 提示：系统采用【蛇形发牌法】，会自动打散同行政班的学生，有效防止作弊。")

uploaded_file = st.file_uploader("请上传考生信息 Excel 文件", type=['xlsx', 'xls', 'csv'])

if uploaded_file:
    # 兼容 CSV 和 Excel
    if uploaded_file.name.endswith('.csv'):
        student_df = pd.read_csv(uploaded_file)
    else:
        student_df = pd.read_excel(uploaded_file)
        
    st.success(f"✅ 成功读取 {len(student_df)} 名考生数据！")
    
    if st.button("🚀 一键生成所有科目考场表", type="primary"):
        with st.spinner("正在提取走班科目，执行防作弊分配算法..."):
            scheduler = ExamScheduler(student_df, room_capacity)
            result = scheduler.arrange()
            
            if not result:
                st.error("❌ 数据解析失败，请检查表格是否有'语种'、'科类'、'选考1'等列名。")
            else:
                st.session_state['exam_result'] = result
                st.session_state['exam_subjects'] = list(result.keys())

# ================= 结果展示与下载 =================
if st.session_state['exam_result'] is not None:
    result = st.session_state['exam_result']
    subjects = st.session_state['exam_subjects']
    
    st.success("✅ 考场编排完毕！由于走班制，每门科目的考场分布均不相同。")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        excel_data = export_to_excel(result)
        st.download_button(
            label="📊 导出全科考场表 (Excel)",
            data=excel_data,
            file_name="英华学校_期末考场编排总表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        components.html("""
            <button onclick="window.parent.print()" style="padding: 8px 15px; background-color: #28a745; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 14px;">
                🖨️ 打印当前查看的科目门贴
            </button>
        """, height=50)

    st.markdown("### 👁️ 考场预览")
    # 创建标签页供教务处查看不同科目
    tabs = st.tabs(subjects)
    for idx, subject in enumerate(subjects):
        with tabs[idx]:
            df = result[subject]
            # 统计数据面板
            room_count = df['考场号'].nunique()
            student_count = len(df)
            st.info(f"📍 **{subject}** 共有 **{student_count}** 人参考，共需安排 **{room_count}** 个考场。")
            
            # 交互式表格展示
            st.dataframe(df, use_container_width=True, height=400)