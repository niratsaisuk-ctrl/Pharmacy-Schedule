import streamlit as st
import pandas as pd
import io
from ortools.sat.python import cp_model

# ==========================================
# ⚙️ ส่วนที่ 1: ฟังก์ชันคำนวณตาราง (AI Logic)
# ==========================================
VALID_TIMES = ["08.30", "09.00", "09.30", "10.00", "10.30", "11.00", "11.30", "12.00",
               "12.30", "13.00", "13.30", "14.00", "14.30", "15.00", "15.30", "16.00", "16.30"]

def time_to_slot(t_str):
    return VALID_TIMES.index(t_str)

def generate_schedule(DAY_OF_WEEK, LEAVES, CUSTOM_TASKS, PART_TIME):
    model = cp_model.CpModel()
    
    ft_pharmacists = ['เต้น', 'แอน', 'แม็ค', 'โบ้ท', 'ไม้เอก', 'กิ๊ฟ', 'ฟอร์จูน', 'มิ้ลค์', 'ริน', 
                      'อ๊อฟฟี่', 'ออย', 'บี', 'มายด์', 'ขิม', 'บีม', 'มิ้น', 'ใบเตย', 'จีน่า', 'ปอนด์']
    pt_pharmacists = [pt['name'] for pt in PART_TIME]
    all_pharmacists = ft_pharmacists + pt_pharmacists
    
    time_slots = [f"{h}:{m:02d}-{h + (m+30)//60}:{(m+30)%60:02d}" 
                  for h in range(8, 16) for m in (30, 0) if not (h == 8 and m == 0) and not (h == 16 and m == 30)]
    time_slots.append("16:00-16:30")
    
    dispensing_tasks = ['จ่ายยา_4', 'จ่ายยา_5', 'จ่ายยา_6', 'จ่ายยา_7', 'จ่ายยา_8', 'จ่ายยา_9', 'จ่ายยา_10', 'จ่ายยา_11']
    ver_cpoe_tasks = ['Ver_1', 'Ver_2', 'Ver_3', 'Ver_4', 'Ver_5', 'Ver_6']
    ver_ps_tasks = ['PS_1', 'PS_2', 'PS_3', 'PS_4', 'PS_5']
    tasks = dispensing_tasks + ver_cpoe_tasks + ver_ps_tasks + ['Match_C', 'พัก', 'งานเฉพาะ', 'ลา', 'นอกเวลา', 'ว่าง']
             
    x = {}
    for p in all_pharmacists:
        for t in range(len(time_slots)):
            for task in tasks:
                x[p, t, task] = model.NewBoolVar(f'x_{p}_{t}_{task}')
                
    for p in all_pharmacists:
        for t in range(len(time_slots)):
            model.AddExactlyOne(x[p, t, task] for task in tasks)

    if DAY_OF_WEEK == 'Wed_Fri':
        break_slots = [6, 7, 8, 9, 10, 11] 
        peak_slots = [3, 4, 12, 13, 14, 15] 
        b_groups = [(6,8), (8,10), (10,12)] 
    else:
        break_slots = [5, 6, 7, 8, 9, 10] 
        peak_slots = [3, 4, 11, 12, 13, 14, 15] 
        b_groups = [(5,7), (7,9), (9,11)]

    # จัดการวันลา (แก้ไขบั๊ก Runtime Error)
    leave_slots = set()
    for p, l_type in LEAVES.items():
        if l_type == 'ทั้งวัน': l_range = range(0, 16)
        elif l_type == 'เช้า': l_range = range(0, 9)
        elif l_type == 'บ่าย': l_range = range(7, 16)
        for t in l_range:
            leave_slots.add((p, t))

    for p in all_pharmacists:
        for t in range(16):
            if (p, t) in leave_slots:
                model.Add(x[p, t, 'ลา'] == 1)
            else:
                model.Add(x[p, t, 'ลา'] == 0)

    # จัดการงานเฉพาะราย
    custom_dict_index = {}
    for (p, start, end), task_name in CUSTOM_TASKS.items():
        s_idx, e_idx = time_to_slot(start), time_to_slot(end)
        for t in range(s_idx, e_idx):
            model.Add(x[p, t, 'งานเฉพาะ'] == 1)
            custom_dict_index[(p, t)] = task_name
            
    for p in all_pharmacists:
        for t in range(16):
            if (p, t) not in custom_dict_index:
                model.Add(x[p, t, 'งานเฉพาะ'] == 0)

    # จัดการ Part-time
    pt_dispense_allowed = ['จ่ายยา_7', 'จ่ายยา_8']
    if len(PART_TIME) > 2: pt_dispense_allowed.extend(['จ่ายยา_6', 'จ่ายยา_9'])
    pt_all_allowed = pt_dispense_allowed + ['Match_C', 'พัก', 'นอกเวลา']

    for pt in PART_TIME:
        p = pt['name']
        s_idx, e_idx = time_to_slot(pt['start']), time_to_slot(pt['end'])
        work_slots_count = e_idx - s_idx
        
        for t in range(16):
            if t < s_idx or t >= e_idx: model.Add(x[p, t, 'นอกเวลา'] == 1)
            else: model.Add(x[p, t, 'นอกเวลา'] == 0)
            model.Add(sum(x[p, t, task] for task in pt_all_allowed) == 1)

        # แก้ไขบั๊กเวลาพัก Part-time
        if pt['has_break'] and s_idx <= 8 and e_idx > 8:
            model.Add(x[p, 8, 'พัก'] == 1) 
            if e_idx > 9: model.Add(x[p, 9, 'Match_C'] == 1) 
            work_slots_count -= 1 
            
        hours = work_slots_count / 2.0
        disp_vars_7 = [x[p, t, 'จ่ายยา_7'] for t in range(max(0, s_idx), min(16, e_idx))]
        disp_vars_8 = [x[p, t, 'จ่ายยา_8'] for t in range(max(0, s_idx), min(16, e_idx))]
        
        if hours > 0:
            if hours <= 4.0:
                model.Add(sum(disp_vars_7) == 2)
                model.Add(sum(disp_vars_8) == 2)
            elif hours <= 5.0:
                model.Add(sum(disp_vars_7) + sum(disp_vars_8) == 6)
                model.Add(sum(disp_vars_7) >= 2)
                model.Add(sum(disp_vars_8) >= 2)
            else:
                model.Add(sum(disp_vars_7) + sum(disp_vars_8) == 8)
                model.Add(sum(disp_vars_7) >= 2)
                model.Add(sum(disp_vars_8) >= 2)

    # กฎ Full-Time
    for p in ft_pharmacists:
        model.Add(sum(x[p, t, 'นอกเวลา'] for t in range(16)) == 0) 
        
        if p not in LEAVES:
            model.Add(sum(x[p, t, 'พัก'] for t in range(16)) == 2)
            b1 = model.NewBoolVar(f'b1_{p}') 
            b2 = model.NewBoolVar(f'b2_{p}') 
            b3 = model.NewBoolVar(f'b3_{p}') 
            model.AddExactlyOne([b1, b2, b3]) 
            
            for t in range(*b_groups[0]): model.Add(x[p, t, 'พัก'] == 1).OnlyEnforceIf(b1)
            for t in range(*b_groups[1]): model.Add(x[p, t, 'พัก'] == 1).OnlyEnforceIf(b2)
            for t in range(*b_groups[2]): model.Add(x[p, t, 'พัก'] == 1).OnlyEnforceIf(b3)
        else:
            model.Add(sum(x[p, t, 'พัก'] for t in range(16)) == 0)

        for t in range(16):
            if t not in break_slots: model.Add(x[p, t, 'พัก'] == 0)

    # กฎตำแหน่งและซ้ำซ้อน
    for t in range(16):
        for task in tasks:
            if task not in ['พัก', 'งานเฉพาะ', 'ลา', 'นอกเวลา', 'ว่าง']:
                model.Add(sum(x[p, t, task] for p in all_pharmacists) <= 1)

        if t < 2: req_dispense = ['จ่ายยา_6', 'จ่ายยา_7', 'จ่ายยา_8', 'จ่ายยา_9']
        else: req_dispense = ['จ่ายยา_5', 'จ่ายยา_6', 'จ่ายยา_7', 'จ่ายยา_8', 'จ่ายยา_9', 'จ่ายยา_10']
            
        for task in req_dispense + ['Ver_1', 'Ver_2', 'Ver_3', 'PS_1', 'Match_C']:
            model.Add(sum(x[p, t, task] for p in all_pharmacists) == 1)
            
        if t not in break_slots:
            model.Add(sum(x[p, t, 'PS_2'] for p in all_pharmacists) == 1)
        else:
            model.Add(sum(x[p, t, 'PS_2'] for p in all_pharmacists) <= 1)

    categories_to_prevent_internal_switch = [dispensing_tasks, ver_cpoe_tasks, ver_ps_tasks]
    work_categories = [dispensing_tasks, ver_cpoe_tasks, ver_ps_tasks, ['Match_C']]
    
    for p in all_pharmacists:
        for t in range(15):
            for cat in categories_to_prevent_internal_switch:
                for task1 in cat:
                    for task2 in cat:
                        if task1 != task2:
                            model.AddImplication(x[p, t, task1], x[p, t+1, task2].Not())
                            
        for cat in work_categories:
            for t in range(14):
                model.Add(sum(x[p, t, task] + x[p, t+1, task] + x[p, t+2, task] for task in cat) <= 2)

    for p in ft_pharmacists:
        model.Add(sum(x[p, t, task] for t in range(16) for task in dispensing_tasks) <= 6)
        b_7 = model.NewBoolVar(f'b_7_{p}')
        b_8 = model.NewBoolVar(f'b_8_{p}')
        model.Add(sum(x[p, t, 'จ่ายยา_7'] for t in range(16)) <= 2 * b_7)
        model.Add(sum(x[p, t, 'จ่ายยา_8'] for t in range(16)) <= 2 * b_8)
        model.Add(b_7 + b_8 <= 1)
        
        for v in ['Ver_1', 'Ver_2', 'Ver_3', 'Match_C']:
            model.Add(sum(x[p, t, v] for t in range(16)) <= 2)

    reward_vars = []
    for p in all_pharmacists:
        for t in range(15):
            for task in tasks:
                if task in ['พัก', 'งานเฉพาะ', 'ลา', 'นอกเวลา', 'ว่าง']: continue
                match_var = model.NewBoolVar(f'match_{p}_{t}_{task}')
                model.AddImplication(match_var, x[p, t, task])
                model.AddImplication(match_var, x[p, t+1, task])
                reward_vars.append(match_var * 50000)

    for t in range(16):
        if t in peak_slots:
            weights = {'จ่ายยา_4': 8000, 'จ่ายยา_11': 7000, 'Ver_4': 6000, 'PS_3': 5000, 'PS_4': 4000, 'Ver_5': 3000, 'PS_5': 2000, 'Ver_6': 1000}
        else:
            weights = {'Ver_4': 8000, 'PS_3': 7000, 'PS_4': 6000, 'Ver_5': 5000, 'PS_5': 4000, 'Ver_6': 3000, 'จ่ายยา_4': 2000, 'จ่ายยา_11': 1000}

        if t in break_slots: weights['PS_2'] = 10000 
        for task, weight in weights.items():
            for p in all_pharmacists:
                reward_vars.append(x[p, t, task] * weight)
                
    for p in ft_pharmacists:
        for t in range(16):
            reward_vars.append(x[p, t, 'ว่าง'] * -5000)

    model.Maximize(sum(reward_vars))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0 
    
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        schedule_data = []
        for p in all_pharmacists:
            row_data = {'ชื่อเภสัชกร': p}
            for t in range(16):
                for task in tasks:
                    if solver.Value(x[p, t, task]) == 1:
                        display_task = task.replace('_', ' ')
                        if task == 'งานเฉพาะ': display_task = custom_dict_index.get((p, t), 'งานเฉพาะ')
                        if task == 'นอกเวลา' or task == 'ว่าง': display_task = '-'
                        row_data[time_slots[t]] = display_task
            schedule_data.append(row_data)
        return pd.DataFrame(schedule_data), "Success"
    else:
        return None, "Infeasible"

# ==========================================
# 🖥️ ส่วนที่ 2: สร้างหน้าเว็บ UI ด้วย Streamlit
# ==========================================
st.set_page_config(page_title="Pharmacy Schedule Manager", layout="wide")
st.title("💊 โปรแกรมจัดตารางเภสัชกรห้องยาอัจฉริยะ")

ft_pharmacists_list = ['เต้น', 'แอน', 'แม็ค', 'โบ้ท', 'ไม้เอก', 'กิ๊ฟ', 'ฟอร์จูน', 'มิ้ลค์', 'ริน', 
                      'อ๊อฟฟี่', 'ออย', 'บี', 'มายด์', 'ขิม', 'บีม', 'มิ้น', 'ใบเตย', 'จีน่า', 'ปอนด์']

with st.sidebar:
    st.header("⚙️ ตั้งค่าตารางประจำวัน")
    
    day_type = st.radio("📅 เลือกประเภทวัน", ["ปกติ (จ,อ,พฤ)", "พุธ หรือ ศุกร์ (ปรับเวลาพัก)"])
    DAY_OF_WEEK = 'Wed_Fri' if day_type == "พุธ หรือ ศุกร์ (ปรับเวลาพัก)" else 'Normal'
    st.divider()
    
    st.subheader("🏖️ ผู้ที่ลาในวันนี้")
    leaves_input = {}
    leave_count = st.number_input("จำนวนคนที่ลา", min_value=0, max_value=10, value=0)
    for i in range(leave_count):
        col1, col2 = st.columns(2)
        with col1:
            p_leave = st.selectbox(f"คนที่ {i+1}", ft_pharmacists_list, key=f"l_name_{i}")
        with col2:
            t_leave = st.selectbox("ประเภท", ["ทั้งวัน", "เช้า", "บ่าย"], key=f"l_type_{i}")
        leaves_input[p_leave] = t_leave

    st.divider()

    st.subheader("📋 ภารกิจพิเศษ / งานเฉพาะราย")
    custom_tasks_input = {}
    task_count = st.number_input("จำนวนภารกิจ", min_value=0, max_value=15, value=0)
    for i in range(task_count):
        st.markdown(f"**ภารกิจที่ {i+1}**")
        c1, c2, c3 = st.columns([1.5, 1, 1])
        with c1: p_task = st.selectbox(f"ชื่อ", ft_pharmacists_list, key=f"t_name_{i}")
        with c2: s_task = st.selectbox(f"เริ่ม", VALID_TIMES, index=0, key=f"t_s_{i}")
        with c3: e_task = st.selectbox(f"สิ้นสุด", VALID_TIMES, index=2, key=f"t_e_{i}")
        n_task = st.text_input(f"ชื่องาน", "ระบุงาน", key=f"t_n_{i}")
        if VALID_TIMES.index(s_task) < VALID_TIMES.index(e_task):
            custom_tasks_input[(p_task, s_task, e_task)] = n_task
        else:
            st.error("เวลาสิ้นสุดต้องมากกว่าเวลาเริ่ม!")

    st.divider()

    st.subheader("🧑‍⚕️ เภสัชกร Part-time")
    pt_input = []
    pt_count = st.number_input("จำนวน Part-time", min_value=0, max_value=5, value=0)
    for i in range(pt_count):
        st.markdown(f"**Part-time คนที่ {i+1}**")
        c1, c2, c3 = st.columns([1.5, 1, 1])
        with c1: pt_name = st.text_input("ชื่อ PT", f"PT_{i+1}", key=f"pt_n_{i}")
        with c2: pt_s = st.selectbox("เริ่ม", VALID_TIMES, index=0, key=f"pt_s_{i}")
        with c3: pt_e = st.selectbox("สิ้นสุด", VALID_TIMES, index=16, key=f"pt_e_{i}")
        pt_b = st.checkbox("จัดพัก 12.30-13.00", value=True, key=f"pt_b_{i}")
        
        if VALID_TIMES.index(pt_s) < VALID_TIMES.index(pt_e):
            pt_input.append({'name': pt_name, 'start': pt_s, 'end': pt_e, 'has_break': pt_b})
        else:
            st.error("เวลาสิ้นสุดต้องมากกว่าเวลาเริ่ม!")

if st.button("🚀 สร้างตารางปฏิบัติงาน (ใช้เวลาคำนวณ 1-2 นาที)", type="primary"):
    with st.spinner('AI กำลังประมวลผลตารางที่สมบูรณ์ที่สุด...'):
        try:
            df_result, status = generate_schedule(DAY_OF_WEEK, leaves_input, custom_tasks_input, pt_input)
            
            if status == "Success":
                st.success("🎉 สร้างตารางเสร็จสมบูรณ์!")
                st.dataframe(df_result, use_container_width=True)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Schedule')
                
                st.download_button(
                    label="📥 ดาวน์โหลดตารางเป็น Excel",
                    data=buffer.getvalue(),
                    file_name="Pharmacy_Schedule_App.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.error("⚠️ Infeasible: ไม่สามารถจัดตารางได้! (คนไม่พอ หรือช่วงเวลางานพิเศษชนกัน) ลองลดภารกิจพิเศษลงครับ")
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในระบบคณิตศาสตร์: {str(e)}")
