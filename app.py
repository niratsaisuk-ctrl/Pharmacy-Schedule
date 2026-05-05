import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta, timezone
import locale
from ortools.sat.python import cp_model
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
import streamlit.components.v1 as components

# ==========================================
# ⚙️ ฟังก์ชันแปลงวันที่เป็นภาษาไทย (สำหรับแสดงผล)
# ==========================================
def get_thai_date(date_obj):
    thai_months = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                   "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    thai_days = ["วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์", "วันอาทิตย์"]
    day_name = thai_days[date_obj.weekday()]
    day = date_obj.day
    month = thai_months[date_obj.month]
    year = date_obj.year + 543 
    return f"{day_name}ที่ {day} {month} {year}"

# ==========================================
# ⚙️ ฟังก์ชันเลือกสีหัวตารางเวลาสำหรับ Excel
# ==========================================
def get_header_color(t_idx, day_of_week):
    if day_of_week == 'Normal':
        if t_idx in [0, 1, 3, 4, 11, 12]: return 'orange' 
        if t_idx in [2]: return 'yellow'                 
        if t_idx in [5, 6, 9, 10]: return 'pink'         
        if t_idx in [7, 8]: return 'purple'              
        if t_idx in [13, 14, 15]: return 'blue'          
    else: 
        if t_idx in [0, 1, 4, 5, 12, 13]: return 'orange' 
        if t_idx in [2, 3]: return 'yellow'              
        if t_idx in [6, 7, 10, 11]: return 'pink'        
        if t_idx in [8, 9]: return 'purple'              
        if t_idx in [14, 15]: return 'blue'              
    return None

header_color_map = {
    'orange': PatternFill(start_color='FFE6CC', end_color='FFE6CC', fill_type='solid'),
    'yellow': PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid'),
    'pink': PatternFill(start_color='F8CECC', end_color='F8CECC', fill_type='solid'),
    'purple': PatternFill(start_color='E1D5E7', end_color='E1D5E7', fill_type='solid'),
    'blue': PatternFill(start_color='DAE8FC', end_color='DAE8FC', fill_type='solid')
}
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# ==========================================
# 🧠 ฟังก์ชันคำนวณตาราง (AI Logic - Core Engine)
# ==========================================
VALID_TIMES = ["08.30", "09.00", "09.30", "10.00", "10.30", "11.00", "11.30", "12.00",
               "12.30", "13.00", "13.30", "14.00", "14.30", "15.00", "15.30", "16.00", "16.30"]

def time_to_slot(t_str): return VALID_TIMES.index(t_str)

def generate_schedule(DAY_OF_WEEK, LEAVES, CUSTOM_TASKS, PART_TIME, FIX_BREAKS, FIXED_MAIN_TASKS, SICK_PEOPLE, IS_MWF):
    model = cp_model.CpModel()
    
    # --- กำหนดรายชื่อบุคลากร ---
    ft_pharmacists = ['เต้น', 'แอน', 'แม็ค', 'โบ้ท', 'ไม้เอก', 'กิ๊ฟ', 'ฟอร์จูน', 'มิ้ลค์', 'ริน', 
                      'อ๊อฟฟี่', 'ออย', 'บี', 'มายด์', 'ขิม', 'บีม', 'มิ้น', 'ใบเตย', 'จีน่า', 'ปอนด์']
    pt_pharmacists = [pt['name'] for pt in PART_TIME]
    all_pharmacists = ft_pharmacists + pt_pharmacists
    
    time_slots = [f"{VALID_TIMES[i]}-{VALID_TIMES[i+1]}" for i in range(16)]
    
    # --- กำหนดหมวดหมู่ภาระงาน ---
    dispensing_tasks = ['จ่ายยา_4', 'จ่ายยา_5', 'จ่ายยา_6', 'จ่ายยา_7', 'จ่ายยา_8', 'จ่ายยา_9', 'จ่ายยา_10', 'จ่ายยา_11']
    ver_cpoe_tasks = ['Ver_1', 'Ver_2', 'Ver_3', 'Ver_4', 'Ver_5', 'Ver_6', 'Ver_7', 'Ver_8', 'Ver_9', 'Ver_10']
    ver_ps_tasks = ['PS_1', 'PS_2', 'PS_3', 'PS_4', 'PS_5', 'PS_6', 'PS_7', 'PS_8', 'PS_9', 'PS_10']
    tasks = dispensing_tasks + ver_cpoe_tasks + ver_ps_tasks + ['Match_C', 'Match_C2', 'Matching', 'พัก', 'งานเฉพาะ', 'ลา', 'นอกเวลา', 'ว่าง']
             
    # --- 1. สร้างตัวแปรการตัดสินใจ ---
    x = {}
    for p in all_pharmacists:
        for t in range(16):
            for task in tasks: 
                x[p, t, task] = model.NewBoolVar(f'x_{p}_{t}_{task}')
            model.AddExactlyOne(x[p, t, task] for task in tasks)
            # ปิดตายช่องว่าง: บังคับว่าต้องไม่มีใครตกหล่นหรือขึ้น 'ว่าง'
            model.Add(x[p, t, 'ว่าง'] == 0)

    # --- กำหนดเวลาพักกลางวันตามวัน ---
    if DAY_OF_WEEK == 'Wed_Fri': break_slots, b_groups = [6, 7, 8, 9, 10, 11], [(6,8), (8,10), (10,12)] 
    else: break_slots, b_groups = [5, 6, 7, 8, 9, 10], [(5,7), (7,9), (9,11)]

    # --- 2. จัดการวันลาของบุคลากร ---
    active_ft = []
    leave_slots = set()
    half_day_leaves = set() 
    for p in ft_pharmacists:
        if p in LEAVES:
            l_type = LEAVES[p]
            l_range = range(0,16) if l_type == 'ทั้งวัน' else (range(0,9) if l_type == 'เช้า' else range(7,16))
            if l_type != 'ทั้งวัน': 
                half_day_leaves.add(p)
                active_ft.append(p)
            for t in l_range: 
                model.Add(x[p, t, 'ลา'] == 1)
                leave_slots.add((p, t))
        else: active_ft.append(p)

    for p in ft_pharmacists:
        for t in range(16):
            if (p, t) not in leave_slots: model.Add(x[p, t, 'ลา'] == 0)

    # --- 3. จัดการภาระงานเฉพาะราย ---
    custom_dict_index = {}
    custom_task_slots_count = {p: 0 for p in ft_pharmacists} 
    for (p, start, end), task_name in CUSTOM_TASKS.items():
        s_idx, e_idx = time_to_slot(start), time_to_slot(end)
        for t in range(s_idx, e_idx):
            model.Add(x[p, t, 'งานเฉพาะ'] == 1)
            custom_dict_index[(p, t)] = task_name
            if p in ft_pharmacists: custom_task_slots_count[p] += 1
            
    for p in all_pharmacists:
        for t in range(16):
            if (p, t) not in custom_dict_index: model.Add(x[p, t, 'งานเฉพาะ'] == 0)

    # --- 4. จัดการคนที่ป่วย ---
    for p in SICK_PEOPLE:
        if p in all_pharmacists:
            for t in range(16):
                for task in dispensing_tasks: model.Add(x[p, t, task] == 0)

    # --- 5. ล็อกภาระงานหลัก (Manual Assignment) ---
    for (p, start, end), task_name in FIXED_MAIN_TASKS.items():
        s_idx, e_idx = time_to_slot(start), time_to_slot(end)
        for t in range(s_idx, e_idx): model.Add(x[p, t, task_name] == 1)

    # --- 6. จัดการตารางงาน Part-time (PT) ---
    for pt in PART_TIME:
        p = pt['name']
        s_idx, e_idx = time_to_slot(pt['start']), time_to_slot(pt['end'])
        
        my_dispense_allowed = ['จ่ายยา_7', 'จ่ายยา_8']
        if len(PART_TIME) >= 2: my_dispense_allowed.extend(['จ่ายยา_6', 'จ่ายยา_9'])
        pt_all_allowed = my_dispense_allowed + ['Matching', 'พัก', 'นอกเวลา']
        
        for t in range(16): 
            if t < s_idx or t >= e_idx: model.Add(x[p, t, 'นอกเวลา'] == 1)
            else: model.Add(x[p, t, 'นอกเวลา'] == 0)
        
        for t in range(max(0, s_idx), min(16, e_idx)):
            model.Add(sum(x[p, t, task] for task in pt_all_allowed) == 1)

        is_group_a = (pt['start'] == "09.30" and pt['end'] in ["16.00", "16.30"])
        is_group_b = (pt['start'] in ["10.30", "11.00", "11.30", "12.00", "12.30"] and pt['end'] in ["16.00", "16.30"])
        is_group_c = (pt['start'] in ["13.00", "13.30"] and pt['end'] in ["16.00", "16.30"])
        is_group_d = (pt['start'] == "09.30" and pt['end'] in ["13.00", "13.30"])
        is_group_e = (pt['start'] == "09.00" and pt['end'] == "14.30")

        if pt['has_break'] and not (is_group_c or is_group_d or is_group_e):
            if s_idx <= 8 and e_idx > 8:
                model.Add(x[p, 8, 'พัก'] == 1) 
                if e_idx > 9: model.Add(x[p, 9, 'Matching'] == 1) 
                    
        for t in range(16):
            if not (pt['has_break'] and not (is_group_c or is_group_d or is_group_e) and t == 8): model.Add(x[p, t, 'พัก'] == 0)

        if is_group_a: 
            model.Add(sum(x[p, t, task] for t in range(16) for task in my_dispense_allowed) == 8)
            model.Add(sum(x[p, t, task] for t in range(0, 8) for task in my_dispense_allowed) == 4)
            model.Add(sum(x[p, t, task] for t in range(9, 16) for task in my_dispense_allowed) == 4)
        elif is_group_b: 
            model.Add(sum(x[p, t, task] for t in range(16) for task in my_dispense_allowed) == 6)
            model.Add(sum(x[p, t, task] for t in range(0, 8) for task in my_dispense_allowed) == 2)
            model.Add(sum(x[p, t, task] for t in range(9, 16) for task in my_dispense_allowed) == 4)
        elif is_group_c or is_group_d: 
            model.Add(sum(x[p, t, task] for t in range(16) for task in my_dispense_allowed) == 4)
        elif is_group_e: 
            model.Add(sum(x[p, t, task] for t in range(16) for task in my_dispense_allowed) == 6)

    # --- 7. จัดการเวลาพักของ Full-time (FT) ---
    b_group_vars_ft = {0: [], 1: [], 2: []}
    for p in all_pharmacists:
        if p in ft_pharmacists:
            model.Add(sum(x[p, t, 'นอกเวลา'] for t in range(16)) == 0) 
            for t in range(16): model.Add(x[p, t, 'Matching'] == 0) # FT ห้ามทำ Matching
        
        if p in active_ft:
            model.Add(sum(x[p, t, 'พัก'] for t in range(16)) == 2)
            choices = [model.NewBoolVar(f'choice_{p}_b{i}') for i in range(3)]
            if p in FIX_BREAKS and p in ft_pharmacists:
                req_b = FIX_BREAKS[p]
                for i in range(3): model.Add(choices[i] == (1 if i == req_b else 0))
            else:
                model.AddExactlyOne(choices) 
            for i in range(3):
                b_group_vars_ft[i].append(choices[i])
                for t in range(*b_groups[i]): model.Add(x[p, t, 'พัก'] == 1).OnlyEnforceIf(choices[i])
            for t in range(16):
                if t not in break_slots: model.Add(x[p, t, 'พัก'] == 0)
        elif p in ft_pharmacists:
            model.Add(sum(x[p, t, 'พัก'] for t in range(16)) == 0)

    total_active_ft = len(active_ft)
    if total_active_ft > 0:
        for i in range(3):
            model.Add(sum(b_group_vars_ft[i]) <= 7) 
            model.Add(sum(b_group_vars_ft[i]) >= max(0, (total_active_ft // 3) - 1))

    reward_vars = []
    
    # --- 8. กฎควบคุมจำนวนคนในแต่ละหน้าที่ ---
    for t in range(16):
        for task in tasks:
            if task not in ['พัก', 'งานเฉพาะ', 'ลา', 'นอกเวลา', 'ว่าง', 'Matching', 'Match_C2']:
                model.Add(sum(x[p, t, task] for p in all_pharmacists) <= 1)
        model.Add(sum(x[p, t, 'Match_C2'] for p in all_pharmacists) <= 1)

        if t < 2: req_core = ['จ่ายยา_6', 'จ่ายยา_7', 'จ่ายยา_8', 'จ่ายยา_9', 'Ver_1', 'Ver_2', 'Ver_3', 'PS_1', 'Match_C']
        elif t == 2: req_core = ['จ่ายยา_5', 'จ่ายยา_6', 'จ่ายยา_7', 'จ่ายยา_8', 'จ่ายยา_9', 'Ver_1', 'Ver_2', 'Ver_3', 'PS_1', 'Match_C']
        else: req_core = ['จ่ายยา_5', 'จ่ายยา_6', 'จ่ายยา_7', 'จ่ายยา_8', 'จ่ายยา_9', 'จ่ายยา_10', 'Ver_1', 'Ver_2', 'Ver_3', 'PS_1', 'Match_C']
            
        for task in req_core: model.Add(sum(x[p, t, task] for p in all_pharmacists) == 1)

        if t not in break_slots:
            model.Add(sum(x[p, t, 'PS_2'] for p in all_pharmacists) == 1)
        else:
            model.Add(sum(x[p, t, 'PS_2'] for p in all_pharmacists) <= 1)
            ps2_sum = sum(x[p, t, 'PS_2'] for p in all_pharmacists)
            reward_vars.append(ps2_sum * 100000)

        if t < 3:
            model.Add(sum(x[p, t, 'จ่ายยา_10'] for p in all_pharmacists) <= 1)
            d10_sum = sum(x[p, t, 'จ่ายยา_10'] for p in all_pharmacists)
            reward_vars.append(d10_sum * 150000)

    for t in range(16):
        for i in range(2, 10): model.Add(sum(x[p, t, f'PS_{i+1}'] for p in all_pharmacists) <= sum(x[p, t, f'PS_{i}'] for p in all_pharmacists))
        for i in range(4, 10): model.Add(sum(x[p, t, f'Ver_{i+1}'] for p in all_pharmacists) <= sum(x[p, t, f'Ver_{i}'] for p in all_pharmacists))
        model.Add(sum(x[p, t, 'จ่ายยา_11'] for p in all_pharmacists) <= sum(x[p, t, 'จ่ายยา_4'] for p in all_pharmacists))

    # --- 9. กฎเหล็ก (Hard Constraints) ---
    
    # 9.1 ห้ามสลับช่องจ่ายยาทันที
    categories_to_prevent_internal_switch = [dispensing_tasks]
    for p in all_pharmacists:
        for t in range(15):
            for cat in categories_to_prevent_internal_switch:
                for task1 in cat:
                    for task2 in cat:
                        if task1 != task2: model.AddImplication(x[p, t, task1], x[p, t+1, task2].Not())

    # 9.2 กฎ 1 ชั่วโมงเฉพาะงานหลัก (บังคับใช้กับ "ทุกคน" เพื่อไม่ให้ PT ยืนจ่ายยาเกิน 1 ชม.)
    for p in all_pharmacists:
        for t in range(14): 
            model.Add(sum(x[p, t+k, task] for task in dispensing_tasks for k in range(3)) <= 2)

    # กฎ 1 ชั่วโมงสำหรับงาน Ver และ Match (บังคับเฉพาะ FT)
    work_categories_ft = [['Ver_1', 'Ver_2', 'Ver_3'], ['PS_1'], ['Match_C', 'Match_C2']]
    for p in ft_pharmacists:                    
        for cat in work_categories_ft:
            for t in range(14): 
                model.Add(sum(x[p, t+k, task] for task in cat for k in range(3)) <= 2)

    # 🌟 9.3 กฎใหม่ V116.2: PT ห้ามทำ Matching ติดกันเกิน 1 ชั่วโมง (2 สล็อต) 🌟
    for p in pt_pharmacists:
        for t in range(14):
            model.Add(sum(x[p, t+k, 'Matching'] for k in range(3)) <= 2)

    # 9.4 ความยุติธรรมและขีดจำกัดการจ่ายยาของ FT
    is_disp_7_vars = []
    for p in ft_pharmacists:
        tot_disp = sum(x[p, t, task] for t in range(16) for task in dispensing_tasks)
        over_3hr_var = model.NewBoolVar(f'over_3hr_{p}')
        model.Add(tot_disp <= 6 + over_3hr_var)
        model.Add(tot_disp <= 7) # เพดานสูงสุดห้ามเกิน 3.5 ชม.
        reward_vars.append(over_3hr_var * -500000) 
        is_disp_7_vars.append(over_3hr_var)
        
        if p in active_ft and p not in SICK_PEOPLE:
            has_heavy_custom_tasks = custom_task_slots_count[p] >= 6 
            is_half_day_leave = p in half_day_leaves
            short_disp = model.NewIntVar(0, 16, f'short_disp_{p}')
            if has_heavy_custom_tasks or is_half_day_leave: model.Add(short_disp >= 2 - tot_disp)
            else: model.Add(short_disp >= 4 - tot_disp)
            model.Add(short_disp >= 0)
            reward_vars.append(short_disp * -500000) 

            under_avg = model.NewIntVar(0, 16, f'under_avg_{p}')
            model.Add(under_avg >= 5 - tot_disp)
            model.Add(under_avg >= 0)
            reward_vars.append(under_avg * -10000) 

        for d in ['จ่ายยา_6', 'จ่ายยา_7', 'จ่ายยา_8', 'จ่ายยา_9']: model.Add(sum(x[p, t, d] for t in range(16)) <= 2)
        for d in ['จ่ายยา_4', 'จ่ายยา_5', 'จ่ายยา_10', 'จ่ายยา_11']:
            total_d = sum(x[p, t, d] for t in range(16))
            over_d = model.NewIntVar(0, 16, f'over_{p}_{d}')
            model.Add(over_d >= total_d - 2)
            model.Add(over_d >= 0) 
            reward_vars.append(over_d * -2500) 

        model.Add(sum(x[p, t, 'Ver_2'] for t in range(16)) <= 4)
        for v in ['Ver_1', 'Ver_3']: model.Add(sum(x[p, t, v] for t in range(16)) <= 2)

        # 🌟 กฎเหล็ก: FT คนเดียวห้ามทำทั้งช่อง 7 และช่อง 8 ในวันเดียวกัน 🌟
        done_disp_7 = model.NewBoolVar(f'done_disp_7_{p}')
        model.Add(sum(x[p, t, 'จ่ายยา_7'] for t in range(16)) > 0).OnlyEnforceIf(done_disp_7)
        model.Add(sum(x[p, t, 'จ่ายยา_7'] for t in range(16)) == 0).OnlyEnforceIf(done_disp_7.Not())

        done_disp_8 = model.NewBoolVar(f'done_disp_8_{p}')
        model.Add(sum(x[p, t, 'จ่ายยา_8'] for t in range(16)) > 0).OnlyEnforceIf(done_disp_8)
        model.Add(sum(x[p, t, 'จ่ายยา_8'] for t in range(16)) == 0).OnlyEnforceIf(done_disp_8.Not())
        model.Add(done_disp_7 + done_disp_8 <= 1)

        model.Add(sum(x[p, t, 'Match_C'] + x[p, t, 'Match_C2'] for t in range(16)) <= 3)

    model.Add(sum(is_disp_7_vars) <= 5) 

    # --- 10. ระบบ Soft Constraints และ Scoring ---
    
    # หักคะแนนอย่างหนักถ้าเว้นพักการจ่ายยาแค่ 30 นาที แล้วกลับมาจ่ายอีก
    for p in all_pharmacists:
        for t in range(14):
            is_disp_t = sum(x[p, t, d] for d in dispensing_tasks)
            is_disp_t1 = sum(x[p, t+1, d] for d in dispensing_tasks)
            is_disp_t2 = sum(x[p, t+2, d] for d in dispensing_tasks)
            too_long = model.NewBoolVar(f'too_long_disp_{p}_{t}')
            model.Add(is_disp_t + is_disp_t1 + is_disp_t2 <= 2 + too_long)
            reward_vars.append(too_long * -100000) 
            short_break = model.NewBoolVar(f'short_break_disp_{p}_{t}')
            model.Add(is_disp_t - is_disp_t1 + is_disp_t2 <= 1 + short_break)
            reward_vars.append(short_break * -100000)

    # ให้โบนัสสูงถ้า AI จับคู่งานเหมือนกันติดกัน 2 สล็อตได้ (ดันให้ทำงานบล็อกละ 1 ชั่วโมง)
    tasks_to_pair = dispensing_tasks + ver_cpoe_tasks + ver_ps_tasks + ['Match_C', 'Match_C2']
    for p in all_pharmacists:
        for t in range(15):
            for task in tasks_to_pair:
                match_var = model.NewBoolVar(f'pair_{p}_{t}_{task}')
                model.AddImplication(match_var, x[p, t, task])
                model.AddImplication(match_var, x[p, t+1, task])
                if task in dispensing_tasks: reward_vars.append(match_var * 500000) 
                else: reward_vars.append(match_var * 150000)

    for pt in PART_TIME:
        p = pt['name']
        for t in range(15):
            match_pair_pt = model.NewBoolVar(f'pair_matching_pt_{p}_{t}')
            model.AddImplication(match_pair_pt, x[p, t, 'Matching'])
            model.AddImplication(match_pair_pt, x[p, t+1, 'Matching'])
            reward_vars.append(match_pair_pt * 150000) 

    # หักคะแนนรุนแรงถ้ามี "เศษงาน 30 นาทีโดดๆ" (เพื่อบีบให้รวมเป็น 1 ชั่วโมงถ้าเป็นไปได้)
    for p in ft_pharmacists:
        ft_iso_disp_vars = []
        for t in range(16):
            for d in dispensing_tasks:
                iso_disp = model.NewBoolVar(f'iso_disp_{p}_{t}_{d}')
                prev_v = x[p, t-1, d] if t > 0 else 0
                next_v = x[p, t+1, d] if t < 15 else 0
                model.Add(x[p, t, d] - prev_v - next_v <= iso_disp)
                ft_iso_disp_vars.append(iso_disp)
                reward_vars.append(iso_disp * -200000) 
        # อนุญาตให้มีเศษ 30 นาทีได้แค่ 2 ครั้งต่อคน
        model.Add(sum(ft_iso_disp_vars) <= 2)

    for p in all_pharmacists:
        for t in range(16):
            for target_task in ['Ver_1', 'Ver_2', 'Ver_3', 'Match_C', 'PS_1']:
                iso_var = model.NewBoolVar(f'iso_{target_task}_{p}_{t}')
                prev_v = x[p, t-1, target_task] if t > 0 else 0
                next_v = x[p, t+1, target_task] if t < 15 else 0
                model.Add(x[p, t, target_task] - prev_v - next_v <= iso_var)
                reward_vars.append(iso_var * -100000)

    for p in ft_pharmacists:
        for t in range(16): reward_vars.append(x[p, t, 'ว่าง'] * -100000) 

    # --- 11. กำหนดความสำคัญ (Weight) ของแต่ละหน้าที่ ---
    for t in range(16):
        weights = {
            'จ่ายยา_4': 300000, 'จ่ายยา_11': 290000, 
            'Ver_4': 50000, 'PS_3': 48000, 
            'Match_C2': 47000, 
            'Ver_5': 46000, 'PS_4': 44000, 
            'Ver_6': 42000, 'PS_5': 40000, 'Ver_7': 38000, 'PS_6': 36000, 
            'Ver_8': 34000, 'PS_7': 32000, 'Ver_9': 30000, 'PS_8': 28000, 
            'Ver_10': 26000, 'PS_9': 24000, 'PS_10': 22000
        }
        if IS_MWF and (t in break_slots):
            weights['จ่ายยา_4'] = -50000 
            weights['จ่ายยา_11'] = -50000

        for task, weight in weights.items():
            for i, p in enumerate(all_pharmacists): reward_vars.append(x[p, t, task] * (weight + i))

    # --- ส่งให้ AI ประมวลผล ---
    model.Maximize(sum(reward_vars))
    solver = cp_model.CpSolver()
    solver.parameters.num_workers = 8
    solver.parameters.max_time_in_seconds = 60.0  

    status = solver.Solve(model)

    # --- จัดการผลลัพธ์ ---
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        schedule_data = []
        for p in all_pharmacists:
            row_data = {'ชื่อ/เวลา': p} 
            for t in range(16):
                for task in tasks:
                    if solver.Value(x[p, t, task]) == 1:
                        if task == 'งานเฉพาะ': display_task = custom_dict_index.get((p, t), 'งานเฉพาะ')
                        elif task in ['นอกเวลา', 'ว่าง']: display_task = '-'
                        elif task == 'Match_C': display_task = 'Match + C'
                        elif task == 'Match_C2': display_task = 'Match + C2'
                        elif task == 'Matching': display_task = 'Matching'
                        elif task == 'Ver_1': display_task = 'Ver 1 INC'
                        elif task == 'Ver_2': display_task = 'Ver 2/ปณ.'
                        elif task == 'Ver_3': display_task = 'Ver 3/A'
                        elif task.startswith('PS_'): display_task = 'Ver ' + task.replace('_', '')
                        elif task.startswith('จ่ายยา_'): display_task = task.replace('จ่ายยา_', 'จ่าย ')
                        else: display_task = task.replace('_', ' ')
                        row_data[time_slots[t]] = display_task
            schedule_data.append(row_data)
        df_result = pd.DataFrame(schedule_data)
        summary_row = {'ชื่อ/เวลา': 'P/C/D'} 
        for t_idx in range(16):
            time_col = time_slots[t_idx]
            w_count, x_count, disp_nums = 0, 0, []
            for p in all_pharmacists:
                val_str = str(df_result.loc[df_result['ชื่อ/เวลา'] == p, time_col].values[0])
                if 'Ver PS' in val_str: w_count += 1
                elif 'Ver' in val_str: x_count += 1
                elif 'จ่าย ' in val_str:
                    try: disp_nums.append(int(val_str.replace('จ่าย ', '').strip()))
                    except: pass
            yz_str = f"{min(disp_nums)}-{max(disp_nums)}" if disp_nums else "-"
            summary_row[time_col] = f"{w_count}/{x_count}/{yz_str}"
        df_result = pd.concat([df_result, pd.DataFrame([summary_row])], ignore_index=True)
        return df_result, "Success", ""
    elif status == cp_model.UNKNOWN: return None, "Timeout", "ระบบคิดไม่ทัน กรุณาลดการล็อกงานหลัก หรือจำนวนคนลาลงครับ"
    else: return None, "Infeasible", "จำนวนคนไม่เพียงพอต่อการจัดตาราง หรือเงื่อนไขตึงเกินไปครับ"

# ==========================================
# 🎨 ฟังก์ชันใส่สีพื้นหลังเว็บ
# ==========================================
def get_color_style(val):
    val_str = str(val)
    base_style = "text-align: center; " 
    if '/' in val_str and '-' in val_str and val_str[0].isdigit(): return base_style + 'background-color: #FFF2CC; color: black; font-weight: bold;' 
    elif 'จ่าย ' in val_str: return base_style + 'background-color: #D5E8D4; color: black;' 
    elif val_str == 'Matching': return base_style + 'background-color: #DAE8FC; color: black;' 
    elif 'Match' in val_str: return base_style + 'background-color: #DAE8FC; color: #FF0000; font-weight: bold;' 
    elif 'Ver PS' in val_str: return base_style + 'background-color: #E1D5E7; color: black;' 
    elif 'Ver' in val_str: return base_style + 'background-color: #FFE6CC; color: black;' 
    elif val_str == 'พัก': return base_style + 'background-color: #F8CECC; color: black;' 
    elif val_str in ['-', 'ว่าง']: return base_style + 'background-color: #F5F5F5; color: black;' 
    else: return base_style + 'background-color: #E6E6E6; color: black;' 

# ==========================================
# 📸 ฟังก์ชันสร้าง HTML Table สำหรับโหลด PNG
# ==========================================
def build_html_table(df, selected_date, DAY_OF_WEEK):
    thai_date_str = get_thai_date(selected_date)
    def get_cell_style(val):
        val_str = str(val)
        bg, color, weight = "#E6E6E6", "black", "normal"
        if '/' in val_str and '-' in val_str and val_str and val_str[0].isdigit(): bg, weight = "#FFF2CC", "bold"
        elif 'จ่าย ' in val_str: bg = "#D5E8D4"
        elif val_str == 'Matching': bg = "#DAE8FC"
        elif 'Match' in val_str: bg, color, weight = "#DAE8FC", "#FF0000", "bold"
        elif 'Ver PS' in val_str: bg = "#E1D5E7"
        elif 'Ver' in val_str: bg = "#FFE6CC"
        elif val_str == 'พัก': bg = "#F8CECC"
        elif val_str in ['-', 'ว่าง']: bg = "#F5F5F5" 
        return f"background-color: {bg}; color: {color}; font-weight: {weight}; border: 1px solid black; padding: 4px 5px; text-align: center; font-size: 17px; white-space: nowrap; height: 50px; box-sizing: border-box;"
        
    def get_head_color_hex(t_idx, day_of_week):
        if day_of_week == 'Normal':
            if t_idx in [0, 1, 3, 4, 11, 12]: return '#FFE6CC' 
            if t_idx in [2]: return '#FFF2CC'                 
            if t_idx in [5, 6, 9, 10]: return '#F8CECC'         
            if t_idx in [7, 8]: return '#E1D5E7'              
            if t_idx in [13, 14, 15]: return '#DAE8FC'          
        else: 
            if t_idx in [0, 1, 4, 5, 12, 13]: return '#FFE6CC' 
            if t_idx in [2, 3]: return '#FFF2CC'              
            if t_idx in [6, 7, 10, 11]: return '#F8CECC'        
            if t_idx in [8, 9]: return '#E1D5E7'              
            if t_idx in [14, 15]: return '#DAE8FC'              
        return '#FFFFFF'

    cols = df.columns.tolist()
    num_cols = len(cols)
    html = f"<div id='capture-area' style='background-color: white; padding: 20px; display: inline-block; font-family: \"Sarabun\", \"TH Sarabun New\", sans-serif;'><table style='border-collapse: collapse; width: 100%;'><tr><td colspan='{num_cols}' style='text-align: center; font-size: 28px; font-weight: bold; border: none; padding-bottom: 5px;'>ตารางปฏิบัติงานเภสัชกร ห้องยาชั้น 1 อาคารสมเด็จพระเทพรัตน์</td></tr><tr><td colspan='{num_cols}' style='text-align: center; font-size: 22px; font-weight: bold; border: none; padding-bottom: 15px;'>ประจำ{thai_date_str}</td></tr><tr>"
    for i, col in enumerate(cols):
        bg = "#FFFFFF" if i == 0 else get_head_color_hex(i - 1, DAY_OF_WEEK)
        html += f"<th style='background-color: {bg}; border: 1px solid black; padding: 6px; font-size: 19px; white-space: nowrap; height: 55px; box-sizing: border-box;'>{col}</th>"
    html += "</tr>"
    for _, row in df.iterrows():
        html += "<tr style='height: 50px;'>"
        for i, col in enumerate(cols):
            val = row[col]
            style = get_cell_style(val)
            if i == 0 or _ == len(df)-1: style = style.replace("font-weight: normal", "font-weight: bold")
            html += f"<td style='{style}'>{val}</td>"
        html += "</tr>"
    html += "</table></div>"
    return html

# ==========================================
# 🖥️ หน้าเว็บ UI ฝั่ง Streamlit
# ==========================================
st.set_page_config(page_title="Pharmacy Schedule App", layout="wide", page_icon="💊")
st.markdown("<style>.block-container { padding-top: 1.5rem !important; padding-bottom: 1rem !important; } .stMarkdown h1 { margin-top: -1rem !important; padding-bottom: 0rem !important; margin-bottom: 0.5rem !important; } .stMarkdown h3 { margin-top: -0.5rem !important; padding-bottom: 0rem !important; margin-bottom: 0.2rem !important; } .stMarkdown p { margin-bottom: 0.5rem !important; } th { text-align: center !important; } hr { margin-top: 0.5rem; margin-bottom: 0.5rem; border-color: #e0e0e0; }</style>", unsafe_allow_html=True)

st.title("💊 จัดตารางปฏิบัติงานเภสัชกร ด้วย AI")
st.subheader("🏥 ห้องยาชั้น 1 อาคารสมเด็จพระเทพรัตน์ โรงพยาบาลรามาธิบดี")
st.markdown("<p style='font-size: 14px; color: gray;'>version 116.2 (06/05/26) พัฒนาโดย Niratsai Sukprasert และ Gemini</p>", unsafe_allow_html=True)

ft_pharmacists_list = ['เต้น', 'แอน', 'แม็ค', 'โบ้ท', 'ไม้เอก', 'กิ๊ฟ', 'ฟอร์จูน', 'มิ้ลค์', 'ริน', 'อ๊อฟฟี่', 'ออย', 'บี', 'มายด์', 'ขิม', 'บีม', 'มิ้น', 'ใบเตย', 'จีน่า', 'ปอนด์']
dropdown_names = ["ไม่มี"] + ft_pharmacists_list

leaves_input, pt_input_list, custom_tasks_input, fixed_main_tasks_input, fix_breaks_input, sick_people_input = {}, [], {}, {}, {}, []

if "schedule_df" not in st.session_state: st.session_state.schedule_df = None
if "run_status" not in st.session_state: st.session_state.run_status = None

with st.sidebar:
    st.markdown("<h2 style='font-size: 24px; font-weight: bold;'>⚙️ ตั้งค่าตารางประจำวัน</h2>", unsafe_allow_html=True)
    tz_bkk = timezone(timedelta(hours=7))
    selected_date = st.date_input("date", datetime.now(tz_bkk).date(), label_visibility="collapsed")
    IS_MWF = selected_date.weekday() in [0, 2, 4]
    DAY_OF_WEEK = 'Wed_Fri' if selected_date.weekday() in [2, 4] else 'Normal'
    st.divider()
    
    st.subheader("🏖️ ผู้ที่ลาในวันนี้")
    with st.expander("คลิกเพื่อระบุผู้ลางาน (สูงสุด 5 คน)", expanded=False):
        for i in range(5):
            c1, c2 = st.columns([3, 2])
            with c1: p_leave = st.selectbox(f"คนที่ {i+1}", dropdown_names, key=f"l_name_{i}")
            with c2: t_leave = st.selectbox("ประเภท", ["ทั้งวัน", "เช้า", "บ่าย"], key=f"l_type_{i}")
            st.divider()
            if p_leave != "ไม่มี": leaves_input[p_leave] = t_leave
    st.divider()

    st.subheader("🧑‍⚕️ เภสัชกร Part-time")
    with st.expander("คลิกเพื่อระบุ Part-time (สูงสุด 5 คน)", expanded=False):
        for i in range(5):
            pt_name = st.text_input(f"ชื่อ PT {i+1}", key=f"pt_n_{i}", label_visibility="collapsed", placeholder="ระบุชื่อ (ถ้ามี)")
            cc1, cc2, cc3 = st.columns([2, 2, 2])
            with cc1: pt_s = st.selectbox(f"เริ่ม{i}", VALID_TIMES, index=0, key=f"pt_s_{i}")
            with cc2: pt_e = st.selectbox(f"สิ้นสุด{i}", VALID_TIMES, index=16, key=f"pt_e_{i}")
            with cc3: pt_b = st.checkbox(f"พัก 12.30 {i}", value=True, key=f"pt_b_{i}")
            st.divider()
            if pt_name.strip() != "": pt_input_list.append({'name': pt_name.strip(), 'start': pt_s, 'end': pt_e, 'has_break': pt_b})
    st.divider()

    st.subheader("📋 ภารกิจพิเศษ")
    with st.expander("คลิกเพื่อระบุภารกิจพิเศษ (สูงสุด 20 งาน)", expanded=False):
        for i in range(20):
            p_task = st.selectbox(f"ชื่อคน {i+1}", dropdown_names, key=f"t_name_{i}", label_visibility="collapsed")
            n_task = st.text_input(f"ชื่องาน {i+1}", key=f"t_n_{i}", placeholder="ระบุชื่องาน")
            c1, c2 = st.columns(2)
            with c1: s_task = st.selectbox(f"เริ่ม {i+1}", VALID_TIMES, index=0, key=f"t_s_{i}")
            with c2: e_task = st.selectbox(f"สิ้นสุด {i+1}", VALID_TIMES, index=2, key=f"t_e_{i}")
            st.divider()
            if p_task != "ไม่มี" and n_task.strip() != "": custom_tasks_input[(p_task, s_task, e_task)] = n_task.strip()
    st.divider()

    st.subheader("📌 ล็อกภาระงานหลัก")
    with st.expander("คลิกเพื่อล็อกภาระงานหลัก (สูงสุด 20 รายการ)", expanded=False):
        opts = ['จ่าย 4', 'จ่าย 5', 'จ่าย 6', 'จ่าย 7', 'จ่าย 8', 'จ่าย 9', 'จ่าย 10', 'จ่าย 11', 'Ver 1 INC', 'Ver 2/ปณ.', 'Ver 3/A', 'Ver 4', 'Ver 5', 'Ver 6', 'Ver 7', 'Ver 8', 'Ver 9', 'Ver 10', 'Ver PS1', 'Ver PS2', 'Ver PS3', 'Ver PS4', 'Ver PS5', 'Ver PS6', 'Ver PS7', 'Ver PS8', 'Ver PS9', 'Ver PS10', 'Match + C', 'Match + C2', 'Matching']
        maps = {'จ่าย 4': 'จ่ายยา_4', 'จ่าย 5': 'จ่ายยา_5', 'จ่าย 6': 'จ่ายยา_6', 'จ่าย 7': 'จ่ายยา_7', 'จ่าย 8': 'จ่ายยา_8', 'จ่าย 9': 'จ่ายยา_9', 'จ่าย 10': 'จ่ายยา_10', 'จ่าย 11': 'จ่ายยา_11', 'Ver 1 INC': 'Ver_1', 'Ver 2/ปณ.': 'Ver_2', 'Ver 3/A': 'Ver_3', 'Ver 4': 'Ver_4', 'Ver 5': 'Ver_5', 'Ver 6': 'Ver_6', 'Ver 7': 'Ver 7', 'Ver 8': 'Ver_8', 'Ver 9': 'Ver_9', 'Ver 10': 'Ver_10', 'Ver PS1': 'PS_1', 'Ver PS2': 'PS_2', 'Ver PS3': 'PS_3', 'Ver PS4': 'PS_4', 'Ver PS5': 'PS_5', 'Ver PS6': 'PS_6', 'Ver PS7': 'PS_7', 'Ver PS8': 'PS_8', 'Ver PS9': 'PS_9', 'Ver PS10': 'PS_10', 'Match + C': 'Match_C', 'Match + C2': 'Match_C2', 'Matching': 'Matching'}
        for i in range(20):
            p_m_task = st.selectbox(f"ชื่อคน {i+1}", dropdown_names, key=f"m_name_{i}", label_visibility="collapsed")
            n_m_task = st.selectbox(f"ภาระงาน {i+1}", ["เลือกภาระงาน"] + opts, key=f"m_task_{i}", label_visibility="collapsed")
            c1, c2 = st.columns(2)
            with c1: s_m_task = st.selectbox(f"เริ่ม {i+1}", VALID_TIMES, index=0, key=f"m_s_{i}")
            with c2: e_m_task = st.selectbox(f"สิ้นสุด {i+1}", VALID_TIMES, index=2, key=f"m_e_{i}")
            st.divider()
            if p_m_task != "ไม่มี" and n_m_task != "เลือกภาระงาน": fixed_main_tasks_input[(p_m_task, s_m_task, e_m_task)] = maps[n_m_task]
    st.divider()

    st.subheader("🤒 คนที่ไม่สบาย")
    with st.expander("คลิกเพื่อระบุผู้ป่วย (สูงสุด 3 คน)", expanded=False):
        for i in range(3):
            p_sick = st.selectbox(f"คนที่ {i+1}", dropdown_names, key=f"sick_{i}")
            st.divider()
            if p_sick != "ไม่มี": sick_people_input.append(p_sick)
    st.divider()

    st.subheader("🍱 ล็อกเวลาพักเฉพาะบุคคล")
    with st.expander("คลิกเพื่อล็อกเวลาพัก (สูงสุด 5 คน)", expanded=False):
        break_choices = ["รอบที่ 1", "รอบที่ 2", "รอบที่ 3"]
        for i in range(5):
            c1, c2 = st.columns([2, 3])
            with c1: p_b = st.selectbox(f"คนที่ {i+1}", dropdown_names, key=f"b_name_{i}")
            with c2: t_b = st.selectbox(f"รอบพัก {i+1}", break_choices, key=f"b_time_{i}")
            st.divider()
            if p_b != "ไม่มี":
                if "รอบที่ 1" in t_b: fix_breaks_input[p_b] = 0
                elif "รอบที่ 2" in t_b: fix_breaks_input[p_b] = 1
                elif "รอบที่ 3" in t_b: fix_breaks_input[p_b] = 2

if st.button("🚀 เริ่มจัดตารางด้วย AI (คลิก)", type="primary", use_container_width=True):
    with st.spinner("กำลังจัดตารางปฏิบัติงานของคุณ... (ใช้เวลาประมาณ 10-30 วินาที)"):
        try:
            df_result, status, msg = generate_schedule(DAY_OF_WEEK, leaves_input, custom_tasks_input, pt_input_list, fix_breaks_input, fixed_main_tasks_input, sick_people_input, IS_MWF)
            if status == "Success":
                st.session_state.schedule_df = df_result
                st.session_state.run_status = "Success"
            else:
                st.error(f"⚠️ {msg}")
                st.session_state.schedule_df = None
                st.session_state.run_status = status
        except Exception as e: st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")

if st.session_state.schedule_df is not None and st.session_state.run_status == "Success":
    st.success("🎉 นี่คือตารางของคุณ ตรวจสอบความถูกต้องและดาวน์โหลดได้เลยครับ")
    df_to_show = st.session_state.schedule_df
    try: styled_df = df_to_show.style.map(get_color_style, subset=df_to_show.columns[1:])
    except AttributeError: styled_df = df_to_show.style.applymap(get_color_style, subset=df_to_show.columns[1:])
    st.dataframe(styled_df, use_container_width=True)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        styled_df.to_excel(writer, index=False, sheet_name='Schedule', startrow=2)
        ws = writer.sheets['Schedule']
        
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1 
        ws.page_setup.fitToHeight = 1 
        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = True
        cm_to_inch = 0.4 / 2.54
        ws.page_margins = PageMargins(left=cm_to_inch, right=cm_to_inch, top=cm_to_inch, bottom=cm_to_inch, header=0, footer=0)
        
        thai_date_str = get_thai_date(selected_date)
        ws['A1'] = "ตารางปฏิบัติงานเภสัชกร ห้องยาชั้น 1 อาคารสมเด็จพระเทพรัตน์"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_to_show.columns))
        ws['A1'].font = Font(name='TH Sarabun New', size=20, bold=True)
        ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
        
        ws['A2'] = f"ประจำ{thai_date_str}"
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(df_to_show.columns))
        ws['A2'].font = Font(name='TH Sarabun New', size=18, bold=True)
        ws['A2'].alignment = Alignment(horizontal="center", vertical="center")
        
        ws.row_dimensions[3].height = 40
        center_aligned_text = Alignment(horizontal="center", vertical="center")
        for col_idx in range(1, len(df_to_show.columns) + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 11.5 
            for row_idx in range(3, len(df_to_show) + 4): 
                if row_idx >= 4: ws.row_dimensions[row_idx].height = 30
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = center_aligned_text
                val_str = str(cell.value)
                is_bold = True if (cell.row == 3 or cell.column == 1) else False
                
                if "Match" in val_str and val_str in ["Match + C", "Match + C2"]:
                    cell.font = Font(name='TH Sarabun New', size=18, bold=True, color="FF0000")
                elif '/' in val_str and '-' in val_str and val_str[0].isdigit():
                    cell.font = Font(name='TH Sarabun New', size=18, bold=True)
                else: cell.font = Font(name='TH Sarabun New', size=18, bold=is_bold)
                cell.border = thin_border
                if cell.row == 3 and col_idx >= 2:
                    c_name = get_header_color(col_idx - 2, DAY_OF_WEEK)
                    if c_name: cell.fill = header_color_map[c_name]
    
    st.download_button("📥 ดาวน์โหลดเป็นไฟล์ Excel", data=buffer.getvalue(), file_name=f"Pharmacy_Schedule_{selected_date.strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    html_table = build_html_table(df_to_show, selected_date, DAY_OF_WEEK)
    file_name_png = f"Pharmacy_Schedule_{selected_date.strftime('%Y-%m-%d')}.png"
    full_html = f"<!DOCTYPE html><html><head><meta charset='utf-8'><link href='https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap' rel='stylesheet'><script src='https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js'></script><style>body {{ font-family: 'Sarabun', 'TH Sarabun New', sans-serif; margin: 0; padding: 0; background: transparent; }} .btn {{ width: 100%; background-color: #f0f2f6; color: #31333F; padding: 0.5rem 1rem; border: 1px solid rgba(49, 51, 63, 0.2); border-radius: 0.5rem; cursor: pointer; font-size: 16px; font-family: 'Sarabun', 'TH Sarabun New', sans-serif; font-weight: 400; line-height: 1.6; transition: all 0.2s ease; display: block; box-sizing: border-box; }} .btn:hover {{ border-color: #FF4B4B; color: #FF4B4B; }} #capture-area-wrapper {{ position: absolute; left: -9999px; top: -9999px; }}</style></head><body><button class='btn' onclick='setTimeout(takeShot, 1000)'>📸 บันทึกเป็นรูปภาพ (PNG)</button><div id='capture-area-wrapper'>{html_table}</div><script>function takeShot() {{ const target = document.getElementById('capture-area'); html2canvas(target, {{ scale: 2, useCORS: true, backgroundColor: '#ffffff' }}).then(canvas => {{ let link = document.createElement('a'); link.download = '{file_name_png}'; link.href = canvas.toDataURL('image/png'); link.click(); }}); }}</script></body></html>"
    components.html(full_html, height=50, scrolling=False)
