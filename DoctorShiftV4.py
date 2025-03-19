import pandas as pd
import random
import json
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill

class DutyScheduler:
    def __init__(self):
        self.doctors = []  # รายชื่อแพทย์
        self.roles = ["อช", "อญ", "สห", "ICU1", "ICU2"]  # หน้าที่เริ่มต้น
        self.holidays = []  # วันหยุดราชการ
        self.constraints = {}  # ข้อจำกัดของแพทย์แต่ละคน
        self.schedule = pd.DataFrame()  # ตารางเวร

    def add_doctor(self, name, max_shifts, unavailable_days, preferred_days, unavailable_roles):
        if name not in self.doctors:
            self.doctors.append(name)
        self.constraints[name] = {
            "max_shifts": max_shifts,
            "unavailable_days": unavailable_days,
            "preferred_days": preferred_days,
            "unavailable_roles": unavailable_roles
        }

    def set_holidays(self, holidays):
        self.holidays = holidays
    
    def generate_schedule(self, num_days):
        self.schedule = pd.DataFrame(index=range(1, num_days+1), columns=self.roles)
        duty_counts = {doc: 0 for doc in self.doctors}
        
        for day in range(1, num_days+1):
            available_doctors = [doc for doc in self.doctors if day not in self.constraints[doc]["unavailable_days"]]
            assigned_doctors = set()  # เก็บแพทย์ที่ได้รับมอบหมายในวันนั้น
            random.shuffle(available_doctors)  # สุ่มแพทย์เพื่อกระจายเวรอย่างยุติธรรม
            
            for role in self.roles:
                valid_doctors = [doc for doc in available_doctors if role not in self.constraints[doc]["unavailable_roles"] 
                                 and duty_counts[doc] < self.constraints[doc]["max_shifts"] 
                                 and doc not in assigned_doctors]
                
                if valid_doctors:
                    selected_doctor = min(valid_doctors, key=lambda d: duty_counts[d])  # เลือกแพทย์ที่มีเวรน้อยที่สุด
                    self.schedule.at[day, role] = selected_doctor
                    duty_counts[selected_doctor] += 1
                    assigned_doctors.add(selected_doctor)
                else:
                    self.schedule.at[day, role] = "-"  # แสดงว่าไม่สามารถจัดเวรได้
        
        return self.schedule

    def validate_schedule(self):
        for day, row in self.schedule.iterrows():
            assigned_doctors = set()
            for role, doctor in row.items():
                if doctor != "-" and doctor in assigned_doctors:
                    return False  # มีแพทย์คนเดียวกันอยู่หลายหน้าที่ในวันเดียวกัน
                assigned_doctors.add(doctor)
        return True
    
    def summarize_schedule(self):
        summary = {doc: {"total": 0, "weekday": 0, "holiday": 0} for doc in self.doctors}
        
        for day, row in self.schedule.iterrows():
            is_holiday = day in self.holidays
            for role, doctor in row.items():
                if doctor != "-":
                    summary[doctor]["total"] += 1
                    if is_holiday:
                        summary[doctor]["holiday"] += 1
                    else:
                        summary[doctor]["weekday"] += 1
        
        return pd.DataFrame.from_dict(summary, orient='index')

# เริ่มต้น Streamlit App
st.title("ระบบจัดเวรแพทย์อัตโนมัติ")

# ใช้ session state เพื่อเก็บข้อมูลแพทย์
if "scheduler" not in st.session_state:
    st.session_state.scheduler = DutyScheduler()

scheduler = st.session_state.scheduler

num_days = st.number_input("จำนวนวันในเดือนนี้", min_value=1, max_value=31, value=30)
holidays = st.multiselect("เลือกวันหยุดราชการ", list(range(1, num_days+1)))
scheduler.set_holidays(holidays)

st.subheader("เพิ่มหรือแก้ไขแพทย์")
new_doc_name = st.text_input("ชื่อแพทย์")
max_shifts = st.number_input("จำนวนเวรสูงสุดในเดือนนี้", min_value=1, max_value=num_days, value=10)
unavailable_days = st.multiselect("เลือกวันที่ไม่สะดวกอยู่เวร", list(range(1, num_days+1)))
preferred_days = st.multiselect("เลือกวันที่สะดวกอยู่เวร", list(range(1, num_days+1)))
unavailable_roles = st.multiselect("เลือกเวรที่ไม่สะดวก", scheduler.roles)

if st.button("บันทึกแพทย์"):
    scheduler.add_doctor(new_doc_name, max_shifts, unavailable_days, preferred_days, unavailable_roles)
    st.success(f"บันทึกข้อมูลของ {new_doc_name} เรียบร้อย")

st.subheader("รายชื่อแพทย์และเงื่อนไข")
for doctor in scheduler.doctors:
    if doctor in scheduler.constraints:
        st.write(f"**{doctor}**: สูงสุด {scheduler.constraints[doctor]['max_shifts']} เวร | ไม่สะดวก: {scheduler.constraints[doctor]['unavailable_days']} | เวรที่ไม่สะดวก: {scheduler.constraints[doctor]['unavailable_roles']}")

st.subheader("จัดตารางเวร")
if st.button("สร้างตารางเวร"):
    schedule = scheduler.generate_schedule(num_days)
    if scheduler.validate_schedule():
        st.write(schedule)
        summary = scheduler.summarize_schedule()
        st.write("สรุปจำนวนเวรของแพทย์:")
        st.write(summary)
    else:
        st.error("ไม่สามารถจัดเวรได้ตามเงื่อนไขทั้งหมด กรุณาปรับเงื่อนไข")
