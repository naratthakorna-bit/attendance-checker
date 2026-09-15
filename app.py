import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="ระบบตรวจสอบเวลาทำงาน", layout="wide")

st.title("ระบบตรวจสอบพนักงานลืมลงเวลา, ขอ OT และจำนวนวันหยุด")
st.write("อัปโหลดไฟล์รายงานตารางเวลา (.xlsx) เพื่อตรวจสอบรายการผิดปกติ")

uploaded_file = st.file_uploader("เลือกไฟล์ Excel รายงานไทม์", type=["xlsx"])

def parse_time(time_str):
    """แปลงสตริงเวลา HH:MM ให้เป็น datetime object สำหรับเปรียบเทียบ"""
    try:
        time_str = str(time_str).strip()
        if not time_str or time_str == 'nan':
            return None
        return datetime.strptime(time_str, "%H:%M")
    except:
        return None

def parse_ot_hours(ot_val):
    """แปลงค่า OT ให้เป็นตัวเลขชั่วโมง"""
    try:
        if pd.isna(ot_val) or str(ot_val).strip() in ['', 'nan']:
            return 0.0
        ot_str = str(ot_val).strip()
        if ':' in ot_str:
            parts = ot_str.split(':')
            return float(parts[0]) + (float(parts[1]) / 60.0)
        return float(ot_str)
    except:
        return 0.0

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        current_emp = None
        records = []

        for idx, row in df_raw.iterrows():
            if idx == 0:
                continue
            col0 = str(row[0]).strip() if pd.notna(row[0]) else ""
            if "แผนก:" in col0:
                current_emp = col0
                continue
            
            date_val = row[1]
            status_val = str(row[2]).strip() if pd.notna(row[2]) else ""
            shift_val = str(row[3]).strip() if pd.notna(row[3]) else ""
            in1 = row[4]
            out1 = row[5]
            ot_reg = row[17] if len(row) > 17 else None
            
            if pd.isna(date_val) or str(date_val).strip() == "วันที่":
                continue
                
            records.append({
                'Emp': current_emp,
                'Date': str(date_val).strip(),
                'Status': status_val,
                'Shift': shift_val,
                'IN': in1,
                'OUT': out1,
                'OT_Regular': ot_reg
            })

        df = pd.DataFrame(records)
        df['Date_dt'] = pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='coerce')
        
        max_date = df['Date_dt'].dropna().max()
        df = df[df['Date_dt'] <= max_date]

        issues = []

        # -------------------------------------------------------------
        # 1. ตรวจสอบจำนวนวันหยุดต่อพนักงาน (8 วัน/รอบตัดวิก)
        # -------------------------------------------------------------
        off_statuses = ['วันหยุดประจำสัปดาห์', 'วันหยุดประเพณี', 'วันหยุด', 'Off', 'OFF', 'Holiday']
        
        for emp, group in df.groupby('Emp'):
            off_days_count = group['Status'].isin(off_statuses).sum()
            
            if off_days_count < 8:
                issues.append({
                    'Emp': emp,
                    'Date': 'ทั้งรอบตัดวิก',
                    'Status': '-',
                    'Shift': '-',
                    'IN': '-',
                    'OUT': '-',
                    'Issue': f'วันหยุดในรอบตัดวิกน้อยกว่า 8 วัน (มี {off_days_count} วัน)'
                })
            elif off_days_count > 8:
                issues.append({
                    'Emp': emp,
                    'Date': 'ทั้งรอบตัดวิก',
                    'Status': '-',
                    'Shift': '-',
                    'IN': '-',
                    'OUT': '-',
                    'Issue': f'วันหยุดในรอบตัดวิกมากกว่า 8 วัน (มี {off_days_count} วัน)'
                })

        # -------------------------------------------------------------
        # 2. ตรวจสอบการลืมลงเวลา, ขอ OT เกินกะ, และ OT เกิน 8 ชม.
        # -------------------------------------------------------------
        for idx, row in df.iterrows():
            in_missing = pd.isna(row['IN']) or str(row['IN']).strip() in ['', 'nan']
            out_missing = pd.isna(row['OUT']) or str(row['OUT']).strip() in ['', 'nan']
            ot_hours = parse_ot_hours(row['OT_Regular'])
            has_ot = ot_hours > 0
            
            # ตรวจสอบลืมสแกนนิ้ววันทำงาน
            if row['Status'] == 'วันทำงาน':
                if not in_missing and out_missing:
                    issues.append({**row, 'Issue': 'ลืมลงเวลาออก (มีแต่เวลาเข้า)'})
                elif in_missing and not out_missing:
                    issues.append({**row, 'Issue': 'ลืมลงเวลาเข้า (มีแต่เวลาออก)'})

            # ตรวจสอบการขอ OT เกิน 8 ชั่วโมงต่อวัน
            if ot_hours > 8.0:
                issues.append({**row, 'Issue': f'มีการขอ OT เกิน 8 ชั่วโมงต่อวัน (ขอไป {row["OT_Regular"]} ชม.)'})

            # ตรวจสอบการเข้าก่อน/ออกหลัง ตามกะทำงาน (เกณฑ์ 30 นาที)
            if not in_missing and not out_missing and '-' in str(row['Shift']):
                try:
                    shift_parts = str(row['Shift']).split('-')
                    shift_start = parse_time(shift_parts[0])
                    shift_end = parse_time(shift_parts[1])
                    
                    actual_in = parse_time(row['IN'])
                    actual_out = parse_time(row['OUT'])

                    if actual_in and actual_out and shift_start and shift_end:
                        # อนุโลม 30 นาที ก่อนกะเริ่ม และ หลังกะเลิก
                        early_in_threshold = shift_start - timedelta(minutes=30)
                        late_out_threshold = shift_end + timedelta(minutes=30)

                        # ออกหลังกะเลิกเกิน 30 นาที โดยไม่มีรายการขอ OT
                        if actual_out > late_out_threshold and not has_ot:
                            over_minutes = int((actual_out - shift_end).total_seconds() / 60)
                            issues.append({**row, 'Issue': f'สแกนออกเลทเกินกะ {over_minutes} นาที แต่ไม่มีรายการขอ OT'})

                        # เข้าก่อนกะเริ่มเกิน 30 นาที โดยไม่มีรายการขอ OT
                        elif actual_in < early_in_threshold and not has_ot:
                            early_minutes = int((shift_start - actual_in).total_seconds() / 60)
                            issues.append({**row, 'Issue': f'สแกนเข้าก่อนกะ {early_minutes} นาที แต่ไม่มีรายการขอ OT'})
                except Exception:
                    pass

        result_df = pd.DataFrame(issues)

        if not result_df.empty:
            display_df = result_df[['Emp', 'Date', 'Status', 'Shift', 'IN', 'OUT', 'Issue']]
            st.error(f"ตรวจพบรายการผิดปกติทั้งหมด {len(display_df)} รายการ")
            st.dataframe(display_df, use_container_width=True)

            csv = display_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="ดาวน์โหลดรายงานสรุป (CSV)",
                data=csv,
                file_name='สรุปรายการผิดปกติ.csv',
                mime='text/csv',
            )
        else:
            st.success("ไม่พบรายการผิดปกติในไฟล์นี้")

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")
