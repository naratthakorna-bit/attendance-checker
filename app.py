{\rtf1\ansi\ansicpg874\cocoartf2870
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 cd /Users/naratthakorn/Desktop/Attendance\
\
cat << 'EOF' > app.py\
import streamlit as st\
import pandas as pd\
from datetime import datetime, timedelta\
\
st.set_page_config(page_title="\uc0\u3619 \u3632 \u3610 \u3610 \u3605 \u3619 \u3623 \u3592 \u3626 \u3629 \u3610 \u3648 \u3623 \u3621 \u3634 \u3607 \u3635 \u3591 \u3634 \u3609 ", layout="wide")\
\
st.title("\uc0\u3619 \u3632 \u3610 \u3610 \u3605 \u3619 \u3623 \u3592 \u3626 \u3629 \u3610 \u3614 \u3609 \u3633 \u3585 \u3591 \u3634 \u3609 \u3621 \u3639 \u3617 \u3621 \u3591 \u3648 \u3623 \u3621 \u3634 , \u3586 \u3629  OT \u3649 \u3621 \u3632 \u3592 \u3635 \u3609 \u3623 \u3609 \u3623 \u3633 \u3609 \u3627 \u3618 \u3640 \u3604 ")\
st.write("\uc0\u3629 \u3633 \u3611 \u3650 \u3627 \u3621 \u3604 \u3652 \u3615 \u3621 \u3660 \u3619 \u3634 \u3618 \u3591 \u3634 \u3609 \u3605 \u3634 \u3619 \u3634 \u3591 \u3648 \u3623 \u3621 \u3634  (.xlsx) \u3648 \u3614 \u3639 \u3656 \u3629 \u3605 \u3619 \u3623 \u3592 \u3626 \u3629 \u3610 \u3619 \u3634 \u3618 \u3585 \u3634 \u3619 \u3612 \u3636 \u3604 \u3611 \u3585 \u3605 \u3636 ")\
\
uploaded_file = st.file_uploader("\uc0\u3648 \u3621 \u3639 \u3629 \u3585 \u3652 \u3615 \u3621 \u3660  Excel \u3619 \u3634 \u3618 \u3591 \u3634 \u3609 \u3652 \u3607 \u3617 \u3660 ", type=["xlsx"])\
\
def parse_time(time_str):\
    """\uc0\u3649 \u3611 \u3621 \u3591 \u3626 \u3605 \u3619 \u3636 \u3591 \u3648 \u3623 \u3621 \u3634  HH:MM \u3651 \u3627 \u3657 \u3648 \u3611 \u3655 \u3609  datetime object \u3626 \u3635 \u3627 \u3619 \u3633 \u3610 \u3648 \u3611 \u3619 \u3637 \u3618 \u3610 \u3648 \u3607 \u3637 \u3618 \u3610 """\
    try:\
        time_str = str(time_str).strip()\
        if not time_str or time_str == 'nan':\
            return None\
        return datetime.strptime(time_str, "%H:%M")\
    except:\
        return None\
\
if uploaded_file is not None:\
    try:\
        df_raw = pd.read_excel(uploaded_file, header=None)\
        \
        current_emp = None\
        records = []\
\
        for idx, row in df_raw.iterrows():\
            if idx == 0:\
                continue\
            col0 = str(row[0]).strip() if pd.notna(row[0]) else ""\
            if "\uc0\u3649 \u3612 \u3609 \u3585 :" in col0:\
                current_emp = col0\
                continue\
            \
            date_val = row[1]\
            status_val = str(row[2]).strip() if pd.notna(row[2]) else ""\
            shift_val = str(row[3]).strip() if pd.notna(row[3]) else ""\
            in1 = row[4]\
            out1 = row[5]\
            ot_reg = row[17]\
            \
            if pd.isna(date_val) or str(date_val).strip() == "\uc0\u3623 \u3633 \u3609 \u3607 \u3637 \u3656 ":\
                continue\
                \
            records.append(\{\
                'Emp': current_emp,\
                'Date': str(date_val).strip(),\
                'Status': status_val,\
                'Shift': shift_val,\
                'IN': in1,\
                'OUT': out1,\
                'OT_Regular': ot_reg\
            \})\
\
        df = pd.DataFrame(records)\
        df['Date_dt'] = pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='coerce')\
        \
        # \uc0\u3585 \u3619 \u3629 \u3591 \u3623 \u3633 \u3609 \u3652 \u3617 \u3656 \u3648 \u3585 \u3636 \u3609 \u3611 \u3633 \u3592 \u3592 \u3640 \u3610 \u3633 \u3609 \u3607 \u3637 \u3656 \u3604 \u3638 \u3591 \u3586 \u3657 \u3629 \u3617 \u3641 \u3621 \
        max_date = df['Date_dt'].dropna().max()\
        df = df[df['Date_dt'] <= max_date]\
\
        issues = []\
\
        # -------------------------------------------------------------\
        # 1. \uc0\u3605 \u3619 \u3623 \u3592 \u3626 \u3629 \u3610 \u3592 \u3635 \u3609 \u3623 \u3609 \u3623 \u3633 \u3609 \u3627 \u3618 \u3640 \u3604 \u3605 \u3656 \u3629 \u3614 \u3609 \u3633 \u3585 \u3591 \u3634 \u3609  (8 \u3623 \u3633 \u3609 /\u3619 \u3629 \u3610 \u3605 \u3633 \u3604 \u3623 \u3636 \u3585 )\
        # -------------------------------------------------------------\
        off_statuses = ['\uc0\u3623 \u3633 \u3609 \u3627 \u3618 \u3640 \u3604 \u3611 \u3619 \u3632 \u3592 \u3635 \u3626 \u3633 \u3611 \u3604 \u3634 \u3627 \u3660 ', '\u3623 \u3633 \u3609 \u3627 \u3618 \u3640 \u3604 \u3611 \u3619 \u3632 \u3648 \u3614 \u3603 \u3637 ', '\u3623 \u3633 \u3609 \u3627 \u3618 \u3640 \u3604 ', 'Off', 'OFF', 'Holiday']\
        \
        for emp, group in df.groupby('Emp'):\
            # \uc0\u3609 \u3633 \u3610 \u3623 \u3633 \u3609 \u3627 \u3618 \u3640 \u3604 \u3607 \u3633 \u3657 \u3591 \u3627 \u3617 \u3604 \u3651 \u3609 \u3619 \u3629 \u3610 \u3605 \u3633 \u3604 \u3623 \u3636 \u3585 \
            off_days_count = group['Status'].isin(off_statuses).sum()\
            \
            if off_days_count < 8:\
                issues.append(\{\
                    'Emp': emp,\
                    'Date': '\uc0\u3607 \u3633 \u3657 \u3591 \u3619 \u3629 \u3610 \u3605 \u3633 \u3604 \u3623 \u3636 \u3585 ',\
                    'Status': '-',\
                    'Shift': '-',\
                    'IN': '-',\
                    'OUT': '-',\
                    'Issue': f'\uc0\u3623 \u3633 \u3609 \u3627 \u3618 \u3640 \u3604 \u3651 \u3609 \u3619 \u3629 \u3610 \u3605 \u3633 \u3604 \u3623 \u3636 \u3585 \u3609 \u3657 \u3629 \u3618 \u3585 \u3623 \u3656 \u3634  8 \u3623 \u3633 \u3609  (\u3617 \u3637  \{off_days_count\} \u3623 \u3633 \u3609 )'\
                \})\
            elif off_days_count > 8:\
                issues.append(\{\
                    'Emp': emp,\
                    'Date': '\uc0\u3607 \u3633 \u3657 \u3591 \u3619 \u3629 \u3610 \u3605 \u3633 \u3604 \u3623 \u3636 \u3585 ',\
                    'Status': '-',\
                    'Shift': '-',\
                    'IN': '-',\
                    'OUT': '-',\
                    'Issue': f'\uc0\u3623 \u3633 \u3609 \u3627 \u3618 \u3640 \u3604 \u3651 \u3609 \u3619 \u3629 \u3610 \u3605 \u3633 \u3604 \u3623 \u3636 \u3585 \u3617 \u3634 \u3585 \u3585 \u3623 \u3656 \u3634  8 \u3623 \u3633 \u3609  (\u3617 \u3637  \{off_days_count\} \u3623 \u3633 \u3609 )'\
                \})\
\
        # -------------------------------------------------------------\
        # 2. \uc0\u3605 \u3619 \u3623 \u3592 \u3626 \u3629 \u3610 \u3585 \u3634 \u3619 \u3621 \u3639 \u3617 \u3621 \u3591 \u3648 \u3623 \u3621 \u3634  \u3649 \u3621 \u3632  \u3585 \u3634 \u3619 \u3586 \u3629  OT (\u3648 \u3585 \u3636 \u3609  30 \u3609 \u3634 \u3607 \u3637 )\
        # -------------------------------------------------------------\
        for idx, row in df.iterrows():\
            in_missing = pd.isna(row['IN']) or str(row['IN']).strip() in ['', 'nan']\
            out_missing = pd.isna(row['OUT']) or str(row['OUT']).strip() in ['', 'nan']\
            has_ot = pd.notna(row['OT_Regular']) and str(row['OT_Regular']).strip() not in ['', 'nan', '0', '0:00']\
            \
            # \uc0\u3605 \u3619 \u3623 \u3592 \u3626 \u3629 \u3610 \u3621 \u3639 \u3617 \u3626 \u3649 \u3585 \u3609 \u3609 \u3636 \u3657 \u3623 \u3623 \u3633 \u3609 \u3607 \u3635 \u3591 \u3634 \u3609 \
            if row['Status'] == '\uc0\u3623 \u3633 \u3609 \u3607 \u3635 \u3591 \u3634 \u3609 ':\
                if not in_missing and out_missing:\
                    issues.append(\{**row, 'Issue': '\uc0\u3621 \u3639 \u3617 \u3621 \u3591 \u3648 \u3623 \u3621 \u3634 \u3629 \u3629 \u3585  (\u3617 \u3637 \u3649 \u3605 \u3656 \u3648 \u3623 \u3621 \u3634 \u3648 \u3586 \u3657 \u3634 )'\})\
                elif in_missing and not out_missing:\
                    issues.append(\{**row, 'Issue': '\uc0\u3621 \u3639 \u3617 \u3621 \u3591 \u3648 \u3623 \u3621 \u3634 \u3648 \u3586 \u3657 \u3634  (\u3617 \u3637 \u3649 \u3605 \u3656 \u3648 \u3623 \u3621 \u3634 \u3629 \u3629 \u3585 )'\})\
\
            # \uc0\u3605 \u3619 \u3623 \u3592 \u3626 \u3629 \u3610 \u3585 \u3634 \u3619 \u3586 \u3629  OT (\u3648 \u3586 \u3657 \u3634 \u3585 \u3656 \u3629 \u3609 /\u3629 \u3629 \u3585 \u3648 \u3621 \u3607  \u3648 \u3585 \u3636 \u3609  30 \u3609 \u3634 \u3607 \u3637 )\
            if not in_missing and not out_missing and '-' in str(row['Shift']):\
                try:\
                    shift_parts = str(row['Shift']).split('-')\
                    shift_start = parse_time(shift_parts[0])\
                    shift_end = parse_time(shift_parts[1])\
                    \
                    actual_in = parse_time(row['IN'])\
                    actual_out = parse_time(row['OUT'])\
\
                    if actual_in and actual_out and shift_start and shift_end:\
                        # \uc0\u3648 \u3586 \u3657 \u3634 \u3585 \u3656 \u3629 \u3609 \u3585 \u3632 \u3648 \u3585 \u3636 \u3609  30 \u3609 \u3634 \u3607 \u3637 \
                        early_in_threshold = shift_start - timedelta(minutes=30)\
                        # \uc0\u3629 \u3629 \u3585 \u3627 \u3621 \u3633 \u3591 \u3585 \u3632 \u3648 \u3585 \u3636 \u3609  30 \u3609 \u3634 \u3607 \u3637 \
                        late_out_threshold = shift_end + timedelta(minutes=30)\
\
                        if (actual_out > late_out_threshold or actual_in < early_in_threshold) and not has_ot:\
                            issues.append(\{**row, 'Issue': '\uc0\u3648 \u3586 \u3657 \u3634 \u3585 \u3656 \u3629 \u3609 /\u3629 \u3629 \u3585 \u3648 \u3621 \u3607 \u3648 \u3585 \u3636 \u3609 \u3585 \u3632 \u3648 \u3585 \u3636 \u3609  30 \u3609 \u3634 \u3607 \u3637  \u3649 \u3605 \u3656 \u3652 \u3617 \u3656 \u3617 \u3637 \u3619 \u3634 \u3618 \u3585 \u3634 \u3619 \u3586 \u3629  OT'\})\
                except Exception:\
                    pass\
\
        result_df = pd.DataFrame(issues)\
\
        if not result_df.empty:\
            display_df = result_df[['Emp', 'Date', 'Status', 'Shift', 'IN', 'OUT', 'Issue']]\
            st.error(f"\uc0\u3605 \u3619 \u3623 \u3592 \u3614 \u3610 \u3619 \u3634 \u3618 \u3585 \u3634 \u3619 \u3612 \u3636 \u3604 \u3611 \u3585 \u3605 \u3636 \u3607 \u3633 \u3657 \u3591 \u3627 \u3617 \u3604  \{len(display_df)\} \u3619 \u3634 \u3618 \u3585 \u3634 \u3619 ")\
            st.dataframe(display_df, use_container_width=True)\
\
            csv = display_df.to_csv(index=False).encode('utf-8-sig')\
            st.download_button(\
                label="\uc0\u3604 \u3634 \u3623 \u3609 \u3660 \u3650 \u3627 \u3621 \u3604 \u3619 \u3634 \u3618 \u3591 \u3634 \u3609 \u3626 \u3619 \u3640 \u3611  (CSV)",\
                data=csv,\
                file_name='\uc0\u3626 \u3619 \u3640 \u3611 \u3619 \u3634 \u3618 \u3585 \u3634 \u3619 \u3612 \u3636 \u3604 \u3611 \u3585 \u3605 \u3636 .csv',\
                mime='text/csv',\
            )\
        else:\
            st.success("\uc0\u3652 \u3617 \u3656 \u3614 \u3610 \u3619 \u3634 \u3618 \u3585 \u3634 \u3619 \u3612 \u3636 \u3604 \u3611 \u3585 \u3605 \u3636 \u3651 \u3609 \u3652 \u3615 \u3621 \u3660 \u3609 \u3637 \u3657 ")\
\
    except Exception as e:\
        st.error(f"\uc0\u3648 \u3585 \u3636 \u3604 \u3586 \u3657 \u3629 \u3612 \u3636 \u3604 \u3614 \u3621 \u3634 \u3604 \u3651 \u3609 \u3585 \u3634 \u3619 \u3611 \u3619 \u3632 \u3617 \u3623 \u3621 \u3612 \u3621 : \{e\}")\
EOF}