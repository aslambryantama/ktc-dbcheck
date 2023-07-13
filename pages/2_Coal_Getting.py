import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import io

# buffer to use for excel writer
buffer = io.BytesIO()

st.set_page_config(page_title="Coal Getting")

st.title("Coal Getting")

def attut(row):
    if row['Time_Out'] < row['Time_In']:
        if (row['Time_In'] - row['Time_Out']) > timedelta(hours=12):
            newout = row["Time_Out"] + timedelta(days=1)
            return [row['Time_In'] ,newout]
        else:
            return [row['Time_In'], row["Time_Out"]]
    else:
        return [row['Time_In'], row["Time_Out"]]

day = [5,6,7,8,9,10,11,12,13,14,15,16,17,18]
night = [17,18,19,20,21,22,23,24,0,1,2,3,4,5,6]

def cekerror_cg(row):
    ksh = []
    if pd.isna(row['Time_In']) or pd.isna(row['Time_Out']):
        ksh.append("Time In/Out Kosong")
    elif row['Shift'] == 'Day' and row['Jam'] not in day:
        ksh.append("Jam tidak sesuai Shift")
    elif row['Shift'] == 'Night' and row['Jam'] not in night:
        ksh.append("Jam tidak sesuai Shift")

    if row['Time_In'] > row['Time_Out']:
        ksh.append("Time In Lebih Besar dari Time Out")
    if row['Previous_Time_Out'] >= row['Time_In']:
        ksh.append("Time In tidak sesuai Time Out sebelumnya")

    if len(ksh) == 0:
        return np.nan
    else:
        return ", ".join(ksh)
    
def convert_to_datetime(time_obj, time_format):
    try:
        time_obj = str(time_obj)
        datetime_obj = datetime.strptime(time_obj, time_format)
        datetime_pd = pd.to_datetime(datetime_obj)
        return datetime_pd
    except:
        return np.nan

def reblnce(row):
    if row['Drivers'] != row['prev_drivers']:
        return row['Time_In'] - timedelta(seconds=1)
    else:
        return row['Previous_Time_Out']

def durasi(row):
    try:
        if row['Time_Out'] - row['Time_In'] >= timedelta(minutes=30):
            return "Over 30 Minutes"
        else:
            return np.nan
    except:
        return np.nan

data_cg = st.file_uploader("Upload Excel Files", type=['xlsx','xls'], key="cg")
if data_cg is not None:
    cg = pd.read_excel(data_cg)
    st.write(cg.head())
    st.write(f"Total {len(cg.index)} Rows & {len(cg.columns)} Columns Uploaded")

    if 'Previous_Time_Out' in cg.columns:
        cg["Tanggal"] = pd.to_datetime(cg["Tanggal"])
        cg["Time_In"] = pd.to_datetime(cg["Time_In"])
        cg["Time_Out"] = pd.to_datetime(cg["Time_Out"])
        cg["Previous_Time_Out"] = pd.to_datetime(cg["Previous_Time_Out"])
    else:
        cg = cg.iloc[:, :28]
        cg = cg.set_axis(['Site', 'Tanggal', 'Supervisor', 'Supervisor_ID', 'Foreman',
       'Foreman_ID', 'Checker', 'Checker_ID', 'Pit', 'Block', 'Dump', 'Seam',
       'Loader_Tipe', 'ID_Loader', 'Operator_ID', 'Nama_Operator', 'Hauler_Tipe',
       'ID_Hauler', 'Driver_ID', 'Nama_Driver', 'Time_In', 'Time_Out', 'Job',
       'Material', 'Shift', 'Ret', 'Cap', 'Produksi'], axis=1)

        try:
            cg["Tanggal"] = pd.to_datetime(cg["Tanggal"])
        except:
            st.error("Format Kolom Tanggal Tidak Valid")

        cg['Shift'] = cg['Shift'].str.title()
        
        try:
            cg[['Time_In','Time_Out']] = cg[['Time_In','Time_Out']].replace([';', '.', ',', '|', '/'] ,':')

            cg["Time_In"] = cg["Time_In"].apply(lambda x: convert_to_datetime(x, '%H:%M') if pd.isna(convert_to_datetime(x, '%H:%M:%S')) else convert_to_datetime(x, '%H:%M:%S'))
            cg["Time_Out"] = cg["Time_Out"].apply(lambda x: convert_to_datetime(x, '%H:%M') if pd.isna(convert_to_datetime(x, '%H:%M:%S')) else convert_to_datetime(x, '%H:%M:%S'))

            cg["Time_In"] = pd.to_timedelta(cg["Time_In"].dt.strftime('%H:%M:%S'))
            cg["Time_Out"] = pd.to_timedelta(cg["Time_Out"].dt.strftime('%H:%M:%S'))

            cg["Time_In"] = cg["Tanggal"] + cg["Time_In"]
            cg["Time_Out"] = cg["Tanggal"] + cg["Time_Out"]
            cg[["Time_In", "Time_Out"]] = cg.apply(attut, axis=1, result_type='expand')
        except:
            st.error('Format Kolom Time In/Out Tidak Valid')
            
    cg["Jam"] = cg["Time_In"].dt.hour
    cg['Pit'] = cg['Pit'].astype(str).str.strip()
    cg['Jam'] = cg['Jam'].replace(0, 24)

    cg['Driver_ID'] = cg['Driver_ID'].astype(str)
    cg['Drivers'] = cg['Driver_ID'] + cg['Nama_Driver']

    cg = cg.sort_values(by=['Site', 'Tanggal', 'Shift', 'Drivers', 'Time_In'])
    
    cg['Previous_Time_Out'] = cg["Time_Out"].shift(1)
    cg['prev_drivers'] = cg['Drivers'].shift(1)
    cg['Previous_Time_Out'] = cg['Previous_Time_Out'].fillna(cg['Time_In'] - timedelta(seconds=1))
    
    cg['Previous_Time_Out'] = cg.apply(reblnce, axis=1)
    cg.drop(columns=['Drivers', 'prev_drivers'], inplace=True)

    cg['Cek_Error'] = cg.apply(cekerror_cg, axis=1)
    cg['Cek_Durasi'] = cg.apply(durasi, axis=1)
    
    cg = cg[['Site', 'Tanggal', 'Supervisor', 'Supervisor_ID', 'Foreman',
       'Foreman_ID', 'Checker', 'Checker_ID', 'Pit', 'Block', 'Dump', 'Seam',
       'Loader_Tipe', 'ID_Loader', 'Operator_ID', 'Nama_Operator',
       'Hauler_Tipe', 'ID_Hauler', 'Driver_ID', 'Nama_Driver', 'Previous_Time_Out', 'Time_In',
       'Time_Out', 'Job', 'Material', 'Shift', 'Ret', 'Cap', 'Produksi', 'Jam', 'Cek_Error']]

    maxcg = max(cg["Tanggal"]).strftime('%d %b %Y')
    site = cg['Site'][0]

    if len(cg['Cek_Error'].value_counts()) >= 1:
        st.error("Error Found !")
        st.write(cg['Cek_Error'].value_counts())
    else:
        st.success("No Problem Found")

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        cg.to_excel(writer, sheet_name='Sheet1', index=False)

    st.download_button(
        label=f"Download File",
        data=buffer,
        file_name=f'{site} Coal Getting DB ({maxcg}).xlsx',
        mime='application/vnd.ms-excel'
    )


    

        

    
