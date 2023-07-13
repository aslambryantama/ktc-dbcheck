import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import io

# buffer to use for excel writer
buffer = io.BytesIO()

st.set_page_config(page_title="Coal Hauling")

st.title("Coal Hauling")

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

dt_supplier = ["BJS", "WMI", "SUBCON"]
dt_ktc = ["KTC", "HANVAN", "HAVAN"]

def cekerror_ch(row):
    ksl = []
    if pd.isna(row['Time_In']) or pd.isna(row['Time_Out']):
        ksl.append("Time In/Out Kosong")
    elif row['Shift'] == 'Day' and row['Jam'] not in day:
        ksl.append("Jam tidak sesuai Shift")
    elif row['Shift'] == 'Night' and row['Jam'] not in night:
        ksl.append("Jam tidak sesuai Shift")
    
    if row['Time_In'] > row['Time_Out']:
        ksl.append("Time In Lebih Besar dari Time Out")
    if row['Previous_Time_Out'] >= row['Time_In']:
        ksl.append("Time In tidak sesuai Time Out sebelumnya")
    if round(row['Berat_Muatan'] - row['Berat_Kosongan'],2) != float(row['Netto']):
        ksl.append("Hasil Timbangan Tidak Sesuai")
    if row['Supplier'] in dt_supplier and row['Driver_ID'] != "0":
        ksl.append("ID Driver tidak sesuai Supplier")
    if row['Supplier'] in dt_ktc and row['Driver_ID'] == "0":
        ksl.append("ID Driver tidak sesuai Supplier")
        
    if len(ksl) == 0:
        return np.nan
    else:
        return ", ".join(ksl)
    
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
    
data_ch = st.file_uploader("Upload Excel Files", type=['xlsx','xls'], key="ch")
if data_ch is not None:
    ch = pd.read_excel(data_ch)
    st.write(ch.head())
    st.write(f"Total {len(ch.index)} Rows & {len(ch.columns)} Columns Uploaded")

    if 'Previous_Time_Out' in ch.columns:
        ch["Tanggal"] = pd.to_datetime(ch["Tanggal"])
        ch["Time_In"] = pd.to_datetime(ch["Time_In"])
        ch["Time_Out"] = pd.to_datetime(ch["Time_Out"])
        ch["Jam_Tambang"] = pd.to_datetime(ch["Jam_Tambang"])
        ch["Previous_Time_Out"] = pd.to_datetime(ch["Previous_Time_Out"])
    else:
        ch = ch.iloc[:, :28]
        ch = ch.set_axis(['Site', 'Tanggal', 'Supervisor', 'Supervisor_ID', 'Foreman',
        'Foreman_ID', 'Checker', 'Checker_ID', 'Pit', 'ID_Hauler', 'Supplier',
        'Jam_Tambang', 'Time_In', 'Time_Out', 'Shift', 'Jenis_Material',
        'Berat_Muatan', 'Berat_Kosongan', 'Netto', 'Ret', 'Driver_ID', 'Nama_Driver',
        'ID_Loader', 'Nama_Operator', 'Operator_ID', 'Tipe_Alat', 'Loading_Area',
        'Dumping_Area'], axis=1)
        
        try:
            ch["Tanggal"] = pd.to_datetime(ch["Tanggal"])
        except:
            st.error("Format Kolom Tanggal Tidak Valid")

        ch[['Jam_Tambang','Time_In','Time_Out']] = ch[['Jam_Tambang','Time_In','Time_Out']].replace([';', '.', ',', '|', '/'] ,':')

        try:
            ch["Jam_Tambang"] = ch["Jam_Tambang"].apply(lambda x: convert_to_datetime(x, '%H:%M') if pd.isna(convert_to_datetime(x, '%H:%M:%S')) else convert_to_datetime(x, '%H:%M:%S'))
            ch["Time_In"] = ch["Time_In"].apply(lambda x: convert_to_datetime(x, '%H:%M') if pd.isna(convert_to_datetime(x, '%H:%M:%S')) else convert_to_datetime(x, '%H:%M:%S'))
            ch["Time_Out"] = ch["Time_Out"].apply(lambda x: convert_to_datetime(x, '%H:%M') if pd.isna(convert_to_datetime(x, '%H:%M:%S')) else convert_to_datetime(x, '%H:%M:%S'))

            ch["Jam_Tambang"] = pd.to_timedelta(ch["Jam_Tambang"].dt.strftime('%H:%M:%S'))
            ch["Jam_Tambang"] = ch["Tanggal"] + ch["Jam_Tambang"]

            ch["Time_In"] = pd.to_timedelta(ch["Time_In"].dt.strftime('%H:%M:%S'))
            ch["Time_In"] = ch["Tanggal"] + ch["Time_In"]

            ch["Time_Out"] = pd.to_timedelta(ch["Time_Out"].dt.strftime('%H:%M:%S'))
            ch["Time_Out"] = ch["Tanggal"] + ch["Time_Out"]

            ch[["Time_In", "Time_Out"]] = ch.apply(attut, axis=1, result_type='expand')
        except:
            st.error('Format Kolom Time In/Out Tidak Valid')

    ch["Jam"] = ch["Time_In"].dt.hour
    ch['Jam'] = ch['Jam'].replace(0, 24)
    ch['Shift'] = ch['Shift'].str.title()

    ch.fillna(value={'Pit':'Unknown'}, inplace=True)

    ch['Pit'] = ch['Pit'].astype(str)
    ch['Pit'] = ch['Pit'].str.strip()

    ch['Supplier'] = ch['Supplier'].str.upper()

    ch['Supplier'] = ch['Supplier'].str.replace('.', '', regex=False)
    ch['Supplier'] = ch['Supplier'].str.replace('PT', '', regex=False)

    ch['Supplier'] = ch['Supplier'].str.strip()

    ch['Driver_ID'] = ch['Driver_ID'].astype(str)
    ch['Driver_ID'] = ch['Driver_ID'].fillna("0")

    ch['Berat_Kosongan'] = round(ch['Berat_Kosongan'],2)
    ch['Berat_Muatan'] = round(ch['Berat_Muatan'],2)
    ch['Netto'] = round(ch['Netto'],2)

    ch['Driver_ID'] = ch['Driver_ID'].astype(str)

    ch['Drivers'] = ch['Driver_ID'] + ch['Nama_Driver']

    ch = ch.sort_values(by=['Site', 'Tanggal', 'Shift', 'Drivers', 'Time_In'])

    ch['Previous_Time_Out'] = ch["Time_Out"].shift(1)

    ch['prev_drivers'] = ch['Drivers'].shift(1)

    ch['Previous_Time_Out'] = ch['Previous_Time_Out'].fillna(ch['Time_In'] - timedelta(seconds=1))

    ch['Previous_Time_Out'] = ch.apply(reblnce, axis=1)

    ch.drop(columns=['Drivers', 'prev_drivers'], inplace=True)

    ch['Cek_Error'] = ch.apply(cekerror_ch, axis=1)

    ch['Cek_Durasi'] = ch.apply(durasi, axis=1)

    ch = ch[['Site', 'Tanggal', 'Supervisor', 'Supervisor_ID', 'Foreman',
        'Foreman_ID', 'Checker', 'Checker_ID', 'Pit', 'ID_Hauler', 'Supplier',
        'Jam_Tambang', 'Previous_Time_Out', 'Time_In', 'Time_Out', 'Shift', 'Jenis_Material',
        'Berat_Muatan', 'Berat_Kosongan', 'Netto', 'Ret', 'Driver_ID',
        'Nama_Driver', 'ID_Loader', 'Nama_Operator', 'Operator_ID', 'Tipe_Alat',
        'Loading_Area', 'Dumping_Area', 'Jam', 'Cek_Error', 'Cek_Durasi']]

    maxch = max(ch["Tanggal"]).strftime('%d %b %Y')
    site = ch['Site'][0]

    if len(ch['Cek_Error'].value_counts()) >= 1:
        st.error("Error Found !")
        st.write(ch['Cek_Error'].value_counts())
    else:
        st.success("No Problem Found")

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        ch.to_excel(writer, sheet_name='Sheet1', index=False)

    st.download_button(
        label=f"Download File",
        data=buffer,
        file_name=f'{site} Coal Hauling DB ({maxch}).xlsx',
        mime='application/vnd.ms-excel'
    )