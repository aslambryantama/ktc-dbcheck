import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import dropbox
import io

st.set_page_config(page_title="KTC | Coal Getting", page_icon="description/logo.png")

st.title("Coal Getting")

def night_adjust_in(row):
    if row['Shift'] == 'Night' and row['Time_In'].hour <= 6:
        new_time = row['Time_In'] + timedelta(days=1)
        return new_time
    else:
        return row['Time_In']

def night_adjust_out(row):
    if row['Shift'] == 'Night' and row['Time_Out'].hour <= 6:
        new_time = row['Time_Out'] + timedelta(days=1)
        return new_time
    else:
        return row['Time_Out']

day = [5,6,7,8,9,10,11,12,13,14,15,16,17,18]
night = [17,18,19,20,21,22,23,24,0,1,2,3,4,5,6]

def cekerror_cg(row):
    ksh = []

    if pd.isna(row['Site']) or row['Site'] not in ['THTW', 'TBL3', 'TNPN', 'SIPK', 'TTLP']:
        ksh.append("Kolom Site Tidak Valid")

    for x in ['Tanggal', 'Shift', 'Produksi', 'Pit',  'ID_Loader', 'Operator_ID', 'Nama_Operator', 'ID_Hauler', 'Driver_ID', 'Nama_Driver',]:
        if pd.isna(row[x]):
            ksh.append(f"Kolom {x} Kosong")

    if pd.isna(row['Time_In']) or pd.isna(row['Time_Out']):
        ksh.append("Format Waktu Tidak Valid")

    if pd.isna(row['Shift']) or row['Shift'] not in ['Day', 'Night']:
        ksh.append("Shift Tidak Valid")

    if row['Shift'] == 'Day' and row['Jam'] not in day:
        ksh.append("Jam tidak sesuai Shift")
    if row['Shift'] == 'Night' and row['Jam'] not in night:
        ksh.append("Jam tidak sesuai Shift")

    if row['Time_In'] >= row['Time_Out']:
        ksh.append("Jarak Time In & Out Tidak Valid")
    if row['Previous_Time_Out'] >= row['Time_In']:
        ksh.append("Time In tidak sesuai Time Out sebelumnya")
    if round(row['Ret'] * row['Cap'], 3) != row['Produksi']:
        ksh.append("Hasil Produksi Tidak Sesuai")

    if len(ksh) == 0:
        return np.nan
    else:
        return ", ".join(ksh)

def convert_to_datetime(time_obj, time_format):
    if isinstance(time_obj, datetime):
        return time_obj
    else:
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

def kemb(row):
    ad = []
    for x in ['Time_In', 'Time_Out']:
        if pd.isna(row[x]):
            ad.append(row[x+'_xy'])
        else:
            ad.append(row[x])
    return ad

def drivers(row):
    if pd.isna(row['Driver_ID']) or row['Driver_ID'] == '0' or row['Driver_ID'] == 0:
        return row['Nama_Driver']
    else:
        return row['Driver_ID']

data_cg = st.file_uploader("Upload Excel Files", type=['xlsx','xls'], key="cg")
if data_cg is not None:
    cg = pd.read_excel(data_cg)
    st.write(cg.head())
    cg.dropna(thresh=5, inplace=True)
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
            st.error(":x: Format Kolom Tanggal Tidak Valid")
            exit()

        cg['Shift'] = cg['Shift'].str.title()

        cg[['Time_In_xy','Time_Out_xy']] = cg[['Time_In','Time_Out']].copy()
        
        try:
            #cg['Time_In'].apply(lambda x: str(x).strip().split(" ")[-1])
            #cg['Time_Out'].apply(lambda x: str(x).strip().split(" ")[-1])

            #cg[['Time_In','Time_Out']] = cg[['Time_In','Time_Out']].replace([';', '.', ',', '|', '/'] ,':')
            
            cg["Time_In"] = cg["Time_In"].apply(lambda x: convert_to_datetime(x, '%H:%M') if pd.isna(convert_to_datetime(x, '%H:%M:%S')) else convert_to_datetime(x, '%H:%M:%S'))
            cg["Time_Out"] = cg["Time_Out"].apply(lambda x: convert_to_datetime(x, '%H:%M') if pd.isna(convert_to_datetime(x, '%H:%M:%S')) else convert_to_datetime(x, '%H:%M:%S'))

            cg["Time_In"] = pd.to_timedelta(cg["Time_In"].dt.strftime('%H:%M:%S'))
            cg["Time_Out"] = pd.to_timedelta(cg["Time_Out"].dt.strftime('%H:%M:%S'))

            cg["Time_In"] = cg["Tanggal"] + cg["Time_In"]
            cg["Time_Out"] = cg["Tanggal"] + cg["Time_Out"]
            
            cg['Shift'] = cg['Shift'].str.title().str.strip()

            cg['Time_In'] = cg.apply(night_adjust_in, axis = 1)
            cg['Time_Out'] = cg.apply(night_adjust_out, axis = 1)
        except:
            st.error(':x: Format Kolom Time In/Out Tidak Valid')
            exit()
            
    cg["Jam"] = cg["Time_In"].dt.hour
    cg['Pit'] = cg['Pit'].astype(str).str.strip()
    cg['Site'] = cg['Site'].str.upper().str.strip()
    cg['Jam'] = cg['Jam'].replace(0, 24)

    cg['Operator_ID'] = cg['Operator_ID'].astype(str)
    cg['Driver_ID'] = cg['Driver_ID'].astype(str)

    cg['Ret'] = round(cg['Ret'], 1)
    cg['Cap'] = round(cg['Cap'], 3)
    cg['Produksi'] = round(cg['Produksi'], 3)
    
    cg['Operator_ID'] = cg['Operator_ID'].str.replace('^0.*', '0', regex=True)
    cg['Driver_ID'] = cg['Driver_ID'].str.replace('^0.*', '0', regex=True)
    cg = cg.replace(['nan', '-', '0', 0, ''], np.nan)
    cg['Jam'] = cg['Jam'].replace(np.nan, 0)

    cg['Drivers'] = cg.apply(drivers, axis=1)
    cg = cg.sort_values(by=['Site', 'Tanggal', 'Shift', 'Drivers', 'Time_In'])

    cg['Previous_Time_Out'] = cg["Time_Out"].shift(1)
    cg['prev_drivers'] = cg['Drivers'].shift(1)
    cg['Previous_Time_Out'] = cg['Previous_Time_Out'].fillna(cg['Time_In'] - timedelta(seconds=1))
    
    cg['Previous_Time_Out'] = cg.apply(reblnce, axis=1)

    cg['Cek_Error'] = cg.apply(cekerror_cg, axis=1)
    cg['Cek_Durasi'] = cg.apply(durasi, axis=1)
    
    cg[['Time_In', 'Time_Out']] = cg.apply(kemb, axis=1, result_type='expand')
    
    cg.drop(columns=['Drivers', 'prev_drivers', 'Time_In_xy', 'Time_Out_xy'], inplace=True)
    
    cg = cg[['Site', 'Tanggal', 'Supervisor', 'Supervisor_ID', 'Foreman',
       'Foreman_ID', 'Checker', 'Checker_ID', 'Pit', 'Block', 'Dump', 'Seam',
       'Loader_Tipe', 'ID_Loader', 'Operator_ID', 'Nama_Operator',
       'Hauler_Tipe', 'ID_Hauler', 'Driver_ID', 'Nama_Driver', 'Previous_Time_Out', 'Time_In',
       'Time_Out', 'Job', 'Material', 'Shift', 'Ret', 'Cap', 'Produksi', 'Jam', 'Cek_Error']]

    # buffer to use for excel writer
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        cg.to_excel(writer, sheet_name='Sheet1', index=False)
    
    maxcg = max(cg["Tanggal"]).strftime('%d %b %Y')
    site = cg['Site'][0]

    if len(cg['Cek_Error'].value_counts()) >= 1:
        st.error("Error Found !")
        st.write(cg['Cek_Error'].value_counts())
    
        st.download_button(
        label=f":bookmark_tabs: Download File",
        data=buffer,
        file_name=f'{site} Coal Getting DB (Koreksi {maxcg}).xlsx',
        mime='application/vnd.ms-excel',
        )
    else:
        st.success("No Problem Found")

        st.download_button(
        label=f":bookmark_tabs: Download File",
        data=buffer,
        file_name=f'{site} Coal Getting DB ({maxcg}).xlsx',
        mime='application/vnd.ms-excel',
        )
 
        #Authenticate with Dropbox
        
        #Cara Lama Harus Generate Access Token setiap saat
        #access_token = st.secrets["api_key"]["token"]
        #dbx = dropbox.Dropbox(access_token)
        
        dbx = dropbox.Dropbox(
            app_key=st.secrets["api_key"]["App_key"],
            app_secret=st.secrets["api_key"]["App_secret"],
            oauth2_refresh_token=st.secrets["api_key"]["refresh_token"]
        )

        # Define the destination path in Dropbox
        dest_path = f'/Production/Coal Getting/{site} Coal Getting DB ({maxcg}).xlsx'  # The file will be uploaded to the root folder
        
        if st.button(':eject: Upload File'):
            with st.spinner('Upload On Process'):
                try:
                    dbx.files_upload(buffer.read(), dest_path, mode=dropbox.files.WriteMode.overwrite)
                    st.write(f':white_check_mark: Upload {site} Coal Getting DB ({maxcg}).xlsx Berhasil')
                except:
                    st.write(f':x: Upload Gagal, Harap Hubungi Admin Untuk Pembaruan')

    

        

    

