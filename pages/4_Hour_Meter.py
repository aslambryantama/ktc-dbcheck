import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import dropbox
import io

st.set_page_config(page_title="KTC | Hour Meter", page_icon="description/logo.png")

st.title("Hour Meter")

def cekerror_hm(row):
    ksl = []
    if pd.isna(row['Shift']) or row['Shift'] not in ['Day', 'Night']:
        ksl.append("Shift Tidak Valid")
    if row['Total_HM'] == 0 and row['Total_KM'] == 0:
        if row['HM_Finish'] > 0 or row['KM_Finish'] > 0:
            ksl.append('HM / KM No Progress')
        else:
            ksl.append('HM / KM Kosong')
    if row['Previous_HM'] != row['HM_Start']:
        ksl.append('HM Awal Tidak sesuai HM Sebelumnya')
    if row['Previous_KM'] != row['KM_Start']:
        ksl.append('KM Awal Tidak sesuai KM Sebelumnya')
    if row['Fill_Interval'] < round(row['Total_HM'],1):
        ksl.append('Total HM Abnormal')
    if round(row['HM_Finish'] - row['HM_Start'], 2) != row['Total_HM'] or row['Total_HM'] < 0:
        ksl.append('Kalkulasi HM Tidak Sesuai')
    if round(row['KM_Finish'] - row['KM_Start'], 2) != row['Total_KM'] or row['Total_KM'] < 0:
        ksl.append('Kalkulasi KM Tidak Sesuai')
    
    if len(ksl) == 0:
        return np.nan
    else:
        return ", ".join(ksl)

def reblnce(row):
    if row['Unit'] != row['Unit_Clone']:
        if row['Total_HM'] == 0:
            now = row['Tanggal'] - timedelta(days=1)
        else:
            prev = round((row['Total_HM'] / 24) + (row['Total_HM'] % 24 > 0), 0)
            now = row['Tanggal'] - timedelta(days=prev)
        return [row['HM_Start'] ,row['KM_Start'], now]
    else:
        return [row['Previous_HM'], row['Previous_KM'], row['Previous_Date']]
    
data_hm = st.file_uploader("Upload Excel Files", type=['xlsx','xls'], key="hm")
if data_hm is not None:
    hm = pd.read_excel(data_hm)
    st.write(hm.head())
    hm.dropna(thresh=5, inplace=True)
    st.write(f"Total {len(hm.index)} Rows & {len(hm.columns)} Columns Uploaded")
    
    file_hm = data_hm.name
    file_hm = file_hm.replace('_', ' ')
    file_hm = file_hm.split('.')[0]

    if 'Cek_Error' in hm.columns:
        pass
    else :
        try:
            hm = hm.iloc[:, :12]
            hm = hm.set_axis(["Tanggal", "Unit", "NIK", "Nama_Operator", "HM_Start", "HM_Finish", "Shift", "Total_HM", "KM_Start", "KM_Finish", "Total_KM", "Remark"], axis=1)
        except:
            st.error(":x: Proses Gagal, Format Laporan HM Salah")
            exit()
    
    try:
        hm["Tanggal"] = pd.to_datetime(hm["Tanggal"])
    except:
        st.error(":x: Format Kolom Tanggal Tidak Valid")
        exit()

    hm['Unit'] = hm['Unit'].astype(str)
    hm['Unit'] = hm['Unit'].apply(lambda x: x.split('.')[0])
    hm = hm[~hm['Unit'].isnull()]

    hm["Shift"] = hm["Shift"].str.title().str.strip()

    hm[["HM_Start", "HM_Finish", "Total_HM", "KM_Start", "KM_Finish", "Total_KM"]] = hm[["HM_Start", "HM_Finish", "Total_HM", "KM_Start", "KM_Finish", "Total_KM"]].fillna(0)

    hm['HM_Start'] = round(hm['HM_Start'], 2)
    hm['HM_Finish'] = round(hm['HM_Finish'], 2)
    hm['KM_Start'] = round(hm['KM_Start'], 2)
    hm['KM_Finish'] = round(hm['KM_Finish'], 2)
    hm['Total_HM'] = round(hm['Total_HM'], 2)
    hm['Total_KM'] = round(hm['Total_KM'], 2)

    hm = hm.sort_values(by=['Unit', 'Tanggal', 'Shift'], ascending=[True, True, True])
    
    hm['Previous_HM'] = hm["HM_Finish"].shift(1)
    hm['Previous_KM'] = hm["KM_Finish"].shift(1)
    hm['Previous_Date'] = hm["Tanggal"].shift(1)
    hm['Unit_Clone'] = hm["Unit"].shift(1)

    hm[['Previous_HM', 'Previous_KM', 'Previous_Date']] = hm.apply(reblnce, axis=1, result_type='expand')

    hm['Fill_Interval'] = pd.to_timedelta(hm['Tanggal'].dt.date - hm['Previous_Date'].dt.date)
    hm['Fill_Interval'] = (hm['Fill_Interval'].dt.days) * 24

    hm['Fill_Interval'] = hm['Fill_Interval'].replace(0, 24)

    hm['Cek_Error'] = hm.apply(cekerror_hm, axis=1)

    hm = hm[['Tanggal', 'Unit', 'NIK', 'Nama_Operator', 'Shift', 
    'Previous_HM', 'HM_Start', 'HM_Finish', 'Total_HM', 
    'Previous_KM', 'KM_Start', 'KM_Finish', 'Total_KM', 
    'Remark', 'Previous_Date', 'Cek_Error']]

    # buffer to use for excel writer
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        hm.to_excel(writer, sheet_name='Sheet1', index=False)
    
    maxhm = max(hm["Tanggal"]).strftime('%d %b %Y')

    if len(hm['Cek_Error'].value_counts()) >= 1:
        st.error("Error Found !")
        st.write(hm['Cek_Error'].value_counts())

        st.download_button(
            label=f":bookmark_tabs: Download File",
            data=buffer,
            file_name=f'{file_hm} (Koreksi {maxhm}).xlsx',
            mime='application/vnd.ms-excel'
            )
    else:
        st.success("No Problem Found")

        st.download_button(
            label=f":bookmark_tabs: Download File",
            data=buffer,
            file_name=f'{file_hm} ({maxhm}).xlsx',
            mime='application/vnd.ms-excel'
            )
        
        dbx = dropbox.Dropbox(
            app_key=st.secrets["api_key"]["App_key"],
            app_secret=st.secrets["api_key"]["App_secret"],
            oauth2_refresh_token=st.secrets["api_key"]["refresh_token"]
        )

        # Define the destination path in Dropbox
        dest_path = f'/Production/Hour Meter/{file_hm} ({maxhm}).xlsx'  
        
        if st.button(':eject: Upload File'):
            with st.spinner('Upload On Process'):
                try:
                    dbx.files_upload(buffer.read(), dest_path, mode=dropbox.files.WriteMode.overwrite)
                    st.write(f':white_check_mark: Upload {file_hm} ({maxhm}).xlsx Berhasil')
                except:
                    st.write(f':x: Upload Gagal, Harap Hubungi Admin Untuk Pembaruan')