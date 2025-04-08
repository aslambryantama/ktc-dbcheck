import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import dropbox
import io

st.set_page_config(page_title="KTC | Fuel Unit", page_icon="description/logo.png")

st.title("Fuel Unit")

site = ['SIPK', 'TNPN', 'THTW', 'TBL3', 'TTLP', 'BCCT', '12BBML', '11KPCT']

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

def night_adjust(row):
    if row['Shift'] == 'Night' and row['Time'].hour <= 6:
        new_time = row['Time'] + timedelta(days=1)
        return new_time
    else:
        return row['Time']

def reblnce(row):
    if row['Unit'] != row['Unit_Clone']:
        if pd.isna(row['HM_Runtime']) or row['HM_Runtime'] == 0:
            now = row['Tanggal'] - timedelta(days=1)
        else:
            prev = round((row['HM_Runtime'] / 24) + (row['HM_Runtime'] % 24 > 0), 0)
            now = row['Tanggal'] - timedelta(days=prev)
        return [row['HM_Start'] ,row['KM_Start'], now]
    else:
        return [row['Previous_HM'], row['Previous_KM'], row['Previous_Date']]
    
def cekerror_fuel(row):
    ksl = []

    if pd.isna(row['Site']) or row['Site'] not in ['THTW', 'TBL3', 'TNPN', 'SIPK', 'TTLP']:
        ksl.append("Kolom Site Tidak Valid")

    if pd.isna(row['Time']):
        ksl.append('Kolom Waktu Tidak Valid')

    for x in ['Tanggal', 'Shift', 'Activity', 'Tangki']:
        if pd.isna(row[x]):
            ksl.append(f"Kolom {x} Kosong")
    try:
        unit = int(row['Unit'])
        
        #if row['HM_Runtime'] == 0 and row['KM_Runtime'] == 0:
        #    ksl.append('HM / KM Kosong')

        if row['HM_Runtime'] == 0 and row['KM_Runtime'] == 0:
            if row['HM_Run'] > 0 or row['KM_Run'] > 0:
                ksl.append('HM / KM No Progress')
            else:
                ksl.append('HM / KM Kosong')
        if row['HM_Start'] > row['HM_Run']:
            ksl.append('HM Terbalik')
        if row['KM_Start'] > row['KM_Run']:
            ksl.append('KM Terbalik')
        if row['Previous_HM'] != row['HM_Start']:
            ksl.append('HM Awal Tidak sesuai HM Sebelumnya')
        if row['Previous_KM'] != row['KM_Start']:
            ksl.append('KM Awal Tidak sesuai KM Sebelumnya')
        if row['Fill_Interval'] < round(row['HM_Runtime'],1):
            ksl.append('Total HM Melebihi Jarak Waktu Pengisian')
        if round(row['HM_Run'] - row['HM_Start'], 2) != row['HM_Runtime'] or row['HM_Runtime'] < 0:
            ksl.append('Kalkulasi HM Tidak Sesuai')
        if round(row['KM_Run'] - row['KM_Start'], 2) != row['KM_Runtime'] or row['KM_Runtime'] < 0:
            ksl.append('Kalkulasi KM Tidak Sesuai')
        if round(row['Flow_Meter_Finish'] - row['Flow_Meter_Start'], 1) != row['Qty_Liter'] or row['Qty_Liter'] < 0:
            ksl.append('Kalkulasi Liter Tidak Sesuai')
        
        if row['HM_Runtime'] >= 1000:
            ksl.append("HM Abnormal Perlu Remark")
        if row['KM_Runtime'] >= 1000:
            ksl.append("KM Abnormal Perlu Remark")
    except:
        pass
    
    if len(ksl) == 0:
        return np.nan
    else:
        return ", ".join(ksl)

def kemb(row):
    if pd.isna(row['Time']):
        return row['Time_xy']
    else:
        return row['Time']

data_fu = st.file_uploader("Upload Excel Files", type=['xlsx','xls'], key="fu")
if data_fu is not None:
    fu = pd.read_excel(data_fu, header=1)
    st.write(fu.head())
    fu.dropna(thresh=5, inplace=True)
    st.write(f"Total {len(fu.index)} Rows & {len(fu.columns)} Columns Uploaded")

    if 'Cek_Error' in fu.columns:
        pass
    else :
        try:
            fu = fu.iloc[:, :23]
            fu = fu.set_axis(['Unit', 'Shift', 'Activity', 'Site', 'HM_Start',
            'HM_Run', 'HM_Runtime', 'Ltr_HM', 'KM_Start', 'KM_Run',
            'KM_Runtime', 'Ltr_KM', 'Time', 'Year', 'Month', 'Tanggal',
            'Flow_Meter_Start', 'Flow_Meter_Finish', 'Qty_Liter', 'Price',
            'Total_Cost', 'Tangki', 'Remark'], axis=1)
        except:
            st.error(":x: Proses Gagal, Format Laporan Fuel Salah")
            exit()
    
    try:
        fu["Tanggal"] = pd.to_datetime(fu["Tanggal"])
    except:
        st.error(":x: Format Kolom Tanggal Tidak Valid")
        exit()
    
    fu['Site'] = fu['Site'].str.upper().str.strip()
    fu['Site'] = fu['Site'].apply(lambda x: x if x in site else np.nan)

    fu['Shift'] = fu['Shift'].str.strip().str.title()

    fu = fu.replace(['nan', '-', '0', 0, ''], np.nan)

    fu['Activity'] = fu['Activity'].str.upper()
    fu['Tangki'] = fu['Tangki'].astype(str)
    fu['Tangki'] = fu['Tangki'].str.upper()
    fu['Tangki'] = fu['Tangki'].str.replace('TS ', '', regex=False).str.strip()

    fu['Time_xy'] = fu['Time']

    try:
        fu['Time'] = fu['Time'].apply(lambda x: convert_to_datetime(x, '%H:%M') if pd.isna(convert_to_datetime(x, '%H:%M:%S')) else convert_to_datetime(x, '%H:%M:%S'))
        fu['Time'] = fu["Tanggal"] + pd.to_timedelta(fu['Time'].dt.strftime('%H:%M:%S'))
        fu['Time'] = fu.apply(night_adjust, axis=1)
    except:
        st.error(':x: Format Kolom Time Tidak Valid, (Format Valid hh:mm)')
        exit()
    
    fu['HM_Start'] = round(fu['HM_Start'], 2)
    fu['HM_Run'] = round(fu['HM_Run'], 2)
    fu['HM_Runtime'] = round(fu['HM_Runtime'], 2)
    fu['KM_Start'] = round(fu['KM_Start'], 2)
    fu['KM_Run'] = round(fu['KM_Run'], 2)
    fu['KM_Runtime'] = round(fu['KM_Runtime'], 2)
    fu['Flow_Meter_Start'] = round(fu['Flow_Meter_Start'], 2)
    fu['Flow_Meter_Finish'] = round(fu['Flow_Meter_Finish'], 2)
    fu['Qty_Liter'] = round(fu['Qty_Liter'], 2)

    fu[['HM_Start','HM_Run', 'HM_Runtime', 'Ltr_HM', 'KM_Start', 'KM_Run', 'KM_Runtime', 'Ltr_KM', 'Flow_Meter_Start', 'Flow_Meter_Finish', 'Qty_Liter']] = \
    fu[['HM_Start','HM_Run', 'HM_Runtime', 'Ltr_HM', 'KM_Start', 'KM_Run', 'KM_Runtime', 'Ltr_KM', 'Flow_Meter_Start', 'Flow_Meter_Finish', 'Qty_Liter']].replace(np.nan, 0)

    fu = fu.sort_values(by=['Unit', 'Tanggal', 'HM_Start'], ascending=True)

    fu['Previous_HM'] = fu["HM_Run"].shift(1)
    fu['Previous_KM'] = fu["KM_Run"].shift(1)
    fu["Previous_Date"] = fu["Tanggal"].shift(1)
    fu['Unit_Clone']= fu['Unit'].shift(1)

    fu[['Previous_HM', 'Previous_KM', 'Previous_Date']] = fu.apply(reblnce, axis=1, result_type='expand')

    fu['Fill_Interval'] = pd.to_timedelta(fu['Tanggal'].dt.date - fu['Previous_Date'].dt.date)
    fu['Fill_Interval'] = (fu['Fill_Interval'].dt.days + 1) * 24

    fu['Cek_Error'] = fu.apply(cekerror_fuel, axis=1)

    fu['Time'] = fu.apply(kemb, axis=1)

    fu = fu[['Unit', 'Shift', 'Activity', 'Site', 'Previous_HM', 'HM_Start',
        'HM_Run', 'HM_Runtime', 'Ltr_HM', 'Previous_KM', 'KM_Start', 'KM_Run',
        'KM_Runtime', 'Ltr_KM', 'Time',  'Year', 'Month', 'Previous_Date', 'Tanggal', 'Fill_Interval',
        'Flow_Meter_Start', 'Flow_Meter_Finish', 'Qty_Liter', 'Price',
        'Total_Cost', 'Tangki', 'Remark', 'Cek_Error']]

    # buffer to use for excel writer
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        fu.to_excel(writer, sheet_name='Sheet1', index=False)
    
    maxfu = max(fu["Tanggal"]).strftime('%d %b %Y')
    site = fu['Site'][0]

    if len(fu['Cek_Error'].value_counts()) >= 1:
        st.error("Error Found !")
        st.write(fu['Cek_Error'].value_counts())

        st.download_button(
            label=f":bookmark_tabs: Download File",
            data=buffer,
            file_name=f'{site} Fuel Unit DB (Koreksi {maxfu}).xlsx',
            mime='application/vnd.ms-excel'
            )
    else:
        st.success("No Problem Found")

        st.download_button(
            label=f":bookmark_tabs: Download File",
            data=buffer,
            file_name=f'{site} Fuel Unit DB ({maxfu}).xlsx',
            mime='application/vnd.ms-excel'
            )
        
        dbx = dropbox.Dropbox(
            app_key=st.secrets["api_key"]["App_key"],
            app_secret=st.secrets["api_key"]["App_secret"],
            oauth2_refresh_token=st.secrets["api_key"]["refresh_token"]
        )

        # Define the destination path in Dropbox
        dest_path = f'/Production/Fuel Unit/{site} Fuel Unit DB ({maxfu}).xlsx'  
        
        if st.button(':eject: Upload File'):
            with st.spinner('Upload On Process'):
                try:
                    dbx.files_upload(buffer.read(), dest_path, mode=dropbox.files.WriteMode.overwrite)
                    st.write(f':white_check_mark: Upload {site} Fuel Unit DB ({maxfu}).xlsx Berhasil')
                except:
                    st.write(f':x: Upload Gagal, Harap Hubungi Admin Untuk Pembaruan')

