import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import dropbox
import io

st.set_page_config(page_title="KTC | OB Removal", page_icon="description/logo.png")

st.title("OB Removal")

day = [5,6,7,8,9,10,11,12,13,14,15,16,17,18]
night = [17,18,19,20,21,22,23,24,0,1,2,3,4,5,6]

def cekerror_ob(row):
    ksl = []

    if pd.isna(row['Site']) or row['Site'] not in ['THTW', 'TBL3', 'TNPN', 'SIPK', 'TTLP', '12BBML', '11KPCT', '11NESB']:
        ksl.append("Kolom Site Tidak Valid")

    for x in ['Tanggal', 'Pit', 'Jam', 'Shift', 'Ret', 'Jarak', 'Vessel', 'ID_Loader', 'Nama_Operator', 'Operator_ID', 'ID_Hauler', 'Nama_Driver', 'Driver_ID']:
        if pd.isna(row[x]):
            ksl.append(f"Kolom {x} Kosong")

    if pd.isna(row['Shift']) or row['Shift'] not in ['Day', 'Night']:
        ksl.append("Shift Tidak Valid")
    if row['Shift'] == 'Day' and row['Jam'] not in day:
        ksl.append("Jam tidak sesuai Shift")
    if row['Shift'] == 'Night' and row['Jam'] not in night:
        ksl.append("Jam tidak sesuai Shift")
    
    if row['Material'] == 'SOIL':
        if row['ID_Hauler'].startswith('22') and row['Vessel'] > 23:
            ksl.append("Vessel Soil HDT Keliru")
        elif row['ID_Hauler'].startswith('25') and row['Vessel'] > 19:
            ksl.append("Vessel Soil ADT Keliru")
    elif row['Material'] == 'MUD':
        if row['ID_Hauler'].startswith('22') and row['Vessel'] > 11:
            ksl.append("Vessel Mud HDT Keliru")
        elif row['ID_Hauler'].startswith('25') and row['Vessel'] > 9:
            ksl.append("Vessel Mud ADT Keliru")
    else:
        if row['ID_Hauler'].startswith('22') and row['Vessel'] != 22:
            ksl.append("Vessel OB HDT Keliru")
        elif row['ID_Hauler'].startswith('25') and row['Vessel'] != 18:
            ksl.append("Vessel OB ADT Keliru")

    if round(row['Ret'] * row['Vessel'], 3) != row['Produksi']:
        ksl.append("Perhitungan Produksi Salah")

    if len(ksl) == 0:
        return np.nan
    else:
        return ", ".join(ksl)

def comb(row):
    if pd.isna(row['Duplicate']):
        return row['Cek_Error']
    else:
        if pd.isna(row['Cek_Error']):
            return row['Duplicate']
        else:
            err = row['Duplicate'] + ', ' + row['Cek_Error']
            return err

data_ob = st.file_uploader("Upload Excel Files", type=['xlsx','xls'], key="ob")
if data_ob is not None:
    try:
        ob = pd.read_excel(data_ob, usecols='A:AB', sheet_name=0, dtype_backend='numpy_nullable')
    except:
        ob = pd.read_excel(data_ob, sheet_name=0, dtype_backend='numpy_nullable')
    st.write(ob.head())
    ob.dropna(thresh=5, inplace=True)
    st.write(f"Total {len(ob.index)} Rows & {len(ob.columns)} Columns Uploaded")

    if len(ob.columns) < 28:
        ob['Material_New'] = np.nan
    else:
        pass

    if 'Cek_Error' in ob.columns:
        ob = ob.iloc[:, :28]
    else:
        ob = ob.iloc[:, :28]
        ob = ob.set_axis(['Tanggal', 'Week', 'Supervisor', 'Supervisor_ID', 'Foreman',
                        'Foreman_ID', 'Checker', 'Pit', 'Block', 'Seam', 'Dump', 'Fleet',
                        'Tipe_Exca', 'ID_Loader', 'Nama_Operator', 'Operator_ID', 'Tipe_Unit',
                        'ID_Hauler', 'Nama_Driver', 'Driver_ID', 'Shift', 'Jam', 'Ret',
                        'Jarak', 'Vessel', 'Produksi', 'Site', 'Material'], axis=1)
    try:
        ob["Tanggal"] = pd.to_datetime(ob["Tanggal"])
    except:
        st.error(":x: Format Kolom Tanggal Tidak Valid")
        exit()
    
    with st.spinner("On Check ..."):
        def try_num(x):
            try:
                if float(x) <= 0:
                    return np.nan
                else:
                    return float(x)
            except:
                return np.nan
        
        ob['Shift'] = ob['Shift'].str.title().str.strip()
        ob['Pit'] = ob['Pit'].astype(str).str.strip()
        ob['Fleet'] = ob['Fleet'].astype(str).str.strip()
        ob['Block'] = ob['Block'].astype(str).str.strip()
        ob['Site'] = ob['Site'].str.upper().str.strip()
        ob['ID_Loader'] = ob['ID_Loader'].astype(str).str.strip()
        ob['ID_Hauler'] = ob['ID_Hauler'].astype(str).str.strip()

        if len(ob[ob['Material'].notnull()]) >= 1:
            ob['Material'] = ob['Material'].str.upper().str.strip()
        else:
            pass
        
        try:
            ob['Ret'] = round(ob['Ret'],1)
            ob['Jarak'] = round(ob['Jarak'],0)
            ob['Vessel'] = round(ob['Vessel'],3)
            ob['Produksi'] = round(ob['Produksi'],3)
        except:
            st.error(":x: Format Angka di Kolom [ Jarak / Ret / Vessel / Produksi ] Tidak Valid (Terdapat Symbol / Text)")
            exit()

        ob['Produksi'] = ob['Produksi'].apply(lambda x: try_num(x))
        ob['Ret'] = ob['Ret'].apply(lambda x: try_num(x))
        ob['Vessel'] = ob['Vessel'].apply(lambda x: try_num(x))

        ob = ob.replace(['nan', '-', '0', 0, ''], np.nan)
        ob['Jam'] = ob['Jam'].replace(np.nan, 0)

        ob['Duplicate'] = ""
        ob.loc[ob.duplicated(), 'Duplicate'] = 'Duplicate Data'
        ob['Duplicate'] = ob['Duplicate'].replace("", np.nan)

        ob['Cek_Error'] = ob.apply(cekerror_ob, axis=1)

        ob['Cek_Error'] = ob.apply(comb, axis=1)
        ob.drop(columns=['Duplicate'], inplace=True)

    # buffer to use for excel writer
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        ob.to_excel(writer, sheet_name='Sheet1', index=False)
    
    maxob = max(ob["Tanggal"]).strftime('%d %b %Y')
    site = ob['Site'][0]

    if len(ob['Cek_Error'].value_counts()) >= 1:
        st.error("Error Found !")
        st.write(ob['Cek_Error'].value_counts())

        st.download_button(
            label=f":bookmark_tabs: Download File",
            data=buffer,
            file_name=f'{site} OB Removal DB (Koreksi {maxob}).xlsx',
            mime='application/vnd.ms-excel'
            )
    else:
        st.success("No Problem Found")

        st.download_button(
            label=f":bookmark_tabs: Download File",
            data=buffer,
            file_name=f'{site} OB Removal DB ({maxob}).xlsx',
            mime='application/vnd.ms-excel'
            )
        
        dbx = dropbox.Dropbox(
            app_key=st.secrets["api_key"]["App_key"],
            app_secret=st.secrets["api_key"]["App_secret"],
            oauth2_refresh_token=st.secrets["api_key"]["refresh_token"]
        )

        # Define the destination path in Dropbox
        dest_path = f'/Production/OB Removal/{site} OB Removal DB ({maxob}).xlsx'  
        
        if st.button(':eject: Upload File'):
            with st.spinner('Upload On Process'):
                try:
                    dbx.files_upload(buffer.read(), dest_path, mode=dropbox.files.WriteMode.overwrite)
                    st.write(f':white_check_mark: Upload {site} OB Removal DB ({maxob}).xlsx Berhasil')
                except:
                    st.write(f':x: Upload Gagal, Harap Hubungi Admin Untuk Pembaruan')
            