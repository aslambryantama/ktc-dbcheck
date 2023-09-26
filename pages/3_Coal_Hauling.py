import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import dropbox
import io

st.set_page_config(page_title="KTC | Coal Hauling", page_icon="description/logo.png")

st.title("Coal Hauling")

def night_adjust_tambang(row):
    if row['Shift'] == 'Night' and row['Jam_Tambang'].hour <= 6:
        new_time = row['Jam_Tambang'] + timedelta(days=1)
        return new_time
    else:
        return row['Jam_Tambang']

def night_adjust_in_2(row):
    if row['Shift'] == 'Night' and row['Time_In'].hour <= 6:
        new_time = row['Time_In'] + timedelta(days=1)
        return new_time
    else:
        return row['Time_In']

def night_adjust_out_2(row):
    if row['Shift'] == 'Night' and row['Time_Out'].hour <= 6:
        new_time = row['Time_Out'] + timedelta(days=1)
        return new_time
    else:
        return row['Time_Out']

day = [5,6,7,8,9,10,11,12,13,14,15,16,17,18]
night = [17,18,19,20,21,22,23,24,0,1,2,3,4,5,6]

dt_supplier = ["BJS", "WMI", "SUBCON"]
dt_ktc = ["HANVAN", "HAVAN"]

def cekerror_ch(row):
    ksl = []
    dt = ['WMI']

    if row['Site'] not in ['THTW', 'TBL3', 'TNPN', 'SIPK', 'TTLP']:
        ksl.append("Site Code Tidak Valid")

    for x in ['Tanggal', 'ID_Loader', 'ID_Hauler', 'Shift', 'Nama_Operator', 'Nama_Driver', 'Netto']:
        if pd.isna(row[x]):
            ksl.append(f"Kolom {x} Kosong")
    try:
        if int(row['ID_Hauler']) < 1:
            ksl.append('No Hauler Tidak Valid')
        else:
            pass
    except:
        ksl.append('No Hauler Tidak Valid')
    
    if row['Supplier'] not in dt:
        if row['Berat_Muatan'] >= 65:
            ksl.append("Berat Muatan > 65")
        if row['Berat_Kosongan'] >= 30:
            ksl.append("Berat Kosongan > 30")
        if row['Netto'] >= 45:
            ksl.append("Netto > 45")
    else:
        pass

    if pd.isna(row['Time_In']) or pd.isna(row['Time_Out']):
        ksl.append("Time In/Out Kosong")
    
    if pd.isna(row['Shift']) or row['Shift'] not in ['Day', 'Night']:
        ksl.append("Shift Tidak Valid")

    if row['Shift'] == 'Day' and row['Jam'] not in day:
        ksl.append("Jam tidak sesuai Shift")
    if row['Shift'] == 'Night' and row['Jam'] not in night:
        ksl.append("Jam tidak sesuai Shift")
    
    if row['Time_In'] >= row['Time_Out']:
        ksl.append("Jarak Time In & Out Tidak Valid")
    if row['Previous_Time_Out'] >= row['Time_In']:
        ksl.append("Time In tidak sesuai Time Out sebelumnya")
    if round(row['Berat_Muatan'] - row['Berat_Kosongan'], 3) != float(row['Netto']):
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
    for x in ['Jam_Tambang', 'Time_In', 'Time_Out']:
        if pd.isna(row[x]):
            ad.append(row[x+'_xy'])
        else:
            ad.append(row[x])
    return ad
    
data_ch = st.file_uploader("Upload Excel Files", type=['xlsx','xls'], key="ch")
if data_ch is not None:
    ch = pd.read_excel(data_ch)
    st.write(ch.head())
    ch.dropna(thresh=5, inplace=True)
    st.write(f"Total {len(ch.index)} Rows & {len(ch.columns)} Columns Uploaded")

    if 'Previous_Time_Out' in ch.columns:
        ch["Tanggal"] = pd.to_datetime(ch["Tanggal"])
        ch["Time_In"] = pd.to_datetime(ch["Time_In"])
        ch["Time_Out"] = pd.to_datetime(ch["Time_Out"])
        ch["Jam_Tambang"] = pd.to_datetime(ch["Jam_Tambang"])
        ch["Previous_Time_Out"] = pd.to_datetime(ch["Previous_Time_Out"])
        ch[['Jam_Tambang_xy','Time_In_xy','Time_Out_xy']] = ch[['Jam_Tambang','Time_In','Time_Out']].copy()
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
            st.error(":x: Format Kolom Tanggal Tidak Valid")
            exit()

        #ch[['Jam_Tambang','Time_In','Time_Out']] = ch[['Jam_Tambang','Time_In','Time_Out']].replace([';', '.', ',', '|', '/'] ,':')
        ch[['Jam_Tambang_xy','Time_In_xy','Time_Out_xy']] = ch[['Jam_Tambang','Time_In','Time_Out']].copy()

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

            ch['Shift'] = ch['Shift'].str.title().str.strip()
            
            ch["Jam_Tambang"] = ch.apply(night_adjust_tambang, axis = 1)
            ch["Time_In"] = ch.apply(night_adjust_in_2, axis = 1)
            ch["Time_Out"] = ch.apply(night_adjust_out_2, axis = 1)
        except:
            st.error(':x: Format Kolom Time In/Out Tidak Valid')
            exit()

    ch["Jam"] = ch["Time_In"].dt.hour
    ch['Jam'] = ch['Jam'].replace(0, 24)

    ch.fillna(value={'Pit':'Unknown'}, inplace=True)

    ch['Site'] = ch['Site'].str.upper().str.strip()
    ch['Pit'] = ch['Pit'].astype(str)
    ch['Pit'] = ch['Pit'].str.strip()

    ch['Supplier'] = ch['Supplier'].str.upper()

    ch['Supplier'] = ch['Supplier'].str.replace('.', '', regex=False)
    ch['Supplier'] = ch['Supplier'].str.replace('PT', '', regex=False)

    ch['Supplier'] = ch['Supplier'].str.strip()

    ch['Driver_ID'] = ch['Driver_ID'].astype(str)
    ch['Driver_ID'] = ch['Driver_ID'].fillna("0")

    ch['Berat_Kosongan'] = round(ch['Berat_Kosongan'],3)
    ch['Berat_Muatan'] = round(ch['Berat_Muatan'],3)
    ch['Netto'] = round(ch['Netto'],3)

    ch['Driver_ID'] = ch['Driver_ID'].astype(str)

    ch['Drivers'] = ch['Driver_ID'] + ch['Nama_Driver']

    ch = ch.sort_values(by=['Site', 'Tanggal', 'Shift', 'Drivers', 'Time_In'])

    ch['Previous_Time_Out'] = ch["Time_Out"].shift(1)

    ch['prev_drivers'] = ch['Drivers'].shift(1)

    ch['Previous_Time_Out'] = ch['Previous_Time_Out'].fillna(ch['Time_In'] - timedelta(seconds=1))

    ch['Previous_Time_Out'] = ch.apply(reblnce, axis=1)

    ch[['ID_Hauler', 'ID_Loader', 'Operator_ID']] = ch[['ID_Hauler', 'ID_Loader', 'Operator_ID']].replace(['nan', '-', '0', 0, ''], np.nan)
    ch['Berat_Kosongan'] = ch['Berat_Kosongan'].replace(np.nan, 0)

    ch['Cek_Error'] = ch.apply(cekerror_ch, axis=1)

    ch['Cek_Durasi'] = ch.apply(durasi, axis=1)

    ch[['Jam_Tambang', 'Time_In', 'Time_Out']] = ch.apply(kemb, axis=1, result_type='expand')

    ch.drop(columns=['Drivers', 'prev_drivers', 'Jam_Tambang_xy', 'Time_In_xy', 'Time_Out_xy'], inplace=True)

    ch = ch[['Site', 'Tanggal', 'Supervisor', 'Supervisor_ID', 'Foreman',
        'Foreman_ID', 'Checker', 'Checker_ID', 'Pit', 'ID_Hauler', 'Supplier',
        'Jam_Tambang', 'Previous_Time_Out', 'Time_In', 'Time_Out', 'Shift', 'Jenis_Material',
        'Berat_Muatan', 'Berat_Kosongan', 'Netto', 'Ret', 'Driver_ID',
        'Nama_Driver', 'ID_Loader', 'Nama_Operator', 'Operator_ID', 'Tipe_Alat',
        'Loading_Area', 'Dumping_Area', 'Jam', 'Cek_Error', 'Cek_Durasi']]
    
    # buffer to use for excel writer
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        ch.to_excel(writer, sheet_name='Sheet1', index=False)

    maxch = max(ch["Tanggal"]).strftime('%d %b %Y')
    site = ch['Site'][0]

    if len(ch['Cek_Error'].value_counts()) >= 1:
        st.error("Error Found !")
        st.write(ch['Cek_Error'].value_counts())

        st.download_button(
        label=f":bookmark_tabs: Download File",
        data=buffer,
        file_name=f'{site} Coal Hauling DB (Koreksi {maxch}).xlsx',
        mime='application/vnd.ms-excel'
        )
    else:
        st.success("No Problem Found")

        st.download_button(
        label=f":bookmark_tabs: Download File",
        data=buffer,
        file_name=f'{site} Coal Hauling DB ({maxch}).xlsx',
        mime='application/vnd.ms-excel'
        )

        dbx = dropbox.Dropbox(
            app_key=st.secrets["api_key"]["App_key"],
            app_secret=st.secrets["api_key"]["App_secret"],
            oauth2_refresh_token=st.secrets["api_key"]["refresh_token"]
        )

        # Define the destination path in Dropbox
        dest_path = f'/Production/Coal Hauling/{site} Coal Hauling DB ({maxch}).xlsx'  # The file will be uploaded to the root folder
        
        if st.button(':eject: Upload File'):
            with st.spinner('Upload On Process'):
                try:
                    dbx.files_upload(buffer.read(), dest_path, mode=dropbox.files.WriteMode.overwrite)
                    st.write(f':white_check_mark: Upload {site} Coal Hauling DB ({maxch}).xlsx Berhasil')
                except:
                    st.write(f':x: Upload Gagal, Harap Hubungi Admin Untuk Pembaruan')

    
    