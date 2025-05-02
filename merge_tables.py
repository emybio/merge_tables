import pandas as pd
import os
import logging
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Logging settings
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

file_path = "template.xlsx"

# Check the file is exist.
if not os.path.exists(file_path):
    raise FileNotFoundError(f"{file_path} no such file!")

# Open the excel file and select the "Kontak Alarmı" sheet
workbook = load_workbook(file_path)
sheet = workbook["Kontak Alarmı"]

# Take the hyperlinks in "Haritada Göster" column
hyperlink_dict = {}  # Dictionary to keep Cihaz Numarası, Tarih ve URL'leri
for row in sheet.iter_rows(min_row=2, values_only=False):  # Skip title line
    cihaz_numarasi = sheet.cell(row=row[0].row, column=2).value  # Cihaz Numarası (1. column)
    tarih = sheet.cell(row=row[0].row, column=4).value  # Tarih (4. column)
    haritada_goster = sheet.cell(row=row[0].row, column=8)  # Haritada Göster (8. column)
    
    if haritada_goster.hyperlink:  # if there is link in the cell
        key = (str(cihaz_numarasi), pd.to_datetime(tarih).date())  # Combining Cihaz Numarası ve Tarih 
        hyperlink_dict[key] = haritada_goster.hyperlink.target  # Save URL

# Close excel file
workbook.close()

# Read excel file with pandas
try:
    logging.info("Reading excel file...")
    xls = pd.ExcelFile(file_path)
except Exception as e:
    raise Exception(f"An error occured while reading excel file: {e}")

# Read data
makine_listesi = pd.read_excel(xls, "Makine Listesi", usecols=["Cihaz Numarası", "Tür"])
kontak_alarm = pd.read_excel(xls, "Kontak Alarmı", usecols=["Cihaz Numarası", "Tarih", "Haritada Göster", "Adres"])
kontak_acik_sure = pd.read_excel(xls, "Kontak Açık Kalma Süresi", usecols=["Cihaz Numarası", "Tarih", "Plaka", "Araç Grubu", "Gün", "Kontak Açık Süresi", "Toplam Mesafe (km)"])

# Clean up unnecessary spaces in column names
makine_listesi.columns = makine_listesi.columns.str.strip()
kontak_alarm.columns = kontak_alarm.columns.str.strip()
kontak_acik_sure.columns = kontak_acik_sure.columns.str.strip()

# Convert device numbers to string format.
makine_listesi["Cihaz Numarası"] = makine_listesi["Cihaz Numarası"].astype(str)
kontak_alarm["Cihaz Numarası"] = kontak_alarm["Cihaz Numarası"].astype(str)
kontak_acik_sure["Cihaz Numarası"] = kontak_acik_sure["Cihaz Numarası"].astype(str)

# Convert date columns
date_formats = ["%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"]

def convert_to_date(date_series):
    for fmt in date_formats:
        try:
            return pd.to_datetime(date_series, format=fmt, errors="raise")
        except ValueError:
            continue
    return pd.to_datetime(date_series, errors="coerce")

kontak_alarm["Tarih"] = convert_to_date(kontak_alarm["Tarih"])
kontak_acik_sure["Tarih"] = convert_to_date(kontak_acik_sure["Tarih"])

# Update the URLs in "Haritada Göster" column
def get_haritada_goster_url(cihaz_numarasi, tarih):
    key = (str(cihaz_numarasi), pd.to_datetime(tarih).date())
    return hyperlink_dict.get(key, "")

kontak_alarm["Haritada Göster"] = kontak_alarm.apply(lambda row: get_haritada_goster_url(row["Cihaz Numarası"], row["Tarih"]), axis=1)

# Filter records matching device number in Machine List
filtered_kontak_acik_sure = kontak_acik_sure[kontak_acik_sure["Cihaz Numarası"].isin(makine_listesi["Cihaz Numarası"])]
filtered_kontak_alarm = kontak_alarm[kontak_alarm["Cihaz Numarası"].isin(makine_listesi["Cihaz Numarası"])]

# Merge tables
try:
    logging.info("Tablolar birleştiriliyor...")
    merged_df = filtered_kontak_acik_sure.merge(filtered_kontak_alarm, on=["Cihaz Numarası", "Tarih"], how="left")
    merged_df = merged_df.merge(makine_listesi, on="Cihaz Numarası", how="left")  # Add "Tür" column
except Exception as e:
    raise Exception(f"Tablolar birleştirilirken bir hata oluştu: {e}")

# Check and merge duplicate records
merged_df = merged_df.groupby(["Cihaz Numarası", "Tarih"], as_index=False).agg({
    "Plaka": "first",
    "Araç Grubu": "first",
    "Gün": "first",
    "Kontak Açık Süresi": "first",
    "Toplam Mesafe (km)": "sum",
    "Haritada Göster": "first",
    "Tür": "first",
    "Adres": "first"
})

# Convert "Kontak Açık Süresi" to hours
def convert_to_hours(time_str):
    try:
        hours = 0
        minutes = 0
        seconds = 0
        
        if "sa" in time_str:
            hours = int(time_str.split("sa")[0].strip())
        if "dk" in time_str:
            minutes = int(time_str.split("dk")[0].split()[-1].strip())
        if "sn" in time_str:
            seconds = int(time_str.split("sn")[0].split()[-1].strip())
        
        total_hours = hours + (minutes / 60) + (seconds / 3600)
        return total_hours
    except:
        return 0

# Add blank line on each machine change
merged_df["Boş Satır"] = merged_df["Cihaz Numarası"] != merged_df["Cihaz Numarası"].shift(1)
final_df = pd.DataFrame()

for _, group in merged_df.groupby("Cihaz Numarası"):
    final_df = pd.concat([final_df, group], ignore_index=True)
    final_df = pd.concat([final_df, pd.DataFrame([{}])], ignore_index=True)  # Add empty line

# Remove empty column Boş sütunu kaldır
final_df = final_df.drop(columns=["Boş Satır"])

# Edit column order
final_df = final_df[["Cihaz Numarası", "Plaka", "Araç Grubu", "Tür", "Tarih", "Gün", "Kontak Açık Süresi", "Toplam Mesafe (km)", "Haritada Göster", "Adres"]]

# Replace NaN values ​​in the "Haritada Göster" column with empty string
final_df["Haritada Göster"] = final_df["Haritada Göster"].fillna("")

# Save excel
wb = Workbook()
ws = wb.active

# Add titles
for col_num, column_title in enumerate(final_df.columns, 1):
    ws.cell(row=1, column=col_num, value=column_title)

# Add datas
for row_num, row in enumerate(dataframe_to_rows(final_df, index=False, header=False), 2):
    for col_num, value in enumerate(row, 1):
        cell = ws.cell(row=row_num, column=col_num, value=value)
        if final_df.columns[col_num - 1] == "Haritada Göster" and value and isinstance(value, str):  # Add link
            cell.value = "Haritada Göster"  # Display text
            cell.hyperlink = value  # Link
            cell.font = Font(color="000000", underline="none")  # Blue and underlined
        elif final_df.columns[col_num - 1] == "Kontak Açık Süresi":  # Kontak Açık Süresi sütunu
            try:
                # Convert time to hours
                total_hours = convert_to_hours(value)
                if 0<total_hours < 5 or 7<total_hours:  # less then 5 or more then 7
                    cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Redd fill
            except:
                pass  # Skip in case of error

# Edit column width
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2) * 1.2
    ws.column_dimensions[column].width = adjusted_width

# Save excel
try:
    wb.save("report.xlsx")
    logging.info("✅ Done, report.xlsx created!")
except Exception as e:
    raise Exception(f"An error occured while saving repor.xlsx: {e}")