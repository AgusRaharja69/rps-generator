import gspread
from google.oauth2.service_account import Credentials
import datetime
import time
import locale
from collections import Counter
import re
from google.colab import files

# Upload credentials.json
uploaded = files.upload()

# Try setting locale to Indonesian
try:
    locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')
except locale.Error:
    print("Locale 'id_ID.UTF-8' not available. Using default.")
    
# Set locale to Indonesian
locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')

# Define the scopes
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Authenticate with Google Sheets
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# Ask user for the Google Sheets URL
sheet_input_url = input("Masukkan URL Google Spreadsheet: ").strip()

# Open the spreadsheet by URL
sheet_input = client.open_by_url(sheet_input_url)

# Select the specific sheets
list_matkul = sheet_input.worksheet("1. List Matkul")
list_cpl = sheet_input.worksheet("2. List CPL")
list_cpmk = sheet_input.worksheet("3. List CPMK")
list_subcpmk = sheet_input.worksheet("4. List SubCPMK")
list_dosen = sheet_input.worksheet("5. List Dosen")
temp_rps_in = sheet_input.worksheet("Template RPS")
temp_kontrak_in = sheet_input.worksheet("Template Kontrak")
temp_rubrik_H1_in = sheet_input.worksheet("Template Rubrik H1")
temp_rubrik_H2_in = sheet_input.worksheet("Template Rubrik H2")
temp_rubrik_H3_in = sheet_input.worksheet("Template Rubrik H3")
temp_rubrik_A1_in = sheet_input.worksheet("Template Rubrik A1")
temp_rubrik_A2_in = sheet_input.worksheet("Template Rubrik A2")
temp_rubrik_A3_in = sheet_input.worksheet("Template Rubrik A3")
temp_rubrik_SP1_in = sheet_input.worksheet("Template Rubrik SP1")
temp_rtm_in = sheet_input.worksheet("Template RTM")
temp_porto_in = sheet_input.worksheet("Template Portofolio Penilaian")
temp_nilai_in = sheet_input.worksheet("Template Nilai Mahasiswa")

# --- dictionary lookup cepat ---
TEMPLATE_MAP = {
    "H1": temp_rubrik_H1_in,
    "H2": temp_rubrik_H2_in,
    "H3": temp_rubrik_H3_in,
    "A1": temp_rubrik_A1_in,
    "A2": temp_rubrik_A2_in,
    "A3": temp_rubrik_A3_in,
    "SP1": temp_rubrik_SP1_in,
}

PENUGASAN_MAP = {
    "H1": "Tugas Pemahaman Materi",
    "H2": "UTS/UAS",
    "H3": "Tugas Proposal",
    "A1": "Tugas Makalah",
    "A2": "Penugasan  PBL",
    "A3": "Laporan Hasil Capstone Project",
    "SP1": "Presentasi Hasil",
}

# Get all data from the worksheets
data_dosen = list_dosen.get_all_values()
data_matkul = list_matkul.get_all_values()
data_cpl = list_cpl.get_all_values()
data_cpmk = list_cpmk.get_all_values()
data_subcpmk = list_subcpmk.get_all_values()

print(data_dosen)
# Extract dosen names for selection
dosen_options = [row[0] for row in data_dosen[1:]]  # Skip header, take first column
print("Pilih nama dosen dari daftar berikut:")
for i, dosen in enumerate(dosen_options, 1):
    print(f"{i}. {dosen}")
dosen_choice = int(input("Masukkan nomor dosen: ")) - 1
if 0 <= dosen_choice < len(dosen_options):
    selected_dosen = dosen_options[dosen_choice]
else:
    print("Pilihan tidak valid. Menggunakan dosen pertama sebagai default.")
    selected_dosen = dosen_options[0] if dosen_options else "I Kadek Agus Wahyu Raharja, S.T., M.T."

# Get matkul input and find data
for row in data_matkul[1:]:
    print(row[3])

mk_input = input("Masukkan nama mata kuliah: ")
matkul_data = {}
for row in data_matkul[1:]:  # Skip header
    if row[3].lower() == mk_input.lower():  # Match based on MATA KULIAH column (index 3)
        matkul_data = {
            "SEMESTER": row[0],
            "KODE": row[1],
            "KODE MK": row[2],
            "MATA KULIAH": row[3],
            "KATEGORI": row[4],
            "SKS": row[5],
            "TEORI": row[6],
            "PRAKTIKUM": row[7],
            "PRAKTIK": row[8],
            "JUMLAH": row[9],
            "KODE DOKUMEN RPS": row[10],
            "KODE DOKUMEN RTM": row[11],
            "KODE DOKUMEN KONTRAK": row[12],
            "KODE DOKUMEN RUBRIK": row[13]
        }
        break

if not matkul_data:
    print(f"Mata Kuliah '{mk_input}' tidak ditemukan.")
else:
    # Display found matkul data
    print("Data Mata Kuliah yang ditemukan:")
    for key, value in matkul_data.items():
        print(f"{key}: {value}")

    # Build arrays for CPL, CPMK, and SubCPMK without duplicates, preserving order
    cpl_list = []
    cpmk_list = []
    subcpmk_list = []
    cpl_set = []
    cpmk_set = []
    subcpmk_set = []

    # Process data from 4. List SubCPMK
    for row in data_subcpmk[1:]:  # Skip header
        if row[0] == matkul_data["KODE"]:  # Match based on KODE
            cpl = row[2]  # CPL (index 2)
            cpmk = row[4]  # CPMK (index 4)
            subcpmk = row[6]  # SubCPMK (index 6)
            if cpl:
                cpl_set.append(cpl)
                if cpl not in cpl_list:
                    cpl_list.append(cpl)
            if cpmk:
                cpmk_set.append(cpmk)
                if cpmk not in cpmk_list:
                    cpmk_list.append(cpmk)
            if subcpmk:
                subcpmk_set.append(subcpmk)
                if subcpmk not in subcpmk_list:
                    subcpmk_list.append(subcpmk)

    print(subcpmk_set)
    # Display the arrays
    print("CPL (tanpa duplikat, urutan asli):", cpl_list)
    print("CPMK (tanpa duplikat, urutan asli):", cpmk_list)
    print("SubCPMK (tanpa duplikat, urutan asli):", subcpmk_list)

    kelas = []
    tahun_ajaran = []
    jumlah_mahasiswa = []
    hari_masuk = []
    lokasi_kelas = []
    deskripsi_matkul = ""
    numbered_materi_tanpa_uts_uas = []
    numbered_pustaka = []
    numbered_kriteria = []
    numbered_indikator = []
    only_numbered_indikator = []
    subcpmk_mingguan_description = []
    bobot = []
    minggu_ke = []
    subcpmk_mingguan = []
    # Function to update the sheet with course data
    def update_rps_sheet(data, cpl_list, cpmk_list, subcpmk_list, dosen):
        global kelas, jumlah_mahasiswa, hari_masuk, lokasi_kelas, subcpmk_mingguan
        global deskripsi_matkul, subcpmk_mingguan_description, tahun_ajaran
        global numbered_materi_tanpa_uts_uas, numbered_pustaka
        global numbered_kriteria, numbered_indikator, bobot, minggu_ke
        # Duplicate Template RPS and rename to RPS
        new_rps_worksheet = sheet_input.duplicate_sheet(temp_rps_in.id, new_sheet_name="RPS")
        worksheet = sheet_input.worksheet("RPS")
        # Update basic course data       
        worksheet.update('C6', [[data['KODE MK']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('A6', [[data['MATA KULIAH']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('J3', [[data['KODE DOKUMEN RPS']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('E6', [[data['KATEGORI']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('F6', [[data['SKS']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('I6', [[data['SEMESTER']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('J6', [[datetime.datetime.now().strftime('%d %B %Y')]])  # Tanggal penetapan
        time.sleep(1)  # Delay to respect quota
        worksheet.update('F8', [[data['TEORI'] or '']])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('G8', [[data['PRAKTIKUM'] or '']])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('H8', [[data['PRAKTIK'] or '']])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('C10', [[dosen]])  
        time.sleep(1)  # Delay to respect quota

        # Update CPL with looping
        cpl_start_row = 12  # Starting row for CPL
        if cpl_list:
            # Insert rows for CPL data
            if len(cpl_list) > 0:
                worksheet.insert_rows([[None] * worksheet.col_count] * len(cpl_list), cpl_start_row)  # Insert rows at the specified row
                time.sleep(1)  # Delay to respect quota
            # Merge cells from C to J for each inserted row
            for i in range(len(cpl_list)):
                worksheet.merge_cells(f'C{cpl_start_row + i}:J{cpl_start_row + i}')
                time.sleep(1)  # Delay to respect quota
            # Write CPL list starting from the next row
            worksheet.update(f'B{cpl_start_row}:B{cpl_start_row + len(cpl_list)}', [[item] for item in cpl_list])
            time.sleep(1)  # Delay to respect quota
            # Get CPL descriptions from data_cpl
            cpl_list_description = []
            for cpl_code in cpl_list:
                for row in data_cpl[1:]:  # Skip header
                    if row[0] == cpl_code:  # Match CPL code (assuming column 0 is CPL code)
                        cpl_list_description.append(row[1] if len(row) > 1 else "")  # Assuming column 1 is description
                        break
            # Write CPL descriptions to column C
            worksheet.update(f'C{cpl_start_row}:C{cpl_start_row + len(cpl_list)}', [[item] for item in cpl_list_description])
            time.sleep(1)  # Delay to respect quota

        # Update CPMK with looping (start 4 rows after CPL)
        cpmk_start_row = cpl_start_row + len(cpmk_list) + 1
        if cpmk_list:
            # Insert rows for CPMK data
            if len(cpmk_list) > 0:
                worksheet.insert_rows([[None] * worksheet.col_count] * len(cpmk_list), cpmk_start_row)
                time.sleep(1)  # Delay to respect quota
            # Merge cells from C to J for each inserted row
            for i in range(len(cpmk_list)):
                worksheet.merge_cells(f'C{cpmk_start_row + i}:J{cpmk_start_row + i}')
                time.sleep(1)  # Delay to respect quota
            # Write CPMK list
            worksheet.update(f'B{cpmk_start_row}:B{cpmk_start_row + len(cpmk_list)}', [[item] for item in cpmk_list])
            time.sleep(1)  # Delay to respect quota
            # Get CPMK descriptions from data_cpmk
            cpmk_list_description = []
            for cpmk_code in cpmk_list:
                for row in data_cpmk[1:]:  # Skip header
                    if row[0] == cpmk_code:  # Match CPMK code (assuming column 0 is CPMK code)
                        cpmk_list_description.append(row[1] if len(row) > 1 else "")  # Assuming column 1 is description
                        break
            # Write CPMK descriptions to column C
            worksheet.update(f'C{cpmk_start_row}:C{cpmk_start_row + len(cpmk_list)}', [[item] for item in cpmk_list_description])
            time.sleep(1)  # Delay to respect quota

        # Update SubCPMK with looping (start 5 rows after CPMK)
        subcpmk_start_row = cpmk_start_row + len(cpmk_list) + 2
        if subcpmk_list:
            # Insert rows for SubCPMK data
            if len(subcpmk_list) > 0:
                worksheet.insert_rows([[None] * worksheet.col_count] * len(subcpmk_list), subcpmk_start_row)
                time.sleep(1)  # Delay to respect quota
            # Merge cells from C to J for each inserted row
            for i in range(len(subcpmk_list)):
                worksheet.merge_cells(f'C{subcpmk_start_row + i}:J{subcpmk_start_row + i}')
                time.sleep(1)  # Delay to respect quota
            # Write SubCPMK list
            worksheet.update(f'B{subcpmk_start_row}:B{subcpmk_start_row + len(subcpmk_list)}', [[item] for item in subcpmk_list])
            time.sleep(1)  # Delay to respect quota
            # Get SubCPMK descriptions from data_subcpmk
            subcpmk_list_description = []
            for subcpmk_code in subcpmk_list:
                for row in data_subcpmk[1:]:  # Skip header
                    if row[6] == subcpmk_code:  # Match SubCPMK code (assuming column 0 is SubCPMK code)
                        subcpmk_list_description.append(row[7] if len(row) > 1 else "")  # Assuming column 1 is description
                        break
            # Write SubCPMK descriptions to column C
            worksheet.update(f'C{subcpmk_start_row}:C{subcpmk_start_row + len(subcpmk_list)}', [[item] for item in subcpmk_list_description])
            time.sleep(1)  # Delay to respect quota

        # Get matkul data from the sheet with matkul name
        try:
            worksheet_matakuliah = sheet_input.worksheet(matkul_data["MATA KULIAH"])
            all_data = worksheet_matakuliah.get_all_values()
            subcpmk_rps = [row[12] for row in all_data[1:] if row[12]]  # Column M (index 12) for subcpmk_rps
            nilai_subcpmk = [row[21] for row in all_data[1:] if row[21]]  # Column V (index 21) for nilai_subcpmk
            pustaka = [row[0] for row in all_data[1:] if row[0]]
            team_teaching = [row[1] for row in all_data[1:] if row[1]]
            matkul_syarat = [row[2] for row in all_data[1:] if row[2]]

            kelas_get =[row[23] for row in all_data[1:] if row[23]] 
            kelas = kelas_get
            jumlah_get =[row[24] for row in all_data[1:] if row[24]] 
            jumlah_mahasiswa = jumlah_get
            hari_masuk_get =[row[25] for row in all_data[1:] if row[25]] 
            hari_masuk = hari_masuk_get
            lokasi_kelas_get =[row[26] for row in all_data[1:] if row[26]] 
            lokasi_kelas = lokasi_kelas_get
            tahun_ajaran_get =[row[27] for row in all_data[1:] if row[27]] 
            tahun_ajaran = tahun_ajaran_get

            # RPS data pertemuan
            minggu_ke = [row[4] for row in all_data[1:] if row[4]]
            subcpmk_mingguan = [row[5] for row in all_data[1:] if row[5]]
            indikator = [row[6] for row in all_data[1:] if row[6]]
            kriteria = [row[7] for row in all_data[1:] if row[7]]
            materi = [row[8] for row in all_data[1:] if row[8]]
            bobot = [float(row[9]) for row in all_data[1:] if row[9]]
            referensi_mingguan = [row[10] for row in all_data[1:] if row[10]]

        except gspread.exceptions.WorksheetNotFound:
            print(f"Sheet '{matkul_data['MATA KULIAH']}' tidak ditemukan. Menggunakan data default kosong.")
            subcpmk_rps = []
            nilai_subcpmk = []
            pustaka = []
            team_teaching = []
            matkul_syarat = []
            indikator = []
            kriteria = []
            materi = []            
            referensi_mingguan = []

        # Update korelasi CPL SubCPMK
        korelasi_start_row = subcpmk_start_row + len(subcpmk_list) + 2
        if subcpmk_list:
            # Insert rows for SubCPMK data
            if len(subcpmk_list) > 0:
                worksheet.insert_rows([[None] * worksheet.col_count] * len(subcpmk_list), korelasi_start_row)
                time.sleep(1)  # Delay to respect quota
            # Write SubCPMK list in column B
            worksheet.update(f'B{korelasi_start_row + 1}:B{korelasi_start_row + len(subcpmk_list) + 1}', [[item] for item in subcpmk_list])
            time.sleep(1)  # Delay to respect quota
            # Write CPL list horizontally from C to J in the same row
            cpl_to_write = cpl_list[:8]  # Take up to 8 elements for C:J
            if cpl_to_write:
                worksheet.update(f'C{korelasi_start_row}:J{korelasi_start_row}', [cpl_to_write])
                time.sleep(1)  # Delay to respect quota

            # Update korelasi with values
            for row_number in range(len(subcpmk_list)):
                if subcpmk_set and list(subcpmk_set)[row_number] == subcpmk_rps[row_number]:
                    value_korelasi = int(nilai_subcpmk[row_number]) if row_number < len(nilai_subcpmk) else ""
                    cpl_korelasi = list(cpl_set)[row_number] if row_number < len(cpl_set) else ""
                    # Find the row number of cpl_list that matches cpl_korelasi
                    if cpl_korelasi in cpl_list:
                        cpl_index = cpl_list.index(cpl_korelasi)
                        # Calculate the target row and update specific column in C:J
                        target_row = korelasi_start_row + 1 + row_number  # Use row_number for target row
                        if target_row <= korelasi_start_row + len(subcpmk_rps):  # Ensure within inserted rows
                            col_offset = cpl_index % 8  # Offset within C:J (0 = C, 1 = D, ..., 7 = J)
                            if col_offset < 8:  # Ensure within C:J range
                                cell = gspread.utils.rowcol_to_a1(target_row, 3 + col_offset)  # 3 = C
                                worksheet.update(cell, [[value_korelasi]])
                                time.sleep(1)  # Delay to respect quota

            # Add SUM formulas with correct format
            if subcpmk_list:  # Ensure there are rows to sum
                print(f"Debug: target_row={target_row}, korelasi_start_row={korelasi_start_row}")
                sum_row = target_row + 2
                if sum_row > korelasi_start_row + 1:  # Ensure there are rows to sum
                    worksheet.update(f'C{sum_row}', [[f'=SUM(C{korelasi_start_row+1}:C{target_row})']], value_input_option='USER_ENTERED')
                    time.sleep(1)  # Delay to respect quota
                    worksheet.update(f'D{sum_row}', [[f'=SUM(D{korelasi_start_row+1}:D{target_row})']], value_input_option='USER_ENTERED')
                    time.sleep(1)  # Delay to respect quota
                    worksheet.update(f'E{sum_row}', [[f'=SUM(E{korelasi_start_row+1}:E{target_row})']], value_input_option='USER_ENTERED')
                    time.sleep(1)  # Delay to respect quota
                    worksheet.update(f'F{sum_row}', [[f'=SUM(F{korelasi_start_row+1}:F{target_row})']], value_input_option='USER_ENTERED')
                    time.sleep(1)  # Delay to respect quota
                    worksheet.update(f'G{sum_row}', [[f'=SUM(G{korelasi_start_row+1}:G{target_row})']], value_input_option='USER_ENTERED')
                    time.sleep(1)  # Delay to respect quota
                    worksheet.update(f'H{sum_row}', [[f'=SUM(H{korelasi_start_row+1}:H{target_row})']], value_input_option='USER_ENTERED')
                    time.sleep(1)  # Delay to respect quota
                    worksheet.update(f'I{sum_row}', [[f'=SUM(I{korelasi_start_row+1}:I{target_row})']], value_input_option='USER_ENTERED')
                    time.sleep(1)  # Delay to respect quota
                    worksheet.update(f'J{sum_row}', [[f'=SUM(J{korelasi_start_row+1}:J{target_row})']], value_input_option='USER_ENTERED')
                    time.sleep(1)  # Delay to respect quota
            # Add Deskripsi Matkul
            materi_tanpa_uts_uas = []            
            numbered_materi = []
            deskripsi_row = sum_row + 1
            materi_row = deskripsi_row + 1
            if materi:
                # Buat materi tanpa "Evaluasi UTS" dan "Evaluasi UAS" tanpa duplikat
                materi_tanpa_uts_uas = list(dict.fromkeys([item for item in materi if item not in ["Evaluasi UTS", "Evaluasi UAS"]]))

                # Buat numbered_materi_tanpa_uts_uas dengan penomoran
                numbered_materi_tanpa_uts_uas = [f"{i+1}. {item}" for i, item in enumerate(materi_tanpa_uts_uas)]

                # Add list of materi                
                if len(numbered_materi_tanpa_uts_uas) > 0:
                    worksheet.insert_rows([[None] * worksheet.col_count] * len(numbered_materi_tanpa_uts_uas), materi_row+1)
                    time.sleep(1)  # Delay to respect quota

                # Buat numbered_materi dengan penomoran, skip "Evaluasi UTS" dan "Evaluasi UAS"
                for i, item in enumerate(materi):
                    if item in ["Evaluasi UTS", "Evaluasi UAS"]:
                        numbered_materi.append(item)  # Tambahkan tanpa nomor
                    else:
                        numbered_materi.append(f"{i+1 - len([x for x in materi[:i] if x in ['Evaluasi UTS', 'Evaluasi UAS']])}. {item}")

                deskripsi_matkul = "Mata kuliah " + matkul_data["MATA KULIAH"] + " membahas konsep teoritis, metode, dan implementasi mengenai materi seperti " + ", ".join(materi_tanpa_uts_uas)
                worksheet.update(f'B{deskripsi_row}', [[deskripsi_matkul]])
                time.sleep(1)  # Delay to respect quota

                worksheet.update(f'B{materi_row}:B{materi_row + len(numbered_materi_tanpa_uts_uas) - 1}', [[item] for item in numbered_materi_tanpa_uts_uas])
                time.sleep(1)  # Delay to respect quota

            pustaka_row = materi_row + len(numbered_materi_tanpa_uts_uas) + 2
            if pustaka:
                # Add list of materi                
                if len(pustaka) > 0:
                    worksheet.insert_rows([[None] * worksheet.col_count] * len(pustaka), pustaka_row+1)
                    time.sleep(1)  # Delay to respect quota
                numbered_pustaka = [f"{i+1}. {item}" for i, item in enumerate(pustaka)]
                worksheet.update(f'B{pustaka_row}:B{pustaka_row + len(pustaka) - 1}', [[item] for item in numbered_pustaka])
                time.sleep(1)  # Delay to respect quota
            
            team_teaching_row = pustaka_row + len(pustaka) + 2
            if team_teaching:               
                if len(team_teaching) > 0:
                    worksheet.insert_rows([[None] * worksheet.col_count] * len(team_teaching), team_teaching_row+1)
                    time.sleep(1)  # Delay to respect quota
                numbered_team = [f"{i+1}. {item}" for i, item in enumerate(team_teaching)]
                worksheet.update(f'B{team_teaching_row}:B{team_teaching_row + len(team_teaching) - 1}', [[item] for item in numbered_team])
                time.sleep(1)  # Delay to respect quota

            matkul_syarat_row = team_teaching_row + len(team_teaching) + 2
            if matkul_syarat:
                if len(matkul_syarat) > 0:
                    worksheet.insert_rows([[None] * worksheet.col_count] * len(matkul_syarat), matkul_syarat_row+1)
                    time.sleep(1)  # Delay to respect quota
                worksheet.update(f'B{matkul_syarat_row}:B{matkul_syarat_row + len(matkul_syarat) - 1}', [[item] for item in matkul_syarat])
                time.sleep(1)  # Delay to respect quota

            mingguan_row = matkul_syarat_row + len(matkul_syarat) + 7
            if minggu_ke:
                if len(minggu_ke) > 0:
                    worksheet.insert_rows([[None] * worksheet.col_count] * (len(minggu_ke)-2), mingguan_row)
                    time.sleep(1)  # Delay to respect quota
                worksheet.update(f'A{mingguan_row-1}:A{mingguan_row + len(minggu_ke) - 2}', [[item] for item in minggu_ke])
                time.sleep(1)  # Delay to respect quota
            
            if subcpmk_mingguan:
                worksheet.update(f'B{mingguan_row-1}:D{mingguan_row + len(subcpmk_mingguan) - 2}', [[item] for item in subcpmk_mingguan])
                time.sleep(1)  # Delay to respect quota

                # Get SubCPMK descriptions from data_subcpmk                
                for subcpmk_code in subcpmk_mingguan:
                    for row in data_subcpmk[1:]:  # Skip header
                        if row[6] == subcpmk_code: 
                            subcpmk_mingguan_description.append(row[7] if len(row) > 7 else "")  # Menggunakan indeks 7 untuk deskripsi
                            break
                
                uts_row = []
                uas_row = []
                for row_num in range(len(subcpmk_mingguan_description)):
                    if "Evaluasi UTS" in materi[row_num]:
                        subcpmk_mingguan_description[row_num] = "Evaluasi UTS"
                        uts_row.append(row_num)
                    elif "Evaluasi UAS" in materi[row_num]:
                        subcpmk_mingguan_description[row_num] = "Evaluasi UAS"
                        uas_row.append(row_num)

                # Write SubCPMK descriptions to column C
                worksheet.update(f'C{mingguan_row-1}:C{mingguan_row + len(subcpmk_mingguan) - 2}', [[item] for item in subcpmk_mingguan_description])
                time.sleep(1)  # Delay to respect quota

            # Create a modified version of subcpmk_mingguan to handle UTS/UAS
            subcpmk_mingguan_modified = subcpmk_mingguan.copy()
            for i in range(len(materi)):
                if "Evaluasi UTS" in materi[i]:
                    subcpmk_mingguan_modified[i] = "uts"
                elif "Evaluasi UAS" in materi[i]:
                    subcpmk_mingguan_modified[i] = "uas"

            # Add indikator with numbering
            if indikator:
                # Use subcpmk_list as the unique SubCPMK
                unique_subcpmk = subcpmk_list                
                current_subcpmk_index = {}  # To track sub-number for each subcpmk

                for i in range(len(indikator)):
                    subcpmk = subcpmk_mingguan_modified[i]  # Use modified version
                    indikator_item = indikator[i]

                    if subcpmk in unique_subcpmk or subcpmk in ["uts", "uas"]:
                        # Initialize sub-number if not present
                        if subcpmk not in current_subcpmk_index:
                            current_subcpmk_index[subcpmk] = 1
                        else:
                            current_subcpmk_index[subcpmk] += 1

                        # Get the main number and sub_number
                        if subcpmk in unique_subcpmk:
                            main_number = unique_subcpmk.index(subcpmk) + 1
                            sub_number = current_subcpmk_index[subcpmk]
                        elif subcpmk == "uts":
                            main_number = len(unique_subcpmk) + 1
                            sub_number = 1  # .1 for UTS
                        elif subcpmk == "uas":
                            main_number = len(unique_subcpmk) + 1
                            sub_number = 2  # .2 for UAS

                        numbered_indikator.append(f"{main_number}.{sub_number} {indikator_item}")

                # Write numbered indikator to column D (adjust column as needed)
                worksheet.update(f'D{mingguan_row-1}:D{mingguan_row + len(indikator) - 2}', [[item] for item in numbered_indikator])
                time.sleep(1)  # Delay to respect quota
            
            if kriteria:                
                criteria_counter = {}  # To track numbering for Diskusi, Kuis, Tugas
                column_f_content = []  # To store content for column F

                for i in range(len(kriteria)):
                    kriteria_item = kriteria[i].strip()

                    # Initialize counter for each type if not present
                    if kriteria_item.lower().startswith(("diskusi", "kuis", "tugas")):
                        criteria_type = kriteria_item.split()[0].lower()  # Get the first word (e.g., "Diskusi")
                        if criteria_type not in criteria_counter:
                            criteria_counter[criteria_type] = 1
                        else:
                            criteria_counter[criteria_type] += 1

                        # Create numbered kriteria
                        number = criteria_counter[criteria_type]
                        numbered_kriteria_item = f"{kriteria_item.split(':', 1)[0]} {number}: {kriteria_item.split(':', 1)[1].strip()}"
                    else:
                        numbered_kriteria_item = kriteria_item  # No numbering for other types

                    numbered_kriteria.append(numbered_kriteria_item)

                    # Determine content for column F
                    sks_value = matkul_data["SKS"]  # Get SKS value
                    if kriteria_item.lower().startswith("diskusi") or kriteria_item.lower().startswith("kuis"):
                        column_f_item = f"Ekspository dan diskusi [TM : {sks_value}x50']"
                    elif kriteria_item.lower().startswith("tugas"):
                        column_f_item = f"Ekspository dan diskusi [TM : {sks_value}x50'] Task Based Learning [TB : {sks_value}x50']"
                    else:
                        column_f_item = ""  # Empty for other types

                    column_f_content.append(column_f_item)

                # Write numbered kriteria to column E
                worksheet.update(f'E{mingguan_row-1}:E{mingguan_row + len(kriteria) - 2}', [[item] for item in numbered_kriteria])
                time.sleep(1)  # Delay to respect quota

                # Write content to column F
                worksheet.update(f'F{mingguan_row-1}:F{mingguan_row + len(kriteria) - 2}', [[item] for item in column_f_content])
                time.sleep(1)  # Delay to respect quota

            if materi:
                materi_mingguan_numbered = []
                for i in range(len(materi)):
                    worksheet.merge_cells(f'H{mingguan_row-1 + i}:I{mingguan_row-1 + i}')
                    time.sleep(1)

                # Create a dictionary to map items to their assigned numbers
                number_assignments = {}
                number = 1

                # Assign numbers to unique non-UTS/UAS items
                for item in materi:
                    if item not in ["Evaluasi UTS", "Evaluasi UAS"] and item not in number_assignments:
                        number_assignments[item] = number
                        number += 1

                # Generate numbered list using the assigned numbers
                for item in materi:
                    if item in number_assignments:
                        numbered_item = f"{number_assignments[item]}. {item}"
                    else:
                        numbered_item = item  # Keep UTS/UAS as is
                    materi_mingguan_numbered.append(numbered_item)

                # Update the worksheet
                worksheet.update(f'H{mingguan_row-1}:H{mingguan_row + len(materi) - 2}', [[item] for item in materi_mingguan_numbered])
                time.sleep(1)  # Delay to respect quota

            if bobot:
                if len(bobot) > 0:
                    worksheet.update(f'J{mingguan_row-1}:J{mingguan_row + len(bobot) - 2}', [[item/100] for item in bobot])
                    time.sleep(1)  # Delay to respect quota
                
                worksheet.update(f'J{mingguan_row + len(bobot) - 1}', [[f'=SUM(J{mingguan_row-1}:J{mingguan_row + len(bobot) - 2})']], value_input_option='USER_ENTERED')
                time.sleep(1)  # Delay to respect quota
        
        # Combine UTS and UAS rows into a set for unique rows
        grey_rows = set(uts_row + uas_row) if (uts_row is not None and uas_row is not None) else set()

        if grey_rows and worksheet:
            requests = []
            for row in grey_rows:
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": worksheet._properties['sheetId'],
                            "startRowIndex": mingguan_row - 2 + row,
                            "endRowIndex": mingguan_row - 1 + row,
                            "startColumnIndex": 0,  # Column A is index 0
                            "endColumnIndex": 9    # Column J is index 9
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 0.9,  # Light grey (RGB: 230, 230, 230)
                                    "green": 0.9,
                                    "blue": 0.9
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })

            # Perform batch update
            if requests:
                body = {"requests": requests}
                worksheet.spreadsheet.batch_update(body)
                time.sleep(1)  # Delay to respect quota after batch update
                    
    # def update_ktrak_sheet(worksheet, data, dosen):
    def update_kontrak_sheet(data, cpl_list, subcpmk_list, dosen):
        # Duplicate Template RPS and rename to RPS
        new_kontrak_worksheet = sheet_input.duplicate_sheet(temp_kontrak_in.id, new_sheet_name="KONTRAK")
        worksheet = sheet_input.worksheet("KONTRAK")

        worksheet.update('I6', [[data['KODE DOKUMEN KONTRAK'] or '']])
        time.sleep(1)

        # Update basic course data 
        worksheet.update('A6', [[data['MATA KULIAH']]])
        time.sleep(1)  # Delay to respect quota      
        worksheet.update('D6', [[data['KODE MK']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('E6', [[data['KATEGORI']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('F6', [[data['SKS']]])
        time.sleep(1)  # Delay to respect quota
        worksheet.update('H6', [[data['SEMESTER']]])
        time.sleep(1)  # Delay to respect quota

        worksheet.update('A8', [[dosen]])  
        time.sleep(1)  # Delay to respect quota
        worksheet.update('E8', [kelas])  
        time.sleep(1)  # Delay to respect quota
        worksheet.update('F8', [jumlah_mahasiswa])  
        time.sleep(1)  # Delay to respect quota
        worksheet.update('G8', [hari_masuk])  
        time.sleep(1)  # Delay to respect quota
        worksheet.update('I8', [lokasi_kelas])  
        time.sleep(1)  # Delay to respect quota

        manfaat_start_row = 10
        if cpl_list:
            # Insert rows for CPL data
            if len(cpl_list) > 0:
                worksheet.insert_rows([[None] * worksheet.col_count] * (len(cpl_list)-2), manfaat_start_row+1)  # Insert rows at the specified row
                time.sleep(1)  # Delay to respect quota
            # Merge cells from C to J for each inserted row
            for i in range(len(cpl_list)):
                worksheet.merge_cells(f'E{manfaat_start_row + i}:I{manfaat_start_row + i}')
                time.sleep(1)  # Delay to respect quota
            # Get CPL descriptions from data_cpl
            cpl_list_description = []
            for cpl_code in cpl_list:
                for row in data_cpl[1:]:  # Skip header
                    if row[0] == cpl_code:  # Match CPL code (assuming column 0 is CPL code)
                        cpl_list_description.append(row[1] if len(row) > 1 else "")  # Assuming column 1 is description
                        break
            # Write CPL descriptions to column C
            worksheet.update(f'E{manfaat_start_row }:E{manfaat_start_row  + len(cpl_list)-1}', [[item] for item in cpl_list_description])
            time.sleep(1)  # Delay to respect quota
            
            deskripsi_kontrak_row = manfaat_start_row +len(cpl_list)
            worksheet.update(f'E{deskripsi_kontrak_row}', [[deskripsi_matkul]])
            time.sleep(1)  # Delay to respect quota

            tujuan_kontrak_row = deskripsi_kontrak_row +1
            # Get SubCPMK descriptions from data_subcpmk
            subcpmk_list_description = []
            for subcpmk_code in subcpmk_list:
                for row in data_subcpmk[1:]:  # Skip header
                    if row[6] == subcpmk_code:  # Match SubCPMK code (assuming column 0 is SubCPMK code)
                        subcpmk_list_description.append(row[7] if len(row) > 1 else "")  # Assuming column 1 is description
                        break
            
            tujuan_kontrak_description = "; ".join([item for item in subcpmk_list_description])
            # Write SubCPMK descriptions to column C
            worksheet.update(f'E{tujuan_kontrak_row}', [[tujuan_kontrak_description]])
            time.sleep(1)  # Delay to respect quota

            materi_kontrak_row = tujuan_kontrak_row + 1
            materi_one_row = "\n".join([item for item in numbered_materi_tanpa_uts_uas])
            # Then update a single cell, e.g.:
            worksheet.update(f'E{materi_kontrak_row}', [[materi_one_row]])
            time.sleep(1)  # Delay to respect quota

            referensi_kontrak_row = materi_kontrak_row + 2
            referensi_one_row = "\n".join([item for item in numbered_pustaka])
            # Then update a single cell, e.g.:
            worksheet.update(f'E{referensi_kontrak_row}', [[referensi_one_row]])
            time.sleep(1)  # Delay to respect quota

            #Tanggal
            date_row = referensi_kontrak_row + 8
            worksheet.update(f'A{date_row}', [[datetime.datetime.now().strftime('%d %B %Y')]])

            kordinator_row = date_row+5
            
            worksheet.update(f'G{kordinator_row}', [[data_dosen[4][0]]])
            worksheet.update(f'G{kordinator_row+1}', [["NIK. " + data_dosen[4][1]]])

    def update_rubrik_sheet(data, kriteria, subcpmk_set, cpl_set, dosen):
        # ――― 1. Siapkan lookup & konstanta ―――
        ALLOWED = set(TEMPLATE_MAP)  # {'H1', …, 'SP1'}
        START_SUB, START_CPL = 15, 20
        sheet_counter = 1
        sub_to_cpl = dict(zip(subcpmk_set, cpl_set))  # SUB‑→ CPL

        rows_by_code = {code: [] for code in ALLOWED}
        for i, text in enumerate(kriteria):
            m = re.search(r"\[(.*?)\]", text)
            if m and (lbl := m.group(1).strip()) in ALLOWED:
                rows_by_code[lbl].append(i)

        print(rows_by_code)

        for code in ALLOWED:
            idx_rows = rows_by_code[code]
            if not idx_rows:
                continue

            # Ambil sub-CPMK dan CPL unik
            sub_lines = [subcpmk_mingguan[i] for i in idx_rows]
            sub_lines_nondup = list(dict.fromkeys(sub_lines))  # Keep order
            cpl_lines = [sub_to_cpl.get(sub, "") for sub in sub_lines]
            cpl_lines_nondup = list(dict.fromkeys(cpl_lines))

            # Ambil deskripsi Sub-CPMK
            sub_desc = []
            for sub_code in sub_lines_nondup:
                for row in data_subcpmk[1:]:  # Skip header
                    if row[6] == sub_code:
                        sub_desc.append(row[7] if len(row) > 7 else "")
                        break

            # Ambil deskripsi CPL
            cpl_desc = []
            for cpl_code in cpl_lines_nondup:
                for row in data_cpl[1:]:  # Skip header
                    if row[0] == cpl_code:
                        cpl_desc.append(row[1] if len(row) > 1 else "")
                        break

            # Skip jika kosong
            if not sub_lines_nondup and not cpl_lines_nondup:
                print(f"[SKIP] RUB {sheet_counter} ({code}) tidak memiliki konten, dilewati.")
                continue

            # Buat sheet rubrik
            new_name = f"RUB {sheet_counter} ({code})"
            template_ws = TEMPLATE_MAP[code]
            penugasan = PENUGASAN_MAP[code]
            sheet_input.duplicate_sheet(template_ws.id, new_sheet_name=new_name)
            ws = sheet_input.worksheet(new_name)
            time.sleep(1)

            # Isi header
            ws.update('K3', [[data['KODE DOKUMEN RUBRIK']]]); time.sleep(1)
            ws.update('C5', [[data['MATA KULIAH']]]);         time.sleep(1)
            ws.update('C6', [[data['KODE MK']]]);             time.sleep(1)
            ws.update('C7', [[data['SKS']]]);                 time.sleep(1)
            ws.update('C8', [[data['SEMESTER']]]);            time.sleep(1)
            ws.update('C9', [[dosen]]);                       time.sleep(1)
            ws.update('C12', [[penugasan]]);                  time.sleep(1)

            # Insert rows untuk Sub-CPMK
            if len(sub_lines_nondup) > 1:
                ws.insert_rows([[None] * ws.col_count] * (len(sub_lines_nondup) - 1), START_SUB + 1)
                time.sleep(1)

            for i in range(len(sub_lines_nondup)):
                ws.merge_cells(f'C{START_SUB+i}:D{START_SUB+i}')
                time.sleep(1)
                ws.merge_cells(f'E{START_SUB+i}:L{START_SUB+i}')
                time.sleep(1)

            ws.update(f'C{START_SUB}:C{START_SUB + len(sub_lines_nondup) - 1}',
                    [[code] for code in sub_lines_nondup])
            time.sleep(1)
            ws.update(f'E{START_SUB}:E{START_SUB + len(sub_desc) - 1}',
                    [[desc] for desc in sub_desc])
            time.sleep(1)

            # Insert rows untuk CPL
            cpl_start = START_SUB + len(sub_lines_nondup) + 4
            if len(cpl_lines_nondup) > 1:
                ws.insert_rows([[None] * ws.col_count] * (len(cpl_lines_nondup) - 1), cpl_start)
                time.sleep(1)

            for i in range(len(cpl_lines_nondup)):
                ws.merge_cells(f'A{cpl_start+i}:C{cpl_start+i}')
                time.sleep(1)
                ws.merge_cells(f'D{cpl_start+i}:J{cpl_start+i}')
                time.sleep(1)
                ws.merge_cells(f'K{cpl_start+i}:L{cpl_start+i}')
                time.sleep(1)

            # Buat mapping SubCPMK → CPL (kebalikan)
            sub_set_label = set(sub_lines_nondup)          # Sub‑CPMK unik utk label ini
            cpl_to_subs   = {cpl: [] for cpl in cpl_lines_nondup}

            for sub in sub_set_label:                      # loop hanya sub yg dipakai
                cpl = sub_to_cpl.get(sub)
                if cpl in cpl_to_subs:                     # pastikan CPL dipakai juga
                    cpl_to_subs[cpl].append(sub)

            # Update ke sheet
            ws.update(f'A{cpl_start}:A{cpl_start + len(cpl_lines_nondup) - 1}',
                    [[code] for code in cpl_lines_nondup])
            time.sleep(1)
            ws.update(f'D{cpl_start}:D{cpl_start + len(cpl_desc) - 1}',
                    [[desc] for desc in cpl_desc])
            time.sleep(1)
            ws.update(f'K{cpl_start}:K{cpl_start + len(cpl_lines_nondup) - 1}',
                    [[", ".join(cpl_to_subs[cpl])] for cpl in cpl_lines_nondup])
            time.sleep(1)
            sheet_counter += 1
    
    def update_rpm_sheets(data, kriteria, dosen):
        kategori_keys = ["Tugas", "Kuis", "Evaluasi UTS", "Evaluasi UAS"]
        kategori_map  = {key: {"label": [], "row": []} for key in kategori_keys}

        for idx, item in enumerate(kriteria):
            for key in kategori_keys:
                if key in item:
                    kategori_map[key]["label"].append(item)
                    kategori_map[key]["row"].append(idx)
                    break   # sebuah kriteria hanya boleh masuk satu kategori

        # ------------------ 2. Duplicated & isi sheet ----------------------
        sheet_counter = 1  # penomoran global RPM 1, 2, 3, ...

        for key in kategori_keys:               # urutan Tugas → Kuis → UTS → UAS
            labels = kategori_map[key]["label"]
            rows   = kategori_map[key]["row"]

            for i, item in enumerate(labels):
                # a) duplikat template RPM
                safe_item = item.split(":", 1)[0].strip()
                new_name = f"RPM {sheet_counter} ({safe_item})"
                new_ws = sheet_input.duplicate_sheet(
                    temp_rtm_in.id,          # atau temp_rtm_in.id jika sama
                    new_sheet_name=new_name
                )
                worksheet = sheet_input.worksheet(new_name)
                time.sleep(1)

                row_idx = rows[i]             # indeks sejajar ke list global

                # b) ISIAN HEADER
                worksheet.update('H3', [[data.get('KODE DOKUMEN RTM', '')]])
                time.sleep(1)

                worksheet.update('C5', [[data['MATA KULIAH']]]);  time.sleep(1)
                worksheet.update('C6', [[data['KODE MK']]]);      time.sleep(1)
                worksheet.update('E6', [[data['SKS']]]);          time.sleep(1)
                worksheet.update('G6', [[data['SEMESTER']]]);     time.sleep(1)
                worksheet.update('C7', [[dosen]]);                time.sleep(1)

                # c) ISIAN RUBRIK SPESIFIK
                worksheet.update('A12', [[item]]);                                time.sleep(1)
                worksheet.update('A14', [[subcpmk_mingguan_description[row_idx]]]);time.sleep(1)
                worksheet.update('A17', [[numbered_indikator[row_idx]]]);          time.sleep(1)
                worksheet.update('A25', [[f"Indikator: {numbered_indikator[row_idx]}"]]); time.sleep(1)
                worksheet.update('A27', [[f"Bobot Penilaian : {bobot[row_idx]} % dari total 100% penilaian mata kuliah"]]); time.sleep(1)
                worksheet.update('A30', [[f"Minggu ke-{minggu_ke[row_idx]}"]]);    time.sleep(1)

                # d) ISIAN PUSTAKA (jika ada)
                if numbered_pustaka:
                    worksheet.insert_rows([[None]*worksheet.col_count]*len(numbered_pustaka), 35)
                    time.sleep(1)
                    for p in range(len(numbered_pustaka)+1):
                        worksheet.merge_cells(f'A{34+p}:H{34+p}')
                    worksheet.update(f'A34:A{34+len(numbered_pustaka)-1}',
                                    [[ref] for ref in numbered_pustaka])
                    time.sleep(1)

                sheet_counter += 1  # naikkan nomor RPM berikutnya

    def update_rtm_sheet(data, kriteria, dosen):
        kriteria_tugas = [item for item in kriteria if "Tugas" in item]
        row_tugas = [i for i, item in enumerate(kriteria) if "Tugas" in item]
        print (kriteria_tugas)
        print(row_tugas)
        for i in range(len(kriteria_tugas)) :
            new_rubrik_worksheet = sheet_input.duplicate_sheet(temp_rtm_in.id, new_sheet_name=f"RPM {i+1}")
            worksheet = sheet_input.worksheet(f"RTM {i+1}")
            time.sleep(1)

            worksheet.update('H3', [[data['KODE DOKUMEN RTM'] or '']])
            time.sleep(1)

            # Update basic course data 
            worksheet.update('C5', [[data['MATA KULIAH']]])
            time.sleep(1)  # Delay to respect quota      
            worksheet.update('C6', [[data['KODE MK']]])
            time.sleep(1)  # Delay to respect quota
            worksheet.update('E6', [[data['SKS']]])
            time.sleep(1)  # Delay to respect quota
            worksheet.update('G6', [[data['SEMESTER']]])
            time.sleep(1)  # Delay to respect quota

            worksheet.update('C7', [[dosen]])  
            time.sleep(1)  # Delay to respect quota

            worksheet.update('A12', [[kriteria_tugas[i]]])  
            time.sleep(1)  # Delay to respect quota

            worksheet.update('A14', [[subcpmk_mingguan_description[row_tugas[i]]]])  
            time.sleep(1)  # Delay to respect quota

            worksheet.update('A17', [[numbered_indikator[row_tugas[i]]]])  
            time.sleep(1)  # Delay to respect quota

            worksheet.update('A25', [["Indikator: " + numbered_indikator[row_tugas[i]]]])  
            time.sleep(1)  # Delay to respect quota

            worksheet.update('A27', [[f"Bobot Penilaian : {bobot[row_tugas[i]]} % dari total 100% penilaian mata kuliah"]])  
            time.sleep(1)  # Delay to respect quota

            worksheet.update('A30', [["Minggu ke-" + minggu_ke[row_tugas[i]]]])  
            time.sleep(1)  # Delay to respect quota
                            
            if len(numbered_pustaka) > 0:
                worksheet.insert_rows([[None] * worksheet.col_count] * len(numbered_pustaka), 35)
                time.sleep(1)
            for pustaka_row in range(len(numbered_pustaka)+1):
                worksheet.merge_cells(f'A{34 + pustaka_row}:H{34 + pustaka_row}')
            worksheet.update(f'A{34}:A{34 + len(numbered_pustaka) - 1}', [[item] for item in numbered_pustaka])
            time.sleep(1)  # Delay to respect quota

    def update_porto_sheet(cpl_list, cpl_set, cpmk_set, subcpmk_set):

        global cpl_sorted, cpmk_sorted, subcpmk_sorted, bentuk_sorted, bobot_sorted
        # Duplicate Template RPS and rename to RPS
        new_porto_worksheet = sheet_input.duplicate_sheet(temp_porto_in.id, new_sheet_name="PORTOFOLIO PENILAIAN")
        worksheet = sheet_input.worksheet("PORTOFOLIO PENILAIAN")
        time.sleep(1)

        if bobot:
            # Insert empty rows if needed
            worksheet.insert_rows([[None] * worksheet.col_count] * (len(bobot) - 2), 3)
            time.sleep(1)

        # Prepare data
        bentuk_penilaian = [item.split(":")[0].strip() for item in numbered_kriteria]
        indikator = [item.split(" ")[0].strip() for item in numbered_indikator]

        # Match CPMK and CPL for each SubCPMK
        cpmk = []
        cpl = []
        for item in subcpmk_mingguan:
            for idx, sub in enumerate(subcpmk_set):
                if sub == item:
                    cpmk.append(cpmk_set[idx])
                    cpl.append(cpl_set[idx])
                    break

        # Combine all columns into rows
        combined_rows = list(zip(
            cpl,
            cpmk,
            subcpmk_mingguan,
            indikator,
            bentuk_penilaian,
            [b / 100 for b in bobot]
        ))

        # Sort rows by Column A (CPL)
        combined_rows.sort(key=lambda x: x[2])  # Sort by CPL (column A)

        # Unpack sorted data
        cpl_sorted, cpmk_sorted, subcpmk_sorted, indikator_sorted, bentuk_sorted, bobot_sorted = zip(*combined_rows)

        # Write sorted data to worksheet
        worksheet.update(f'A2:A{len(bobot) + 1}', [[item] for item in cpl_sorted])
        time.sleep(1)

        worksheet.update(f'B2:B{len(bobot) + 1}', [[item] for item in cpmk_sorted])
        time.sleep(1)

        worksheet.update(f'C2:C{len(bobot) + 1}', [[item] for item in subcpmk_sorted])
        time.sleep(1)

        worksheet.update(f'D2:D{len(bobot) + 1}', [[item] for item in indikator_sorted])
        time.sleep(1)

        worksheet.update(f'E2:E{len(bobot) + 1}', [[item] for item in bentuk_sorted])
        time.sleep(1)

        worksheet.update(f'F2:F{len(bobot) + 1}', [[item] for item in bobot_sorted])
        time.sleep(1)

        for i in range(len(bobot)):
            worksheet.update(f'H{i + 2}', [[80]])
            time.sleep(1)
            worksheet.update(f'I{i + 2}', [[f"=H{i + 2}*F{i + 2}"]], value_input_option='USER_ENTERED')
            time.sleep(1)

        for cpl_name in cpl_list:
            matching_rows = [i for i, val in enumerate(cpl_sorted) if val == cpl_name]
            if matching_rows:
                row_start = matching_rows[0] + 2  # +2 to convert index to sheet row (starts from row 2)
                row_end = matching_rows[-1] + 2
                sum_formula = f'=SUM(F{row_start}:F{row_end})'
                percent_formula = f'=SUM(I{row_start}:I{row_end})/I{len(bobot)+2}'
                worksheet.update(f'G{row_start}', [[sum_formula]], value_input_option='USER_ENTERED')
                time.sleep(1)
                worksheet.update(f'J{row_start}', [[percent_formula]], value_input_option='USER_ENTERED')
                time.sleep(1)

    def col_index_to_letter(n):
        """Convert column index (1-based) to letter (e.g., 4 -> D)."""
        result = ''
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def update_nilai_sheet(data, dosen, cpl_list, cpmk_list, subcpmk_list, cpl_nilai, cpmk_nilai, subcpmk_nilai, bentuk_nilai, bobot_nilai):
        # Duplicate Template RPS and rename to RPS
        sheet_name = f"NILAI {(data['MATA KULIAH']).upper()} - TA: {tahun_ajaran[0]}"
        new_porto_worksheet = sheet_input.duplicate_sheet(temp_nilai_in.id, new_sheet_name= sheet_name)
        worksheet = sheet_input.worksheet(sheet_name)
        time.sleep(1)

        worksheet.update(f'C3', [[f": {data['MATA KULIAH']}"]])
        time.sleep(1)

        worksheet.update(f'C4', [[f": {data['KODE MK']}"]])
        time.sleep(1)

        worksheet.update(f'C5', [[f": {kelas[0]}"]])
        time.sleep(1)

        worksheet.update(f'C6', [[f": {tahun_ajaran[0]}"]])
        time.sleep(1)

        worksheet.update(f'C7', [[f": {dosen}"]])
        time.sleep(1)

        start_col = 4  # Kolom D = 4 (1-based indexing)
        row = 13       # Header row to write labels

        cpl_counts = Counter(cpl_nilai)  # Hitung jumlah CPL
        cpmk_counts = Counter(cpmk_nilai)
        subcpmk_counts = Counter(subcpmk_nilai)
        merge_cpl_start_col = start_col
        merge_cpmk_start_col = start_col
        merge_subcpmk_start_col = start_col

        # Hitung total kolom yang dibutuhkan
        total_cols_to_insert = len(cpl_nilai) * 4

        # Sisipkan kolom kosong dari kolom D (1-based)
        worksheet.insert_cols([[None]] * total_cols_to_insert, start_col)
        time.sleep(1)

        # Tulis header setelah menyisipkan
        current_col = start_col
        for idx, cpl_value in enumerate(cpl_nilai):
            headers = [
                f"NILAI",
                f"Tambahan",
                f"NILAI AKHIR",
                f"SUB BOBOT",
            ]
            for j, header in enumerate(headers):
                col_number = current_col + j
                col_letter = col_index_to_letter(col_number)
                worksheet.update(f"{col_letter}{row}", [[header]])
                time.sleep(1)
                if col_number < current_col + 3:
                    worksheet.merge_cells(f"{col_letter}{row}:{col_letter}{row+1}")
                    time.sleep(1)
            col_letter_cpmk = col_index_to_letter(current_col)
            end_col_bentuk = col_index_to_letter(current_col+3)
            col_letter_bobot = col_index_to_letter(current_col+3)
            worksheet.update(f"{col_letter_cpmk}{row-4}", [[cpl_value]])
            time.sleep(1)
            worksheet.update(f"{col_letter_cpmk}{row-3}", [[cpmk_nilai[idx]]])
            time.sleep(1)
            worksheet.update(f"{col_letter_cpmk}{row-2}", [[subcpmk_nilai[idx]]])
            time.sleep(1)
            worksheet.update(f"{col_letter_cpmk}{row-1}", [[bentuk_nilai[idx]]])
            time.sleep(1)
            worksheet.update(f"{col_letter_bobot}{row+1}", [[bobot_nilai[idx]]])
            time.sleep(1)

            worksheet.merge_cells(f"{col_letter_cpmk}{row-1}:{end_col_bentuk}{row-1}")
            time.sleep(1)

            current_col += 4  # Lanjut ke blok berikutnya
        
        for cpl_value in cpl_list:  # Urutkan berdasarkan kemunculan
            count = cpl_counts[cpl_value]
            merge_end_col = merge_cpl_start_col + (count * 4) - 1  # 4 kolom per CPL

            start_letter = col_index_to_letter(merge_cpl_start_col)
            end_letter = col_index_to_letter(merge_end_col)

            # Baris untuk CPL (row-4)
            merge_range = f"{start_letter}{row-4}:{end_letter}{row-4}"
            worksheet.merge_cells(merge_range)
            time.sleep(1)

            merge_cpl_start_col = merge_end_col + 1  # Geser untuk CPL berikutnya

        for cpmk_value in cpmk_list:  # Urutkan berdasarkan kemunculan
            count = cpmk_counts[cpmk_value]
            merge_end_col = merge_cpmk_start_col + (count * 4) - 1  # 4 kolom per CPL

            start_letter = col_index_to_letter(merge_cpmk_start_col)
            end_letter = col_index_to_letter(merge_end_col)

            # Baris untuk CPL (row-4)
            merge_range = f"{start_letter}{row-3}:{end_letter}{row-3}"
            worksheet.merge_cells(merge_range)
            time.sleep(1)

            merge_cpmk_start_col = merge_end_col + 1  # Geser untuk CPL berikutnya
        
        for subcpmk_value in subcpmk_list:  # Urutkan berdasarkan kemunculan
            count = subcpmk_counts[subcpmk_value]
            merge_end_col = merge_subcpmk_start_col + (count * 4) - 1  # 4 kolom per CPL

            start_letter = col_index_to_letter(merge_subcpmk_start_col)
            end_letter = col_index_to_letter(merge_end_col)

            # Baris untuk CPL (row-4)
            merge_range = f"{start_letter}{row-2}:{end_letter}{row-2}"
            worksheet.merge_cells(merge_range)
            time.sleep(1)

            merge_subcpmk_start_col = merge_end_col + 1  # Geser untuk CPL berikutnya


        formula_row = row + 2  # baris ke-15
        bobot_row = row + 1    # baris ke-14
        current_col = start_col
        for _ in range(len(cpl_nilai)):
            col_1 = col_index_to_letter(current_col)
            col_2 = col_index_to_letter(current_col + 1)
            col_3 = col_index_to_letter(current_col + 2)
            col_4 = col_index_to_letter(current_col + 3)

            # Kolom 3 = SUM(kolom1:kolom2)
            sum_formula = f"=SUM({col_1}{formula_row}:{col_2}{formula_row})"
            worksheet.update(f"{col_3}{formula_row}", [[sum_formula]], value_input_option='USER_ENTERED')
            time.sleep(1)

            # Kolom 4 = kolom3 * kolom4_di_baris14
            mult_formula = f"={col_3}{formula_row}*${col_4}${bobot_row}"
            worksheet.update(f"{col_4}{formula_row}", [[mult_formula]], value_input_option='USER_ENTERED')
            time.sleep(1)

            current_col += 4
        
        #Tolong insert 1 column untuk setiap cpl yang berbeda
        nilai_per_cpl_letter = []
        add_start = start_col
        for cpl_value in cpl_list:  # Urutkan berdasarkan kemunculan
            count = cpl_counts[cpl_value]
            last_bobot_col_idx = add_start + (count * 4) - 1
            insert_col_idx = last_bobot_col_idx + 1

            # Sisipkan satu kolom kosong setelah blok CPL
            worksheet.insert_cols([[None]], insert_col_idx)
            time.sleep(1)

            # Hitung kolom-kolom SUB BOBOT untuk CPL ini (kolom ke-4 dari setiap 4 kolom)
            sub_bobot_cols = [
                col_index_to_letter(add_start + i * 4 + 3)
                for i in range(count)
            ]

            formula_row = row + 2   # baris 15
            start_merge_row = row - 3  # baris 10
            end_merge_row = row + 1     # baris 13

            insert_letter = col_index_to_letter(insert_col_idx)

            # 1. Header: merge row 10:14, tulis 'NILAI PER CPL'
            merge_range = f"{insert_letter}{start_merge_row}:{insert_letter}{end_merge_row}"
            worksheet.merge_cells(merge_range)
            time.sleep(1)
            worksheet.update(f"{insert_letter}{start_merge_row}", [["NILAI PER CPL"]])
            time.sleep(1)

            # 2. Formula: jumlahkan semua SUB BOBOT CPL ini di baris 15
            sum_formula = f"=SUM({','.join([col + str(formula_row) for col in sub_bobot_cols])})"
            worksheet.update(f"{insert_letter}{formula_row}", [[sum_formula]], value_input_option='USER_ENTERED')
            time.sleep(1)

            nilai_per_cpl_letter.append(insert_letter)

            # Geser pointer ke CPL berikutnya (melewati kolom yang baru diinsert)
            add_start = insert_col_idx + 1
        
        # === Jumlahkan seluruh "NILAI PER CPL" ===
        if nilai_per_cpl_letter:
            total_col_letter = nilai_per_cpl_letter[-1]
            total_col_index = sum([(ord(char) - 64) * (26 ** i) for i, char in enumerate(reversed(total_col_letter))])
            result_col_letter = col_index_to_letter(total_col_index + 1)  # Kolom setelah terakhir

            # Bangun formula penjumlahan
            final_sum_formula = f"=SUM({','.join([col + str(row + 2) for col in nilai_per_cpl_letter])})"
            worksheet.update(f"{result_col_letter}{row + 2}", [[final_sum_formula]], value_input_option='USER_ENTERED')
            time.sleep(1)        

            ketercapaian_col_start = total_col_index + 5

            # Konversi kolom angka ke huruf
            col_ket_letter = col_index_to_letter(ketercapaian_col_start)
            col_thresh_letter = col_index_to_letter(ketercapaian_col_start + 1)
            col_avg_letter = col_index_to_letter(ketercapaian_col_start + 2)

            # Update kolom KETERANGAN CPL
            worksheet.update(f'{col_ket_letter}15:{col_ket_letter}{15 + len(cpl_list) - 1}', [[item] for item in cpl_list])
            time.sleep(1)

            threshold_cpl = []

            for cpl_name in cpl_list:
                # Temukan indeks baris yang cocok dengan CPL saat ini
                matching_rows = [i for i, val in enumerate(cpl_nilai) if val == cpl_name]
                
                if matching_rows:
                    # Jumlahkan nilai bobot untuk baris-baris tersebut
                    threshold = sum([bobot_nilai[i] for i in matching_rows])
                    threshold_cpl.append(threshold)
                else:
                    threshold_cpl.append(0)  # Jika tidak ditemukan, bisa diisi 0 atau None

            # Update kolom THRESHOLD CPL
            worksheet.update(
                f'{col_thresh_letter}15:{col_thresh_letter}{15 + len(threshold_cpl) - 1}',
                [[item] for item in threshold_cpl]
            )
            time.sleep(1)

            # Update kolom AVERAGE
            for i, letter in enumerate(nilai_per_cpl_letter):
                average_formula = f"=AVERAGE({letter}15:{letter}17)"
                worksheet.update(f"{col_avg_letter}{15 + i}", [[average_formula]], value_input_option='USER_ENTERED')
                time.sleep(1)


    # Call the update function
    update_rps_sheet(matkul_data, cpl_list, cpmk_list, subcpmk_list, selected_dosen)
    update_kontrak_sheet(matkul_data, cpl_list, subcpmk_list, selected_dosen)
    update_rubrik_sheet(matkul_data, numbered_kriteria, subcpmk_set, cpl_set, selected_dosen)
    # update_rtm_sheet(matkul_data, numbered_kriteria, selected_dosen)
    update_rpm_sheets(matkul_data, numbered_kriteria, selected_dosen)
    update_porto_sheet(cpl_list, cpl_set, cpmk_set, subcpmk_set)
    update_nilai_sheet(matkul_data, selected_dosen, cpl_list, cpmk_list, subcpmk_list, cpl_sorted, cpmk_sorted, subcpmk_sorted, bentuk_sorted, bobot_sorted)

    print(f"RPS telah dibuat dan diperbarui di sheet 'RPS' dalam spreadsheet: {sheet_input_url}")