import pandas as pd
import re
from dateutil.parser import parse

# Nama file Excel
# file_path = "penjualan_dqmart_01-beta.xlsx"
# file_out_path = "penjualan_dqmart_clean.xlsx"

# Dictionary untuk mengganti nama bulan Indonesia/singkatan ke Inggris standar
bulan_map : dict[str, str] = {
    'Januari': 'Jan', 'Februari': 'Feb', 'Maret': 'Mar', 'Mei': 'May',
    'Juni': 'Jun', 'Juli': 'Jul', 'Agustus': 'Aug', 'September': 'Sep',
    'Oktober': 'Oct', "Okt" : "Oct", 'November': 'Nov', 'Desember': 'Dec'
}

# Fungsi Pembersihan Kustom
def clean_and_parse_date(date_str):
    if pd.isna(date_str) or date_str == "":
        return date_str  # Pertahankan nilai hilang jika sudah ada

    # 1. Normalisasi string: hapus spasi berlebih, ubah ke title case
    s = str(date_str).strip().title()
    
    # 2. Ganti nama bulan/singkatan Indonesia dengan standar Inggris
    for indo, eng in bulan_map.items():
        s = s.replace(indo, eng)

    # 3. Perbaiki singkatan tahun (misal: '24 menjadi 2024)
    s = re.sub(r"'(\d{2})$", r"20\1", s)
    
    # 4. Coba mengurai dengan dateutil.parser (sangat fleksibel)
    try:
        # Coba parsing
        parsed_date = parse(s, dayfirst=True)
        # Kembalikan dalam format string dd-MM-yyyy
        return parsed_date.strftime('%d-%m-%Y')
        
    except Exception:
        # Jika parsing gagal, kembalikan string aslinya
        return date_str

# Fungsi Utama
def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    df = pd.read_excel(input_xlsx_path)
    #
    df["Tanggal Transaksi"] = df["Tanggal Transaksi"].apply(clean_and_parse_date)
    #
    df.to_excel( output_xlsx_path, index=False )    

