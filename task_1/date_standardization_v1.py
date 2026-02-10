import pandas as pd
# Library 're' adalah bawaan (built-in) Python, jadi diizinkan
import re 

# Dictionary untuk mengganti nama bulan dan format bahasa asing ke Inggris standar
bulan_map: dict[str, str] = {
    'Januari': 'Jan', 'Februari': 'Feb', 'Maret': 'Mar', 'Mei': 'May',
    'Juni': 'Jun', 'Juli': 'Jul', 'Agustus': 'Aug', 'September': 'Sep',
    'Oktober': 'Oct', "Okt": "Oct", 'November': 'Nov', 'Desember': 'Dec',
    
    # Bulan Asing Umum
    'Februar': 'Feb', 'Marzo': 'Mar', 'Aprile': 'Apr', 'Giugno': 'Jun',
    'Luglio': 'Jul', 'Agosto': 'Aug', 'Settembre': 'Sep', 'Ottobre': 'Oct',
    'Novembre': 'Nov', 'Dicembre': 'Dec', 'Janvier': 'Jan', 'Fevrier': 'Feb',
    'Mars': 'Mar', 'Avril': 'Apr', 'Juin': 'Jun', 'Juillet': 'Jul', 
    'Aout': 'Aug', 'Septembre': 'Sep', 'Octobre': 'Oct', 'Novembre': 'Nov',
    'Decembre': 'Dec', 'Enero': 'Jan', 'Febrero': 'Feb', 'Marzo': 'Mar',
    'Abril': 'Apr', 'Mayo': 'May', 'Junio': 'Jun', 'Julio': 'Jul', 
    'Agosto': 'Aug', 'Setiembre': 'Sep', 'Noviembre': 'Nov', 'Diciembre': 'Dec',
    'März': 'Mar', 'Mai': 'May', 'Juni': 'Jun', 'Juli': 'Jul', 
    'August': 'Aug', 'September': 'Sep', 'Oktober': 'Oct', 'November': 'Nov',
    'Dezember': 'Dec', 'Februrari': 'Feb', 'Februarie': 'Feb',
    'Feber': 'Feb', 'Februarii': 'Feb', 'Feb': 'Feb', 
}

# Himpunan (Set) singkatan bulan Inggris standar (digunakan untuk deteksi Month-First)
ENGLISH_MONTH_ABBREVIATIONS = set(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])

# Fungsi Pembersihan Kustom (Final, bebas warning)
def clean_and_parse_date(date_str):
    if pd.isna(date_str) or date_str == "":
        return ""

    s = str(date_str).strip()
    
    # 0. Pembersihan Agresif 
    s = re.sub(r'\(.*?\)', '', s) 
    s = re.sub(r'\b(de|del|a|la|of|marca)\b', ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'(\d)(st|nd|rd|th)', r'\1', s, flags=re.IGNORECASE) 
    s = re.sub(r'(\d)er\b', r'\1', s, flags=re.IGNORECASE) 
    s = re.sub(r'[TZ]', ' ', s) 
    s = re.sub(r'[\u4E00-\u9FFF]', ' ', s) 
    
    # 1. Ganti nama bulan/singkatan ke Inggris
    s_lower = s.lower()
    temp_s = s_lower
    for indo, eng in bulan_map.items():
        temp_s = re.sub(r'\b' + re.escape(indo.lower()) + r'\b', eng, temp_s, flags=re.IGNORECASE)
        
    s = temp_s.strip()
        
    # 2. Normalisasi Pemisah dan Spasi
    s = re.sub(r'[\\/\-.,]', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    s = s.strip()
    
    # 3. Perbaiki singkatan tahun
    s = re.sub(r"'(\d{2})$", r"20\1", s)
    
    # 4. Tangani format YYYYMMDD
    if re.fullmatch(r'\d{8}', s.replace(' ', '')):
         s = re.sub(r'(\d{4})(\d{2})(\d{2})', r'\1-\2-\3', s)

    # 5. Coba mengurai ke datetime (Koreksi Logika Parsing Terakhir)
    try:
        parsed_date = pd.NaT
        components = s.split() 
        
        # --- Strategi Baru: Menggunakan format eksplisit atau membiarkan inferensi Pandas ---
        
        # Deteksi Format yang Jelas Numerik DD/MM/YYYY (2 atau 4 digit tahun)
        # Format ini adalah sumber utama konflik dayfirst=True
        # Kita akan secara eksplisit memberitahu Pandas bahwa ini adalah DD/MM/YYYY
        if re.fullmatch(r'\d{1,2}\s\d{1,2}\s(\d{4}|\d{2})', s):
            # Coba parsing sebagai DD MM YYYY. Jika gagal, biarkan Pandas menginfersi
            try:
                # Menggunakan format eksplisit menghilangkan konflik dan warning
                parsed_date = pd.to_datetime(s, format='%d %m %Y', errors='coerce')
            except ValueError:
                 pass # Biarkan Pandas mencoba inferensi di bawah

        # Deteksi Year-First (YYYY-MM-DD atau YYYY MM DD, termasuk waktu/zona waktu)
        elif re.match(r'^\d{4}', s): 
            # Jika dimulai dengan 4 digit, gunakan dayfirst=False (ISO/Year-First)
            # Ini mengatasi warning 2 dan 3
             parsed_date = pd.to_datetime(s, errors='coerce', dayfirst=False)
        
        # Deteksi Month-First Teks (Misal: "Jan 12 2023")
        elif len(components) >= 2 and components[0] in ENGLISH_MONTH_ABBREVIATIONS:
            # Jika dimulai dengan bulan, gunakan dayfirst=False (Month-First)
            # Ini mengatasi warning 1
             parsed_date = pd.to_datetime(s, errors='coerce', dayfirst=False)
        
        # Inferensi Standar
        if pd.isna(parsed_date):
            # Untuk semua format lain yang sangat kotor, biarkan Pandas menginfersi
            # TANPA dayfirst=True/False untuk menghindari warning
            parsed_date = pd.to_datetime(s, errors='coerce') 
        
        if pd.isna(parsed_date):
             return s 
        
        # FINAL: Mengembalikan STRING BERFORMAT dd-mm-yyyy
        return parsed_date.strftime("%d-%m-%Y")
        
    except Exception:
        return s

# Fungsi Utama
def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    df = pd.read_excel(input_xlsx_path, sheet_name="transaksi")
    # 1. Terapkan cleaning, menghasilkan string terformat "dd-mm-yyyy"
    df["Tanggal Transaksi"] = df["Tanggal Transaksi"].apply(clean_and_parse_date)
    # tulis ulang
    with pd.ExcelWriter(
        output_xlsx_path,
        engine='xlsxwriter',
        datetime_format="dd-mm-yyyy",
        date_format="dd/mm/yyyy",
    ) as writer:
        df.to_excel(writer, sheet_name='transaksi', index=False)

normalize_tanggal_transaksi("penjualan_dqmart_01-beta.xlsx", "penjualan_dqmart_clean.xlsx")
