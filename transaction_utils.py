import pandas as pd
import locale

# Set locale ke Indonesia untuk format mata uang yang benar
try:
    locale.setlocale(locale.LC_ALL, 'id_ID.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'id_ID')
    except locale.Error:
        locale.setlocale(locale.LC_ALL, '')  # Gunakan default sistem

def count_transactions(file, vendor=None):
    """
    Menganalisis transaksi purchase order dari file Excel.
    
    Args:
        file: Path file Excel atau objek file
        vendor: Nama vendor opsional untuk filter transaksi
        
    Returns:
        tuple: (total_transaksi, total_jumlah, count_0_100k, count_100k_500k, 
                count_500k_1M, count_1M_5M, count_5M_10M, count_10M_50M, count_above_50M)
    """
    # Memuat file Excel
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names
    
    if not sheet_names:
        raise ValueError("Tidak ada sheet yang ditemukan dalam file Excel.")
    
    # Membaca sheet pertama
    df = pd.read_excel(file, sheet_name=sheet_names[0])
    
    # Nama kolom
    total_column = 'jumlah'
    vendor_column = 'vendor_name'
    status_column = 'po_status_approval'
    
    # Validasi kolom yang diperlukan ada
    if total_column not in df.columns:
        raise ValueError(f"Kolom '{total_column}' tidak ditemukan. Kolom yang tersedia: {', '.join(df.columns)}")
    if vendor_column not in df.columns:
        raise ValueError(f"Kolom '{vendor_column}' tidak ditemukan. Kolom yang tersedia: {', '.join(df.columns)}")
    
    # Membersihkan data: hapus nilai null dan konversi ke numerik
    df[total_column] = pd.to_numeric(df[total_column], errors='coerce')
    df = df.dropna(subset=[total_column])
    
    # Filter: ONLY "Approved"
    if status_column in df.columns:
        df = df[df[status_column] == 'Approved']
    
    # Filter berdasarkan vendor jika ditentukan
    if vendor:
        df = df[df[vendor_column] == vendor]
        if df.empty:
            raise ValueError(f"Tidak ada transaksi yang ditemukan untuk vendor: {vendor}")
    
    # Hitung total transaksi
    total_transactions = len(df)
    
    # Hitung total jumlah
    total_amount = df[total_column].sum()
    
    # Definisi rentang untuk maintainability yang lebih baik
    ranges = [
        (0, 100000),
        (100001, 500000),
        (500001, 1000000),
        (1000001, 5000000),
        (5000001, 10000000),
        (10000001, 50000000),
    ]
    
    # Hitung transaksi di setiap rentang
    counts = []
    for min_val, max_val in ranges:
        count = len(df[(df[total_column] >= min_val) & (df[total_column] <= max_val)])
        counts.append(count)
    
    # Hitung transaksi di atas 50M
    count_above_50M = len(df[df[total_column] > 50000000])
    counts.append(count_above_50M)
    
    return (total_transactions, total_amount, *counts)


def count_unique_po_by_range(file, vendor=None):
    """
    Menghitung jumlah PO unik berdasarkan po_code untuk setiap rentang.
    Jika po_code sama, dihitung sebagai 1 PO.
    
    Args:
        file: Path file Excel atau objek file
        vendor: Nama vendor opsional untuk filter transaksi
        
    Returns:
        tuple: (total_unique_po, count_0_100k, count_100k_500k, 
                count_500k_1M, count_1M_5M, count_5M_10M, count_10M_50M, count_above_50M)
    """
    xls = pd.ExcelFile(file)
    df = pd.read_excel(file, sheet_name=xls.sheet_names[0])
    
    total_column = 'jumlah'
    vendor_column = 'vendor_name'
    status_column = 'po_status_approval'
    po_code_column = 'po_code'
    
    # Validasi kolom
    if total_column not in df.columns:
        raise ValueError(f"Kolom '{total_column}' tidak ditemukan.")
    if po_code_column not in df.columns:
        raise ValueError(f"Kolom '{po_code_column}' tidak ditemukan.")
    
    # Membersihkan data
    df[total_column] = pd.to_numeric(df[total_column], errors='coerce')
    df = df.dropna(subset=[total_column, po_code_column])
    
    # Filter: ONLY "Approved"
    if status_column in df.columns:
        df = df[df[status_column] == 'Approved']
    
    # Filter berdasarkan vendor
    if vendor:
        df = df[df[vendor_column] == vendor]
    
    # Agregasi berdasarkan po_code - ambil total jumlah per PO
    df_po = df.groupby(po_code_column)[total_column].sum().reset_index()
    
    # Hitung total unique PO
    total_unique_po = len(df_po)
    
    # Definisi rentang
    ranges = [
        (0, 100000),
        (100001, 500000),
        (500001, 1000000),
        (1000001, 5000000),
        (5000001, 10000000),
        (10000001, 50000000),
    ]
    
    # Hitung PO unik di setiap rentang
    counts = []
    for min_val, max_val in ranges:
        count = len(df_po[(df_po[total_column] >= min_val) & (df_po[total_column] <= max_val)])
        counts.append(count)
    
    # Hitung PO di atas 50M
    count_above_50M = len(df_po[df_po[total_column] > 50000000])
    counts.append(count_above_50M)
    
    return (total_unique_po, *counts)


def get_po_breakdown(file, vendor=None):
    """
    Mendapatkan breakdown detail untuk setiap PO unik.
    
    Args:
        file: Path file Excel atau objek file
        vendor: Nama vendor opsional untuk filter transaksi
        
    Returns:
        pd.DataFrame: DataFrame dengan kolom [PO Code, Jumlah Transaksi, Total Nilai, Rentang]
    """
    xls = pd.ExcelFile(file)
    df = pd.read_excel(file, sheet_name=xls.sheet_names[0])
    
    total_column = 'jumlah'
    vendor_column = 'vendor_name'
    status_column = 'po_status_approval'
    po_code_column = 'po_code'
    
    # Validasi kolom
    if total_column not in df.columns:
        raise ValueError(f"Kolom '{total_column}' tidak ditemukan.")
    if po_code_column not in df.columns:
        raise ValueError(f"Kolom '{po_code_column}' tidak ditemukan.")
    
    # Membersihkan data
    df[total_column] = pd.to_numeric(df[total_column], errors='coerce')
    df = df.dropna(subset=[total_column, po_code_column])
    
    # Filter: ONLY "Approved"
    if status_column in df.columns:
        df = df[df[status_column] == 'Approved']
    
    # Filter berdasarkan vendor
    if vendor:
        df = df[df[vendor_column] == vendor]
    
    # Agregasi per PO
    po_summary = df.groupby(po_code_column).agg({
        total_column: ['sum', 'count']
    }).reset_index()
    
    po_summary.columns = ['PO Code', 'Total Nilai', 'Jumlah Transaksi']
    
    # Tambahkan kolom rentang
    def get_range(value):
        if value <= 100000:
            return "0 - 100 Ribu"
        elif value <= 500000:
            return "100 Ribu - 500 Ribu"
        elif value <= 1000000:
            return "500 Ribu - 1 Juta"
        elif value <= 5000000:
            return "1 Juta - 5 Juta"
        elif value <= 10000000:
            return "5 Juta - 10 Juta"
        elif value <= 50000000:
            return "10 Juta - 50 Juta"
        else:
            return "Di atas 50 Juta"
    
    po_summary['Rentang'] = po_summary['Total Nilai'].apply(get_range)
    
    # Urutkan berdasarkan Total Nilai descending
    po_summary = po_summary.sort_values('Total Nilai', ascending=False).reset_index(drop=True)
    
    # Reorder kolom
    po_summary = po_summary[['PO Code', 'Jumlah Transaksi', 'Total Nilai', 'Rentang']]
    
    return po_summary


def get_unique_po_statistics(file, vendor=None):
    """
    Mendapatkan statistik PO unik berdasarkan rentang sebagai DataFrame.
    
    Args:
        file: Path file Excel atau objek file
        vendor: Nama vendor opsional untuk filter transaksi
        
    Returns:
        pd.DataFrame: Statistik PO unik berdasarkan rentang
    """
    counts = count_unique_po_by_range(file, vendor)
    
    ranges_labels = [
        "0 - 100 Ribu",
        "100 Ribu - 500 Ribu",
        "500 Ribu - 1 Juta",
        "1 Juta - 5 Juta",
        "5 Juta - 10 Juta",
        "10 Juta - 50 Juta",
        "Di atas 50 Juta"
    ]
    
    stats_df = pd.DataFrame({
        'Rentang': ranges_labels,
        'Jumlah PO': counts[1:]  # Lewati total_unique_po
    })
    
    # Hitung persentase
    total = counts[0]
    stats_df['Persentase'] = (stats_df['Jumlah PO'] / total * 100).round(2) if total > 0 else 0
    
    return stats_df


def get_transaction_dataframe(file, vendor=None):
    """
    Mendapatkan data transaksi detail sebagai DataFrame untuk visualisasi.
    
    Args:
        file: Path file Excel atau objek file
        vendor: Nama vendor opsional untuk filter transaksi
        
    Returns:
        pd.DataFrame: Data transaksi yang sudah difilter
    """
    xls = pd.ExcelFile(file)
    df = pd.read_excel(file, sheet_name=xls.sheet_names[0])
    
    total_column = 'jumlah'
    vendor_column = 'vendor_name'
    status_column = 'po_status_approval'
    
    # Membersihkan data
    df[total_column] = pd.to_numeric(df[total_column], errors='coerce')
    df = df.dropna(subset=[total_column])
    
    # Filter: ONLY "Approved"
    if status_column in df.columns:
        df = df[df[status_column] == 'Approved']
    
    # Filter berdasarkan vendor jika ditentukan
    if vendor:
        df = df[df[vendor_column] == vendor]
    
    return df


def get_range_statistics(file, vendor=None):
    """
    Mendapatkan statistik transaksi berdasarkan rentang sebagai DataFrame untuk visualisasi mudah.
    
    Args:
        file: Path file Excel atau objek file
        vendor: Nama vendor opsional untuk filter transaksi
        
    Returns:
        pd.DataFrame: Statistik berdasarkan rentang transaksi
    """
    counts = count_transactions(file, vendor)
    
    ranges_labels = [
        "0 - 100 Ribu",
        "100 Ribu - 500 Ribu",
        "500 Ribu - 1 Juta",
        "1 Juta - 5 Juta",
        "5 Juta - 10 Juta",
        "10 Juta - 50 Juta",
        "Di atas 50 Juta"
    ]
    
    # Membuat DataFrame untuk visualisasi
    stats_df = pd.DataFrame({
        'Rentang': ranges_labels,
        'Jumlah': counts[2:]  # Lewati total_transaksi dan total_jumlah
    })
    
    # Hitung persentase
    total = counts[0]
    stats_df['Persentase'] = (stats_df['Jumlah'] / total * 100).round(2)
    
    return stats_df