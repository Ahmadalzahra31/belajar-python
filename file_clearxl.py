import os
import sys
import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS
import re

# ===================== FUNGSI RENAME FOTO =====================
def get_exif_date(file_path):
    """Ambil tanggal dari EXIF (DateTimeOriginal) file gambar."""
    try:
        img = Image.open(file_path)
        exif = img._getexif()
        if exif:
            for tag_id, value in exif.items():
                tag = TAGS.get(tag_id, tag_id)
                if tag == "DateTimeOriginal":
                    return value  # Format: "YYYY:MM:DD HH:MM:SS"
    except Exception as e:
        print(f"  Gagal membaca EXIF {file_path}: {e}")
    return None

def rename_photos(directory, prefix, dry_run=False):
    """Rename file gambar berdasarkan tanggal EXIF."""
    extensions = ('.jpg', '.jpeg', '.png', '.tiff', '.bmp')
    files = [f for f in os.listdir(directory) if f.lower().endswith(extensions)]
    if not files:
        print("Tidak ada file gambar yang ditemukan.")
        return

    renamed_count = 0
    date_counter = {}

    for filename in files:
        file_path = os.path.join(directory, filename)
        date_str = get_exif_date(file_path)
        if not date_str:
            print(f"  {filename}: tidak ada EXIF, lewati.")
            continue

        # Parsing tanggal "YYYY:MM:DD HH:MM:SS" -> YYYY, MM
        match = re.match(r"(\d{4}):(\d{2}):\d{2}", date_str)
        if not match:
            print(f"  {filename}: format tanggal tidak dikenal: {date_str}")
            continue
        year, month = match.groups()
        base_new_name = f"{prefix}-{year}-{month}.jpg"
        new_name = base_new_name

        # Handle duplikat nama
        if new_name in date_counter:
            date_counter[new_name] += 1
            name_without_ext = new_name[:-4]  # hapus .jpg
            new_name = f"{name_without_ext}_{date_counter[new_name]}.jpg"
        else:
            date_counter[new_name] = 1

        new_path = os.path.join(directory, new_name)
        if dry_run:
            print(f"{filename} -> {new_name}")
        else:
            os.rename(file_path, new_path)
            print(f"Renamed: {filename} -> {new_name}")
        renamed_count += 1

    print(f"\nSelesai. {renamed_count} file di-rename.")

# ===================== FUNGSI BERSIHKAN EXCEL/CSV =====================
def clean_excel(file_path, output_path=None, remove_duplicates=True,
                date_columns=None, strip_whitespace=True):
    """Bersihkan file Excel/CSV: hapus duplikat, format tanggal, hapus spasi."""
    # Deteksi ekstensi
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    awal = len(df)
    print(f"Jumlah baris awal: {awal}")

    # Hapus duplikat
    if remove_duplicates:
        df.drop_duplicates(inplace=True)
        print(f"Duplikat dihapus. Baris sekarang: {len(df)}")

    # Bersihkan spasi pada kolom string
    if strip_whitespace:
        str_cols = df.select_dtypes(include=['object']).columns
        for col in str_cols:
            df[col] = df[col].astype(str).str.strip()
        print("Spasi berlebih pada kolom teks telah dibersihkan.")

    # Format kolom tanggal
    if date_columns:
        for col in date_columns:
            if col in df.columns:
                # Ubah ke datetime, format jadi YYYY-MM-DD
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
                print(f"Kolom '{col}' diformat ke YYYY-MM-DD.")
            else:
                print(f"Kolom '{col}' tidak ditemukan.")

    # Simpan
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_clean{ext}"

    if output_path.endswith('.csv'):
        df.to_csv(output_path, index=False)
    else:
        df.to_excel(output_path, index=False)

    print(f"File bersih disimpan sebagai: {output_path}")

# ===================== MENU UTAMA =====================
def main():
    print("=" * 50)
    print("   PROGRAM MULTIGUNA: RENAME FOTO & BERSIHKAN EXCEL")
    print("=" * 50)
    while True:
        print("\nPilih mode:")
        print("1. Rename file foto berdasarkan EXIF")
        print("2. Bersihkan file Excel/CSV (hapus duplikat, spasi, format tanggal)")
        print("3. Keluar")
        pilihan = input("Masukkan pilihan (1/2/3): ").strip()

        if pilihan == '1':
            # Mode rename foto
            dir_path = input("Masukkan path direktori foto: ").strip()
            if not os.path.isdir(dir_path):
                print("Direktori tidak valid.")
                continue
            prefix = input("Masukkan prefix nama baru (misal: Liburan): ").strip()
            if not prefix:
                prefix = "Foto"
            dry = input("Jalankan dalam mode dry-run? (y/n, hanya tampilkan tanpa rename): ").strip().lower()
            dry_run = (dry == 'y')
            rename_photos(dir_path, prefix, dry_run)

        elif pilihan == '2':
            # Mode bersihkan Excel/CSV
            file_path = input("Masukkan path file Excel/CSV: ").strip()
            if not os.path.isfile(file_path):
                print("File tidak ditemukan.")
                continue
            output_path = input("Masukkan path output (kosongkan untuk otomatis): ").strip()
            if output_path == "":
                output_path = None

            dup = input("Hapus duplikat? (y/n, default y): ").strip().lower()
            remove_duplicates = (dup != 'n')

            spasi = input("Bersihkan spasi berlebih pada teks? (y/n, default y): ").strip().lower()
            strip_whitespace = (spasi != 'n')

            date_cols_input = input("Masukkan nama kolom tanggal (pisahkan dengan koma jika lebih, kosongkan jika tidak ada): ").strip()
            date_columns = [c.strip() for c in date_cols_input.split(',')] if date_cols_input else None

            clean_excel(file_path, output_path, remove_duplicates, date_columns, strip_whitespace)

        elif pilihan == '3':
            print("Terima kasih. Sampai jumpa!")
            break
        else:
            print("Pilihan tidak valid. Coba lagi.")

if __name__ == "__main__":
    main()