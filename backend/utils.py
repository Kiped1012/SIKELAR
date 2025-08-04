"""
Utility functions for SIKELAR application
Contains helper functions for Excel processing, formatting, file operations, and validation
"""

import re
import os
import colorsys
from datetime import datetime


class ExcelUtils:
    """Utility class for Excel data extraction operations"""
    
    @staticmethod
    def is_valid_kegiatan_format(kode_kegiatan):
        """Cek apakah kode kegiatan sesuai format xx.xx.xx (2 digit, titik, 2 digit, titik, 2 digit, titik)"""
        if not kode_kegiatan:
            return False
        
        # Pattern untuk format xx.xx.xx.
        pattern = r'^\d{2}\.\d{2}\.\d{2}\.$'
        return bool(re.match(pattern, kode_kegiatan))
    
    @staticmethod
    def extract_merged_text_strict(sheet, row_idx, col_range):
        """Ekstrak teks dari kolom yang di-merge - VERSI STRICT tanpa fallback ke baris lain"""
        text_parts = []
        
        # Hanya baca dari baris yang tepat saja
        for col_idx in col_range:
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if cell_value and str(cell_value).strip() and str(cell_value).strip() != "None":
                text_parts.append(str(cell_value).strip())
        
        # Gabungkan teks yang ditemukan
        combined_text = " ".join(text_parts).strip()
        
        return combined_text
    
    @staticmethod
    def extract_merged_text(sheet, row_idx, col_range):
        """Ekstrak teks dari kolom yang di-merge"""
        text_parts = []
        
        for col_idx in col_range:
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if cell_value and str(cell_value).strip() and str(cell_value).strip() != "None":
                text_parts.append(str(cell_value).strip())
        
        # Gabungkan teks yang ditemukan
        combined_text = " ".join(text_parts).strip()
        
        # Jika tidak ditemukan teks di baris yang sama, coba baris sebelum/sesudah (untuk handling merge cell)
        if not combined_text:
            for offset in [-1, 1, -2, 2]:
                if row_idx + offset > 0:
                    for col_idx in col_range:
                        cell_value = sheet.cell(row=row_idx + offset, column=col_idx).value
                        if cell_value and str(cell_value).strip() and str(cell_value).strip() != "None":
                            combined_text = str(cell_value).strip()
                            break
                    if combined_text:
                        break
        
        return combined_text
    
    @staticmethod
    def extract_merged_number(sheet, row_idx, col_range):
        """Ekstrak angka dari kolom yang di-merge"""
        for col_idx in col_range:
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if isinstance(cell_value, (int, float)) and cell_value > 0:
                return int(cell_value)
        
        # Jika tidak ditemukan di baris yang sama, coba baris sebelum/sesudah
        for offset in [-1, 1, -2, 2]:
            if row_idx + offset > 0:
                for col_idx in col_range:
                    cell_value = sheet.cell(row=row_idx + offset, column=col_idx).value
                    if isinstance(cell_value, (int, float)) and cell_value > 0:
                        return int(cell_value)
        
        return 0

    @staticmethod
    def extract_merged_text_bku(sheet, row_idx, col_range):
        """Ekstrak teks dari kolom yang di-merge khusus untuk BKU (lebih fleksibel)"""
        text_parts = []
        
        # Coba dari baris yang tepat dulu
        for col_idx in col_range:
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if cell_value and str(cell_value).strip() and str(cell_value).strip() != "None":
                text_parts.append(str(cell_value).strip())
        
        combined_text = " ".join(text_parts).strip()
        
        # Jika tidak ditemukan, coba baris sekitar (untuk merged cells)
        if not combined_text:
            for offset in [-1, 1, -2, 2, -3, 3]:
                if row_idx + offset > 0:
                    for col_idx in col_range:
                        try:
                            cell_value = sheet.cell(row=row_idx + offset, column=col_idx).value
                            if cell_value and str(cell_value).strip() and str(cell_value).strip() != "None":
                                combined_text = str(cell_value).strip()
                                break
                        except:
                            continue
                    if combined_text:
                        break
        
        return combined_text

    @staticmethod
    def extract_merged_number_bku(sheet, row_idx, col_range):
        """Ekstrak angka dari kolom yang di-merge khusus untuk BKU"""
        # Coba dari baris yang tepat dulu
        for col_idx in col_range:
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if isinstance(cell_value, (int, float)) and cell_value > 0:
                return int(cell_value)
        
        # Jika tidak ditemukan, coba baris sekitar (untuk merged cells)
        for offset in [-1, 1, -2, 2, -3, 3]:
            if row_idx + offset > 0:
                for col_idx in col_range:
                    try:
                        cell_value = sheet.cell(row=row_idx + offset, column=col_idx).value
                        if isinstance(cell_value, (int, float)) and cell_value > 0:
                            return int(cell_value)
                    except:
                        continue
        
        return 0

    @staticmethod
    def parse_bku_date(date_value):
        """Parse tanggal dari berbagai format yang mungkin ada di BKU"""
        import datetime
        
        if not date_value:
            return None
        
        # Jika sudah berupa datetime object
        if isinstance(date_value, datetime.datetime):
            return date_value.date()
        elif isinstance(date_value, datetime.date):
            return date_value
        
        # Convert ke string dan bersihkan
        date_str = str(date_value).strip()
        
        # Format yang akan dicoba
        date_formats = [
            '%d-%m-%Y',    # 27-02-2025
            '%d/%m/%Y',    # 27/02/2025
            '%Y-%m-%d',    # 2025-02-27
            '%d.%m.%Y',    # 27.02.2025
            '%d %m %Y',    # 27 02 2025
        ]
        
        for fmt in date_formats:
            try:
                return datetime.datetime.strptime(date_str, fmt).date()
            except:
                continue
        
        # Jika semua format gagal, coba parsing manual
        try:
            # Coba split dengan berbagai separator
            for sep in ['-', '/', '.', ' ']:
                if sep in date_str:
                    parts = date_str.split(sep)
                    if len(parts) == 3:
                        # Assume DD-MM-YYYY format first
                        try:
                            day, month, year = map(int, parts)
                            if 1 <= day <= 31 and 1 <= month <= 12 and year > 1900:
                                return datetime.date(year, month, day)
                        except:
                            pass
                        
                        # Try YYYY-MM-DD format
                        try:
                            year, month, day = map(int, parts)
                            if 1 <= day <= 31 and 1 <= month <= 12 and year > 1900:
                                return datetime.date(year, month, day)
                        except:
                            pass
                    break
        except:
            pass
        
        print(f"Warning: Could not parse date: {date_value}")
        return None


class FormatUtils:
    """Utility class for formatting and data cleaning operations"""
    
    @staticmethod
    def clean_number(num_str):
        """Membersihkan format angka dari string"""
        if not num_str:
            return 0
        
        cleaned = str(num_str).replace('Rp', '').replace(' ', '').strip()
        
        if ',' in cleaned:
            parts = cleaned.split(',')
            integer_part = parts[0]
            decimal_part = parts[1] if len(parts) > 1 else '0'
            integer_part = integer_part.replace('.', '')
            cleaned = integer_part + '.' + decimal_part
        else:
            cleaned = cleaned.replace('.', '')
        
        try:
            return float(cleaned)
        except ValueError:
            return 0
    
    @staticmethod
    def format_currency(amount):
        """Format number as Indonesian currency"""
        return f"Rp {amount:,.0f}".replace(',', '.')
    
    @staticmethod
    def darken_color(hex_color, factor=0.65):
        """Turunkan kecerahan warna hex (factor < 1.0)"""
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        h, l, s = colorsys.rgb_to_hls(r/255.0, g/255.0, b/255.0)
        l = max(0, min(1, l * factor))
        r, g, b = colorsys.hls_to_rgb(h, l, s)
        return f'#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}'
    
    @staticmethod
    def extract_school_name(raw_data):
        """Mengekstrak nama sekolah dari data input"""
        lines = raw_data.split('\n')
        
        for line in lines:
            # Cari baris yang mengandung "Nama:"
            if "Nama:" in line:
                # Ekstrak nama sekolah setelah "Nama:"
                name_match = re.search(r'Nama:\s*(.+)', line.strip())
                if name_match:
                    return name_match.group(1).strip()
        
        # Jika tidak ditemukan, set default
        return "Nama Sekolah Tidak Ditemukan"
    
    @staticmethod
    def create_table_display(title, data_dict, found_codes, total_budget, school_name):
        """Membuat tampilan tabel yang rapi"""
        if not found_codes:
            return f"\n{title}\n" + "="*150 + "\nTidak ada data yang ditemukan.\n" + "="*150
        
        output = f"\n{title}\n"
        output += "="*150 + "\n"
        
        # Header tabel
        output += f"{'Kode':<15} | {'Uraian':<80} | {'Jumlah (Rp)':<20}\n"
        output += "-"*150 + "\n"
        
        total_alokasi = 0
        for kode in sorted(found_codes):
            if kode in data_dict:
                data = data_dict[kode]
                uraian = data['uraian']
                if len(uraian) > 78:
                    uraian = uraian[:75] + "..."
                
                jumlah = FormatUtils.format_currency(data['jumlah'])
                output += f"{kode:<15} | {uraian:<80} | {jumlah:<20}\n"
                total_alokasi += data['jumlah']
        
        output += "-"*150 + "\n"
        output += f"{'TOTAL':<15} | {'':<80} | {FormatUtils.format_currency(total_alokasi):<20}\n"
        
        if total_budget > 0:
            persentase = (total_alokasi / total_budget) * 100
            output += f"{'PERSENTASE':<15} | {'':<80} | {persentase:.2f}%\n"
        
        # Tambahkan nama sekolah di bawah persentase
        output += f"{'NAMA SEKOLAH':<15} | {'':<80} | {school_name}\n"
        
        output += "="*150
        return output
    
    @staticmethod
    def create_summary_display(data_processor):
        """Membuat tampilan ringkasan data yang telah diproses"""
        summary = data_processor.get_summary()
        
        output = "="*120 + "\n"
        output += f"{'RINGKASAN DATA YANG DIPROSES':^120}\n"
        output += "="*120 + "\n"
        output += f"Nama Sekolah: {summary['school_name']}\n"
        output += f"Total item yang diproses: {summary['total_items']}\n"
        output += f"Total anggaran: {FormatUtils.format_currency(summary['total_budget'])}\n"
        output += f"Waktu pemrosesan: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        output += "\n"
        
        # Breakdown per kategori
        output += "BREAKDOWN PER KATEGORI:\n"
        output += f"- BUKU (05.02): {summary['buku_count']} item\n"
        output += f"- SARANA DAN PRASARANA (05.08): {summary['sarana_count']} item\n"
        output += f"- HONOR (07.12): {summary['honor_count']} item\n"
        output += "\n"
        
        # Tampilkan kode-kode yang berhasil ditemukan
        output += "KODE YANG BERHASIL DITEMUKAN:\n"
        for kode in sorted(data_processor.processed_data.keys()):
            uraian = data_processor.processed_data[kode]['uraian']
            if len(uraian) > 60:
                uraian = uraian[:57] + "..."
            output += f"- {kode}: {uraian}\n"
        
        output += "="*120 + "\n"
        output += "Silakan pilih kategori yang ingin ditampilkan menggunakan tombol di atas.\n"
        output += "="*120
        
        return output


class FileUtils:
    """Utility class for file operations"""
    
    @staticmethod
    def validate_excel_file(file_path):
        """Validasi apakah file adalah Excel yang valid"""
        if not os.path.exists(file_path):
            return False, "File tidak ditemukan"
        
        if not os.path.isfile(file_path):
            return False, "Path bukan file"
        
        # Cek ekstensi file
        valid_extensions = ['.xlsx', '.xls']
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext not in valid_extensions:
            return False, f"Format file tidak didukung. Gunakan {', '.join(valid_extensions)}"
        
        # Cek ukuran file (maksimal 10MB)
        file_size = os.path.getsize(file_path)
        max_size = 10 * 1024 * 1024  # 10MB
        
        if file_size > max_size:
            return False, "File terlalu besar (maksimal 10MB)"
        
        return True, "File valid"
    
    @staticmethod
    def get_file_info(file_path):
        """Mendapatkan informasi file"""
        if not os.path.exists(file_path):
            return None
        
        stat = os.stat(file_path)
        
        return {
            'name': os.path.basename(file_path),
            'size': stat.st_size,
            'size_formatted': FileUtils.format_file_size(stat.st_size),
            'modified': datetime.fromtimestamp(stat.st_mtime),
            'extension': os.path.splitext(file_path)[1].lower()
        }
    
    @staticmethod
    def format_file_size(size_bytes):
        """Format ukuran file menjadi string yang mudah dibaca"""
        if size_bytes == 0:
            return "0B"
        
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.1f}{size_names[i]}"


class ValidationUtils:
    """Utility class for data validation"""
    
    @staticmethod
    def validate_kode_format(kode):
        """Validasi format kode anggaran"""
        # Format: XX.XX.XX atau XX.XX
        pattern = r'^\d{2}\.\d{2}(\.\d{2})?$'
        return re.match(pattern, kode) is not None
    
    @staticmethod
    def validate_currency_format(currency_str):
        """Validasi format mata uang"""
        # Format: Rp 1.000.000 atau Rp 1,000,000 atau 1000000
        pattern = r'^(Rp\s*)?[\d.,]+$'
        return re.match(pattern, currency_str.strip()) is not None
    
    @staticmethod
    def validate_school_name(name):
        """Validasi nama sekolah"""
        if not name or name.strip() == "":
            return False, "Nama sekolah tidak boleh kosong"
        
        if len(name.strip()) < 3:
            return False, "Nama sekolah terlalu pendek"
        
        if len(name.strip()) > 100:
            return False, "Nama sekolah terlalu panjang"
        
        return True, "Nama sekolah valid"
    
    @staticmethod
    def validate_processed_data(processed_data):
        """Validasi data yang telah diproses"""
        if not processed_data:
            return False, "Tidak ada data yang diproses"
        
        required_fields = ['uraian', 'volume', 'satuan', 'harga_satuan', 'jumlah']
        
        for kode, data in processed_data.items():
            # Validasi format kode
            if not ValidationUtils.validate_kode_format(kode):
                return False, f"Format kode tidak valid: {kode}"
            
            # Validasi field yang diperlukan
            for field in required_fields:
                if field not in data:
                    return False, f"Field '{field}' tidak ditemukan untuk kode {kode}"
            
            # Validasi jumlah tidak negatif
            if data['jumlah'] < 0:
                return False, f"Jumlah tidak boleh negatif untuk kode {kode}"
        
        return True, "Data valid"


# Helper functions untuk kompatibilitas dengan kode lama
def is_valid_kegiatan_format(kode_kegiatan):
    """Wrapper function untuk kompatibilitas"""
    return ExcelUtils.is_valid_kegiatan_format(kode_kegiatan)

def extract_merged_text(sheet, row_idx, col_range):
    """Wrapper function untuk kompatibilitas"""
    return ExcelUtils.extract_merged_text(sheet, row_idx, col_range)

def extract_merged_number(sheet, row_idx, col_range):
    """Wrapper function untuk kompatibilitas"""
    return ExcelUtils.extract_merged_number(sheet, row_idx, col_range)

def format_currency(amount):
    """Wrapper function untuk kompatibilitas"""
    return FormatUtils.format_currency(amount)

def clean_number(num_str):
    """Wrapper function untuk kompatibilitas"""
    return FormatUtils.clean_number(num_str)


# Contoh penggunaan
if __name__ == "__main__":
    # Test format currency
    print(FormatUtils.format_currency(1500000))  # Rp 1.500.000
    
    # Test clean number
    print(FormatUtils.clean_number("Rp 1.500.000"))  # 1500000.0
    
    # Test validate kode format
    print(ValidationUtils.validate_kode_format("05.02.01"))  # True
    print(ValidationUtils.validate_kode_format("05.02"))     # True
    print(ValidationUtils.validate_kode_format("5.2.1"))     # False
    
    # Test school name validation
    is_valid, message = ValidationUtils.validate_school_name("SD Negeri 1 Jakarta")
    print(f"Valid: {is_valid}, Message: {message}")
    
    # Test file size formatting
    print(FileUtils.format_file_size(1500000))  # 1.4MB