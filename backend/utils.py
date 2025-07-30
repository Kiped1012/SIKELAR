"""
Utility functions for SIKELAR application
Contains helper functions for Excel processing
"""

import re

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
    """Utility class for formatting operations"""
    
    @staticmethod
    def format_currency(amount):
        """Format number as Indonesian currency"""
        return f"Rp {amount:,}".replace(',', '.')
    
    @staticmethod
    def darken_color(hex_color, factor=0.65):
        """Turunkan kecerahan warna hex (factor < 1.0)"""
        import colorsys
        
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        h, l, s = colorsys.rgb_to_hls(r/255.0, g/255.0, b/255.0)
        l = max(0, min(1, l * factor))
        r, g, b = colorsys.hls_to_rgb(h, l, s)
        return f'#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}'