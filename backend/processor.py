"""
Combined Data processing module for SIKELAR application
Handles parsing and processing of RKAS and BKU data
"""

import re
from .utils import FormatUtils
from .rkas_processor import RKASDataProcessor  
from .bku_processor import BKUDataProcessor

 
class BOSDataProcessor:
    """
    Main Data processor for SIKELAR application
    Combines RKAS and BKU data processing
    """
    def __init__(self):
        self.rkas_processor = RKASDataProcessor()
        self.bku_processor = BKUDataProcessor()
            
        self.raw_data = ""
        self.processed_data = {}
        self.total_budget = 0
        self.school_name = ""
        
        # Definisikan kode-kode yang diinginkan secara spesifik
        self.target_codes = {
            '05.02.01', '05.02.02', '05.02.03', '05.02.04', '05.02.05',  # BUKU
            '05.08.01', '05.08.08', '05.08.12',  # SARANA DAN PRASARANA
            '07.12'  # HONOR
        }

        
    def reset_data(self):
        """Reset all data to initial state"""
        self.rkas_processor.reset_data()
        self.bku_processor.reset_data()

    def extract_excel_data(self, file_path):
        """Ekstrak data dari file Excel dengan struktur spesifik RKAS dan BKU"""
        print("Debug: Starting data extraction...")
        
        # Reset semua data
        self.reset_data()
        
        # Ekstrak data RKAS
        self.rkas_processor.extract_rkas_data(file_path)
        
        # Ekstrak data BKU  
        self.bku_processor.extract_bku_data(file_path)
        
        # Print summary
        print(f"Debug: RKAS - Total Penerimaan: Rp {self.rkas_processor.total_penerimaan:,}")
        print(f"Debug: RKAS - Found {len(self.rkas_processor.belanja_persediaan_items)} belanja persediaan items")
        print(f"Debug: RKAS - Found {len(self.rkas_processor.belanja_jasa_items)} belanja jasa items")
        print(f"Debug: RKAS - Found {len(self.rkas_processor.belanja_pemeliharaan_items)} belanja pemeliharaan items")
        print(f"Debug: RKAS - Found {len(self.rkas_processor.belanja_perjalanan_items)} belanja perjalanan items")
        print(f"Debug: RKAS - Found {len(self.rkas_processor.peralatan_items)} peralatan items")
        print(f"Debug: RKAS - Found {len(self.rkas_processor.aset_tetap_items)} aset tetap items")
        print(f"Debug: BKU - Data available: {self.bku_processor.bku_data_available}")

    # Properties untuk kompatibilitas dengan kode existing
    @property
    def excel_data(self):
        """Compatibility property untuk check apakah data sudah dimuat"""
        return self.rkas_processor.excel_data if hasattr(self.rkas_processor, 'excel_data') else None
    
    @property
    def total_penerimaan(self):
        return self.rkas_processor.total_penerimaan
    
    @property
    def nama_sekolah(self):
        return self.rkas_processor.nama_sekolah
    
    @property
    def budget_items(self):
        return self.rkas_processor.budget_items
    
    @property
    def belanja_persediaan_items(self):
        return self.rkas_processor.belanja_persediaan_items
    
    @property
    def belanja_jasa_items(self):
        return self.rkas_processor.belanja_jasa_items
    
    @property
    def belanja_pemeliharaan_items(self):
        return self.rkas_processor.belanja_pemeliharaan_items
    
    @property
    def belanja_perjalanan_items(self):
        return self.rkas_processor.belanja_perjalanan_items
    
    @property
    def peralatan_items(self):
        return self.rkas_processor.peralatan_items
    
    @property
    def aset_tetap_items(self):
        return self.rkas_processor.aset_tetap_items
    
    @property
    def bku_data_available(self):
        return self.bku_processor.bku_data_available

    # RKAS Methods - delegate to RKAS processor
    def filter_budget_by_codes(self, codes):
        """Filter budget berdasarkan kode kategori"""
        return self.rkas_processor.filter_budget_by_codes(codes)

    def get_summary_data(self):
        """Generate summary data for ringkasan RKAS"""
        return self.rkas_processor.get_summary_data()

    # BKU Methods - delegate to BKU processor  
    def get_bku_belanja_persediaan_by_triwulan(self, triwulan):
        """Get data realisasi belanja persediaan berdasarkan triwulan"""
        return self.bku_processor.get_bku_belanja_persediaan_by_triwulan(triwulan)

    def get_bku_belanja_pemeliharaan_by_triwulan(self, triwulan):
        """Get data realisasi belanja pemeliharaan berdasarkan triwulan"""
        return self.bku_processor.get_bku_belanja_pemeliharaan_by_triwulan(triwulan)

    def get_bku_belanja_perjalanan_by_triwulan(self, triwulan):
        """Get data realisasi belanja perjalanan berdasarkan triwulan"""
        return self.bku_processor.get_bku_belanja_perjalanan_by_triwulan(triwulan)

    def get_bku_peralatan_by_triwulan(self, triwulan):
        """Get data realisasi peralatan berdasarkan triwulan"""
        return self.bku_processor.get_bku_peralatan_by_triwulan(triwulan)

    def get_bku_aset_tetap_by_triwulan(self, triwulan):
        """Get data realisasi aset tetap lainnya berdasarkan triwulan"""
        return self.bku_processor.get_bku_aset_tetap_by_triwulan(triwulan)

    def get_bku_belanja_jasa_by_triwulan(self, triwulan):
        """Get data realisasi belanja jasa berdasarkan triwulan"""
        return self.bku_processor.get_bku_belanja_jasa_by_triwulan(triwulan)

    def get_bku_summary_data_by_triwulan(self, triwulan):
        """Generate summary data BKU untuk triwulan tertentu"""
        return self.bku_processor.get_bku_summary_data_by_triwulan(triwulan)

    def get_all_triwulan_summary(self):
        """Get ringkasan untuk semua triwulan sekaligus"""
        return self.bku_processor.get_all_triwulan_summary()

    # Compatibility methods untuk backward compatibility
    @property
    def kategori_kode(self):
        """Get kategori kode dari RKAS processor"""
        return self.rkas_processor.kategori_kode

    def process_rkas_data(self, sheet):
        """Process RKAS data from sheet - delegate to RKAS processor"""
        return self.rkas_processor.process_rkas_data(sheet)

    def process_bku_data(self, sheet):
        """Process BKU data from sheet - delegate to BKU processor"""
        return self.bku_processor.process_bku_data(sheet)

    def extract_nama_sekolah(self, sheet):
        """Extract nama sekolah - delegate to RKAS processor"""
        return self.rkas_processor.extract_nama_sekolah(sheet)

    def extract_total_penerimaan(self, sheet):
        """Extract total penerimaan - delegate to RKAS processor"""
        return self.rkas_processor.extract_total_penerimaan(sheet)

    def extract_budget_data(self, sheet):
        """Extract budget data - delegate to RKAS processor"""
        return self.rkas_processor.extract_budget_data(sheet)

    def extract_belanja_persediaan_data(self, sheet):
        """Extract belanja persediaan data - delegate to RKAS processor"""
        return self.rkas_processor.extract_belanja_persediaan_data(sheet)

    def extract_belanja_jasa_data(self, sheet):
        """Extract belanja jasa data - delegate to RKAS processor"""
        return self.rkas_processor.extract_belanja_jasa_data(sheet)

    def extract_belanja_pemeliharaan_data(self, sheet):
        """Extract belanja pemeliharaan data - delegate to RKAS processor"""
        return self.rkas_processor.extract_belanja_pemeliharaan_data(sheet)

    def extract_belanja_perjalanan_data(self, sheet):
        """Extract belanja perjalanan data - delegate to RKAS processor"""
        return self.rkas_processor.extract_belanja_perjalanan_data(sheet)

    def extract_peralatan_data(self, sheet):
        """Extract peralatan data - delegate to RKAS processor"""
        return self.rkas_processor.extract_peralatan_data(sheet)

    def extract_aset_tetap_data(self, sheet):
        """Extract aset tetap data - delegate to RKAS processor"""
        return self.rkas_processor.extract_aset_tetap_data(sheet)

    # BKU extraction methods - delegate to BKU processor
    def extract_bku_belanja_persediaan_data(self, sheet):
        """Extract BKU belanja persediaan data - delegate to BKU processor"""
        return self.bku_processor.extract_bku_belanja_persediaan_data(sheet)

    def extract_bku_belanja_pemeliharaan_data(self, sheet):
        """Extract BKU belanja pemeliharaan data - delegate to BKU processor"""
        return self.bku_processor.extract_bku_belanja_pemeliharaan_data(sheet)

    def extract_bku_belanja_perjalanan_data(self, sheet):
        """Extract BKU belanja perjalanan data - delegate to BKU processor"""
        return self.bku_processor.extract_bku_belanja_perjalanan_data(sheet)

    def extract_bku_peralatan_data(self, sheet):
        """Extract BKU peralatan data - delegate to BKU processor"""
        return self.bku_processor.extract_bku_peralatan_data(sheet)

    def extract_bku_aset_tetap_data(self, sheet):
        """Extract BKU aset tetap data - delegate to BKU processor"""
        return self.bku_processor.extract_bku_aset_tetap_data(sheet)

    def extract_bku_belanja_jasa_data(self, sheet):
        """Extract BKU belanja jasa data - delegate to BKU processor"""
        return self.bku_processor.extract_bku_belanja_jasa_data(sheet)
    
    def process_data(self, raw_input):
        """Memproses data input dan mengekstrak informasi yang diperlukan"""
        self.raw_data = raw_input.strip()
        
        if not self.raw_data:
            raise ValueError("Data input kosong!")
        
        # Reset data
        self.processed_data = {}
        self.total_budget = 0
        self.school_name = ""
        
        # Proses data
        self.school_name = FormatUtils.extract_school_name(self.raw_data)
        self.parse_rkas_data()
        self.calculate_total_budget()
        
        return {
            'processed_data': self.processed_data,
            'total_budget': self.total_budget,
            'school_name': self.school_name,
            'total_items': len(self.processed_data)
        }
    
    def process_excel_file(self, file_path):
        """Memproses file Excel dan mengekstrak data RKAS/BKU"""
        # TODO: Implement Excel file processing
        # This is a placeholder method for future Excel processing functionality
        
        try:
            # Import pandas untuk membaca Excel
            # import pandas as pd
            
            # Reset data
            self.processed_data = {}
            self.total_budget = 0
            self.school_name = ""
            
            # Placeholder processing
            # df = pd.read_excel(file_path)
            # Process the DataFrame and extract relevant information
            
            # For now, return dummy data
            self.school_name = "Sekolah dari File Excel"
            self.processed_data = {
                '05.02.01': {
                    'uraian': 'Buku Teks Pelajaran (dari Excel)',
                    'volume': '100',
                    'satuan': 'buku',
                    'harga_satuan': 25000,
                    'jumlah': 2500000
                }
            }
            self.total_budget = 2500000
            
            return {
                'processed_data': self.processed_data,
                'total_budget': self.total_budget,
                'school_name': self.school_name,
                'total_items': len(self.processed_data),
                'file_path': file_path
            }
            
        except Exception as e:
            raise ValueError(f"Gagal memproses file Excel: {str(e)}")
    
    def parse_rkas_data(self):
        """Parsing data dari input text format tabel RKAS"""
        lines = self.raw_data.split('\n')
        processed_lines = set()  # Track baris yang sudah diproses
        
        # Cari total anggaran dari header
        for line in lines:
            if "Total Anggaran:" in line:
                total_match = re.search(r'Rp\s*([\d.,]+)', line)
                if total_match:
                    self.total_budget = FormatUtils.clean_number(total_match.group(1))
                    break
        
        # Parse data baris per baris
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line or i in processed_lines:
                i += 1
                continue
            
            # Cek apakah ini adalah kode standar yang terpisah - HANYA untuk kode yang diinginkan
            kode_found = None
            for target_code in self.target_codes:
                # Pattern untuk kode yang berdiri sendiri
                if target_code == '07.12':
                    pattern = r'^07\.12\.'
                else:
                    pattern = rf'^{re.escape(target_code)}\.'
                
                if re.match(pattern, line):
                    kode_found = target_code
                    break
            
            if kode_found:
                processed_lines.add(i)  # Tandai baris ini sudah diproses
                
                if i + 1 < len(lines):
                    uraian_line = lines[i + 1].strip()
                    processed_lines.add(i + 1)  # Tandai baris uraian sudah diproses
                    
                    volume = "0"
                    satuan = "-"
                    harga_satuan = 0
                    jumlah = 0
                    
                    # Cari nilai di baris-baris berikutnya
                    for j in range(i + 2, min(i + 8, len(lines))):
                        check_line = lines[j].strip()
                        
                        jumlah_match = re.search(r'Rp\s*([\d.,]+)', check_line)
                        if jumlah_match:
                            potential_jumlah = FormatUtils.clean_number(jumlah_match.group(1))
                            if potential_jumlah > jumlah:
                                jumlah = potential_jumlah
                                processed_lines.add(j)  # Tandai baris dengan nilai sudah diproses
                        
                        volume_match = re.match(r'^(\d+)', check_line)
                        if volume_match:
                            volume = volume_match.group(1)
                            processed_lines.add(j)
                        
                        if re.match(r'^(buku|unit|orang|buah|lusin|meter|rim|botol|jam)', check_line.lower()):
                            satuan = check_line
                            processed_lines.add(j)
                    
                    # Jika tidak ada nilai terpisah, cari di uraian
                    if jumlah == 0:
                        jumlah_match = re.search(r'Rp\s*([\d.,]+)', uraian_line)
                        if jumlah_match:
                            jumlah = FormatUtils.clean_number(jumlah_match.group(1))
                            uraian_line = re.sub(r'Rp\s*[\d.,]+', '', uraian_line).strip()
                    
                    if jumlah > 0:
                        self.processed_data[kode_found] = {
                            'uraian': uraian_line,
                            'volume': volume,
                            'satuan': satuan,
                            'harga_satuan': harga_satuan,
                            'jumlah': jumlah
                        }
            
            # Cek apakah ini adalah kode standar dalam satu baris - HANYA untuk kode yang diinginkan
            elif i not in processed_lines:
                kode_found = None
                for target_code in self.target_codes:
                    if target_code == '07.12':
                        pattern = r'^07\.12\.\s+(.+)'
                    else:
                        pattern = rf'^{re.escape(target_code)}\.\s+(.+)'
                    
                    match = re.match(pattern, line)
                    if match:
                        kode_found = target_code
                        data_part = match.group(1)
                        break
                
                if kode_found:
                    processed_lines.add(i)  # Tandai baris ini sudah diproses
                    
                    # Hanya proses jika kode belum ada di processed_data
                    if kode_found not in self.processed_data:
                        # Untuk kode 07.12, pastikan bukan sub-kode
                        if kode_found == '07.12':
                            if not re.match(r'^\d{2}\.', data_part):
                                jumlah_matches = list(re.finditer(r'Rp\s*([\d.,]+)', data_part))
                                if jumlah_matches:
                                    jumlah = FormatUtils.clean_number(jumlah_matches[-1].group(1))
                                    uraian = re.split(r'\d+\s+\w+\s+Rp|\d+\s+Rp|Rp', data_part)[0].strip()
                                    
                                    self.processed_data[kode_found] = {
                                        'uraian': uraian,
                                        'volume': "0",
                                        'satuan': "-",
                                        'harga_satuan': 0,
                                        'jumlah': jumlah
                                    }
                        else:
                            jumlah_matches = list(re.finditer(r'Rp\s*([\d.,]+)', data_part))
                            if jumlah_matches:
                                jumlah = FormatUtils.clean_number(jumlah_matches[-1].group(1))
                                uraian = re.split(r'\d+\s+\w+\s+Rp|\d+\s+Rp|Rp', data_part)[0].strip()
                                
                                self.processed_data[kode_found] = {
                                    'uraian': uraian,
                                    'volume': "0",
                                    'satuan': "-",
                                    'harga_satuan': 0,
                                    'jumlah': jumlah
                                }
            
            i += 1
    
    def calculate_total_budget(self):
        """Hitung total budget dari semua item yang diparsing"""
        if self.total_budget == 0:
            total = 0
            for kode, data in self.processed_data.items():
                total += data['jumlah']
            self.total_budget = total
    
    def get_buku_data(self):
        """Mengembalikan data BUKU (05.02.01. - 05.02.05.) hanya yang ditemukan"""
        target_codes = ['05.02.01', '05.02.02', '05.02.03', '05.02.04', '05.02.05']
        found_codes = [kode for kode in target_codes if kode in self.processed_data]
        return found_codes
    
    def get_sarana_data(self):
        """Mengembalikan data SARANA & PRASARANA hanya yang ditemukan"""
        target_codes = ['05.08.01', '05.08.08', '05.08.12']
        found_codes = [kode for kode in target_codes if kode in self.processed_data]
        return found_codes
    
    def get_honor_data(self):
        """Mengembalikan data HONOR hanya yang ditemukan"""
        found_codes = ['07.12'] if '07.12' in self.processed_data else []
        return found_codes
    
    def get_all_data(self):
        """Mengembalikan semua data yang telah diproses"""
        return self.processed_data
    
    def get_summary(self):
        """Mengembalikan ringkasan data yang telah diproses"""
        return {
            'school_name': self.school_name,
            'total_items': len(self.processed_data),
            'total_budget': self.total_budget,
            'buku_count': len(self.get_buku_data()),
            'sarana_count': len(self.get_sarana_data()),
            'honor_count': len(self.get_honor_data())
        }
    
    def clear_data(self):
        """Membersihkan semua data"""
        self.raw_data = ""
        self.processed_data = {}
        self.total_budget = 0
        self.school_name = ""

    