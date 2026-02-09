# ERPAgent.py - معالجة كل أنماط اللوت والكتابة الصحيحة على الشهادة
import pandas as pd
import os
import yaml
from datetime import datetime
import logging

logger = logging.getLogger('CertPrintAgent')

class ERPAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.excel_path = self.get_excel_path()
        self.sheets = self.config.get('excel', {}).get('sheets', ["2026", "2025", "2024", "2023"])
        self.column_names = self.config.get('excel', {}).get('columns', {
            'cert_lot': 'NO',
            'internal_lot': 'Lot Num.',
            'supplier': 'Supplier'
        })
        self.excel_cache = {}
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def get_excel_path(self):
        base_dir = self.config.get('paths', {}).get('base_dir', '.')
        erp_file = self.config.get('paths', {}).get('erp_file', 'Raw_Warehouses.xlsx')
        return os.path.join(base_dir, erp_file)
    
    def load_excel_sheet(self, sheet_name):
        try:
            if sheet_name in self.excel_cache:
                return self.excel_cache[sheet_name]
            
            logger.info(f"Loading sheet: {sheet_name}")
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            df.columns = df.columns.str.strip()
            
            required = [
                self.column_names.get('cert_lot', 'NO'),
                self.column_names.get('internal_lot', 'Lot Num.'),
                self.column_names.get('supplier', 'Supplier')
            ]
            
            missing = [col for col in required if col not in df.columns]
            if missing:
                logger.warning(f"Missing columns in {sheet_name}: {missing}")
                return None
            
            cert_col = self.column_names.get('cert_lot', 'NO')
            # تنظيف الأرقام
            df[cert_col] = df[cert_col].astype(str).str.strip().str.replace('.0', '', regex=False)
            
            self.excel_cache[sheet_name] = df
            logger.info(f"Sheet {sheet_name} loaded: {len(df)} rows")
            return df
            
        except Exception as e:
            logger.error(f"Error loading sheet {sheet_name}: {e}")
            return None
    
    def search_lot_in_sheet(self, cert_lot_number, sheet_name):
        try:
            df = self.load_excel_sheet(sheet_name)
            if df is None:
                return None
            
            cert_col = self.column_names.get('cert_lot', 'NO')
            lot_str = str(cert_lot_number).strip()
            
            # بحث مطابق تماماً
            matches = df[df[cert_col] == lot_str]
            
            if not matches.empty:
                return matches.iloc[0]
            
            # بحث بدون أصفار في البداية
            matches = df[df[cert_col].str.lstrip('0') == lot_str.lstrip('0')]
            if not matches.empty:
                return matches.iloc[0]
            
            return None
            
        except Exception as e:
            logger.error(f"Error searching in {sheet_name}: {e}")
            return None
    
    def search_lot(self, cert_lot_number):
        """البحث عن لوت واحد في كل الشيتات"""
        logger.info(f"Searching ERP for lot: {cert_lot_number}")
        
        result = {
            'cert_lot': cert_lot_number,
            'found': False,
            'supplier': None,
            'internal_lot': None,
            'sheet_found': None
        }
        
        for sheet in self.sheets:
            row = self.search_lot_in_sheet(cert_lot_number, sheet)
            if row is not None:
                result['found'] = True
                result['supplier'] = str(row[self.column_names.get('supplier', 'Supplier')])
                result['internal_lot'] = str(row[self.column_names.get('internal_lot', 'Lot Num.')])
                result['sheet_found'] = sheet
                logger.info(f"Found in {sheet}: Supplier={result['supplier']}, Internal Lot={result['internal_lot']}")
                return result
        
        logger.warning(f"Lot {cert_lot_number} not found in ERP")
        return result
    
    # ERPAgent.py - التأكد من البحث عن كل الأرقام
    def search_multiple_lots(self, extraction_result):
        """البحث عن كل الأرقام"""
        lot_numbers = extraction_result.get('lot_numbers', [])
        lot_info_list = extraction_result.get('lot_info', [])
        
        if not lot_numbers:
            return []
        
        logger.info(f"Searching for {len(lot_numbers)} lot(s): {lot_numbers}")
        results = []
        
        for i, lot_num in enumerate(lot_numbers):
            logger.info(f"Searching for lot {i+1}/{len(lot_numbers)}: {lot_num}")
            result = self.search_lot(lot_num)
            
            # إضافة معلومات إضافية من lot_info
            if i < len(lot_info_list):
                info = lot_info_list[i]
                result['type'] = info.get('type', 'single')
                result['annotation_hint'] = info.get('annotation_hint')
                result['count'] = info.get('count', 1)
            
            results.append(result)
            logger.info(f"Result for {lot_num}: found={result['found']}, supplier={result.get('supplier')}, internal_lot={result.get('internal_lot')}")
        
        return results
    
    def generate_annotation_text(self, lot_results, annotation_hint=None):
        """
        توليد نص التعليق للكتابة على الشهادة
        """
        if not lot_results:
            return "غير مسجل في النظام"
        
        # لو مفيش ولا واحد موجود
        if not any(r.get('found') for r in lot_results):
            lots_str = " - ".join([r['cert_lot'] for r in lot_results])
            return f"غير مسجل في النظام / {lots_str} N/A"
        
        # تجميع النتائج حسب المورد
        supplier_groups = {}
        not_found_lots = []
        
        for result in lot_results:
            if result.get('found'):
                supplier = result.get('supplier', '').strip()
                internal_lot = result.get('internal_lot', '').strip()
                
                if supplier not in supplier_groups:
                    supplier_groups[supplier] = []
                supplier_groups[supplier].append(internal_lot)
            else:
                not_found_lots.append(result['cert_lot'])
        
        # بناء النص
        parts = []
        
        # الحالة 1: مورد واحد (مع implicit multi مثل +1, +2)
        if len(supplier_groups) == 1:
            supplier = list(supplier_groups.keys())[0]
            lots = supplier_groups[supplier]
            
            if len(lots) == 1:
                # لوط واحد أو implicit multi
                lot_text = lots[0]
                if annotation_hint:
                    lot_text = f"{lot_text} {annotation_hint}"
                parts.append(f"{supplier} - Lot  {lot_text}")
            else:
                # لوطين显式 (مثل 2601 و 2602)
                lots_text = " - Lot  ".join(lots)
                parts.append(f"{supplier} - Lot {lots_text}")
        
        # الحالة 2: موردين مختلفين
        else:
            for supplier, lots in supplier_groups.items():
                if len(lots) == 1:
                    parts.append(f"{supplier} - Lot  {lots[0]}")
                else:
                    lots_text = " - Lot  ".join(lots)
                    parts.append(f"{supplier} - Lot {lots_text}")
        
        # إضافة اللوتات اللي ملقتش
        if not_found_lots:
            parts.append(f"{' - '.join(not_found_lots)} N/A")
        
        return " | ".join(parts) if len(parts) > 1 else parts[0]
    
    def process_certificate(self, extraction_result):
        cert_number = extraction_result.get('certification_number', 'UNKNOWN')
        logger.info(f"Processing cert: {cert_number}")
        
        # البحث عن كل الأرقام
        lot_results = self.search_multiple_lots(extraction_result)
        
        # التحقق من النتائج
        found_count = sum(1 for r in lot_results if r.get('found'))
        total_lots = len(lot_results)
        
        # توليد النص
        annotation_hint = extraction_result.get('annotation_hint')
        annotation_text = self.generate_annotation_text(lot_results, annotation_hint)
        
        result = {
            'cert_number': cert_number,
            'file_path': extraction_result.get('file_path', ''),
            'file_name': extraction_result.get('file_name', ''),
            'product': extraction_result.get('product_name', 'UNKNOWN'),
            'lot_results': lot_results,
            'annotation_text': annotation_text,
            'all_found': found_count == total_lots and total_lots > 0,
            'partial_found': found_count > 0,
            'found_count': found_count,
            'total_lots': total_lots,
            'processing_time': datetime.now().isoformat()
        }
        
        logger.info(f"ERP complete: {found_count}/{total_lots} found")
        logger.info(f"Annotation: {annotation_text}")
        
        return result
    
    def process_all(self, extraction_results):
        logger.info(f"Processing {len(extraction_results)} certificates")
        return [self.process_certificate(ext) for ext in extraction_results]
    
    def run(self, extraction_results=None):
        logger.info("Starting ERPAgent...")
        if extraction_results:
            return self.process_all(extraction_results)
        return []

def process_erp_data(extraction_results, config_path="config.yaml"):
    agent = ERPAgent(config_path)
    return agent.run(extraction_results)