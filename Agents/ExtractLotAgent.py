# ExtractLotAgent.py - استخراج رقم اللوت من اسم الملف مباشرة (بدون OCR)
import os
import re
import yaml
from datetime import datetime
import logging

logger = logging.getLogger('CertPrintAgent')

class ExtractLotAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except:
            return {}
    
    def extract_lot_from_filename(self, filename):
        """استخراج رقم اللوت من اسم الملف مباشرة"""
        logger.info(f"Extracting lot from filename: {filename}")
        
        # إزالة الامتداد
        name_without_ext = os.path.splitext(filename)[0]
        
        # البحث عن رقم مكون من 5-7 أرقام (الأكتر شيوعاً في اللوتات)
        match = re.search(r'(\d{5,7})', name_without_ext)
        if match:
            lot_num = match.group(1)
            logger.info(f"✓✓✓ FOUND LOT FROM FILENAME: {lot_num} ✓✓✓")
            return {
                "lot_raw": lot_num,
                "lot_structured": {
                    "type": "single",
                    "base_lot": lot_num,
                    "count": 1,
                    "expanded_lots": [lot_num],
                    "annotation_hint": None
                }
            }
        
        logger.error(f"✗✗✗ NO LOT NUMBER FOUND IN FILENAME: {filename} ✗✗✗")
        return None
    
    def extract_product_name(self, filename):
        """استخراج اسم المنتج من اسم الملف"""
        name_without_ext = os.path.splitext(filename)[0]
        
        # قائمة المنتجات المعروفة
        products = ['Basil', 'Fennel', 'Peppermint', 'Marjoram', 'Sage', 'Thyme', 
                   'Rosemary', 'Oregano', 'Parsley', 'Cilantro', 'Dill', 'Chamomile',
                   'Hibiscus', 'Calendula', 'Lavender', 'Melissa']
        
        for product in products:
            if product.lower() in name_without_ext.lower():
                return product
        
        # لو ملقناش، خد أول كلمة قبل الرقم
        match = re.search(r'^([A-Za-z]+)', name_without_ext)
        if match:
            return match.group(1)
        
        return "UNKNOWN"
    
    def process_certificate(self, cert_path):
        logger.info(f"Processing: {os.path.basename(cert_path)}")
        
        filename = os.path.basename(cert_path)
        
        # استخراج اللوت من الاسم مباشرة (بدون OCR)
        lot_data = self.extract_lot_from_filename(filename)
        
        if not lot_data:
            logger.error("FAILED - No lot found in filename")
            return None
        
        product_name = self.extract_product_name(filename)
        
        result = {
            "file_path": cert_path,
            "file_name": filename,
            "certification_number": "UNKNOWN",
            "product_name": product_name,
            "lot_numbers": lot_data["lot_structured"]["expanded_lots"],
            "lot_info": [{"num": lot_data["lot_raw"], "type": "single"}],
            "lot_structure": "single",
            "extraction_time": datetime.now().isoformat(),
        }
        
        logger.info(f"SUCCESS: Lot={result['lot_numbers']}, Product={product_name}")
        return result
    
    def run(self):
        logger.info("=== ExtractLotAgent (Filename-Based - No OCR) ===")
        
        cert_inbox = self.config.get('paths', {}).get('cert_inbox', 'InPut/Cert_Inbox')
        
        if not os.path.exists(cert_inbox):
            logger.error(f"Inbox not found: {cert_inbox}")
            return []
        
        cert_files = [f for f in os.listdir(cert_inbox) if f.lower().endswith('.pdf')]
        logger.info(f"Found {len(cert_files)} PDF(s)")
        
        results = []
        for filename in cert_files:
            cert_path = os.path.join(cert_inbox, filename)
            result = self.process_certificate(cert_path)
            if result:
                results.append(result)
        
        logger.info(f"=== COMPLETED: {len(results)} successful ===")
        return results


def extract_lots_from_certificates(config_path="config.yaml"):
    agent = ExtractLotAgent(config_path)
    return agent.run()