# ExtractLotAgent.py - النسخة المصححة
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
    
    def extract_lot_numbers(self, lot_string):
        """
        استخراج كل أرقام اللوت من النص
        تدعم: 139385, 139912/139913, 139912-139913, 139865/2, SFP228, إلخ
        """
        lot_string = lot_string.strip()
        logger.info(f"Parsing lot string: '{lot_string}'")
        
        # تنظيف النص من علامات التنصيص الغريبة
        lot_string = lot_string.replace("'", "").replace('"', '').replace('`', '')
        
        # النمط 1: رقمين مفصولين بـ / (مثال: 139912/139913)
        if '/' in lot_string:
            parts = lot_string.split('/')
            # لو الطرفين أرقام وطولهم 5-6 أرقام، يبقى دول لوطين منفصلين
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                if 5 <= len(parts[0]) <= 6 and 5 <= len(parts[1]) <= 6:
                    logger.info(f"Found explicit multi (slash): {parts}")
                    return {
                        "type": "explicit_multi",
                        "lots": [parts[0], parts[1]],
                        "count": 2,
                        "annotation_hint": None
                    }
                # لو الطرف الثاني رقم صغير (1-9)، يبقى implicit multi
                elif 5 <= len(parts[0]) <= 6 and 1 <= len(parts[1]) <= 2 and int(parts[1]) <= 10:
                    logger.info(f"Found implicit multi: base={parts[0]}, count={parts[1]}")
                    return {
                        "type": "implicit_multi",
                        "base_lot": parts[0],
                        "lots": [parts[0]],
                        "count": int(parts[1]),
                        "annotation_hint": f"+{int(parts[1])-1}"
                    }
        
        # النمط 2: رقمين مفصولين بـ - (مثال: 139859-139860)
        if '-' in lot_string:
            parts = lot_string.split('-')
            # لو كل الأجزاء أرقام وطول كل جزء 5-6 أرقام
            if all(p.isdigit() and 5 <= len(p) <= 6 for p in parts):
                logger.info(f"Found explicit multi (dash): {parts}")
                return {
                    "type": "explicit_multi",
                    "lots": parts,
                    "count": len(parts),
                    "annotation_hint": None
                }
        
        # النمط 3: رقم واحد (مثال: 139385, 91191, SFP228, إلخ)
        # ندور على أي رقم 5-6 أرقام
        number_match = re.search(r'(\d{5,6})', lot_string)
        if number_match:
            lot_num = number_match.group(1)
            logger.info(f"Found single lot: {lot_num}")
            return {
                "type": "single",
                "lots": [lot_num],
                "count": 1,
                "annotation_hint": None
            }
        
        # النمط 4: أحرف وأرقام (مثل SFP228, DH956-TX/2025)
        if re.match(r'^[A-Za-z0-9\-\/]+$', lot_string):
            logger.info(f"Found alphanumeric lot: {lot_string}")
            return {
                "type": "single",
                "lots": [lot_string],
                "count": 1,
                "annotation_hint": None
            }
        
        logger.error(f"Could not parse lot string: '{lot_string}'")
        return None
    
    def extract_lot_from_filename(self, filename):
        """استخراج أرقام اللوت من اسم الملف"""
        logger.info(f"Extracting lot from filename: {filename}")
        
        # إزالة الامتداد
        name_without_ext = os.path.splitext(filename)[0]
        
        # البحث عن نمط "Lot Number : XXX" في اسم الملف
        # نمط مرن يقبل مسافات وعلامات مختلفة
        patterns = [
            r'Lot\s*Number\s*[:_\s-]*\s*([A-Za-z0-9\-\/]+)',  # Lot Number : 139859-139860
            r'Lot\s*[:_\s-]*\s*([A-Za-z0-9\-\/]+)',           # Lot : 139859-139860
            r'([A-Za-z]+)\s+(\d{5,6}[\/\-]\d{5,6})',          # Basil 139859-139860
            r'([A-Za-z]+)\s+(\d{5,6})',                       # Basil 139385
        ]
        
        for pattern in patterns:
            match = re.search(pattern, name_without_ext, re.IGNORECASE)
            if match:
                # لو النمط فيه مجموعتين (منتج + رقم)، خد الرقم
                if len(match.groups()) >= 2 and match.group(2):
                    lot_string = match.group(2).strip()
                else:
                    lot_string = match.group(1).strip()
                
                logger.info(f"Matched pattern '{pattern}': {lot_string}")
                
                parsed = self.extract_lot_numbers(lot_string)
                if parsed:
                    return {
                        "lot_raw": lot_string,
                        "lot_parsed": parsed
                    }
        
        # محاولة أخيرة: دور على أي رقم 5-6 أرقام في الاسم
        numbers = re.findall(r'\d{5,6}', name_without_ext)
        if numbers:
            logger.info(f"Found numbers in filename: {numbers}")
            if len(numbers) >= 2:
                # لو لقينا رقمين، يبقى explicit multi
                return {
                    "lot_raw": f"{numbers[0]}-{numbers[1]}",
                    "lot_parsed": {
                        "type": "explicit_multi",
                        "lots": [numbers[0], numbers[1]],
                        "count": 2,
                        "annotation_hint": None
                    }
                }
            else:
                # رقم واحد
                return {
                    "lot_raw": numbers[0],
                    "lot_parsed": {
                        "type": "single",
                        "lots": [numbers[0]],
                        "count": 1,
                        "annotation_hint": None
                    }
                }
        
        logger.error(f"✗✗✗ NO LOT NUMBER FOUND IN FILENAME: {filename} ✗✗✗")
        return None
    
    def extract_product_name(self, filename):
        """استخراج اسم المنتج من اسم الملف"""
        name_without_ext = os.path.splitext(filename)[0]
        
        products = ['Basil', 'Fennel', 'Peppermint', 'Marjoram', 'Sage', 'Thyme', 
                   'Rosemary', 'Oregano', 'Parsley', 'Cilantro', 'Dill', 'Chamomile',
                   'Hibiscus', 'Calendula', 'Lavender', 'Melissa']
        
        for product in products:
            if product.lower() in name_without_ext.lower():
                return product
        
        match = re.search(r'^([A-Za-z]+)', name_without_ext)
        if match:
            return match.group(1)
        
        return "UNKNOWN"
    
    def process_certificate(self, cert_path):
        logger.info(f"Processing: {os.path.basename(cert_path)}")
        
        filename = os.path.basename(cert_path)
        lot_data = self.extract_lot_from_filename(filename)
        
        if not lot_data:
            logger.error("FAILED - No lot found in filename")
            return None
        
        product_name = self.extract_product_name(filename)
        parsed = lot_data["lot_parsed"]
        
        # إعداد lot_info للـ ERP
        lot_info_list = []
        for lot in parsed["lots"]:
            lot_info_list.append({
                "num": lot,
                "type": parsed["type"],
                "base_lot": parsed.get("base_lot"),
                "count": parsed.get("count", 1),
                "annotation_hint": parsed.get("annotation_hint")
            })
        
        result = {
            "file_path": cert_path,
            "file_name": filename,
            "certification_number": "UNKNOWN",
            "product_name": product_name,
            "lot_numbers": parsed["lots"],
            "lot_info": lot_info_list,
            "lot_structure": parsed["type"],
            "total_count": parsed["count"],
            "annotation_hint": parsed.get("annotation_hint"),
            "extraction_time": datetime.now().isoformat(),
        }
        
        logger.info(f"SUCCESS: Lots={result['lot_numbers']}, Type={parsed['type']}, Product={product_name}")
        return result
    
    def run(self):
        logger.info("=== ExtractLotAgent (Filename-Based - Multi-Lot Support) ===")
        
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