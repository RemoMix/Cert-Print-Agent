
# Create emergency extraction using exact text match
import os
import re
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import yaml
from datetime import datetime
import logging

logger = logging.getLogger('CertPrintAgent')

class ExtractLotAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.setup_tesseract()
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except:
            return {}
    
    def setup_tesseract(self):
        try:
            tesseract_path = self.config.get('paths', {}).get('tesseract_path', 'Tesseract/tesseract.exe')
            if os.path.exists(tesseract_path):
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
        except:
            pass
    
    def extract_text_from_pdf(self, pdf_path):
        try:
            poppler_path = self.config.get('paths', {}).get('poppler_path')
            kwargs = {'dpi': 300}
            if poppler_path and os.path.exists(poppler_path):
                kwargs['poppler_path'] = poppler_path
            
            images = convert_from_path(pdf_path, **kwargs)
            all_text = ""
            for image in images:
                text = pytesseract.image_to_string(image, lang='eng')
                all_text += text + "\\n"
            return all_text
        except Exception as e:
            logger.error(f"OCR error: {e}")
            return ""
    
    def extract_lot_from_text(self, text):
        """Emergency extraction - direct pattern match"""
        text_lower = text.lower()
        
        # Direct extraction using the known pattern from debug file
        # Pattern: "Lot Number : 139928"
        match = re.search(r'lot\\s+number\\s*:\\s*(\\d{5,7})', text_lower)
        if match:
            lot_num = match.group(1)
            logger.info(f"✓✓✓ FOUND LOT NUMBER: {lot_num} ✓✓✓")
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
        
        logger.error("✗✗✗ LOT NUMBER NOT FOUND ✗✗✗")
        return None
    
    def extract_certification_number(self, text):
        match = re.search(r'(Dokki[-–]\\d+)', text, re.IGNORECASE)
        if match:
            return match.group(1)
        return "UNKNOWN"
    
    def extract_product_name(self, text):
        match = re.search(r'Sample\\s*:\\s*([A-Za-z]+)', text, re.IGNORECASE)
        if match:
            return match.group(1)
        return "UNKNOWN"
    
    def process_certificate(self, cert_path):
        logger.info(f"Processing: {os.path.basename(cert_path)}")
        
        text = self.extract_text_from_pdf(cert_path)
        if not text:
            logger.error("No text extracted")
            return None
        
        # Show first 500 chars for debug
        logger.info(f"Text sample: {repr(text[:200])}")
        
        lot_data = self.extract_lot_from_text(text)
        cert_number = self.extract_certification_number(text)
        product_name = self.extract_product_name(text)
        
        if not lot_data:
            logger.error("FAILED - No lot found")
            return None
        
        result = {
            "file_path": cert_path,
            "file_name": os.path.basename(cert_path),
            "certification_number": cert_number,
            "product_name": product_name,
            "lot_numbers": lot_data["lot_structured"]["expanded_lots"],
            "lot_info": [{"num": lot_data["lot_raw"], "type": "single"}],
            "lot_structure": "single",
            "extraction_time": datetime.now().isoformat(),
        }
        
        logger.info(f"SUCCESS: Lot={result['lot_numbers']}, Cert={cert_number}, Product={product_name}")
        return result
    
    def run(self):
        logger.info("=== STARTING EMERGENCY EXTRACTION ===")
        cert_inbox = self.config.get('paths', {}).get('cert_inbox', 'GetCertAgent/Cert_Inbox')
        
        if not os.path.exists(cert_inbox):
            logger.error("Inbox not found")
            return []
        
        cert_files = [f for f in os.listdir(cert_inbox) if f.lower().endswith('.pdf')]
        logger.info(f"Found {len(cert_files)} PDFs")
        
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
