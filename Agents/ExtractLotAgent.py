
# Create simplified version that directly searches for lot number
import os
import re
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance
import yaml
from datetime import datetime
from pathlib import Path
import logging

try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False

logger = logging.getLogger('CertPrintAgent')

class ExtractLotAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.setup_tesseract()
        self.setup_poppler()
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def setup_tesseract(self):
        try:
            tesseract_path = self.config.get('paths', {}).get('tesseract_path', 'Tesseract/tesseract.exe')
            if os.path.exists(tesseract_path):
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                logger.info(f"Tesseract set from: {tesseract_path}")
            else:
                pytesseract.pytesseract.tesseract_cmd = 'tesseract'
        except Exception as e:
            logger.error(f"Error setting up Tesseract: {e}")
    
    def setup_poppler(self):
        try:
            poppler_path = self.config.get('paths', {}).get('poppler_path')
            if poppler_path and os.path.exists(poppler_path):
                self.poppler_path = poppler_path
                logger.info(f"Poppler set from: {poppler_path}")
            else:
                self.poppler_path = None
        except Exception as e:
            logger.error(f"Error setting up Poppler: {e}")
            self.poppler_path = None
    
    def preprocess_image(self, image):
        try:
            if CV2_AVAILABLE:
                img_array = np.array(image)
                img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
                gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
                denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                enhanced = clahe.apply(denoised)
                _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                return Image.fromarray(binary)
            else:
                gray = image.convert('L')
                enhancer = ImageEnhance.Contrast(gray)
                return enhancer.enhance(2.0)
        except Exception as e:
            logger.warning(f"Image preprocessing failed: {e}, using original")
            return image
    
    def pdf_to_images(self, pdf_path):
        try:
            logger.info(f"Converting PDF to images: {os.path.basename(pdf_path)}")
            kwargs = {'dpi': 400}
            if self.poppler_path and os.path.exists(self.poppler_path):
                kwargs['poppler_path'] = self.poppler_path
            images = convert_from_path(pdf_path, **kwargs)
            logger.info(f"Converted PDF to {len(images)} images")
            return images
        except Exception as e:
            logger.error(f"Error converting PDF: {e}")
            return []
    
    def extract_text_from_image(self, image):
        try:
            processed_img = self.preprocess_image(image)
            languages = ['eng', 'eng+ara']
            best_text = ""
            for lang in languages:
                try:
                    text = pytesseract.image_to_string(processed_img, lang=lang)
                    if len(text) > len(best_text):
                        best_text = text
                except:
                    continue
            if not best_text.strip():
                best_text = pytesseract.image_to_string(image, lang='eng')
            return best_text
        except Exception as e:
            logger.error(f"OCR error: {e}")
            return ""
    
    def extract_text_from_pdf(self, pdf_path):
        logger.info(f"Extracting text from: {os.path.basename(pdf_path)}")
        images = self.pdf_to_images(pdf_path)
        if not images:
            return ""
        all_text = ""
        for i, image in enumerate(images):
            logger.info(f"Processing page {i+1}/{len(images)}")
            text = self.extract_text_from_image(image)
            all_text += text + "\\n"
        self.save_extracted_text(pdf_path, all_text)
        return all_text
    
    def save_extracted_text(self, pdf_path, text):
        try:
            debug_dir = Path(self.config.get('paths', {}).get('base_dir', '.')) / 'debug_texts'
            debug_dir.mkdir(exist_ok=True)
            base_name = Path(pdf_path).name
            text_file = debug_dir / f"{base_name}.txt"
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write(text)
            logger.info(f"Debug text saved: {text_file}")
        except:
            pass
    
    def extract_lot_from_text(self, text):
        """Simplified lot extraction - directly search for patterns"""
        logger.info("=== SEARCHING FOR LOT NUMBER ===")
        
        # Normalize text
        text_lower = text.lower()
        
        # Pattern 1: "Lot Number : 139928" (most common)
        pattern1 = r'lot\\s+number\\s*[:：]\\s*(\\d{5,7})\\b'
        match = re.search(pattern1, text_lower)
        if match:
            lot_num = match.group(1)
            logger.info(f"✓ Found lot number (pattern 1): {lot_num}")
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
        
        # Pattern 2: "Lot: 139928"
        pattern2 = r'lot\\s*[:：]\\s*(\\d{5,7})\\b'
        match = re.search(pattern2, text_lower)
        if match:
            lot_num = match.group(1)
            logger.info(f"✓ Found lot number (pattern 2): {lot_num}")
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
        
        # Pattern 3: "Lot No: 139928" or "Lot No.: 139928"
        pattern3 = r'lot\\s+no\\.?\\s*[:：]\\s*(\\d{5,7})\\b'
        match = re.search(pattern3, text_lower)
        if match:
            lot_num = match.group(1)
            logger.info(f"✓ Found lot number (pattern 3): {lot_num}")
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
        
        # Pattern 4: "Lot# 139928"
        pattern4 = r'lot\\s*#\\s*(\\d{5,7})\\b'
        match = re.search(pattern4, text_lower)
        if match:
            lot_num = match.group(1)
            logger.info(f"✓ Found lot number (pattern 4): {lot_num}")
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
        
        logger.warning("✗ No lot number found")
        return None
    
    def extract_certification_number(self, text):
        patterns = [
            r'certificate\\s*(?:number|no\\.?)?\\s*[:：]\\s*([a-za-z]+[-–]\\d+)',
            r'cert\\.?\\s*#?\\s*[:：]?\\s*([a-za-z0-9-]+)',
            r'(dokki[-–]\\d+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                cert_num = match.group(1).strip()
                if cert_num and len(cert_num) > 3:
                    logger.info(f"Found certification number: {cert_num}")
                    return cert_num
        return "UNKNOWN"
    
    def extract_product_name(self, text):
        match = re.search(r'sample\\s*[:：]\\s*([a-za-z]{3,20})', text, re.IGNORECASE)
        if match:
            product = match.group(1).strip()
            logger.info(f"Found product name: {product}")
            return product
        return "UNKNOWN"
    
    def process_certificate(self, cert_path):
        logger.info(f"Processing certificate: {os.path.basename(cert_path)}")
        
        text = self.extract_text_from_pdf(cert_path)
        if not text:
            logger.error(f"No text extracted from: {cert_path}")
            return None
        
        lot_data = self.extract_lot_from_text(text)
        cert_number = self.extract_certification_number(text)
        product_name = self.extract_product_name(text)
        
        lot_numbers = []
        lot_info = []
        lot_type = "unknown"
        
        if lot_data:
            structured = lot_data["lot_structured"]
            lot_numbers = structured["expanded_lots"]
            lot_type = structured["type"]
            for lot in lot_numbers:
                lot_info.append({
                    "num": lot,
                    "type": structured["type"],
                    "hint": structured.get("annotation_hint")
                })
        
        result = {
            "file_path": cert_path,
            "file_name": os.path.basename(cert_path),
            "certification_number": cert_number,
            "product_name": product_name,
            "lot_numbers": lot_numbers,
            "lot_info": lot_info,
            "lot_structure": lot_type,
            "extraction_time": datetime.now().isoformat(),
        }
        
        logger.info(f"Extraction complete:")
        logger.info(f"  - Cert: {cert_number}")
        logger.info(f"  - Product: {product_name}")
        logger.info(f"  - Lots: {lot_numbers}")
        logger.info(f"  - Structure: {lot_type}")
        
        return result
    
    def run(self):
        logger.info("Starting ExtractLotAgent...")
        cert_inbox = self.config.get('paths', {}).get('cert_inbox', 'GetCertAgent/Cert_Inbox')
        
        if not os.path.exists(cert_inbox):
            logger.error(f"Folder not found: {cert_inbox}")
            return []
        
        cert_files = []
        for ext in ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.bmp']:
            cert_files.extend([
                os.path.join(cert_inbox, f)
                for f in os.listdir(cert_inbox)
                if f.lower().endswith(ext)
            ])
        
        if not cert_files:
            logger.info("No certificates found in Cert_Inbox")
            return []
        
        logger.info(f"Found {len(cert_files)} certificates to process")
        
        results = []
        for cert_path in cert_files:
            result = self.process_certificate(cert_path)
            if result:
                results.append(result)
        
        return results

def extract_lots_from_certificates(config_path="config.yaml"):
    agent = ExtractLotAgent(config_path)
    return agent.run()
