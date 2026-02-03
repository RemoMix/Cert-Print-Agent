
import os
import re
import json
import yaml
from datetime import datetime
from pathlib import Path
import logging

logger = logging.getLogger('CertPrintAgent')

STOP_WORDS = {
    "kg", "weight", "variety", "phone", "fax",
    "total", "sample", "analysis", "package", "packge", 
    "size", "number", "date", "protocol", "customer", "address"
}

LOT_VALUE = re.compile(r"[A-Z0-9][A-Z0-9\\-\\/]{2,}", re.I)


class JSONExtractLotAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def load_json(self, json_path):
        """Load OCR JSON file"""
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            logger.info(f"Loaded JSON: {json_path} ({len(data)} pages)")
            return data
        except Exception as e:
            logger.error(f"Error loading JSON: {e}")
            return None
    
    def parse_lot(self, raw):
        """Parse lot number structure"""
        if "-" in raw and all(p.isdigit() for p in raw.split("-") if p):
            parts = raw.split("-")
            return {
                "type": "explicit_multi",
                "base_lot": None,
                "count": len(parts),
                "expanded_lots": parts,
                "annotation_hint": None,
            }
        
        if "/" in raw:
            parts = raw.split("/")
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                base, cnt = parts[0], int(parts[1])
                return {
                    "type": "implicit_multi",
                    "base_lot": base,
                    "count": cnt,
                    "expanded_lots": [base],
                    "annotation_hint": f"+{cnt-1}",
                }
        
        return {
            "type": "single",
            "base_lot": raw,
            "count": 1,
            "expanded_lots": [raw],
            "annotation_hint": None,
        }
    
    def extract_lot_from_text(self, text):
        """Extract lot from OCR text using the WORKING logic"""
        text = re.sub(r"\\s+", " ", text)
        text_lower = text.lower()
        
        logger.info("=== EXTRACTING LOT FROM OCR TEXT ===")
        
        # Find "lot number : XXX" pattern
        match = re.search(
            r"\\blot\\s+number\\s*[:：](.+?)(?=\\b(?:number|total|weight|variety|packge|package|size|sample|protocol|customer|address)\\b)",
            text_lower
        )
        
        if not match:
            logger.warning("✗ Pattern 'lot number :' not found")
            # Try simpler pattern
            match = re.search(r"lot\\s+number\\s*[:：]\\s*(\\d{5,7})", text_lower)
            if match:
                lot_num = match.group(1)
                logger.info(f"✓ Found (simple pattern): {lot_num}")
                return {
                    "lot_raw": lot_num,
                    "lot_structured": self.parse_lot(lot_num)
                }
            return None
        
        segment = match.group(1).strip()
        logger.info(f"Found segment: {repr(segment)}")
        
        # Extract tokens
        tokens = re.split(r"[^\\w\\/\\-]+", segment)
        logger.info(f"Tokens: {tokens}")
        
        for tok in tokens:
            tok = tok.strip()
            if not tok:
                continue
            if tok in STOP_WORDS:
                logger.info(f"Stop word: {tok}")
                break
            if LOT_VALUE.fullmatch(tok) and tok.isdigit() and len(tok) >= 5:
                logger.info(f"✓✓✓ LOT NUMBER FOUND: {tok} ✓✓✓")
                return {
                    "lot_raw": tok,
                    "lot_structured": self.parse_lot(tok)
                }
        
        logger.warning("✗ No valid lot in segment")
        return None
    
    def extract_certification_number(self, text):
        patterns = [
            r'Certificate\\s*(?:Number|No\\.?)?\\s*[:：]\\s*([A-Za-z]+[-–]\\d+)',
            r'(Dokki[-–]\\d+)',
            r'(ISM[-–]\\d+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                cert_num = match.group(1).strip()
                logger.info(f"Found cert number: {cert_num}")
                return cert_num
        return "UNKNOWN"
    
    def extract_product_name(self, text):
        match = re.search(r'Sample\\s*[:：]\\s*([A-Za-z]{3,20})', text, re.IGNORECASE)
        if match:
            product = match.group(1).strip()
            logger.info(f"Found product: {product}")
            return product
        return "UNKNOWN"
    
    def process_json(self, json_path):
        """Process single JSON file"""
        logger.info(f"\\n{'='*60}")
        logger.info(f"Processing JSON: {os.path.basename(json_path)}")
        logger.info(f"{'='*60}")
        
        data = self.load_json(json_path)
        if not data:
            return None
        
        # Use first page (usually certificates are 1 page)
        page_data = data[0] if data else None
        if not page_data:
            logger.error("No page data in JSON")
            return None
        
        text = page_data.get("ocr_text", "")
        pdf_file = page_data.get("pdf_file", "unknown.pdf")
        
        if not text:
            logger.error("No OCR text in JSON")
            return None
        
        # Show text sample
        logger.info(f"OCR text sample: {repr(text[:300])}")
        
        # Extract data
        lot_data = self.extract_lot_from_text(text)
        cert_number = self.extract_certification_number(text)
        product_name = self.extract_product_name(text)
        
        if not lot_data:
            logger.error("✗✗✗ EXTRACTION FAILED ✗✗✗")
            return None
        
        result = {
            "file_path": page_data.get("pdf_file", ""),
            "file_name": pdf_file,
            "certification_number": cert_number,
            "product_name": product_name,
            "lot_numbers": lot_data["lot_structured"]["expanded_lots"],
            "lot_info": [{"num": lot_data["lot_raw"], "type": lot_data["lot_structured"]["type"]}],
            "lot_structure": lot_data["lot_structured"]["type"],
            "extraction_time": datetime.now().isoformat(),
        }
        
        logger.info(f"\\n✓✓✓ SUCCESS ✓✓✓")
        logger.info(f"  Lot: {result['lot_numbers']}")
        logger.info(f"  Cert: {cert_number}")
        logger.info(f"  Product: {product_name}")
        
        return result
    
    def process_all(self, json_dir=None):
        """Process all JSON files"""
        if json_dir is None:
            json_dir = Path("temp_images") / "json"
        
        json_files = list(Path(json_dir).glob("*_ocr.json"))
        logger.info(f"\\nFound {len(json_files)} JSON file(s)")
        
        results = []
        for json_file in json_files:
            result = self.process_json(str(json_file))
            if result:
                results.append(result)
        
        logger.info(f"\\n{'='*60}")
        logger.info(f"Total successful: {len(results)}/{len(json_files)}")
        logger.info(f"{'='*60}")
        
        return results
    
    def run(self):
        """Run the agent"""
        logger.info("=== JSON Extract Lot Agent ===")
        return self.process_all()


def extract_from_json(config_path="config.yaml"):
    agent = JSONExtractLotAgent(config_path)
    return agent.run()
