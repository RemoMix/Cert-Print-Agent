
#!/usr/bin/env python3
"""
Cert-Print-Agent v3.0 - Complete Certificate Processing System
New Folder Structure:
- Source_Cert: Original certificates archive
- Annotated_Certificates: PDFs with supplier/lot annotations
- Printed_Annotated_Cert: Successfully printed certificates
- Processed: Legacy processed folder
"""

import os
import sys
import time
import yaml
import schedule
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from Agents.LoggingAgent import get_logger
from Agents.OutlookAgent import OutlookAgent
from Agents.ExtractLotAgent import ExtractLotAgent
from Agents.ERPAgent import ERPAgent
from Agents.AnnotatePrintAgent import AnnotatePrintAgent

class CertPrintOrchestrator:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.logger = get_logger(config_path)
        self.setup_directories()
        
        # Initialize agents
        self.outlook_agent = None
        self.extract_agent = ExtractLotAgent(config_path)
        self.erp_agent = ERPAgent(config_path)
        self.print_agent = AnnotatePrintAgent(config_path)
        
        self.stats = {
            'total_processed': 0,
            'successful': 0,
            'failed': 0,
            'start_time': datetime.now()
        }
    
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            print(f"Config error: {e}")
            return {}
    
    def setup_directories(self):
        """Create all required directories"""
        base_dir = self.config.get('paths', {}).get('base_dir', 'Cert-Print-Agent')
        paths = self.config.get('paths', {})
        
        dirs = [
            paths.get('source_cert', 'GetCertAgent/Source_Cert'),
            paths.get('annotated_cert', 'GetCertAgent/Annotated_Certificates'),
            paths.get('printed_cert', 'GetCertAgent/Printed_Annotated_Cert'),
            paths.get('emails_dir', 'GetCertAgent/MyEmails'),
            paths.get('cert_inbox', 'GetCertAgent/Cert_Inbox'),
            paths.get('processed', 'GetCertAgent/Processed'),
            paths.get('temp_images', 'GetCertAgent/TempImages'),
            paths.get('logs_dir', 'logs'),
            'Data'
        ]
        
        for d in dirs:
            path = os.path.join(base_dir, d)
            os.makedirs(path, exist_ok=True)
    
    def log_banner(self, message):
        self.logger.info("=" * 70)
        self.logger.info(message)
        self.logger.info("=" * 70)
    
    def process_single_certificate(self, cert_path):
        """Process one certificate through all stages"""
        filename = os.path.basename(cert_path)
        self.logger.info(f"\\n{'─'*70}")
        self.logger.info(f"Processing: {filename}")
        self.logger.info(f"{'─'*70}")
        
        try:
            # Stage 1: OCR Extraction
            self.logger.info("[1/4] OCR: Extracting lot numbers...")
            extraction = self.extract_agent.process_certificate(cert_path)
            
            if not extraction:
                raise Exception("OCR extraction failed")
            
            lot_numbers = extraction.get('lot_numbers', [])
            if not lot_numbers:
                raise Exception("No lot numbers found")
            
            cert_num = extraction.get('certification_number', 'UNKNOWN')
            product = extraction.get('product_name', 'UNKNOWN')
            self.logger.info(f"✓ Cert: {cert_num} | Product: {product}")
            self.logger.info(f"✓ Lots: {lot_numbers}")
            
            # Stage 2: ERP Lookup
            self.logger.info("[2/4] ERP: Looking up supplier information...")
            erp_result = self.erp_agent.process_certificate(extraction)
            
            annotation = erp_result.get('annotation_text', '')
            found_all = erp_result.get('all_found', False)
            
            if not found_all:
                missing = [r['cert_lot'] for r in erp_result.get('lot_results', []) if not r['found']]
                self.logger.warning(f"⚠ Not found in ERP: {missing}")
            
            self.logger.info(f"✓ Annotation: {annotation}")
            
            # Stage 3: Annotate & Print
            self.logger.info("[3/4] Print: Annotating and printing...")
            print_success = self.print_agent.process_certificate(erp_result, cert_path)
            
            status = "✓ PRINTED" if print_success else "✓ ANNOTATED (print pending)"
            self.logger.info(f"{status}")
            
            # Stage 4: Archive (handled by print_agent)
            self.logger.info("[4/4] Archive: Files organized")
            
            self.stats['successful'] += 1
            return True
            
        except Exception as e:
            self.logger.error(f"✗ FAILED: {e}")
            self.stats['failed'] += 1
            return False
    
    def process_inbox(self):
        """Process all certificates in inbox"""
        cert_inbox = self.config.get('paths', {}).get('cert_inbox', 'GetCertAgent/Cert_Inbox')
        
        if not os.path.exists(cert_inbox):
            self.logger.error(f"Inbox not found: {cert_inbox}")
            return
        
        # Find all certificate files
        cert_files = []
        for ext in ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.bmp']:
            cert_files.extend([
                os.path.join(cert_inbox, f)
                for f in os.listdir(cert_inbox)
                if f.lower().endswith(ext)
            ])
        
        if not cert_files:
            self.logger.info("No certificates in inbox")
            return
        
        self.logger.info(f"\\nFound {len(cert_files)} certificate(s)")
        
        for cert_path in cert_files:
            self.process_single_certificate(cert_path)
            self.stats['total_processed'] += 1
            time.sleep(1)
    
    def check_outlook(self):
        """Check Outlook for new emails"""
        self.logger.info("\\nChecking Outlook...")
        
        try:
            if self.outlook_agent is None:
                self.outlook_agent = OutlookAgent()
            
            new_certs = self.outlook_agent.run()
            
            if new_certs:
                self.logger.info(f"✓ Downloaded {len(new_certs)} new certificate(s)")
            else:
                self.logger.info("No new emails")
                
        except Exception as e:
            self.logger.error(f"Outlook error: {e}")
    
    def run_cycle(self):
        """Run one complete cycle"""
        self.log_banner("STARTING PROCESSING CYCLE")
        
        try:
            self.check_outlook()
        except Exception as e:
            self.logger.error(f"Outlook step failed: {e}")
        
        try:
            self.process_inbox()
        except Exception as e:
            self.logger.error(f"Processing failed: {e}")
        
        # Statistics
        self.log_banner("CYCLE COMPLETE")
        self.logger.info(f"Session Statistics:")
        self.logger.info(f"  Total processed: {self.stats['total_processed']}")
        self.logger.info(f"  Successful: {self.stats['successful']}")
        self.logger.info(f"  Failed: {self.stats['failed']}")
        uptime = datetime.now() - self.stats['start_time']
        self.logger.info(f"  Uptime: {uptime}")
    
    def run_continuous(self):
        """Run continuously"""
        interval = self.config.get('monitoring', {}).get('check_interval_minutes', 5)
        
        self.log_banner(f"Cert-Print-Agent v3.0 - Running every {interval} minutes")
        self.logger.info("Press Ctrl+C to stop")
        
        schedule.every(interval).minutes.do(self.run_cycle)
        self.run_cycle()  # Run immediately
        
        try:
            while True:
                schedule.run_pending()
                time.sleep(1)
        except KeyboardInterrupt:
            self.log_banner("STOPPING")
    
    def run_once(self):
        """Run single cycle"""
        self.run_cycle()

def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Cert-Print-Agent v3.0')
    parser.add_argument('--once', action='store_true', help='Run once')
    parser.add_argument('--config', default='config.yaml', help='Config file')
    args = parser.parse_args()
    
    orchestrator = CertPrintOrchestrator(args.config)
    
    if args.once:
        orchestrator.run_once()
    else:
        orchestrator.run_continuous()

if __name__ == "__main__":
    main()
