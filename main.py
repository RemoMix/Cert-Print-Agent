#!/usr/bin/env python3
"""
Cert-Print-Agent - Continuous Monitoring & Processing System
نظام مراقبة ومعالجة مستمر للشهادات
"""

import os
import sys
import time
import yaml
import shutil
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from Agents.LoggingAgent import get_logger
from Agents.OutlookAgent import OutlookAgent
from Agents.ExtractLotAgent import ExtractLotAgent  
from Agents.ERPAgent import ERPAgent
from Agents.AnnotatePrintAgent import AnnotatePrintAgent


class CertPrintOrchestrator:
    """النظام الرئيسي - شغال على طول"""
    
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.logger = get_logger(config_path)
        self.running = True
        
        # Initialize all agents
        self.outlook_agent = OutlookAgent(config_path)
        self.extract_agent = ExtractLotAgent(config_path)  # ✅ النسخة الجديدة
        self.erp_agent = ERPAgent(config_path)
        self.print_agent = AnnotatePrintAgent(config_path)
        
        # Get check interval
        self.check_interval = self.config.get('monitoring', {}).get('check_interval_minutes', 5)
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            print(f"Config error: {e}")
            return {}
    
    def process_certificates(self):
        """معالجة الشهادات من البداية للنهاية"""
        try:
            self.logger.info("\n" + "="*60)
            self.logger.info("Start new process cycle...")
            self.logger.info("="*60)
            
            # Stage 1: Extract Lot from filename (بدل OCR)
            self.logger.info("\n--- Phase 1 : Extract lot number ---")
            extraction_results = self.extract_agent.run()
            
            if not extraction_results:
                self.logger.info("No PDFs for process")
                return True
            
            self.logger.info(f"✓  {len(extraction_results)} Cert Extracted")
            
            # Stage 2: ERP Lookup
            self.logger.info("\n--- Phase 2 : search in ERP ---")
            erp_results = self.erp_agent.run(extraction_results)
            
            if not erp_results:
                self.logger.error("Error in ERP")
                return False
            
            # Stage 3: Annotate & Print
            self.logger.info("\n--- Phase 3 : Writing on Certification ---")
            print_results = self.print_agent.run(erp_results)
            
            if print_results:
                printed_count = print_results.get('printed', 0)
                not_found_count = print_results.get('not_found', 0)
                annotated_count = print_results.get('annotated_only', 0)
                self.logger.info(f"✓ Printed: {printed_count}, Not Found: {not_found_count}, تعليق فقط: {annotated_count}")
            
            # Cleanup: Move processed PDFs to Source_Cert
            self.archive_processed_pdfs()
            
            return True
            
        except Exception as e:
            self.logger.error(f"Processing Error {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False
    
    def archive_processed_pdfs(self):
        """نقل ملفات PDF المعالجة للأرشيف"""
        try:
            cert_inbox = self.config.get('paths', {}).get('cert_inbox', 'InPut/Cert_Inbox')
            source_cert = self.config.get('paths', {}).get('source_cert', 'InPut/Source_Cert')
            
            if not os.path.exists(cert_inbox):
                return
            
            os.makedirs(source_cert, exist_ok=True)
            
            pdf_files = [f for f in os.listdir(cert_inbox) if f.lower().endswith('.pdf')]
            
            for pdf_file in pdf_files:
                src = os.path.join(cert_inbox, pdf_file)
                dst = os.path.join(source_cert, pdf_file)
                
                # لو الملف موجود، ضيف timestamp
                if os.path.exists(dst):
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_")
                    dst = os.path.join(source_cert, timestamp + pdf_file)
                
                shutil.move(src, dst)
                self.logger.info(f"{pdf_file} Transfered to archive")
                
        except Exception as e:
            self.logger.error(f"Archive Error: {e}")
    
    def check_outlook(self):
        """فحص الإيميل وجلب الشهادات الجديدة"""
        try:
            self.logger.info("\n--- Check Email ---")
            new_certs = self.outlook_agent.run()
            
            if new_certs:
                self.logger.info(f"✓  {len(new_certs)} New Certificates from Email")
                return True
            else:
                self.logger.info("No New Emails")
                return False
                
        except Exception as e:
            self.logger.error(f"Check Email Error: {e}")
            return False
    
    def run_continuous(self):
        """التشغيل المستمر - شغال على طول"""
        self.logger.info("\n" + "="*60)
        self.logger.info("Cert-Print-Agent - نظام المراقبة المستمر")
        self.logger.info(" Press Ctrl+C For Stopping")
        self.logger.info("="*60)
        
        cycle_count = 0
        
        while self.running:
            cycle_count += 1
            start_time = datetime.now()
            
            try:
                self.logger.info(f"\n{'='*60}")
                self.logger.info(f"دورة #{cycle_count} - {start_time.strftime('%H:%M:%S')}")
                self.logger.info(f"{'='*60}")
                
                # 1. فحص الإيميل أولاً
                has_new_emails = self.check_outlook()
                
                # 2. معالجة الشهادات (سواء من إيميل أو موجودة)
                self.process_certificates()
                
                # 3. انتظر للدورة الجديدة
                elapsed = (datetime.now() - start_time).total_seconds()
                wait_time = max(0, (self.check_interval * 60) - elapsed)
                
                if wait_time > 0 and self.running:
                    self.logger.info(f"\n Waiting {int(wait_time)} second for next cycle...")
                    time.sleep(wait_time)
                
            except KeyboardInterrupt:
                self.logger.info("\n Stopped By User")
                self.running = False
                break
            except Exception as e:
                self.logger.error(f"Error in cycle: {e}")
                time.sleep(60)  # انتظر دقيقة لو حصل خطأ
        
        self.logger.info("\n" + "="*60)
        self.logger.info("System stoped")
        self.logger.info("="*60)
    
    def run_once(self):
        """تشغيل دورة واحدة فقط"""
        self.logger.info("\n" + "="*60)
        self.logger.info("Start one cycle")
        self.logger.info("="*60)
        
        self.check_outlook()
        self.process_certificates()


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Cert-Print-Agent')
    parser.add_argument('--config', default='config.yaml', help='ملف الإعدادات')
    parser.add_argument('--once', action='store_true', help='تشغيل دورة واحدة فقط')
    parser.add_argument('--no-outlook', action='store_true', help='تجاهل الإيميل، معالجة الملفات الموجودة فقط')
    args = parser.parse_args()
    
    orchestrator = CertPrintOrchestrator(args.config)
    
    if args.once:
        orchestrator.run_once()
    else:
        orchestrator.run_continuous()


if __name__ == "__main__":
    main()