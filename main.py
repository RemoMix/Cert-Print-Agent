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
from Agents.PDFtoImageAgent import PDFtoImageAgent
from Agents.JSONExtractLotAgent import JSONExtractLotAgent
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
        self.pdf_agent = PDFtoImageAgent(config_path)
        self.json_extract_agent = JSONExtractLotAgent(config_path)
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
            self.logger.info("\\n" + "="*60)
            self.logger.info("بدء دورة معالجة جديدة...")
            self.logger.info("="*60)
            
            # Stage 1: PDF → JSON (OCR)
            self.logger.info("\\n--- المرحلة 1: استخراج النص من PDF ---")
            pdf_results = self.pdf_agent.run()
            
            if not pdf_results:
                self.logger.info("مفيش PDFs جديدة للمعالجة")
                return True
            
            self.logger.info(f"✓ تم معالجة {len(pdf_results)} PDF")
            
            # Stage 2: JSON → Extract Lot
            self.logger.info("\\n--- المرحلة 2: استخراج أرقام اللوت ---")
            extraction_results = self.json_extract_agent.run()
            
            if not extraction_results:
                self.logger.error("مفيش لوتات مستخرجة")
                return False
            
            self.logger.info(f"✓ تم استخراج {len(extraction_results)} شهادة")
            
            # Stage 3: ERP Lookup
            self.logger.info("\\n--- المرحلة 3: البحث في ERP ---")
            erp_results = self.erp_agent.run(extraction_results)
            
            if not erp_results:
                self.logger.error("فشل البحث في ERP")
                return False
            
            # Stage 4: Annotate & Print (الخطوة المهمة!)
            self.logger.info("\\n--- المرحلة 4: الكتابة على الشهادات والطباعة ---")
            print_results = self.print_agent.run(erp_results)
            
            if print_results:
                printed_count = print_results.get('printed', 0)
                annotated_count = print_results.get('annotated', 0)
                self.logger.info(f"✓ تم طباعة: {printed_count}, تعليق فقط: {annotated_count}")
            
            # Cleanup: Move processed PDFs to Source_Cert
            self.archive_processed_pdfs()
            
            return True
            
        except Exception as e:
            self.logger.error(f"خطأ في المعالجة: {e}")
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
                self.logger.info(f"تم نقل {pdf_file} للأرشيف")
                
        except Exception as e:
            self.logger.error(f"خطأ في الأرشفة: {e}")
    
    def check_outlook(self):
        """فحص الإيميل وجلب الشهادات الجديدة"""
        try:
            self.logger.info("\\n--- فحص الإيميل ---")
            new_certs = self.outlook_agent.run()
            
            if new_certs:
                self.logger.info(f"✓ تم جلب {len(new_certs)} شهادة جديدة من الإيميل")
                return True
            else:
                self.logger.info("مفيش إيميلات جديدة")
                return False
                
        except Exception as e:
            self.logger.error(f"خطأ في فحص الإيميل: {e}")
            return False
    
    def run_continuous(self):
        """التشغيل المستمر - شغال على طول"""
        self.logger.info("\\n" + "="*60)
        self.logger.info("Cert-Print-Agent - نظام المراقبة المستمر")
        self.logger.info("الضغط على Ctrl+C للإيقاف")
        self.logger.info("="*60)
        
        cycle_count = 0
        
        while self.running:
            cycle_count += 1
            start_time = datetime.now()
            
            try:
                self.logger.info(f"\\n{'='*60}")
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
                    self.logger.info(f"\\nانتظار {int(wait_time)} ثانية للدورة التالية...")
                    time.sleep(wait_time)
                
            except KeyboardInterrupt:
                self.logger.info("\\nتم إيقاف النظام بواسطة المستخدم")
                self.running = False
                break
            except Exception as e:
                self.logger.error(f"خطأ في الدورة: {e}")
                time.sleep(60)  # انتظر دقيقة لو حصل خطأ
        
        self.logger.info("\\n" + "="*60)
        self.logger.info("النظام توقف")
        self.logger.info("="*60)
    
    def run_once(self):
        """تشغيل دورة واحدة فقط"""
        self.logger.info("\\n" + "="*60)
        self.logger.info("تشغيل دورة واحدة")
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
