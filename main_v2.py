
#!/usr/bin/env python3
"""
Cert-Print-Agent v3.1 - Two-Stage Processing
Stage 1: PDF → Image → OCR → JSON
Stage 2: JSON → Extract Lot → ERP → Print
"""

import os
import sys
import yaml
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from Agents.LoggingAgent import get_logger
from Agents.PDFtoImageAgent import PDFtoImageAgent
from Agents.JSONExtractLotAgent import JSONExtractLotAgent
from Agents.ERPAgent import ERPAgent
from Agents.AnnotatePrintAgent import AnnotatePrintAgent

class CertPrintOrchestratorV2:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.logger = get_logger(config_path)
        
        # Initialize agents
        self.pdf_agent = PDFtoImageAgent(config_path)
        self.json_extract_agent = JSONExtractLotAgent(config_path)
        self.erp_agent = ERPAgent(config_path)
        self.print_agent = AnnotatePrintAgent(config_path)
    
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            print(f"Config error: {e}")
            return {}
    
    def run_stage1(self):
        """Stage 1: PDF → Image → OCR → JSON"""
        self.logger.info("\\n" + "="*60)
        self.logger.info("STAGE 1: PDF → Image → OCR → JSON")
        self.logger.info("="*60)
        
        results = self.pdf_agent.run()
        
        if not results:
            self.logger.error("Stage 1 failed - no PDFs processed")
            return False
        
        self.logger.info(f"✓ Stage 1 complete: {len(results)} PDF(s) → JSON")
        return True
    
    def run_stage2(self):
        """Stage 2: JSON → Extract → ERP → Print"""
        self.logger.info("\\n" + "="*60)
        self.logger.info("STAGE 2: JSON → Extract → ERP → Print")
        self.logger.info("="*60)
        
        # Step 1: Extract from JSON
        extraction_results = self.json_extract_agent.run()
        
        if not extraction_results:
            self.logger.error("Stage 2 failed - no lots extracted")
            return False
        
        self.logger.info(f"✓ Extraction complete: {len(extraction_results)} certificate(s)")
        
        # Step 2: ERP Lookup
        self.logger.info("\\n--- ERP Lookup ---")
        erp_results = self.erp_agent.run(extraction_results)
        
        # Step 3: Annotate & Print
        self.logger.info("\\n--- Annotate & Print ---")
        print_results = self.print_agent.run(erp_results)
        
        return True
    
    def run(self):
        """Run complete workflow"""
        self.logger.info("\\n" + "="*60)
        self.logger.info("Cert-Print-Agent v3.1 - Starting")
        self.logger.info("="*60)
        
        start_time = datetime.now()
        
        # Stage 1
        if not self.run_stage1():
            self.logger.error("Workflow failed at Stage 1")
            return
        
        # Stage 2
        if not self.run_stage2():
            self.logger.error("Workflow failed at Stage 2")
            return
        
        # Complete
        elapsed = datetime.now() - start_time
        self.logger.info("\\n" + "="*60)
        self.logger.info("WORKFLOW COMPLETE")
        self.logger.info(f"Elapsed time: {elapsed}")
        self.logger.info("="*60)


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Cert-Print-Agent v3.1')
    parser.add_argument('--config', default='config.yaml', help='Config file')
    parser.add_argument('--stage1-only', action='store_true', help='Only run Stage 1 (PDF→JSON)')
    parser.add_argument('--stage2-only', action='store_true', help='Only run Stage 2 (JSON→Print)')
    args = parser.parse_args()
    
    orchestrator = CertPrintOrchestratorV2(args.config)
    
    if args.stage1_only:
        orchestrator.run_stage1()
    elif args.stage2_only:
        orchestrator.run_stage2()
    else:
        orchestrator.run()


if __name__ == "__main__":
    main()
