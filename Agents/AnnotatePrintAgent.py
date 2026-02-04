import os
import io
import time
import shutil
from datetime import datetime
import yaml
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import logging

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    ARABIC_AVAILABLE = True
except ImportError:
    ARABIC_AVAILABLE = False

try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

logger = logging.getLogger('CertPrintAgent')

# ==================================================
# FONT
# ==================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT_PATH = os.path.join(BASE_DIR, "fonts", "arial.ttf")

# Register font
pdfmetrics.registerFont(TTFont("Arabic", FONT_PATH))


class AnnotatePrintAgent:
    def __init__(self, config_path="config.yaml"):
        self.config = self.load_config(config_path)
        self.printer_name = self.config.get('printing', {}).get('printer_name', '')
        self.retry_attempts = self.config.get('printing', {}).get('retry_attempts', 3)
        self.retry_delay = self.config.get('printing', {}).get('retry_delay_seconds', 10)
        self.setup_paths()
        
    def load_config(self, config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            return {}
    
    def setup_paths(self):
        base_dir = self.config.get('paths', {}).get('base_dir', '.')
        paths_config = self.config.get('paths', {})
        
        self.source_cert_dir = os.path.join(base_dir, paths_config.get('source_cert', 'GetCertAgent/Source_Cert'))
        self.annotated_dir = os.path.join(base_dir, paths_config.get('annotated_cert', 'GetCertAgent/Annotated_Certificates'))
        self.printed_dir = os.path.join(base_dir, paths_config.get('printed_cert', 'GetCertAgent/Printed_Annotated_Cert'))
        self.cert_inbox = os.path.join(base_dir, paths_config.get('cert_inbox', 'GetCertAgent/Cert_Inbox'))
        
        for d in [self.source_cert_dir, self.annotated_dir, self.printed_dir]:
            os.makedirs(d, exist_ok=True)
    
    def build_annotated_pdf(self, pdf_path, annotation_text):
        """Build annotated PDF using reportlab - نفس الكود بالظبط"""
        try:
            # تقسيم النص
            parts = annotation_text.split(' lot ')
            if len(parts) == 2:
                supplier = parts[0].strip()
                internal_lots_text = 'lot ' + parts[1].strip()
            else:
                supplier = annotation_text
                internal_lots_text = ''
            
            # قراءة PDF الأصلي
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            
            # إنشاء overlay
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            
            # تجهيز النص العربي
            if ARABIC_AVAILABLE:
                supplier_ar = get_display(arabic_reshaper.reshape(supplier))
            else:
                supplier_ar = supplier
            
            # النص الكامل
            if internal_lots_text:
                text = f"{internal_lots_text}  {supplier_ar}"
            else:
                text = supplier_ar
            
            font = "Arabic"
            size = 14
            can.setFont(font, size)
            
            # حساب عرض النص
            width = pdfmetrics.stringWidth(text, font, size)
            x = 560
            y = 810
            
            # رسم خلفية رمادية
            can.setFillColorRGB(0.85, 0.85, 0.85)
            can.rect(x - width - 12, y - 4, width + 12, size + 8, fill=1, stroke=0)
            
            # كتابة النص
            can.setFillColorRGB(0, 0, 0)
            can.drawRightString(x - 6, y, text)
            
            can.save()
            packet.seek(0)
            
            # دمج الـ overlay
            overlay = PdfReader(packet)
            page = reader.pages[0]
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
            
            # باقي الصفحات
            for i in range(1, len(reader.pages)):
                writer.add_page(reader.pages[i])
            
            # حفظ
            filename = os.path.basename(pdf_path)
            base_name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_pdf = os.path.join(self.annotated_dir, f"{base_name}_{timestamp}_ANNOTATED{ext}")
            
            with open(out_pdf, "wb") as f:
                writer.write(f)
            
            logger.info(f"✓ Annotated PDF saved: {out_pdf}")
            return out_pdf
            
        except Exception as e:
            logger.error(f"Error annotating PDF: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None
    
    def is_printer_available(self):
        if not WIN32_AVAILABLE:
            return False
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            if self.printer_name in printers:
                return True
            default = win32print.GetDefaultPrinter()
            if default:
                self.printer_name = default
                return True
            return False
        except:
            return False
    
    def print_pdf(self, pdf_path, retry=0):
        if not WIN32_AVAILABLE:
            return False
        try:
            result = win32api.ShellExecute(0, "print", pdf_path, f'/d:"{self.printer_name}"', ".", 0)
            return result > 32
        except:
            return False
    
    def print_with_retry(self, pdf_path):
        for attempt in range(self.retry_attempts):
            if self.print_pdf(pdf_path, attempt):
                return True
            time.sleep(self.retry_delay)
        return False
    
    def find_pdf_file(self, filename):
        for path in [self.cert_inbox, self.source_cert_dir, '.']:
            full = os.path.join(path, filename)
            if os.path.exists(full):
                return full
        return None
    
    def process_certificate(self, erp_result, original_pdf_path):
        try:
            cert_number = erp_result.get('cert_number', 'UNKNOWN')
            annotation_text = erp_result.get('annotation_text', '')
            
            logger.info(f"Processing: {cert_number}")
            logger.info(f"Annotation text: {annotation_text}")
            
            pdf_path = self.find_pdf_file(os.path.basename(original_pdf_path))
            if not pdf_path:
                logger.error(f"PDF not found: {original_pdf_path}")
                return False
            
            # بناء PDF مكتوب عليه
            annotated = self.build_annotated_pdf(pdf_path, annotation_text)
            if not annotated:
                return False
            
            # طباعة
            printed = False
            if self.is_printer_available():
                printed = self.print_with_retry(annotated)
            
            # نقل الملفات
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_{timestamp}.pdf"
            shutil.move(pdf_path, os.path.join(self.source_cert_dir, new_name))
            
            if printed:
                printed_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_{timestamp}_printed.pdf"
                shutil.copy(annotated, os.path.join(self.printed_dir, printed_name))
            
            return printed
            
        except Exception as e:
            logger.error(f"Error: {e}")
            return False
    
    def process_all(self, erp_results):
        logger.info(f"Processing {len(erp_results)} certificates")
        results = {'total': len(erp_results), 'printed': 0, 'annotated': 0, 'failed': 0}
        
        for erp in erp_results:
            success = self.process_certificate(erp, erp.get('file_path', ''))
            if success:
                results['printed'] += 1
            else:
                results['annotated'] += 1
        
        return results
    
    def run(self, erp_results=None):
        if not erp_results:
            return None
        return self.process_all(erp_results)


def annotate_and_print(erp_results, config_path="config.yaml"):
    agent = AnnotatePrintAgent(config_path)
    return agent.run(erp_results)
