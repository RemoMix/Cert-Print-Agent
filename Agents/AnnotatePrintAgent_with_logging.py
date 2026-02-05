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

# مكتبات لتشكيل النص العربي بشكل صحيح
try:
    from arabic_reshaper import reshape
    # Try new bidi location first (v0.6+), then fall back to old location
    try:
        from bidi.bidi import get_display
    except ImportError:
        from bidi.algorithm import get_display
    ARABIC_SUPPORT = True
except ImportError:
    ARABIC_SUPPORT = False
    print("⚠️ Warning: arabic_reshaper and python-bidi not installed")
    print("Run: pip install arabic-reshaper python-bidi")

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

FONT_PATHS = [
    os.path.join(BASE_DIR, "fonts", "arial.ttf"),
    "C:/Windows/Fonts/arial.ttf",
    "C:/Windows/Fonts/tahoma.ttf",
]

FONT_PATH = None
for fp in FONT_PATHS:
    if os.path.exists(fp):
        FONT_PATH = fp
        break

if FONT_PATH:
    pdfmetrics.registerFont(TTFont("ArabicFont", FONT_PATH))
    logger.info(f"Font: {FONT_PATH}")
else:
    logger.error("No font found!")


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
    
    def prepare_arabic_text(self, text):
        """
        تحضير النص العربي للطباعة بشكل صحيح
        يربط الحروف مع بعضها ويعكس اتجاه الكتابة
        """
        if not ARABIC_SUPPORT:
            # إذا لم تكن المكتبات متوفرة، نرجع النص كما هو
            return text
        
        try:
            # تشكيل النص العربي (ربط الحروف)
            reshaped_text = reshape(text)
            # عكس اتجاه الكتابة (من اليمين لليسار)
            bidi_text = get_display(reshaped_text)
            return bidi_text
        except Exception as e:
            logger.warning(f"Error preparing Arabic text: {e}")
            return text
    
    def build_annotated_pdf(self, pdf_path, annotation_text):
        """
        بناء PDF مع التعليقات التوضيحية
        سطر واحد: اسم المورد بالعربي + رقم اللوط
        """
        try:
            # تقسيم النص
            parts = annotation_text.split(' lot ')
            if len(parts) == 2:
                arabic_name = parts[0].strip()  # مثال: عزمي ابراهيم
                lot_number = parts[1].strip()   # مثال: 2601
                # دمج الاسم واللوط في سطر واحد
                full_text = f'{arabic_name} - Lot {lot_number}'
            else:
                full_text = annotation_text
            
            # تحضير النص العربي للطباعة الصحيحة
            if ARABIC_SUPPORT:
                full_text_display = self.prepare_arabic_text(full_text)
            else:
                full_text_display = full_text
            
            # قراءة PDF الأصلي
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            
            # إنشاء طبقة التعليقات
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            
            font = "ArabicFont"
            size = 14  # حجم الخط
            can.setFont(font, size)
            
            # المواقع (أعلى يمين الصفحة)
            x_right = 560  # الحافة اليمنى
            y_position = 815  # موقع السطر الواحد
            
            # === السطر الواحد: الاسم + اللوط ===
            if full_text_display:
                # قياس عرض النص
                text_width = pdfmetrics.stringWidth(full_text_display, font, size)
                
                # رسم خلفية رمادية فاتحة
                can.setFillColorRGB(0.9, 0.9, 0.9)
                padding = 8
                can.rect(x_right - text_width - padding*2, y_position - 3, 
                        text_width + padding*2, 20, fill=1, stroke=0)
                
                # كتابة النص (من اليمين)
                can.setFillColorRGB(0, 0, 0)
                can.drawRightString(x_right - padding, y_position, full_text_display)
            
            can.save()
            packet.seek(0)
            
            # دمج التعليقات مع الصفحة الأولى
            overlay = PdfReader(packet)
            page = reader.pages[0]
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
            
            # إضافة باقي الصفحات
            for i in range(1, len(reader.pages)):
                writer.add_page(reader.pages[i])
            
            # حفظ الملف
            filename = os.path.basename(pdf_path)
            base_name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_pdf = os.path.join(self.annotated_dir, f"{base_name}_ANNOTATED{ext}")
            
            with open(out_pdf, "wb") as f:
                writer.write(f)
            
            logger.info(f"✓ Annotated: {out_pdf}")
            logger.info(f"  - Text: {full_text}")
            return out_pdf
            
        except Exception as e:
            logger.error(f"Error building annotated PDF: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None
    
    def is_printer_available(self):
        """Check if printer is available"""
        if not WIN32_AVAILABLE:
            logger.warning("⚠️ win32print not available - cannot print on this system")
            logger.warning("   Install with: pip install pywin32")
            return False
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            logger.info(f"Available printers: {printers}")
            
            if self.printer_name in printers:
                logger.info(f"✓ Using configured printer: {self.printer_name}")
                return True
            
            default = win32print.GetDefaultPrinter()
            if default:
                logger.info(f"⚠️ Configured printer not found, using default: {default}")
                self.printer_name = default
                return True
            
            logger.error("✗ No printer found!")
            return False
        except Exception as e:
            logger.error(f"✗ Error checking printer: {e}")
            return False
    
    def print_pdf(self, pdf_path, retry=0):
        """Print PDF file"""
        if not WIN32_AVAILABLE:
            logger.warning("Cannot print - win32api not available")
            return False
        try:
            logger.info(f"🖨️  Printing attempt {retry + 1}: {os.path.basename(pdf_path)}")
            logger.info(f"   Printer: {self.printer_name}")
            
            result = win32api.ShellExecute(0, "print", pdf_path, f'/d:"{self.printer_name}"', ".", 0)
            
            if result > 32:
                logger.info("✓ Print command sent successfully")
                return True
            else:
                logger.error(f"✗ Print command failed with code: {result}")
                return False
        except Exception as e:
            logger.error(f"✗ Print error: {e}")
            return False
    
    def print_with_retry(self, pdf_path):
        """Print with retry logic"""
        logger.info(f"Starting print with {self.retry_attempts} attempts, {self.retry_delay}s delay")
        
        for attempt in range(self.retry_attempts):
            if self.print_pdf(pdf_path, attempt):
                return True
            
            if attempt < self.retry_attempts - 1:
                logger.warning(f"⏳ Waiting {self.retry_delay}s before retry...")
                time.sleep(self.retry_delay)
        
        logger.error("✗ All print attempts failed")
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
            
            logger.info("=" * 60)
            logger.info(f"📄 Processing certificate: {cert_number}")
            logger.info(f"   Annotation: {annotation_text}")
            
            pdf_path = self.find_pdf_file(os.path.basename(original_pdf_path))
            if not pdf_path:
                logger.error(f"✗ PDF not found: {original_pdf_path}")
                return False
            
            logger.info(f"✓ Found PDF: {pdf_path}")
            
            # إنشاء PDF معلّم
            annotated = self.build_annotated_pdf(pdf_path, annotation_text)
            if not annotated:
                logger.error("✗ Failed to create annotated PDF")
                return False
            
            # محاولة الطباعة
            printed = False
            logger.info("🖨️  Checking printer availability...")
            
            if self.is_printer_available():
                logger.info("✓ Printer is available - starting print...")
                printed = self.print_with_retry(annotated)
                
                if printed:
                    logger.info("✅ PRINTED SUCCESSFULLY!")
                else:
                    logger.warning("⚠️ Printing failed - file annotated but not printed")
            else:
                logger.warning("⚠️ No printer available - file annotated only")
            
            # نقل الملفات
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}.pdf"
            dest_path = os.path.join(self.source_cert_dir, new_name)
            
            logger.info(f"📁 Moving source PDF to: {dest_path}")
            shutil.move(pdf_path, dest_path)
            
            if printed:
                printed_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_printed.pdf"
                printed_path = os.path.join(self.printed_dir, printed_name)
                logger.info(f"📁 Copying to printed folder: {printed_path}")
                shutil.copy(annotated, printed_path)
            
            logger.info("=" * 60)
            return printed
            
        except Exception as e:
            logger.error(f"✗ Error processing certificate: {e}")
            import traceback
            logger.error(traceback.format_exc())
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
        
        logger.info(f"Results: {results}")
        return results
    
    def run(self, erp_results=None):
        if not erp_results:
            logger.warning("No ERP results provided")
            return None
        return self.process_all(erp_results)


def annotate_and_print(erp_results, config_path="config.yaml"):
    """
    دالة مساعدة لاستدعاء الـ Agent
    """
    agent = AnnotatePrintAgent(config_path)
    return agent.run(erp_results)


# ============================================
# للاختبار المباشر
# ============================================
if __name__ == "__main__":
    # مثال للاختبار
    logging.basicConfig(level=logging.INFO)
    
    test_data = [
        {
            'cert_number': 'CERT001',
            'annotation_text': 'عزمي ابراهيم lot 2601',
            'file_path': 'test_certificate.pdf'
        }
    ]
    
    if ARABIC_SUPPORT:
        print("✓ Arabic support enabled")
    else:
        print("⚠️ Arabic support disabled - install required packages:")
        print("  pip install arabic-reshaper python-bidi")
    
    agent = AnnotatePrintAgent()
    # agent.run(test_data)
