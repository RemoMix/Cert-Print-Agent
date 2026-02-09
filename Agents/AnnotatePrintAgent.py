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
import subprocess

# مكتبات لتشكيل النص العربي بشكل صحيح
try:
    from arabic_reshaper import reshape
    try:
        from bidi.bidi import get_display
    except ImportError:
        from bidi.algorithm import get_display
    ARABIC_SUPPORT = True
except ImportError:
    ARABIC_SUPPORT = False
    print("⚠️ Warning: arabic_reshaper and python-bidi not installed")

try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("⚠️ Warning: pywin32 not installed. Printing will be disabled.")

logger = logging.getLogger('CertPrintAgent')

# ==================================================
# FONT
# ==================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FONT_PATHS = [
    os.path.join(BASE_DIR, "fonts", "arial.ttf"),
    "C:/Windows/Fonts/arial.ttf",
    "C:/Windows/Fonts/tahoma.ttf",
    "C:/Windows/Fonts/arialuni.ttf",
]

FONT_PATH = None
for fp in FONT_PATHS:
    if os.path.exists(fp):
        FONT_PATH = fp
        break

if FONT_PATH:
    try:
        pdfmetrics.registerFont(TTFont("ArabicFont", FONT_PATH))
        logger.info(f"Font loaded: {FONT_PATH}")
    except Exception as e:
        logger.error(f"Error loading font: {e}")
else:
    logger.error("No Arabic font found!")


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
        
        self.source_cert_dir = os.path.join(base_dir, paths_config.get('source_cert', 'InPut/Source_Cert'))
        self.annotated_dir = os.path.join(base_dir, paths_config.get('annotated_cert', 'OutPut/Annotated_Certificates'))
        self.printed_dir = os.path.join(base_dir, paths_config.get('printed_cert', 'OutPut/Printed_Annotated_Cert'))
        self.cert_inbox = os.path.join(base_dir, paths_config.get('cert_inbox', 'InPut/Cert_Inbox'))
        # مجلد جديد للشهادات اللي ملقتش في Excel
        self.not_found_dir = os.path.join(base_dir, 'OutPut', 'Not_Founded_In_Excel')
        
        for d in [self.source_cert_dir, self.annotated_dir, self.printed_dir, self.not_found_dir]:
            os.makedirs(d, exist_ok=True)
    
    def prepare_arabic_text(self, text):
        """تحضير النص العربي للطباعة بشكل صحيح"""
        if not ARABIC_SUPPORT:
            return text
        
        try:
            reshaped_text = reshape(text)
            bidi_text = get_display(reshaped_text)
            return bidi_text
        except Exception as e:
            logger.warning(f"Error preparing Arabic text: {e}")
            return text
    
    # AnnotatePrintAgent.py - تحسين الكتابة على PDF
    def build_annotated_pdf(self, pdf_path, annotation_text, is_not_found=False):
        """بناء PDF مع التعليقات التوضيحية"""
        try:
            logger.info(f"Building annotated PDF for: {os.path.basename(pdf_path)}")
            logger.info(f"Annotation text: {annotation_text}")
            
            if ARABIC_SUPPORT:
                full_text_display = self.prepare_arabic_text(annotation_text)
            else:
                full_text_display = annotation_text
            
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            
            # الإعدادات
            font = "ArabicFont" if FONT_PATH else "Helvetica"
            size = 17
            can.setFont(font, size)
            
            # حساب عرض النص
            try:
                text_width = pdfmetrics.stringWidth(full_text_display, font, size)
            except:
                text_width = len(full_text_display) * 10
            
            # موضع الكتابة (أعلى اليمين)
            x_right = 580
            y_position = 820
            
            # خلفية رمادية
            can.setFillColorRGB(0.6, 0.6, 0.6)
            padding = 10
            can.rect(x_right - text_width - padding*2, y_position - 5, 
                    text_width + padding*2, 30, fill=1, stroke=0)
            
            # النص بالأسود
            can.setFillColorRGB(0, 0, 0)
            can.drawRightString(x_right - padding, y_position, full_text_display)
            
            can.save()
            packet.seek(0)
            
            overlay = PdfReader(packet)
            page = reader.pages[0]
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
            
            for i in range(1, len(reader.pages)):
                writer.add_page(reader.pages[i])
            
            filename = os.path.basename(pdf_path)
            base_name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%d-%b-%y_%H%M%S")
            
            if is_not_found:
                out_pdf = os.path.join(self.not_found_dir, f"{base_name}_NOT_FOUND_{timestamp}{ext}")
            else:
                out_pdf = os.path.join(self.annotated_dir, f"{base_name}_ANNOTATED_{timestamp}{ext}")
            
            with open(out_pdf, "wb") as f:
                writer.write(f)
            
            logger.info(f"✓ Annotated PDF created: {out_pdf}")
            return out_pdf
            
        except Exception as e:
            logger.error(f"Error building annotated PDF: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None
    
    def is_printer_available(self):
        """التحقق من توفر الطابعة"""
        if not WIN32_AVAILABLE:
            logger.warning("Win32 not available")
            return False
        
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            logger.info(f"Available printers: {printers}")
            
            if self.printer_name in printers:
                logger.info(f"Using printer: {self.printer_name}")
                return True
            
            default = win32print.GetDefaultPrinter()
            if default:
                logger.info(f"Using default printer: {default}")
                self.printer_name = default
                return True
            
            logger.error("No printer found!")
            return False
            
        except Exception as e:
            logger.error(f"Error checking printer: {e}")
            return False
    
    def get_sumatra_path(self):
        """البحث عن SumatraPDF في كل الأماكن المحتملة"""
        search_dirs = []
        
        script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        search_dirs.append(script_dir)
        
        cwd = os.getcwd()
        search_dirs.append(cwd)
        
        config_base = self.config.get('paths', {}).get('base_dir', '')
        if config_base and config_base != '.':
            search_dirs.append(config_base)
        
        env_path = os.environ.get('PATH', '')
        search_dirs.extend(env_path.split(os.pathsep))
        
        possible_names = [
            "SumatraPDF-3.5.2-64.exe",
            "SumatraPDF.exe",
        ]
        
        for directory in search_dirs:
            if not directory or not os.path.isdir(directory):
                continue
                
            for name in possible_names:
                full_path = os.path.join(directory, name)
                if os.path.exists(full_path):
                    logger.info(f"✓ Found SumatraPDF at: {full_path}")
                    return full_path
                
                for subdir in os.listdir(directory):
                    subdir_path = os.path.join(directory, subdir)
                    if os.path.isdir(subdir_path):
                        full_path = os.path.join(subdir_path, name)
                        if os.path.exists(full_path):
                            logger.info(f"✓ Found SumatraPDF at: {full_path}")
                            return full_path
        
        program_files_paths = [
            r"C:\\Program Files\\SumatraPDF\\SumatraPDF.exe",
            r"C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe",
        ]
        
        for path in program_files_paths:
            if os.path.exists(path):
                logger.info(f"✓ Found SumatraPDF at: {path}")
                return path
        
        logger.warning("✗ SumatraPDF not found in any location")
        return None
    
    def print_pdf(self, pdf_path):
        """طباعة ملف PDF"""
        logger.info(f"Attempting to print: {os.path.basename(pdf_path)}")
        
        if self.print_with_sumatra(pdf_path):
            return True
        
        if self.print_with_adobe(pdf_path):
            return True
        
        if self.print_with_default(pdf_path):
            return True
        
        logger.error("All printing methods failed")
        return False
    
    def print_with_sumatra(self, pdf_path):
        """طباعة باستخدام SumatraPDF"""
        try:
            sumatra = self.get_sumatra_path()
            
            if not sumatra:
                return False
            
            logger.info(f"Using SumatraPDF: {sumatra}")
            
            cmd = [
                sumatra,
                "-print-to", self.printer_name,
                "-print-settings", "fit",
                "-exit-when-done",
                pdf_path
            ]
            
            logger.info(f"Command: {' '.join(cmd)}")
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                timeout=60,
                shell=False
            )
            
            logger.info(f"Return code: {result.returncode}")
            
            if result.stdout:
                logger.info(f"stdout: {result.stdout.decode()}")
            if result.stderr:
                logger.warning(f"stderr: {result.stderr.decode()}")
            
            if result.returncode == 0:
                logger.info("✓ SumatraPDF print completed successfully")
                time.sleep(3)
                return True
            else:
                logger.error(f"SumatraPDF failed with code: {result.returncode}")
                return False
                
        except subprocess.TimeoutExpired:
            logger.error("SumatraPDF timeout (60 seconds)")
            return False
        except Exception as e:
            logger.error(f"SumatraPDF error: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def print_with_adobe(self, pdf_path):
        """طباعة باستخدام Adobe Acrobat"""
        try:
            adobe_paths = [
                r"C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe",
                r"C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe",
                r"C:\\Program Files\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe",
            ]
            
            for adobe in adobe_paths:
                if os.path.exists(adobe):
                    logger.info(f"Using Adobe: {adobe}")
                    
                    cmd = [adobe, "/t", pdf_path, self.printer_name]
                    logger.info(f"Command: {' '.join(cmd)}")
                    
                    subprocess.Popen(cmd)
                    
                    logger.info("Waiting 10 seconds for Adobe to print...")
                    time.sleep(10)
                    
                    logger.info("✓ Adobe print command sent")
                    return True
            
            logger.warning("Adobe not found")
            return False
            
        except Exception as e:
            logger.error(f"Adobe error: {e}")
            return False
    
    def print_with_default(self, pdf_path):
        """استخدام الطابعة الافتراضية لـ Windows"""
        try:
            if not WIN32_AVAILABLE:
                return False
            
            logger.info("Using Windows default print method...")
            
            result = win32api.ShellExecute(
                0,
                "print",
                pdf_path,
                None,
                ".",
                0
            )
            
            logger.info(f"ShellExecute result: {result}")
            
            if result > 32:
                logger.info("✓ Default print command sent")
                time.sleep(5)
                return True
            else:
                logger.error(f"ShellExecute failed with code: {result}")
                return False
                
        except Exception as e:
            logger.error(f"Default print error: {e}")
            return False
    
    def print_with_retry(self, pdf_path):
        """محاولة الطباعة مع إعادة المحاولة"""
        for attempt in range(self.retry_attempts):
            logger.info(f"Print attempt {attempt + 1}/{self.retry_attempts}")
            
            if self.print_pdf(pdf_path):
                return True
            
            if attempt < self.retry_attempts - 1:
                logger.info(f"Waiting {self.retry_delay} seconds before retry...")
                time.sleep(self.retry_delay)
        
        logger.error(f"Failed to print after {self.retry_attempts} attempts")
        return False
    
    def find_pdf_file(self, filename):
        """البحث عن ملف PDF في المسارات المختلفة"""
        search_paths = [
            self.cert_inbox,
            self.source_cert_dir,
            '.',
            os.path.join('.', 'InPut', 'Cert_Inbox'),
        ]
        
        for path in search_paths:
            if path and os.path.exists(path):
                full = os.path.join(path, filename)
                if os.path.exists(full):
                    return full
        
        return None
    
    def process_certificate(self, erp_result):
        """معالجة شهادة واحدة"""
        try:
            cert_number = erp_result.get('cert_number', 'UNKNOWN')
            annotation_text = erp_result.get('annotation_text', '')
            file_path = erp_result.get('file_path', '')
            file_name = erp_result.get('file_name', '')
            all_found = erp_result.get('all_found', False)
            lot_results = erp_result.get('lot_results', [])
            
            # التحقق لو ملقتش في Excel
            is_not_found = not all_found or not any(r.get('found') for r in lot_results)
            
            logger.info(f"\\n{'='*50}")
            logger.info(f"Processing certificate: {cert_number}")
            logger.info(f"File: {file_name}")
            logger.info(f"Annotation: {annotation_text}")
            logger.info(f"Found in Excel: {not is_not_found}")
            logger.info(f"{'='*50}")
            
            # البحث عن ملف PDF
            if file_path and os.path.exists(file_path):
                pdf_path = file_path
            else:
                pdf_path = self.find_pdf_file(file_name)
            
            if not pdf_path or not os.path.exists(pdf_path):
                logger.error(f"PDF not found: {file_name}")
                return {'success': False, 'printed': False, 'error': 'File not found'}
            
            logger.info(f"Found PDF: {pdf_path}")
            
            # إنشاء PDF مكتوب عليه
            annotated_path = self.build_annotated_pdf(pdf_path, annotation_text, is_not_found)
            if not annotated_path:
                logger.error("Failed to create annotated PDF")
                return {'success': False, 'printed': False, 'error': 'Annotation failed'}
            
            # لو ملقتش في Excel، ماتطبعش، حفظ بس
            if is_not_found:
                logger.warning(f"⚠ Certificate NOT FOUND in Excel - saved to: {self.not_found_dir}")
                
                # نقل الملف الأصلي للأرشيف برضه
                try:
                    if os.path.exists(pdf_path):
                        archive_name = os.path.basename(pdf_path)
                        archive_path = os.path.join(self.source_cert_dir, archive_name)
                        
                        if os.path.exists(archive_path):
                            timestamp = datetime.now().strftime("%d-%b-%y_%H%M%S")
                            archive_path = os.path.join(self.source_cert_dir, f"{timestamp}_{archive_name}")
                        
                        shutil.move(pdf_path, archive_path)
                        logger.info(f"Archived original to: {archive_path}")
                except Exception as e:
                    logger.error(f"Error archiving: {e}")
                
                return {
                    'success': True,
                    'printed': False,
                    'not_found': True,
                    'annotated_path': annotated_path,
                    'cert_number': cert_number
                }
            
            # محاولة الطباعة (للشهادات اللي اتلقت في Excel)
            printed = False
            if self.is_printer_available():
                printed = self.print_with_retry(annotated_path)
                if printed:
                    logger.info("✓✓✓ Certificate printed successfully! ✓✓✓")
                else:
                    logger.warning("⚠ Certificate annotated but NOT printed")
            else:
                logger.warning("⚠ Printer not available - saved for manual printing")
            
            # نقل الملف الأصلي للأرشيف
            try:
                if os.path.exists(pdf_path):
                    archive_name = os.path.basename(pdf_path)
                    archive_path = os.path.join(self.source_cert_dir, archive_name)
                    
                    if os.path.exists(archive_path):
                        timestamp = datetime.now().strftime("%d-%b-%y_%H%M%S")
                        archive_path = os.path.join(self.source_cert_dir, f"{timestamp}_{archive_name}")
                    
                    shutil.move(pdf_path, archive_path)
                    logger.info(f"Archived original to: {archive_path}")
            except Exception as e:
                logger.error(f"Error archiving: {e}")
            
            # لو اتطبع، انسخ للمجلد المطبوع
            if printed:
                try:
                    printed_name = os.path.basename(annotated_path).replace('_ANNOTATED_', '_PRINTED_')
                    printed_path = os.path.join(self.printed_dir, printed_name)
                    shutil.copy2(annotated_path, printed_path)
                    logger.info(f"Copied to printed folder: {printed_path}")
                except Exception as e:
                    logger.error(f"Error copying to printed folder: {e}")
            
            return {
                'success': True,
                'printed': printed,
                'not_found': False,
                'annotated_path': annotated_path,
                'cert_number': cert_number
            }
            
        except Exception as e:
            logger.error(f"Error processing certificate: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return {'success': False, 'printed': False, 'error': str(e)}
    
    def process_all(self, erp_results):
        """معالجة جميع الشهادات"""
        if not erp_results:
            logger.warning("No ERP results to process")
            return None
        
        logger.info(f"\\n{'='*60}")
        logger.info(f"Processing {len(erp_results)} certificate(s)")
        logger.info(f"{'='*60}")
        
        results = {
            'total': len(erp_results),
            'printed': 0,
            'not_found': 0,
            'annotated_only': 0,
            'failed': 0,
            'details': []
        }
        
        for i, erp_result in enumerate(erp_results, 1):
            logger.info(f"\\n--- Certificate {i}/{len(erp_results)} ---")
            result = self.process_certificate(erp_result)
            results['details'].append(result)
            
            if result.get('success'):
                if result.get('not_found'):
                    results['not_found'] += 1
                elif result.get('printed'):
                    results['printed'] += 1
                else:
                    results['annotated_only'] += 1
            else:
                results['failed'] += 1
        
        logger.info(f"\\n{'='*60}")
        logger.info(f"Summary: {results['printed']} printed, {results['not_found']} not found, {results['annotated_only']} annotated, {results['failed']} failed")
        logger.info(f"{'='*60}")
        
        return results
    
    def run(self, erp_results=None):
        """تشغيل الوكيل"""
        if not erp_results:
            logger.warning("No ERP results provided")
            return None
        
        logger.info("=== Annotate & Print Agent ===")
        return self.process_all(erp_results)


def annotate_and_print(erp_results, config_path="config.yaml"):
    """دالة مساعدة"""
    agent = AnnotatePrintAgent(config_path)
    return agent.run(erp_results)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    test_data = [
        {
            'cert_number': 'CERT001',
            'annotation_text': 'عزمي ابراهيم - Lot 2601',
            'file_name': 'test_certificate.pdf',
            'file_path': 'test_certificate.pdf',
            'all_found': True,
            'lot_results': [{'found': True}]
        }
    ]
    
    agent = AnnotatePrintAgent()
    agent.run(test_data)
