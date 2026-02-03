import os
import time
import shutil
from datetime import datetime
import yaml
import fitz
from PIL import Image, ImageDraw, ImageFont
import io
import logging
import re

try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

logger = logging.getLogger('CertPrintAgent')


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
        self.processed_dir = os.path.join(base_dir, paths_config.get('processed', 'GetCertAgent/Processed'))
        self.cert_inbox = os.path.join(base_dir, paths_config.get('cert_inbox', 'GetCertAgent/Cert_Inbox'))
        
        for d in [self.source_cert_dir, self.annotated_dir, self.printed_dir, self.processed_dir]:
            os.makedirs(d, exist_ok=True)
    
    def get_font(self):
        """الحصول على خط مناسب"""
        font_paths = [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/tahoma.ttf",
            "arial.ttf",
        ]
        
        for font_path in font_paths:
            try:
                if os.path.exists(font_path):
                    return ImageFont.truetype(font_path, 16)
            except:
                continue
        
        return ImageFont.load_default()
    
    def parse_annotation(self, text):
        """تقسيم النص لعربي وإنجليزي"""
        # مثال: "ماهر سعد lot 2601"
        # نبحث عن "lot" ونفصل
        match = re.search(r'^(.+?)\s+lot\s+(\d+)$', text, re.IGNORECASE)
        if match:
            arabic_name = match.group(1).strip()  # ماهر سعد
            number = match.group(2).strip()       # 2601
            return arabic_name, number
        
        # لو مفيش lot، نرجع النص كله عربي
        return text, ""
    
    def create_text_image(self, text, width=220, height=40):
        """إنشاء صورة فيها النص - العربي من اليمين"""
        img = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(img)
        
        # رسم حدود
        draw.rectangle([0, 0, width-1, height-1], outline='black', width=2)
        
        font = self.get_font()
        
        # تقسيم النص
        arabic_name, number = self.parse_annotation(text)
        
        if number:
            # عندنا "lot XXXX" - نكتبهم منفصلين
            # الجزء العربي (نكتبه من اليمين للشمال)
            arabic_text = arabic_name
            english_text = f"lot {number}"
            
            # نرسم العربي على اليمين
            bbox_ar = draw.textbbox((0, 0), arabic_text, font=font)
            ar_width = bbox_ar[2] - bbox_ar[0]
            
            # نرسم الإنجليزي على الشمال
            bbox_en = draw.textbbox((0, 0), english_text, font=font)
            en_width = bbox_en[2] - bbox_en[0]
            
            # حساب المواقع
            total_width = ar_width + en_width + 10  # مسافة بينهم
            start_x = (width - total_width) // 2
            
            y = (height - (bbox_ar[3] - bbox_ar[1])) // 2 - 2
            
            # رسم العربي (على اليمين في الصورة = آخر حاجة نرسمها)
            draw.text((start_x + en_width + 10, y), arabic_text, fill='black', font=font)
            
            # رسم الإنجليزي (على الشمال)
            draw.text((start_x, y), english_text, fill='black', font=font)
        else:
            # كله عربي - في المنتصف
            bbox = draw.textbbox((0, 0), text, font=font)
            x = (width - (bbox[2] - bbox[0])) // 2
            y = (height - (bbox[3] - bbox[1])) // 2 - 2
            draw.text((x, y), text, fill='black', font=font)
        
        return img
    
    def annotate_pdf(self, pdf_path, annotation_text, output_dir=None):
        """كتابة النص على PDF"""
        try:
            logger.info(f"Annotating: {os.path.basename(pdf_path)}")
            logger.info(f"Text: {annotation_text}")
            
            if output_dir is None:
                output_dir = self.annotated_dir
            
            filename = os.path.basename(pdf_path)
            base_name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(output_dir, f"{base_name}_{timestamp}_annotated{ext}")
            
            # فتح PDF
            doc = fitz.open(pdf_path)
            page = doc[0]
            
            # إنشاء صورة النص
            text_img = self.create_text_image(annotation_text, width=220, height=40)
            
            # حفظ الصورة مؤقتاً
            img_bytes = io.BytesIO()
            text_img.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            # وضع الصورة على PDF (أعلى يمين)
            page_width = page.rect.width
            margin = 20
            box_width = 220
            box_height = 40
            
            x_pos = page_width - box_width - margin
            y_pos = margin
            
            rect = fitz.Rect(x_pos, y_pos, x_pos + box_width, y_pos + box_height)
            page.insert_image(rect, stream=img_bytes.read())
            
            # حفظ
            doc.save(output_path)
            doc.close()
            
            logger.info(f"✓ Annotated PDF saved: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Error annotating PDF: {e}")
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
            
            pdf_path = self.find_pdf_file(os.path.basename(original_pdf_path))
            if not pdf_path:
                logger.error(f"PDF not found")
                return False
            
            annotated = self.annotate_pdf(pdf_path, annotation_text)
            if not annotated:
                return False
            
            printed = False
            if self.is_printer_available():
                printed = self.print_with_retry(annotated)
            
            # نقل الملفات
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_{timestamp}.pdf"
            shutil.move(pdf_path, os.path.join(self.source_cert_dir, new_name))
            
            if printed:
                printed_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_{timestamp}_printed.pdf"
                shutil.move(annotated, os.path.join(self.printed_dir, printed_name))
            
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

