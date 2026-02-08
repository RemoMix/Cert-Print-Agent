# LoggingAgent.py - النسخة المعدلة
import logging
import logging.handlers
import os
import yaml
from datetime import datetime

class LoggingAgent:
    _instance = None
    
    def __new__(cls, config_path="config.yaml"):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance.initialized = False
        return cls._instance
    
    def __init__(self, config_path="config.yaml"):
        if self.initialized:
            return
            
        self.config = self.load_config(config_path)
        self.logger = self.setup_logger()
        self.initialized = True
    
    def load_config(self, config_path):
        """Load configuration file"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            # Default config if file not found
            return {
                'paths': {'logs_dir': 'logs'},
                'logging': {
                    'log_file': 'logs/processing.log',
                    'level': 'INFO',
                    'max_size_mb': 10,
                    'backup_count': 5
                }
            }
    
    def setup_logger(self):
        """Setup logging system"""
        # Get logs directory with fallback to default
        try:
            logs_dir = self.config.get('paths', {}).get('logs_dir', 'logs')
        except:
            logs_dir = 'logs'
        
        # Create logs directory if it doesn't exist
        os.makedirs(logs_dir, exist_ok=True)
        
        log_file = os.path.join(logs_dir, 'processing.log')
        
        # Setup logger
        logger = logging.getLogger('CertPrintAgent')
        logger.setLevel(logging.INFO)
        
        # Prevent duplicate handlers
        if logger.handlers:
            return logger
        
        # Message format
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # File handler with log rotation
        try:
            max_bytes = self.config.get('logging', {}).get('max_size_mb', 10) * 1024 * 1024
            backup_count = self.config.get('logging', {}).get('backup_count', 5)
        except:
            max_bytes = 10 * 1024 * 1024
            backup_count = 5
            
        file_handler = logging.handlers.RotatingFileHandler(
            log_file,
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.INFO)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger
    
    def log_processing_start(self, cert_number):
        """Log start of certificate processing"""
        self.logger.info(f"Starting certificate processing: {cert_number}")
    
    def log_lot_extraction(self, lot_numbers):
        """Log extracted lot numbers"""
        self.logger.info(f"Extracted lot numbers: {lot_numbers}")
    
    def log_erp_search(self, lot_number, found, supplier=None, internal_lot=None):
        """Log ERP search results"""
        if found:
            self.logger.info(f"Lot {lot_number} found: Supplier={supplier}, Internal Lot={internal_lot}")
        else:
            self.logger.warning(f"Lot {lot_number} not found in ERP")
    
    def log_printing(self, cert_number, success, retry_count=0):
        """Log printing status"""
        if success:
            self.logger.info(f"Certificate {cert_number} printed successfully (Attempt: {retry_count + 1})")
        else:
            self.logger.error(f"Failed to print certificate {cert_number} after {retry_count + 1} attempts")
    
    def log_error(self, error_message, error_details=None):
        """Log errors"""
        self.logger.error(f"Error: {error_message}")
        if error_details:
            self.logger.error(f"Error details: {error_details}")
    
    def log_info(self, message):
        """Log informational messages"""
        self.logger.info(message)
    
    def log_warning(self, message):
        """Log warning messages"""
        self.logger.warning(message)
    
    def log_cycle_start(self):
        """Log start of processing cycle"""
        self.logger.info("=" * 50)
        self.logger.info(f"Starting processing cycle - {datetime.now()}")
        self.logger.info("=" * 50)
    
    def log_cycle_end(self):
        """Log end of processing cycle"""
        self.logger.info("=" * 50)
        self.logger.info(f"Ending processing cycle - {datetime.now()}")
        self.logger.info("=" * 50)

# Global logger instance
logger = None

def get_logger(config_path="config.yaml"):
    """Get or create global logger instance"""
    global logger
    if logger is None:
        agent = LoggingAgent(config_path)
        logger = agent.logger
    return logger