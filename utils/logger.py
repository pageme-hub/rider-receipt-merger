import logging
from config.settings import LOG_FILE

def setup_logger() -> logging.Logger:
    logger = logging.getLogger('excel_merger')
    logger.setLevel(logging.ERROR)
    
    formatter = logging.Formatter('[%(asctime)s] %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    
    if not logger.handlers:
        file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
        file_handler.setFormatter(formatter)
    
        logger.addHandler(file_handler)
    
    return logger

def log_error(message: str) -> None:
    setup_logger().error(message)