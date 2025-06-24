import logging
import os
import traceback
from datetime import datetime
from init.init import LOG_FOLDER

os.makedirs(LOG_FOLDER, exist_ok=True)
timestamp = datetime.now().strftime('%Y%m%d %H%M%S')

logger = logging.getLogger(__name__)
logger.setLevel(10)
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
formatter = logging.Formatter('%(asctime)s - %(message)s')

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
console_handler.setFormatter(formatter)

file_handler = logging.FileHandler(f'{LOG_FOLDER}/{timestamp}.log')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)


error_file_handler = logging.FileHandler(f'{LOG_FOLDER}/{timestamp} ERROR.log')
error_file_handler.setLevel(logging.ERROR)
error_file_handler.setFormatter(formatter)
logger.addHandler(error_file_handler)

# Sample of using
if __name__ == '__main__':
    logger.debug('This is a debug message')
    logger.info('This is an info message')
    logger.warning('This is a warning message')
    logger.error('This is an error message')
    logger.critical('This is a critical message')