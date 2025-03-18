"""AI GENERATED LOGGING FILE"""

import logging
import os

# Define log file path
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, 'sharepoint_operations.log')

# Set up a logger
def setup_logger():
    logger = logging.getLogger('SharePointLogger')
    logger.setLevel(logging.DEBUG)
    
    # Create file handler to log to file
    file_handler = logging.FileHandler(LOG_FILE)
    file_handler.setLevel(logging.DEBUG)
    
    # Create console handler to print log messages to console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Define log format
    log_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(log_format)
    console_handler.setFormatter(log_format)
    
    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# Initialize logger
logger = setup_logger()

# Log initialization message
logger.info("SharePoint CRUD operations logging initialized.")

