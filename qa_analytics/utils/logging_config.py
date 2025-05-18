import os
import logging


def setup_logging(log_level=logging.INFO):
    """Set up logging configuration"""

    # Get the project root directory (parent of qa_analytics)
    project_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../'))

    # Create logs directory if it doesn't exist
    logs_dir = os.path.join(project_dir, 'logs')
    os.makedirs(logs_dir, exist_ok=True)

    # Configure log file path
    log_file = os.path.join(logs_dir, 'qa_analytics.log')

    # Set up logging
    handlers = [
        logging.StreamHandler(),
        logging.FileHandler(log_file)
    ]

    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=handlers
    )

    # Get the logger for the qa_analytics package
    logger = logging.getLogger('qa_analytics')

    return logger