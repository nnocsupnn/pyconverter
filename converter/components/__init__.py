import logging
from .custom_formatter import CustomFormatter

def getLogger(appName):
    logger = logging.getLogger(appName)
    logger.setLevel(logging.DEBUG)

    # create console handler with a higher log level
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)

    ch.setFormatter(CustomFormatter())

    logger.addHandler(ch)
    
    return logger