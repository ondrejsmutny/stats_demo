from sys import executable
from os import path, environ
from datetime import datetime
from time import time
from logging import basicConfig, FileHandler, StreamHandler, DEBUG, getLogger
from PyQt6 import QtWidgets
from ui import App

APP_NAME = "StatS: 2023 čtvrtletní P 6-04 (a)"
VERSION = "1.1.0"
PY_FILENAME = "StatS_2023_quarter_P_6_04_a.py"
EXE_FILENAME = "StatS_2023_quarter_P_6_04_a.exe"

if __name__ == "__main__":
    # Getting paths
    application_path = path.dirname(executable)

    timestamp = datetime.fromtimestamp(time()).strftime("%Y-%m-%d_%H-%M-%S")
    log_path = f"logs/log_{timestamp}.txt"

    pyqt_path = path.dirname(QtWidgets.__file__)
    plugin_path = path.join(pyqt_path, "Qt6\plugins")
    environ['QT_PLUGIN_PATH'] = plugin_path

    # Creating logger
    log_format = "[%(levelname)s] - %(asctime)s - %(name)s - : %(message)s in %(pathname)s:%(lineno)d"
    basicConfig(handlers=[FileHandler(log_path), StreamHandler()], level=DEBUG,
                        format=log_format)
    logger = getLogger(__name__)
    logger.info(f"Log file created with timestamp: {timestamp}\n")
    logger.info(f"App name: {APP_NAME}\n")
    logger.info(f"Version: {VERSION}\n")

    # Running application
    try:
        root = App(logger, timestamp)
        root.build("reports/2023/quarter/P_6-04_a.json", "reports/2023/quarter/P_6-04_a.pdf")

        timestamp_end = datetime.fromtimestamp(time()).strftime("%Y-%m-%d_%H-%M-%S")
        logger.info(f"Logging ended with timestamp end: {timestamp_end}\n")
    except:
        logger.exception("")
