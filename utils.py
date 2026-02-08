import sys
import os
import logging

def setup_logging():
    logging.basicConfig(
        filename='app_debug.log',
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.getcwd(), relative_path)


def months_between(m1, y1, m2, y2):
    return abs((y1 - y2) * 12 + (m1 - m2))


def is_shift_compatible(a, b, target):
    return target in (a, b)