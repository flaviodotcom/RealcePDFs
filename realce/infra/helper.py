import os
import sys


def resource_path(relative_path):
    current_dir = os.path.abspath(__file__)
    images_path = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))

    base_path = getattr(sys, "_MEIPASS", images_path)
    return os.path.join(base_path, relative_path)
