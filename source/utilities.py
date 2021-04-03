import os


def get_path_name_ext(path):
    folder = os.path.dirname(path)
    file = os.path.basename(path)
    base, ext = os.path.splitext(file)

    return folder, base, ext