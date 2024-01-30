import os
import shutil

folder_path = r"C:\Users\Mihai\Desktop\MyFolder"
src = r"C:\Users\Mihai\Desktop\MyFolder\myExcel.xlsx"


def delete_file(src):
    os.remove(src)
    print(f"{src} was deleted.")


def on_move(src, dst):
    delete_file(src)


def on_copy(src, dst):
    delete_file(src)


def on_delete(src):
    delete_file(src)


# The following line should be replaced with actual source and destination paths
shutil.move("src", "dst", copy_function=on_copy)
