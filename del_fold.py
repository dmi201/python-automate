import os

folder_path = r"C:\Users\Mihai\Desktop\del\New folder"

for root, dirs, files in os.walk(folder_path):
    for f in files:
        os.unlink(os.path.join(root, f))
    for d in dirs:
        os.rmdir(os.path.join(root, d))

print("All contents of the folder have been deleted.")
