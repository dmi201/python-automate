import os
import time

# replace with the path of the folder you want to monitor
folder_path = "C:\\path\\to\\folder\\"

while True:
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            current_modified = os.path.getmtime(file_path)
            try:
                last_modified = os.path.getmtime(file_path + ".log")
                with open(file_path + ".log", "r") as f:
                    last_action = f.read()
            except FileNotFoundError:
                last_modified = 0
                last_action = "None"

            if last_modified != current_modified:
                if last_action == "Modified":
                    print(f"{file_name} was modified")
                elif last_action == "Created":
                    print(f"{file_name} was created")
                elif last_action == "Deleted":
                    print(f"{file_name} was deleted")
                elif last_action == "Moved":
                    print(f"{file_name} was moved")

                with open(file_path + ".log", "w") as f:
                    f.write("Modified")

        else
