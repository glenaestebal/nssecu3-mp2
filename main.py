import subprocess   # this module will execute a child program in a new process
import csv
import os
import shutil


def display():
    print(" ___ ___  ___          _     _     ___                _   ___ _ _      ___         _        ___                      ")
    print("| _ \\ _ )/ __|_ __  __| |  _| |_  | _ \\___ __ ___ _ _| |_| __(_) |___ / __|__ _ __| |_  ___| _ \\__ _ _ _ ___ ___ _ _ ")
    print("|   / _ \\ (__| '  \\/ _` | |_   _| |   / -_) _/ -_) ' \\  _| _|| | / -_) (__/ _` / _| ' \\/ -_)  _/ _` | '_(_-</ -_) '_|")
    print("|_|_\\___/\\___|_|_|_\\__,_|   |_|   |_|_\\___\\__\\___|_||_\\__|_| |_|_\\___|\\___\\__,_\\__|_||_\\___|_| \\__,_|_| /__/\\___|_|  ")
    print("\nRBCmd is Windows Recycle Bin artifact parser.")
    print("RecentFileCacheParser parses RecentFileCacheParser.bcf files.")
    print("This tool will help analyze the user's deleted files and recently accessed files.\n")
    
        
def run_RBCmd(output_folder_path):

    print("Running RBCmd.exe...")

    if os.path.exists(output_folder_path):
        shutil.rmtree(output_folder_path)
        print(f"Existing folder removed: {output_folder_path}")

    try:

        os.makedirs(output_folder_path)
        print(f"Directory created: {output_folder_path}")
        RBCmd_process = subprocess.Popen(
            ["RBCmd.exe", "-d", "C:\\$Recycle.Bin","--csv", output_folder_path],
            # stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            shell=True
        )

        RBCmd_process.communicate()

        if RBCmd_process.returncode != 0:
            print(f"Command failed with error")
    
        else:
            print(f"Output saved to {output_folder_path}.")
        
    except Exception as e:
        print("Error:", e)


# def run_RecentFileCacheParser():
#     RecentFileCacheParser_process = subprocess.Popen(
#         ["RecentFileCacheParser.exe"],
#         shell=True
#     )

def main():

    display()

    current_directory = os.getcwd()
    output_folder = "Output"
    output_folder_path = os.path.join(current_directory, output_folder)
    RBCmd_output = run_RBCmd(output_folder_path)
    # RecentFileCacheParser_output = run_RecentFileCacheParser()

    # with open(output_path, 'w') as file:
    #     file.write(RBCmd_output)

    # with open(output_path, 'w') as file:
    #     file.write(RecentFileCacheParser_output)

if __name__ == "__main__":
    main()

# # Check the return code
# if RBCmd_process.returncode != 0:
#     print(f"Command failed with return code {RBCmd_process.returncode}")
