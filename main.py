import subprocess   # this module will execute a child program in a new process
import csv
import os
import shutil


def display():
    print("   ___  ___  _____         __    __    ___                      __       ___                      ")
    print("  / _ \\/ _ )/ ___/_ _  ___/ / __/ /_  / _ | __ _  _______ _____/ /  ___ / _ \\___ ________ ___ ____")
    print(" / , _/ _  / /__/  ' \\/ _  / /_  __/ / __ |/  ' \\/ __/ _ `/ __/ _ \\/ -_) ___/ _ `/ __(_-</ -_) __/")
    print("/_/|_/____/\\___/_/_/_/\\_,_/   /_/   /_/ |_/_/_/_/\\__/\\_,_/\\__/_//_/\\__/_/   \\_,_/_/ /___/\\__/_/   ")

    print("\nRBCmd is Windows Recycle Bin artifact parser.")
    print("AmcacheParser parses Amcache.hve files, which contain metadata related to applications, their installation paths, and execution history.")
    print("This tool will help analyze when the user installed and deleted applications and/or files.\n")
    
        
def run_RBCmd(output_folder_path):

    print("Running RBCmd.exe...")

    # if os.path.exists(output_folder_path):
    #     shutil.rmtree(output_folder_path)

    try:

        # os.makedirs(output_folder_path)
        # print(f"Directory created: {output_folder_path}")
        RBCmd_process = subprocess.Popen(
            ["RBCmd.exe", "-d", "C:\\$Recycle.Bin","--csv", output_folder_path],
            # stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            shell=True
        )

        stdout, stderr = RBCmd_process.communicate()

        if RBCmd_process.returncode != 0:
            print(f"Command failed with error")

        else:
            print("RBCmd execution completed successfully.\n")
            
    except Exception as e:
        print("Error:", e)


def run_AmcacheParser(output_folder_path):

    print("Running AmcacheParser.exe...")

    # if os.path.exists(output_folder_path):
    #     shutil.rmtree(output_folder_path)

    try:
        # os.makedirs(output_folder_path)
        # print(f"Directory created: {output_folder_path}")
        AmcacheParser_process = subprocess.Popen(
            ["AmcacheParser.exe", "-f", "%WINDIR%\\appcompat\\Programs\\Amcache.hve","--csv", output_folder_path],
            # stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            shell=True
        )

        stdout, stderr = AmcacheParser_process.communicate()

        if AmcacheParser_process.returncode != 0:
            print(f"Command failed with error")
    
        else:
            print("AmcacheParser execution completed successfully.")
        
    except Exception as e:
        print("Error:", e)


def main():

    display()

    current_directory = os.getcwd()
    output_folder = "Output"
    output_folder_path = os.path.join(current_directory, output_folder)

    if os.path.exists(output_folder_path):
        shutil.rmtree(output_folder_path)
        print(f"Directory exists. Existing folder removed: {output_folder_path}")

    try:
        os.makedirs(output_folder_path)
        print(f"Output folder created: {output_folder_path}\n")
    except FileExistsError:
        print(f"Output folder already exists: {output_folder_path}")

    run_RBCmd(output_folder_path)
    run_AmcacheParser(output_folder_path)

if __name__ == "__main__":
    main()

