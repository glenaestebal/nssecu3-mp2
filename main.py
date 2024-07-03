import subprocess   # this module will execute a child program in a new process
import csv
import os
import shutil
import pandas as pd


def display():
    print("   ___  ___  _____         __    __    ___                      __       ___                      ")
    print("  / _ \\/ _ )/ ___/_ _  ___/ / __/ /_  / _ | __ _  _______ _____/ /  ___ / _ \\___ ________ ___ ____")
    print(" / , _/ _  / /__/  ' \\/ _  / /_  __/ / __ |/  ' \\/ __/ _ `/ __/ _ \\/ -_) ___/ _ `/ __(_-</ -_) __/")
    print("/_/|_/____/\\___/_/_/_/\\_,_/   /_/   /_/ |_/_/_/_/\\__/\\_,_/\\__/_//_/\\__/_/   \\_,_/_/ /___/\\__/_/   ")

    print("\nRBCmd is Windows Recycle Bin artifact parser.")
    print("AmcacheParser parses Amcache.hve files, which contain metadata related to applications, their installation paths, and execution history.")
    print("This tool will help analyze when the user installed and deleted applications and/or files.")

    print("\nMake sure to run your command prompt as an Administrator or you won't see the Amcache file!\n")
    
        
def run_RBCmd(individual_output_folder_path):

    print("Running RBCmd.exe...")

    try:

        RBCmd_process = subprocess.Popen(
            ["RBCmd.exe", "-d", "C:\\$Recycle.Bin","--csv", individual_output_folder_path],
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


def run_AmcacheParser(individual_output_folder_path):

    print("Running AmcacheParser.exe...")

    try:
        AmcacheParser_process = subprocess.Popen(
            ["AmcacheParser.exe", "-f", "%WINDIR%\\appcompat\\Programs\\Amcache.hve", "-i", "--csv", individual_output_folder_path],
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

def find_csv_file(directory, pattern):
    for file in os.listdir(directory):
        if file.endswith(".csv") and pattern in file:
            return os.path.join(directory, file)
    return None


def combine_selected_columns(combined_output_folder_path):
    rb_cmd_csv = find_csv_file(individual_output_folder, "RBCmd_Output")
    amcache_csv = find_csv_file(individual_output_folder, "ProgramEntries")
    combined_excel_path = os.path.join(combined_output_folder_path, "Combined Output.csv")

    rb_cmd_cols = ['FileName', 'DeletedOn']  
    amcache_cols = ['Name', 'InstallDate'] 

    combined_data = pd.DataFrame()

    if os.path.exists(rb_cmd_csv):
        rb_cmd_df = pd.read_csv(rb_cmd_csv, usecols=rb_cmd_cols)
        combined_data = pd.concat([combined_data, rb_cmd_df], axis=1)
    else:
        print(f"RBCmd output file not found: {rb_cmd_csv}")

    if os.path.exists(amcache_csv):
        amcache_df = pd.read_csv(amcache_csv, usecols=amcache_cols)
        combined_data = pd.concat([combined_data, amcache_df], axis=1)
    else:
        print(f"AmcacheParser output file not found: {amcache_csv}")

    combined_data.to_excel(combined_excel_path, index=False)
    print(f"Combined Excel file created: {combined_excel_path}")

def main():

    display()

    current_directory = os.getcwd()
    individual_output_folder = "Individual Outputs"
    combined_output_folder = "Combined Output"
    individual_output_folder_path = os.path.join(current_directory, individual_output_folder)
    combined_output_folder_path = os.path.join(current_directory, combined_output_folder)

    if os.path.exists(individual_output_folder_path):
        shutil.rmtree(individual_output_folder_path)
        print(f"Directory exists. Existing folder removed: {individual_output_folder_path}")

    try:
        os.makedirs(individual_output_folder)
        print(f"Output folder created: {individual_output_folder_path}\n")
    except FileExistsError:
        print(f"Output folder already exists: {individual_output_folder_path}")

    run_RBCmd(individual_output_folder_path)
    run_AmcacheParser(individual_output_folder_path)


if __name__ == "__main__":
    main()

