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
    
        
def run_RBCmd(output_folder):

    print("Running RBCmd.exe...")

    try:

        RBCmd_process = subprocess.Popen(
            ["RBCmd.exe", "-d", "C:\\$Recycle.Bin","--csv", output_folder],
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


def run_AmcacheParser(output_folder):

    print("Running AmcacheParser.exe...")

    try:
        AmcacheParser_process = subprocess.Popen(
            ["AmcacheParser.exe", "-f", "%WINDIR%\\appcompat\\Programs\\Amcache.hve", "-i", "--csv", output_folder],
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


def combine_selected_columns(output_folder):
    rb_cmd_csv = find_csv_file(output_folder, "RBCmd_Output")
    amcache_csv = find_csv_file(output_folder, "ProgramEntries")

    combined_excel_path = os.path.join(output_folder, "Combined Output.xlsx")

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

    with pd.ExcelWriter(combined_excel_path, engine='openpyxl') as writer:
        rb_cmd_df.to_excel(writer, sheet_name='RBCmd', index=False)
        amcache_df.to_excel(writer, sheet_name='Amcache', index=False)
    print(f"Combined Excel file created: {combined_excel_path}")

def main():

    display()

    current_directory = os.getcwd()
    output_folder = "Output"
    # individual_output_folder_path = os.path.join(current_directory, individual_output_folder)
    output_folder_path = os.path.join(current_directory, output_folder)

    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
        print(f"Directory exists. Existing folder removed: {output_folder}")

    try:
        os.makedirs(output_folder)
        print(f"Output folder created: {output_folder}\n")
    except FileExistsError:
        print(f"Output folder already exists: {output_folder}")

    run_RBCmd(output_folder)
    run_AmcacheParser(output_folder)
    combine_selected_columns(output_folder_path)


if __name__ == "__main__":
    main()

