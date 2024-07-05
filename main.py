import subprocess   # this module will execute a child program in a new process
import os           # module for interacting with the operating system
import shutil       # high-level file operations
import pandas as pd # a library for data manipulation and analysis
from datetime import datetime, timezone # module for working with dates and times

# display() function, for displaying the project description
def display():
    print("   ___  ___  _____         __    __    ___                      __       ___                      ")
    print("  / _ \\/ _ )/ ___/_ _  ___/ / __/ /_  / _ | __ _  _______ _____/ /  ___ / _ \\___ ________ ___ ____")
    print(" / , _/ _  / /__/  ' \\/ _  / /_  __/ / __ |/  ' \\/ __/ _ `/ __/ _ \\/ -_) ___/ _ `/ __(_-</ -_) __/")
    print("/_/|_/____/\\___/_/_/_/\\_,_/   /_/   /_/ |_/_/_/_/\\__/\\_,_/\\__/_//_/\\__/_/   \\_,_/_/ /___/\\__/_/   ")

    print("\nRBCmd is Windows Recycle Bin artifact parser.")
    print("AmcacheParser parses Amcache.hve files, which contain metadata related to applications, their installation paths, and execution history.")
    print("This tool will help analyze when the user installed and deleted applications and/or files.")

    print("\nMake sure to run your command prompt as an Administrator or you won't see the Amcache file!\n")
    
# run_RBCmd() function        
def run_RBCmd(output_folder):
    print("Running RBCmd.exe...")

    try:

        RBCmd_process = subprocess.Popen(
            ["RBCmd.exe", "-d", "C:\\$Recycle.Bin","--csv", output_folder],
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


# run_AmcacheParser() function
def run_AmcacheParser(output_folder):
    print("Running AmcacheParser.exe...")

    try:
        AmcacheParser_process = subprocess.Popen(
            ["AmcacheParser.exe", "-f", "%WINDIR%\\appcompat\\Programs\\Amcache.hve", "-i", "--csv", output_folder],
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

# find_csv_file() function
def find_csv_file(directory, pattern):
    for file in os.listdir(directory):
        if file.endswith(".csv") and pattern in file:
            return os.path.join(directory, file)
    return None

# combine_selected_columns () function
def combine_selected_columns(output_folder):
    rb_cmd_csv = find_csv_file(output_folder, "RBCmd_Output")
    amcache_program_entries_csv = find_csv_file(output_folder, "ProgramEntries")
    amcache_associated_files_csv = find_csv_file(output_folder, "AssociatedFileEntries")
    amcache_unassociated_files_csv = find_csv_file(output_folder, "UnassociatedFileEntries")
    
    combined_excel_path = os.path.join(output_folder, "Combined Output.xlsx")

    rb_cmd_cols = ['FileName', 'DeletedOn']  
    amcache_program_entries_cols = ['Name', 'Publisher', 'InstallDate', 'RegistryKeyPath', 'RootDirPath']
    amcache_associated_cols = ['ApplicationName', 'FullPath', 'Name', 'FileExtension', 'LinkDate'] 
    amcache_unassociated_cols = ['FullPath', 'Name', 'FileExtension', 'LinkDate']

    combined_data = pd.DataFrame() 

    if os.path.exists(rb_cmd_csv):
        rb_cmd_df = pd.read_csv(rb_cmd_csv, usecols=rb_cmd_cols)
        combined_data = pd.concat([combined_data, rb_cmd_df], axis=1)
    else:
        print(f"RBCmd output file not found: {rb_cmd_csv}")

    if os.path.exists(amcache_program_entries_csv):
        amcache_program_entries_df = pd.read_csv(amcache_program_entries_csv, usecols=amcache_program_entries_cols)
        combined_data = pd.concat([combined_data, amcache_program_entries_df], axis=1)
    else:
        print(f"AmcacheParser ProgramEntries output file not found: {amcache_program_entries_csv}")

    if os.path.exists(amcache_associated_files_csv):
        amcache_associated_df = pd.read_csv(amcache_associated_files_csv, usecols=amcache_associated_cols)
        combined_data = pd.concat([combined_data, amcache_associated_df], axis=1)
    else:
        print(f"AmcacheParser AssociatedFileEntries output file not found: {amcache_associated_files_csv}")

    if os.path.exists(amcache_unassociated_files_csv):
        amcache_unassociated_df = pd.read_csv(amcache_unassociated_files_csv, usecols=amcache_unassociated_cols)
        combined_data = pd.concat([combined_data, amcache_unassociated_df], axis=1)
    else:
        print(f"AmcacheParser UnassociatedFileEntries output file not found: {amcache_unassociated_files_csv}")

    combined_data = combined_data.loc[:, ~combined_data.columns.duplicated()]

    rb_cmd_df['DeletedOn'] = pd.to_datetime(rb_cmd_df['DeletedOn']).dt.tz_localize('Asia/Manila').dt.tz_convert('UTC').dt.tz_localize(None)
    amcache_program_entries_df['InstallDate'] = pd.to_datetime(amcache_program_entries_df['InstallDate']).dt.tz_localize('Asia/Manila').dt.tz_convert('UTC').dt.tz_localize(None)
    amcache_associated_df['LinkDate'] =  pd.to_datetime(amcache_associated_df['LinkDate']).dt.tz_localize('Asia/Manila').dt.tz_convert('UTC').dt.tz_localize(None)
    amcache_unassociated_df['LinkDate'] =  pd.to_datetime(amcache_unassociated_df['LinkDate']).dt.tz_localize('Asia/Manila').dt.tz_convert('UTC').dt.tz_localize(None)

    rb_cmd_df.rename(columns={'DeletedOn': 'DeletedOn (Normalized UTC+0)'}, inplace=True)
    amcache_program_entries_df.rename(columns={'InstallDate': 'InstallDate (Normalized UTC+0)'}, inplace=True)
    amcache_associated_df.rename(columns={'LinkDate': 'LinkDate (Normalized UTC+0)'}, inplace=True)
    amcache_unassociated_df.rename(columns={'LinkDate': 'LinkDate (Normalized UTC+0)'}, inplace=True)


    with pd.ExcelWriter(combined_excel_path, engine='openpyxl') as writer:
        if 'rb_cmd_df' in locals():
            rb_cmd_df.to_excel(writer, sheet_name='RBCmd', index=False)
        if 'amcache_program_entries_df' in locals():
            amcache_program_entries_df.to_excel(writer, sheet_name='Amcache Program Entries', index=False)
        if 'amcache_associated_df' in locals():
            amcache_associated_df.to_excel(writer, sheet_name='Amcache Associated Files', index=False)
        if 'amcache_unassociated_df' in locals():
            amcache_unassociated_df.to_excel(writer, sheet_name='Amcache Unassociated Files', index=False)

    print(f"Combined Excel file created: {combined_excel_path}")

# main() function
def main():
    display()

    current_directory = os.getcwd()
    output_folder = "Output"
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

