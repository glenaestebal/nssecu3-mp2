import subprocess   #this module will execute a child program in a new process


def display():
    print(" ___ ___  ___          _     _     ___                _   ___ _ _      ___         _        ___                      ")
    print("| _ \ _ )/ __|_ __  __| |  _| |_  | _ \___ __ ___ _ _| |_| __(_) |___ / __|__ _ __| |_  ___| _ \__ _ _ _ ___ ___ _ _ ")
    print("|   / _ \ (__| '  \/ _` | |_   _| |   / -_) _/ -_) ' \  _| _|| | / -_) (__/ _` / _| ' \/ -_)  _/ _` | '_(_-</ -_) '_|")
    print("|_|_\___/\___|_|_|_\__,_|   |_|   |_|_\___\__\___|_||_\__|_| |_|_\___|\___\__,_\__|_||_\___|_| \__,_|_| /__/\___|_|  ")
    print("\nRBCmd is Windows Recycle Bin artifact parser.")
    print("RecentFileCacheParser parses RecentFileCacheParser.bcf files.")
    print("This tool will help analyze the user's deleted files and recently accessed files.\n")
    

def run_RBCmd():
    RBCmd_process = subprocess.Popen(
        ["RBCmd.exe"],
        shell=True
    )


def run_RecentFileCacheParser():
    RecentFileCacheParser_process = subprocess.Popen(
        ["RecentFileCacheParser.exe"],
        shell=True
    )

def main():
    RBCmd_output = run_RBCmd()
    RecentFileCacheParser_output = run_RecentFileCacheParser()

    with open(output_path, 'w') as file:
        file.write(RBCmd_output)

    with open(output_path, 'w') as file:
        file.write(RecentFileCacheParser_output)

if __name__ == "__main__":
    display()
    main()

# # Check the return code
# if RBCmd_process.returncode != 0:
#     print(f"Command failed with return code {RBCmd_process.returncode}")
