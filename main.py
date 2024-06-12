import subprocess # this module will execute a child program in a new process
import os

RBCmd_process = subprocess.Popen(
    ["RBCmd.exe"],
    shell=True
)

# # Communicate with the process (send input and receive output)
# stdout, stderr = RBCmd.communicate()

# # print the output
# print("Output:", stdout)
# print("Errors:", stderr)

# Check the return code
if RBCmd_process.returncode != 0:
    print(f"Command failed with return code {RBCmd_process.returncode}")



