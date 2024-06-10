import subprocess


# Run RBCmd.exe with arguments and capture output
# result = subprocess.Popen(
#     ['RBCmd'], 
#     capture_output=True, 
#     text=True, 
#     cwd='/Dependencies/RBCmd'  # Set the working directory
# )

process = subprocess.Popen(
    ['RBCmd'], 
    cwd="/Dependencies/RBCmd", 
    stdout=subprocess.PIPE, 
    stderr=subprocess.PIPE, 
    shell=True
)

# Communicate with the process (send input and receive output)
stdout, stderr = process.communicate()

# Print the output
print("Output:", stdout.decode())
print("Errors:", stderr.decode())

print(result.stdout)



