import os
import signal
import subprocess

# Get the list of all running processes
process = subprocess.Popen(['tasklist'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
stdout, stderr = process.communicate()

# Parse the output to find the PID of the specific process
process_name = 'EXCEL.EXE'
lines = stdout.splitlines()
for line in lines:
    if process_name in line:
        parts = line.split()
        pid = int(parts[1])  # The PID is the second element in the split line
        break
else:
    print(f"{process_name} is not running.")
    exit()

# Terminate the process
try:
    os.kill(pid, signal.SIGTERM)
    print(f"Process {process_name} (PID: {pid}) has been terminated.")
except ProcessLookupError:
    print(f"Process with PID {pid} not found.")
except PermissionError:
    print(f"Permission denied to terminate process with PID {pid}.")
