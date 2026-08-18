import os
import subprocess

try:
    cmd = 'wmic process get processid,commandline'
    out = subprocess.check_output(cmd, shell=True, text=True, errors='ignore')
    for line in out.splitlines():
        if 'worker.py' in line:
            parts = line.strip().split()
            pid = parts[-1]
            if pid.isdigit():
                print(f"Terminando worker PID {pid}...")
                subprocess.run(f"taskkill /F /PID {pid}", shell=True)
except Exception as e:
    print("Erro:", e)
