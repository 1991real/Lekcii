import wmi

import subprocess
import os
import time
# Path relative to the Python script
ps_script = os.path.join(os.path.dirname(__file__), "SaveAll.ps1")

# Run PowerShell script

def on_word_close():
    subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_script])

c = wmi.WMI()

# Wait until WINWORD.EXE terminates (blocks until it happens)
watcher = c.Win32_Process.watch_for(
    notification_type="Deletion",
    Name="WINWORD.EXE"
)

try:
        while True:
            watcher()
            time.sleep(3)
            on_word_close()
except Exception as e:
    print(e)
    time.sleep(5)
