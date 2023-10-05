import subprocess
import re

proc = subprocess.Popen(['powershell', 'Get-WmiObject win32_desktopmonitor;'], stdout=subprocess.PIPE)
res = proc.communicate()
monitors = re.findall('(?s)\r\nName\s+:\s(.*?)\r\n', res[0].decode("utf-8"))
print(monitors)