import urllib.request
import urllib.parse
import json

raw_id = "BC9424-TCL%205G%20CPE-869299070835408"
encoded_id = urllib.parse.quote(raw_id, safe='')
url = f"http://127.0.0.1:7547/devices/{encoded_id}/tasks"

task = {
    "name": "refreshObject",
    "objectName": "Device.IP.Diagnostics.DownloadDiagnostics."
}

req = urllib.request.Request(
    url,
    data=json.dumps(task).encode("utf-8"),
    headers={"Content-Type": "application/json"},
    method="POST"
)

try:
    with urllib.request.urlopen(req) as resp:
        print(f"[+] Task refreshObject accodato (HTTP {resp.status})")
        print("   ", resp.read().decode())
except Exception as e:
    print(f"[-] Errore: {e}")
