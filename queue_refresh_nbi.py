import urllib.request
import urllib.parse
import json

raw_id = "BC9424-TCL%205G%20CPE-869299070835408"
encoded_id = urllib.parse.quote(raw_id, safe='')
url = f"http://127.0.0.1:7547/devices/{encoded_id}/tasks"

objects = [
    "Device.IP.Diagnostics.DownloadDiagnostics.",
    "InternetGatewayDevice.DownloadDiagnostics."
]

for obj in objects:
    task = {
        "name": "refreshObject",
        "objectName": obj
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(task).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    try:
        with urllib.request.urlopen(req) as resp:
            print(f"[+] Task refreshObject accodato per: {obj} (HTTP {resp.status})")
            print("   ", resp.read().decode())
    except urllib.error.HTTPError as e:
        print(f"[-] Errore HTTP {e.code} per {obj}: {e.read().decode()}")
    except Exception as e:
        print(f"[-] Errore: {e}")
