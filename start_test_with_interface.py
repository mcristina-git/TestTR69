import urllib.request
import urllib.parse
import json

raw_id = "BC9424-TCL%205G%20CPE-869299070835408"
encoded_id = urllib.parse.quote(raw_id, safe='')
url = f"http://127.0.0.1:7547/devices/{encoded_id}/tasks"

task = {
    "name": "setParameterValues",
    "parameterValues": [
        ["Device.IP.Diagnostics.DownloadDiagnostics.Interface", "", "xsd:string"],
        ["Device.IP.Diagnostics.DownloadDiagnostics.DownloadURL", "http://192.168.1.219:8080/test_50MB.bin", "xsd:string"],
        ["Device.IP.Diagnostics.DownloadDiagnostics.DiagnosticsState", "Requested", "xsd:string"]
    ]
}

req = urllib.request.Request(
    url,
    data=json.dumps(task).encode("utf-8"),
    headers={"Content-Type": "application/json"},
    method="POST"
)

try:
    with urllib.request.urlopen(req) as resp:
        print(f"[+] Task accodato (HTTP {resp.status}):")
        print("   ", resp.read().decode())
except urllib.error.HTTPError as e:
    print(f"[-] Errore HTTP {e.code}: {e.read().decode()}")
except Exception as e:
    print(f"[-] Errore: {e}")
