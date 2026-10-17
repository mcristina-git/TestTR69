import urllib.request, urllib.parse, json

raw_id = "BC9424-TCL%205G%20CPE-869299070835408"
encoded_id = urllib.parse.quote(raw_id, safe='')
url = f"http://127.0.0.1:7547/devices/{encoded_id}/tasks"

# 1. Trigger DownloadDiagnostics su Device.IP.Interface.1.
task_set = {
    "name": "setParameterValues",
    "parameterValues": [
        ["Device.IP.Diagnostics.DownloadDiagnostics.Interface", "Device.IP.Interface.1.", "xsd:string"],
        ["Device.IP.Diagnostics.DownloadDiagnostics.DownloadURL", "http://192.168.1.219:8080/test_50MB.bin", "xsd:string"],
        ["Device.IP.Diagnostics.DownloadDiagnostics.DiagnosticsState", "Requested", "xsd:string"]
    ]
}

# 2. Richiesta lettura risultati e metriche
task_get = {
    "name": "getParameterValues",
    "parameterNames": [
        "Device.IP.Diagnostics.DownloadDiagnostics.DiagnosticsState",
        "Device.IP.Diagnostics.DownloadDiagnostics.BOMTime",
        "Device.IP.Diagnostics.DownloadDiagnostics.EOMTime",
        "Device.IP.Diagnostics.DownloadDiagnostics.TestBytesReceived",
        "Device.IP.Diagnostics.DownloadDiagnostics.TotalBytesReceived"
    ]
}

for t in [task_set, task_get]:
    req = urllib.request.Request(
        url,
        data=json.dumps(t).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        print(f"[+] Task accodato ({t['name']}): HTTP {resp.status}")
