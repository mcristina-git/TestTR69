import urllib.request, urllib.parse, json

raw_id = "BC9424-TCL%205G%20CPE-869299070835408"
encoded_id = urllib.parse.quote(raw_id, safe='')
url = f"http://127.0.0.1:7547/devices/{encoded_id}/tasks"

# Chiede i parametri dell'interfaccia 1 e 2
task = {
    "name": "getParameterValues",
    "parameterNames": [
        "Device.IP.Interface.1.Name",
        "Device.IP.Interface.1.Enable",
        "Device.IP.Interface.1.IPv4Address.1.IPAddress",
        "Device.IP.Interface.2.Name",
        "Device.IP.Interface.2.Enable"
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
        print(f"[+] Task discovery interfacce accodato: HTTP {resp.status}")
except Exception as e:
    print(f"[-] Errore: {e}")
