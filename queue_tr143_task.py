import urllib.request
import urllib.parse
import json
import subprocess

# 1. Pulisce la coda tasks da MongoDB
cmd_clear = "docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval \"db.getSiblingDB('genieacs').tasks.deleteMany({})\""
subprocess.run(cmd_clear, shell=True, check=True)
print("[+] Coda tasks azzerata su MongoDB.")

# 2. Prepara la richiesta NBI
device_id = "BC9424-TCL%205G%20CPE-869299070835408"
encoded_id = urllib.parse.quote(device_id, safe='')
url = f"http://127.0.0.1:7547/devices/{encoded_id}/tasks"

task = {
    "name": "getParameterNames",
    "parameterPath": "Device.IP.Diagnostics.",
    "nextLevel": False
}

req = urllib.request.Request(
    url,
    data=json.dumps(task).encode("utf-8"),
    headers={"Content-Type": "application/json"},
    method="POST"
)

try:
    with urllib.request.urlopen(req) as resp:
        print(f"[+] Task accodato via NBI (HTTP {resp.status})")
        print(resp.read().decode())
except urllib.error.HTTPError as e:
    print(f"[-] Errore HTTP NBI {e.code}: {e.read().decode()}")
except Exception as e:
    print(f"[-] Errore: {e}")
