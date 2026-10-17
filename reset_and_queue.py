import urllib.request
import json
import subprocess

# 1. Pulizia tasks vecchi direttamente da Mongo
clear_cmd = "docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval \"db.getSiblingDB('genieacs').tasks.deleteMany({})\""
subprocess.run(clear_cmd, shell=True, check=True)
print("[+] Coda tasks azzerata su MongoDB")

# 2. Accoda refresh per TR-181 e TR-098
device_param = "BC9424-TCL%25205G%2520CPE-869299070835408"
targets = ["Device.IP.Diagnostics.", "InternetGatewayDevice.DownloadDiagnostics."]

for target in targets:
    url = f"http://127.0.0.1:7547/devices/{device_param}/tasks"
    payload = json.dumps({"name": "refreshObject", "objectName": target}).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req) as r:
            print(f"[+] Task accodato per {target} (HTTP {r.status})")
    except Exception as e:
        print(f"[-] Errore per {target}: {e}")
