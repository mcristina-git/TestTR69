import subprocess
import json

cmd = "docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval \"JSON.stringify(Object.keys(db.getSiblingDB('genieacs').devices.findOne({_id: 'BC9424-TCL%205G%20CPE-869299070835408'}) || {}))\""
res = subprocess.run(cmd, shell=True, capture_output=True, text=True)

try:
    keys = json.loads(res.stdout.strip())
    print("[+] Chiavi presenti nel record del dispositivo:")
    for k in sorted(keys):
        print(f"    - {k}")
except Exception as e:
    print(f"[-] Errore: {e}\nOutput: {res.stdout}")
