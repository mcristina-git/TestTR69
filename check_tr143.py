import subprocess
import json

cmd = "docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval \"JSON.stringify(db.getSiblingDB('genieacs').devices.find().sort({_lastInform: -1}).limit(1).toArray())\""
res = subprocess.run(cmd, shell=True, capture_output=True, text=True)

try:
    arr = json.loads(res.stdout.strip())
    if not arr:
        print("[-] Nessun dispositivo trovato a database.")
    else:
        dev = arr[0]
        dev_id = dev.get("_id", "Sconosciuto")
        print(f"[+] Dispositivo trovato: {dev_id}")
        
        matches = [k for k in dev.keys() if "diagnostics" in k.lower() or "download" in k.lower()]
        if matches:
            print("[+] Parametri di diagnostica/download presenti:")
            for m in sorted(matches):
                print(f"    - {m}")
        else:
            print("[-] Nessun parametro DownloadDiagnostics presente a database.")
            print("[i] I nodi non sono ancora stati letti tramite GetParameterNames/Values.")
except Exception as e:
    print(f"[-] Errore: {e}")
    print(f"Stdout: {res.stdout[:200]}")
    print(f"Stderr: {res.stderr[:200]}")
