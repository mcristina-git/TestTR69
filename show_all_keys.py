import subprocess
import json

cmd = """docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval '
  const d = db.getSiblingDB("genieacs").devices.findOne({_id: /BC9424/});
  if (!d) {
    print("[]");
  } else {
    print(JSON.stringify(Object.keys(d)));
  }
'"""

res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
try:
    keys = json.loads(res.stdout.strip())
    # Separa metadati interni da parametri TR-069
    tr_params = [k for k in keys if not k.startswith("_")]
    meta_params = [k for k in keys if k.startswith("_")]
    
    print(f"[+] Metadati interni GenieACS ({len(meta_params)}):")
    for m in meta_params:
        print(f"    {m}")
        
    print(f"\n[+] Parametri dati registrati dalla CPE ({len(tr_params)}):")
    if not tr_params:
        print("    [-] Nessun parametro censito: la CPE invia solo l'Inform minimale.")
    else:
        for p in sorted(tr_params)[:30]:
            print(f"    {p}")
except Exception as e:
    print(f"[-] Errore parsing: {e}\nRaw output: {res.stdout}")
