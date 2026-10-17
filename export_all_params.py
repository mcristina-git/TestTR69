import subprocess
import json

script = """
const d = db.getSiblingDB('genieacs').devices.findOne({_id: /BC9424/});

function extractAll(obj, prefix = '') {
  let res = {};
  for (let k in obj) {
    if (k.startsWith('_')) continue;
    const currentPath = prefix ? prefix + '.' + k : k;
    const item = obj[k];
    
    if (item && typeof item === 'object') {
      // Se ha _value (anche se stringa vuota, null o false) oppure ha metadati di tipo CWMP
      if (item._value !== undefined) {
        res[currentPath] = item._value;
      } else if (item._type !== undefined || item._writable !== undefined) {
        res[currentPath] = "(non valorizzato)";
      } else {
        // È un nodo intermedio, continua la ricorsione
        Object.assign(res, extractAll(item, currentPath));
      }
    }
  }
  return res;
}

print(JSON.stringify(extractAll(d)));
"""

cmd = f"docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval \"{script}\""
res = subprocess.run(cmd, shell=True, capture_output=True, text=True)

try:
    all_params = json.loads(res.stdout.strip())
    print(f"[+] Parametri estratti dal database: {len(all_params)}")
    
    # Salva la lista completa in JSON
    with open('/home/cristina/cpe-testbed/all_cpe_params.json', 'w') as f:
        json.dump(all_params, f, indent=2)
    print("[+] Salvati in: /home/cristina/cpe-testbed/all_cpe_params.json")
    
except Exception as e:
    print(f"[-] Errore estrazione: {e}\n{res.stdout}")
