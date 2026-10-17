import subprocess
import json

script = """
const d = db.getSiblingDB('genieacs').devices.findOne({_id: /BC9424/});
function extract(obj, prefix = '') {
  let res = {};
  for (let k in obj) {
    if (k.startsWith('_')) continue;
    const path = prefix ? prefix + '.' + k : k;
    if (obj[k] && typeof obj[k] === 'object' && obj[k]._value !== undefined) {
      res[path] = obj[k]._value;
    } else if (obj[k] && typeof obj[k] === 'object') {
      Object.assign(res, extract(obj[k], path));
    }
  }
  return res;
}
print(JSON.stringify(extract(d)));
"""

cmd = f"docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval \"{script}\""
res = subprocess.run(cmd, shell=True, capture_output=True, text=True)

try:
    data = json.loads(res.stdout.strip())
    diag_metrics = {k: v for k, v in data.items() if any(term in k.lower() for term in ['diagnostics', 'download', 'romtime', 'bomtime', 'eomtime'])}
    if diag_metrics:
        print(f"[+] Trovati {len(diag_metrics)} parametri diagnostici TR-143:")
        for k, v in sorted(diag_metrics.items()):
            print(f"    * {k} = {v}")
    else:
        print("[-] Nessun parametro diagnostico trovato nei valori salvati.")
except Exception as e:
    print(f"[-] Errore: {e}\nRaw output: {res.stdout}")
