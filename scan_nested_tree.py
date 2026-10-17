import subprocess
import json

script = """
const d = db.getSiblingDB('genieacs').devices.findOne({_id: /BC9424/});
function flatten(obj, prefix = '') {
  let res = [];
  for (let k in obj) {
    if (k.startsWith('_')) continue;
    const path = prefix ? prefix + '.' + k : k;
    if (obj[k] && typeof obj[k] === 'object' && !obj[k]._value && !obj[k]._type) {
      res = res.concat(flatten(obj[k], path));
    } else {
      res.push(path);
    }
  }
  return res;
}
print(JSON.stringify(flatten(d)));
"""

cmd = f"docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval \"{script}\""
res = subprocess.run(cmd, shell=True, capture_output=True, text=True)

try:
    paths = json.loads(res.stdout.strip())
    print(f"[+] Totale parametri scoperti nell'albero: {len(paths)}")
    diag = [p for p in paths if 'diagnostics' in p.lower() or 'download' in p.lower()]
    print(f"[+] Parametri Download / Diagnostics trovati ({len(diag)}):")
    for p in sorted(diag):
        print(f"    * {p}")
    if not diag and paths:
        print("[i] Primi 15 parametri disponibili nell'albero:")
        for p in sorted(paths)[:15]:
            print(f"    - {p}")
except Exception as e:
    print(f"[-] Errore: {e}\nRaw output: {res.stdout}")
