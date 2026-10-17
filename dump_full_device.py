import json
import subprocess
import sys

# Estrae l'intero documento JSON del device da mongosh formattandolo esplicitamente
cmd = [
    "docker", "exec", "-i",
    subprocess.check_output("docker ps -q -f name=_mongo_", shell=True).decode().strip(),
    "mongosh", "--quiet", "--eval",
    'EJSON.stringify(db.getSiblingDB("genieacs").devices.findOne({_id: /BC9424/}))'
]

res = subprocess.run(cmd, capture_output=True, text=True)

if res.returncode != 0 or not res.stdout.strip():
    print("[-] Errore nel recupero da mongosh:", res.stderr)
    sys.exit(1)

# Pulisce l'output da eventuali righe di warning iniziali
raw_text = res.stdout.strip()
json_start = raw_text.find('{')
if json_start == -1:
    print("[-] Nessun JSON trovato nell'output:", raw_text)
    sys.exit(1)

doc = json.loads(raw_text[json_start:])

# Ricorsione su tutti i nodi dell'albero CWMP
all_params = {}

def walk(node, prefix=""):
    if not isinstance(node, dict):
        return
    for k, v in node.items():
        if k.startswith("_"):
            continue
        current_path = f"{prefix}.{k}" if prefix else k
        if isinstance(v, dict):
            # Se ha un campo _value o è un nodo foglia censito
            if "_value" in v:
                val = v["_value"]
                # Gestione oggetti Extended JSON ($date, $numberLong, ecc.)
                if isinstance(val, dict):
                    val = next(iter(val.values()))
                all_params[current_path] = str(val) if val is not None else ""
            else:
                # Controlla se ha metadati di tipo/writable ma senza _value
                has_meta = any(m in v for m in ("_type", "_writable", "_timestamp"))
                sub_keys = [sk for sk in v.keys() if not sk.startswith("_")]
                if has_meta and not sub_keys:
                    all_params[current_path] = "-"
                else:
                    walk(v, current_path)
        else:
            all_params[current_path] = str(v)

walk(doc)

out_file = "/home/cristina/cpe-testbed/all_cpe_params.json"
with open(out_file, "w") as f:
    json.dump(all_params, f, indent=2)

print(f"[+] Estratti con successo {len(all_params)} parametri TR-181!")
print(f"[+] File salvato: {out_file}")
print("\nPrimi 10 parametri rilevati:")
for k in sorted(all_params.keys())[:10]:
    print(f"  * {k} = {all_params[k]}")
