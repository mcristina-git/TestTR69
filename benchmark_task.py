import subprocess
import json
import time
import urllib.request

# 1. Recupero ID reale da Mongo
cmd = """docker exec 1fc52538d6f7 mongosh --quiet --eval '
  const d = db.getSiblingDB("genieacs").devices.findOne({_id: /BC9424/});
  print(d._id);
'"""
dev_id = subprocess.check_output(cmd, shell=True).decode().strip()
print(f"[*] ID rilevato nel DB: [{dev_id}]")

# 2. Pulizia coda task
clean_cmd = f"""docker exec 1fc52538d6f7 mongosh --quiet --eval '
  db.getSiblingDB("genieacs").tasks.deleteMany({{device: "{dev_id}"}});
'"""
subprocess.check_call(clean_cmd, shell=True)
print("[*] Coda task azzerata.")

# 3. Invio task via porta NBI REST 7557
url = f"http://127.0.0.1:7557/devices/{dev_id}/tasks"

task = {
    "name": "getParameterValues",
    "parameterNames": ["Device.DeviceInfo.SoftwareVersion"]
}

t_start = time.time()
req = urllib.request.Request(
    url,
    data=json.dumps(task).encode("utf-8"),
    headers={"Content-Type": "application/json"},
    method="POST"
)

try:
    with urllib.request.urlopen(req) as resp:
        print(f"[+] Task accodato via NBI (porta 7557) alle {time.strftime('%H:%M:%S')} (HTTP {resp.status})")
except Exception as e:
    print(f"[-] Errore chiamata API: {e}")
    exit(1)

# 4. Polling continuo su MongoDB per cronometrare la presa in carico reale
print("[*] In ascolto: in attesa che la CPE prelevi il task...")
print("[*] (Puoi forzare l'Inform immediato con flap Ethernet da un altro terminale)")

while True:
    check_cmd = f"""docker exec 1fc52538d6f7 mongosh --quiet --eval '
      print(db.getSiblingDB("genieacs").tasks.countDocuments({{device: "{dev_id}"}}));
    '"""
    count_str = subprocess.check_output(check_cmd, shell=True).decode().strip()
    try:
        count = int(count_str)
    except ValueError:
        count = 1

    if count == 0:
        elapsed = time.time() - t_start
        print("\n==============================================")
        print(f"[OK] TASK PRESO IN CARICO DALLA CPE!")
        print(f"     Tempo reale di attesa: {elapsed:.2f} secondi")
        print("==============================================")
        break
    
    print(".", end="", flush=True)
    time.sleep(2)
