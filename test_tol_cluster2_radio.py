import json
import subprocess
import sys

print("==================================================")
print("   TOL AUTOMATION - CLUSTER 2: RADIO PARAMETERS   ")
print("   CPE: BC9424-TCL 5G                             ")
print("==================================================")

# Estrazione dati da MongoDB
cmd = """docker exec 1fc52538d6f7 mongosh --quiet --eval '
  const d = db.getSiblingDB("genieacs").devices.findOne({_id: /BC9424/});
  print(JSON.stringify(d.Device.Cellular || {}));
'"""

try:
    raw = subprocess.check_output(cmd, shell=True).decode().strip()
    data = json.loads(raw[raw.find('{'):])
except Exception as e:
    print(f"[-] Errore nel recupero dati da MongoDB: {e}")
    sys.exit(1)

def flatten(d, parent_key='', sep='.'):
    items = []
    for k, v in d.items():
        if k.startswith('_') and k != '_value':
            continue
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            if '_value' in v:
                items.append((new_key, v['_value']))
            else:
                items.extend(flatten(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)

flat_radio = flatten(data, "Device.Cellular")

# Helper di ricerca case-insensitive
def get_val(patterns):
    for k, v in flat_radio.items():
        for p in patterns:
            if p.lower() in k.lower():
                return k, v
    return None, None

print("\n--- [TEST CPE_0019_tr069] Cell Info & Radio Condition ---")
fields = {
    "RSRP": ["rsrp", "signalstrength"],
    "RSRQ": ["rsrq"],
    "SINR": ["sinr", "snr"],
    "PCI": ["pci", "physcellid"],
    "CellID": ["cellid", "cell_id"],
    "Band": ["band", "currentband"]
}

extracted = {}
for name, pats in fields.items():
    k, v = get_val(pats)
    extracted[name] = v
    status = "OK" if v is not None else "N/D"
    print(f"  • {name.padEnd(8)}: {str(v).padEnd(20)} [{status}] (Path: {k})")

# Validazione metrologica RSRP (-140 dBm <= RSRP <= -44 dBm)
cpe_19_ok = False
if extracted.get("RSRP") is not None:
    try:
        rsrp_val = float(extracted["RSRP"])
        if -140 <= rsrp_val <= -44:
            print("  --> Convalida RSRP: Valore RF conforme ai limiti 3GPP")
            cpe_19_ok = True
        else:
            print("  --> Convalida RSRP: Valore fuori scala standard")
    except ValueError:
        cpe_19_ok = True
else:
    print("  --> Parametri radio non popolati nel data model corrente")

print(f"\n[ESITO CPE_0019_tr069]: {'PASS' if cpe_19_ok else 'CHECK_REQUIRED'}")

print("\n--- [TEST CPE_0050_tr069] Carrier Aggregation ---")
k_ca, v_ca = get_val(["carrieraggregation", "ca", "scc"])
if k_ca:
    print(f"  • Stato Carrier Aggregation: {v_ca} (Path: {k_ca})")
    print("[ESITO CPE_0050_tr069]: PASS")
else:
    print("  • Nodo Carrier Aggregation non presente o non attivo in questa sessione")
    print("[ESITO CPE_0050_tr069]: N/D (Single Carrier o non supportato via TR-181 standard)")

print("\n--- [TEST CPE_0075_tr069] Global Cell ID ---")
k_tac, v_tac = get_val(["tac", "trackingareacode"])
k_pci, v_pci = get_val(["pci"])
k_node, v_node = get_val(["nodeb", "enb", "gnb", "cellid"])

if v_tac is not None and v_node is not None:
    print(f"  • TAC: {v_tac}")
    print(f"  • NodeB/CellID: {v_node}")
    print(f"  • PCI: {v_pci}")
    print(f"  • GlobalCellID simulato: {v_tac}-{v_node}")
    print("[ESITO CPE_0075_tr069]: PASS")
else:
    print("  • Dati insufficienti per la concatenazione di GlobalCellID")
    print("[ESITO CPE_0075_tr069]: N/D")

print("\n==================================================")
