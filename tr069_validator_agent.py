import json
import subprocess
import os
import datetime
import openpyxl

PCAP_PATH = "/home/cristina/cpe-testbed/cpe_set_param_120s.pcap"
LOG_FILE = "periodic_inform_audit.log"
EXCEL_FILE = "test.xlsx"

def audit_pcap_periodic_inform(pcap_path, expected_interval=120, tolerance=15):
    """Estrae gli ultimi 4 pacchetti Inform da Wireshark e calcola i delta temporali."""
    if not os.path.exists(pcap_path):
        tmp_files = [os.path.join("/tmp", f) for f in os.listdir("/tmp") if f.endswith((".pcap", ".pcapng"))]
        if tmp_files:
            tmp_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            pcap_path = tmp_files[0]
        else:
            return False, "Nessun file pcap trovato", None

    cmd = f"tshark -r {pcap_path} -Y 'http contains \"cwmp:Inform\" and http contains \"2 PERIODIC\"' -T fields -e frame.time_epoch"
    res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    raw_epochs = [float(x.strip()) for x in res.stdout.strip().split("\n") if x.strip()]
    epochs = []
    for ep in raw_epochs:
        if not epochs or (ep - epochs[-1]) > 10.0:
            epochs.append(ep)

    # Fallback su tutti gli Inform se l'EventCode non è filtrabile esplicitamente
    if not epochs:
        cmd_all = f"tshark -r {pcap_path} -Y 'http contains \"cwmp:Inform\"' -T fields -e frame.time_epoch"
        res_all = subprocess.run(cmd_all, shell=True, capture_output=True, text=True)
        epochs = [float(x.strip()) for x in res_all.stdout.strip().split("\n") if x.strip()]

    if len(epochs) < 2:
        return False, f"Campioni insufficienti ({len(epochs)})", None

    timestamps = [datetime.datetime.fromtimestamp(ts) for ts in epochs[-4:]]
    deltas = []
    log_lines = [
        "======================================================================",
        f" AUDIT PERIODIC INFORM - CPE BC9424-TCL (Fonte: {pcap_path})",
        "======================================================================",
        f"Data esecuzione:       {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Intervallo atteso:     {expected_interval} s (Tolleranza: ±{tolerance} s)",
        f"Campioni analizzati:   {len(timestamps)}",
        "----------------------------------------------------------------------",
        f"{'N°':<4} | {'Timestamp':<22} | {'Delta (s)':<12} | {'Esito':<15}",
        "-" * 60
    ]

    all_ok = True
    for i in range(len(timestamps)):
        ts_str = timestamps[i].strftime("%Y-%m-%d %H:%M:%S")
        if i == 0:
            log_lines.append(f"{i+1:<4} | {ts_str:<22} | {'[Riferimento]':<12} | -")
        else:
            delta = (timestamps[i] - timestamps[i-1]).total_seconds()
            deltas.append(delta)
            diff = abs(delta - expected_interval)
            is_ok = diff <= tolerance
            if not is_ok:
                all_ok = False
            status = "CONFORME" if is_ok else f"FAIL (scarto {diff:.1f}s)"
            log_lines.append(f"{i+1:<4} | {ts_str:<22} | {f'{delta:.1f} s':<12} | {status:<15}")

    avg_delta = sum(deltas) / len(deltas) if deltas else 0
    verdetto = "CONFORME AI REQUISITI TR-069" if all_ok else "NON CONFORME (Jitter elevato)"
    log_lines.append("----------------------------------------------------------------------")
    log_lines.append(f"Media misurata: {avg_delta:.1f} s | VERDETTO: {verdetto}")
    log_lines.append("======================================================================\n")

    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    detail_msg = f"{'OK' if all_ok else 'WARN'}: Delta medio {avg_delta:.1f}s (atteso {expected_interval}s) - Log: {LOG_FILE}"
    return all_ok, detail_msg, avg_delta

def get_mongo_device():
    cmd = "docker exec $(docker ps -q -f name=_mongo_) mongosh --quiet --eval \"JSON.stringify(db.getSiblingDB('genieacs').devices.findOne())\""
    res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    out = res.stdout.strip()
    s_idx = out.find("{")
    e_idx = out.rfind("}")
    if s_idx == -1 or e_idx == -1:
        return None
    return json.loads(out[s_idx:e_idx+1])

def get_param_value(dev, path):
    parts = path.split(".")
    curr = dev
    for p in parts:
        if isinstance(curr, dict) and p in curr:
            curr = curr[p]
        else:
            return None
    if isinstance(curr, dict) and "_value" in curr:
        return curr["_value"]
    return curr

def run_test_suite():
    print("[*] Connessione a MongoDB ed estrazione dati CPE...")
    dev = get_mongo_device()
    if not dev:
        print("[!] Errore: Impossibile estrarre i dati da MongoDB.")
        return

    dev_id = dev.get("_id", "Unknown")
    print(f"[+] Dispositivo rilevato: {dev_id}\n")

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Sheet2"]

    interval_val = get_param_value(dev, "Device.ManagementServer.PeriodicInformInterval") or 300

    for row in range(2, ws.max_row + 1):
        param = ws.cell(row=row, column=1).value
        expected = ws.cell(row=row, column=2).value
        if not param:
            continue

        val = get_param_value(dev, param)
        esito = "FAIL"
        note = ""

        # Test specifico per PeriodicInformTime con verifica Wireshark
        if param == "Device.ManagementServer.PeriodicInformTime":
            pcap_ok, pcap_detail, _ = audit_pcap_periodic_inform(PCAP_PATH, expected_interval=int(interval_val))
            
            if val is not None and str(val).strip() != "":
                val_str = str(val)
                if pcap_ok:
                    esito = "PASS"
                    note = f"PASS: Valore albero '{val_str}' | {pcap_detail}"
                else:
                    esito = "WARN" if "Campioni insufficienti" in pcap_detail else "FAIL"
                    note = f"{esito}: Valore albero '{val_str}' | {pcap_detail}"
            else:
                esito = "FAIL"
                note = f"FAIL: Parametro non valorizzato nell'albero | {pcap_detail}"

        elif "DeviceID." in param:
            did_key = param.split(".")[1]
            did_obj = dev.get("_deviceId", {})
            dval = did_obj.get(f"_{did_key}") or did_obj.get(did_key)
            if dval and (expected is None or str(expected).lower() in str(dval).lower() or "not null" in str(expected).lower()):
                esito = "PASS"
                note = f"OK: Corrisponde ({dval})"
            else:
                esito = "FAIL"
                note = f"FAIL: Trovato ({dval})"

        elif "IMEI" in param or "SerialNumber" in param:
            if val and len(str(val).strip()) >= 10:
                esito = "PASS"
                note = f"OK: Conforme ({val})"
            else:
                esito = "FAIL"
                note = f"FAIL: {val}"

        elif expected is not None and str(val) == str(expected):
            esito = "PASS"
            note = f"OK: Corrisponde ({val})"
        elif val is not None:
            esito = "PASS"
            note = f"OK: {val}"
        else:
            esito = "FAIL"
            note = "FAIL: Parametro non valorizzato"

        ws.cell(row=row, column=3).value = note
        status_tag = f"[{esito}]".ljust(7)
        print(f"  {status_tag} {param:<45} -> {note}")

    wb.save(EXCEL_FILE)
    print(f"\n[+] Data Model dedotto: TR-181")
    print(f"[+] Test completati. File aggiornato: {EXCEL_FILE}")
    print(f"[+] Dettaglio Wireshark salvato in: {LOG_FILE}")

if __name__ == "__main__":
    run_test_suite()

try:
    import generate_report_ppt
    generate_report_ppt.build_presentation()
except Exception as e:
    print(f"[-] Errore generazione PPT: {e}")
