import os, re, openpyxl
from datetime import datetime

PCAP_PATH = "/home/cristina/cpe-testbed/cpe_set_param_120s.pcap"
EXCEL_PATH = "/home/cristina/cpe-testbed/test.xlsx"
LOG_DIR = "/home/cristina/cpe-testbed/logs"
os.makedirs(LOG_DIR, exist_ok=True)

if os.path.exists(PCAP_PATH):
    with open(PCAP_PATH, "rb") as f:
        data = f.read().decode("latin1", errors="ignore")
    reqs = re.findall(r"(<cwmp:SetParameterValues\b.*?</cwmp:SetParameterValues>)", data, re.DOTALL)
    resps = re.findall(r"(<cwmp:SetParameterValuesResponse\b.*?</cwmp:SetParameterValuesResponse>|<soapenv:Fault\b.*?</soapenv:Fault>)", data, re.DOTALL)
    
    if reqs:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
        for i, req in enumerate(reqs):
            resp = resps[i] if i < len(resps) else "ATTESA_RISPOSTA"
            status = "PASS" if "<Status>0</Status>" in resp else ("PASS" if "<Status>1</Status>" in resp else "FAIL")
            note = "Status 0 (OK)" if status == "PASS" else "Fault o Errore"
            for name, val in re.findall(r"<Name>(.*?)</Name>\s*<Value[^>]*>(.*?)</Value>", req, re.DOTALL):
                name, val = name.strip(), val.strip()
                log_name = f"set_param_{name.replace(".", "_")}.log"
                log_path = os.path.join(LOG_DIR, log_name)
                with open(log_path, "w", encoding="utf-8") as lf:
                    lf.write(f"=== SET PARAM AUDIT ===\nParametro: {name}\nValore: {val}\nEsito: {status}\n\n--- ACS -> CPE ---\n{req}\n\n--- CPE -> ACS ---\n{resp}\n")
                print(f"[+] Generato log: logs/{log_name}")
                found = False
                for r in ws.iter_rows(min_row=2):
                    if r[0].value == name and r[1].value == "SET":
                        r[2].value, r[3].value, r[4].value = val, status, f"logs/{log_name} | {note}"
                        found = True
                        break
                if not found:
                    ws.append([name, "SET", val, status, f"logs/{log_name} | {note}"])
        wb.save(EXCEL_PATH)
        print(f"[+] Excel {EXCEL_PATH} aggiornato con successo.")
    else:
        print("[-] Nessuna richiesta SetParameterValues trovata nel PCAP.")
