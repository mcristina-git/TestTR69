import json
import re
import os
import pandas as pd
from datetime import datetime

class TRStandardsComplianceAuditor:
    def __init__(self, json_path="cpe_datamodel_full.json"):
        self.json_path = json_path
        self.anomalies = []
        self.warnings = []
        self.passed = []
        self.params = {}
        self.model = "UNKNOWN"

    def load_data(self):
        if not os.path.exists(self.json_path):
            raise FileNotFoundError(f"File {self.json_path} non trovato.")
        
        with open(self.json_path, "r", encoding="utf-8") as f:
            dev = json.load(f)

        # Appiattimento parametri
        def walk(prefix, node):
            if isinstance(node, dict):
                if "_value" in node:
                    self.params[prefix] = {
                        "value": node.get("_value"),
                        "type": node.get("_type", ""),
                        "writable": node.get("_writable", False)
                    }
                for k, v in node.items():
                    if not k.startswith("_"):
                        walk(f"{prefix}.{k}" if prefix else k, v)

        for root in ["Device", "InternetGatewayDevice"]:
            if root in dev:
                walk(root, dev[root])

        if any(k.startswith("Device.") for k in self.params.keys()):
            self.model = "TR-181 Issue 2 (Device:2)"
        elif any(k.startswith("InternetGatewayDevice.") for k in self.params.keys()):
            self.model = "TR-098 (InternetGatewayDevice:1)"

    def log(self, level, code, param, detail):
        entry = {
            "Severità": level,
            "Regola": code,
            "Parametro": param,
            "Dettaglio": detail
        }
        if level == "CRITICAL":
            self.anomalies.append(entry)
        elif level == "WARNING":
            self.warnings.append(entry)
        else:
            self.passed.append(entry)

    def check_root_and_collisions(self):
        has_dev = any(k.startswith("Device.") for k in self.params.keys())
        has_igd = any(k.startswith("InternetGatewayDevice.") for k in self.params.keys())

        if has_dev and has_igd:
            self.log("CRITICAL", "BBF-STRUCT-01", "ROOT", "Collisione rilevata: presenti sia nodi Device. sia InternetGatewayDevice.")
        elif not has_dev and not has_igd:
            self.log("CRITICAL", "BBF-STRUCT-02", "ROOT", "Nessun albero TR standard valido rilevato.")
        else:
            self.log("INFO", "BBF-STRUCT-OK", "ROOT", f"Struttura coerente con {self.model}")

    def check_device_info_mandatory(self):
        root = "Device.DeviceInfo." if self.model.startswith("TR-181") else "InternetGatewayDevice.DeviceInfo."
        mandatory = [
            ("Manufacturer", str),
            ("ManufacturerOUI", str),
            ("ProductClass", str),
            ("SerialNumber", str),
            ("HardwareVersion", str),
            ("SoftwareVersion", str),
            ("UpTime", (int, float, str))
        ]

        for p_name, exp_type in mandatory:
            full_path = root + p_name
            if full_path not in self.params:
                self.log("CRITICAL", "TR181-MANDATORY-MISSING", full_path, "Parametro obbligatorio mancante nel data model.")
            else:
                val = self.params[full_path]["value"]
                if val is None or str(val).strip() == "":
                    self.log("CRITICAL", "TR181-VAL-EMPTY", full_path, "Valore nullo o stringa vuota per parametro obbligatorio.")
                else:
                    self.log("INFO", "TR181-OK", full_path, f"Presente e valorizzato ({val})")

        # Verifica formato OUI (deve essere 6 caratteri esadecimali)
        oui_param = root + "ManufacturerOUI"
        if oui_param in self.params:
            oui_val = str(self.params[oui_param]["value"])
            if not re.match(r"^[0-9A-Fa-f]{6}$", oui_val):
                self.log("WARNING", "BBF-FORMAT-OUI", oui_param, f"ManufacturerOUI non è un codice esadecimale a 6 cifre valido: '{oui_val}'")

    def check_management_server_mandatory(self):
        root = "Device.ManagementServer." if self.model.startswith("TR-181") else "InternetGatewayDevice.ManagementServer."
        
        # 1. URL Connection Request
        cr_url = root + "ConnectionRequestURL"
        if cr_url in self.params:
            val = str(self.params[cr_url]["value"] or "")
            if not re.match(r"^https?://[a-zA-Z0-9.\-_]+:[0-9]+(/.*)?$", val):
                self.log("CRITICAL", "TR069-CR-URL-INVALID", cr_url, f"Formato ConnectionRequestURL non conforme a RFC 3986: '{val}'")
            else:
                self.log("INFO", "TR069-OK", cr_url, f"URL valido: {val}")
        else:
            self.log("CRITICAL", "TR069-CR-URL-MISSING", cr_url, "ConnectionRequestURL non presente.")

        # 2. Credenziali Connection Request
        cr_user = root + "ConnectionRequestUsername"
        cr_pwd = root + "ConnectionRequestPassword"
        for cr_p in [cr_user, cr_pwd]:
            if cr_p not in self.params or not str(self.params[cr_p]["value"]):
                self.log("WARNING", "TR069-CR-AUTH", cr_p, "Credenziale Connection Request vuota o non valorizzata (rischio sicurezza HTTP 401).")

        # 3. Inform Periodico
        inf_en = root + "PeriodicInformEnable"
        inf_int = root + "PeriodicInformInterval"
        inf_time = root + "PeriodicInformTime"

        if inf_int in self.params:
            try:
                sec = int(self.params[inf_int]["value"])
                if sec < 5:
                    self.log("WARNING", "TR069-INFORM-INTERVAL", inf_int, f"Intervallo troppo basso ({sec}s), rischio saturazione ACS.")
                elif sec > 86400:
                    self.log("WARNING", "TR069-INFORM-INTERVAL", inf_int, f"Intervallo molto elevato ({sec}s > 24h).")
            except Exception:
                self.log("CRITICAL", "TR069-INFORM-TYPE", inf_int, "PeriodicInformInterval non è un intero.")

        if inf_time in self.params:
            t_val = str(self.params[inf_time]["value"])
            # Controllo formato ISO 8601 o epoch base
            if t_val.startswith("1970-01-01"):
                self.log("WARNING", "TR069-TIME-UNSYNC", inf_time, f"Orologio non sincronizzato (Epoch 1970): '{t_val}'. La CPE necessita di NTP.")

    def check_instance_indexing(self):
        # Verifica che le istanze {i} siano numeriche positive
        regex = re.compile(r"^Device\.([A-Za-z0-9]+)\.([0-9]+)\.")
        instances = {}
        for p in self.params.keys():
            match = regex.match(p)
            if match:
                obj, idx = match.group(1), int(match.group(2))
                instances.setdefault(obj, set()).add(idx)

        for obj, idx_set in instances.items():
            if 0 in idx_set:
                self.log("CRITICAL", "TR181-INDEX-0", f"Device.{obj}.0.", "Indice d'istanza non valido: gli indici TR-181 partono da 1.")
            sorted_idx = sorted(list(idx_set))
            if sorted_idx and sorted_idx[0] != 1:
                self.log("WARNING", "TR181-INDEX-GAP", f"Device.{obj}", f"Le istanze non partono da 1 (rilevato primo indice {sorted_idx[0]}).")

    def run_audit(self):
        self.load_data()
        self.check_root_and_collisions()
        self.check_device_info_mandatory()
        self.check_management_server_mandatory()
        self.check_instance_indexing()

        all_results = self.anomalies + self.warnings + self.passed
        df = pd.DataFrame(all_results)
        df.to_excel("bbf_compliance_report.xlsx", index=False)

        print("\n" + "="*70)
        print(f" BBF STANDARDS COMPLIANCE AUDIT - {self.model}")
        print("="*70)
        print(f"Parametri totali scansionati: {len(self.params)}")
        print(f"Anomalie Critiche (FAIL):    {len(self.anomalies)}")
        print(f"Avvisi/Discrepanze (WARN):   {len(self.warnings)}")
        print(f"Controlli Superati (PASS):   {len(self.passed)}")
        print("="*70)

        if self.anomalies:
            print("\n[!] CRITICAL ANOMALIES (Non conformi a TR-069/181):")
            for a in self.anomalies:
                print(f"  * [{a['Regola']}] {a['Parametro']} -> {a['Dettaglio']}")

        if self.warnings:
            print("\n[?] WARNINGS (Comportamenti anomali o rischio disallineamento):")
            for w in self.warnings:
                print(f"  * [{w['Regola']}] {w['Parametro']} -> {w['Dettaglio']}")

        print(f"\n[+] Report salvato: bbf_compliance_report.xlsx\n")

if __name__ == "__main__":
    auditor = TRStandardsComplianceAuditor("cpe_datamodel_full.json")
    auditor.run_audit()
