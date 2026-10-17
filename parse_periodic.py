import subprocess
import datetime

pcap = "/tmp/wireshark_enp3s0W0ZOW3.pcapng"
log_file = "periodic_inform_audit.log"
expected = 300
tol = 15

cmd = f"tshark -r {pcap} -Y 'http contains \"cwmp:Inform\" and http contains \"2 PERIODIC\"' -T fields -e frame.time_epoch"
res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
epochs = [float(x.strip()) for x in res.stdout.strip().split("\n") if x.strip()]

if not epochs:
    cmd_all = f"tshark -r {pcap} -Y 'http contains \"cwmp:Inform\"' -T fields -e frame.time_epoch"
    res_all = subprocess.run(cmd_all, shell=True, capture_output=True, text=True)
    epochs = [float(x.strip()) for x in res_all.stdout.strip().split("\n") if x.strip()]
    tipo = "Tutti gli Inform (inclusi Boot/Bootstrap)"
else:
    tipo = "2 PERIODIC"

timestamps = [datetime.datetime.fromtimestamp(ts) for ts in epochs[-4:]]
num = len(timestamps)

out = []
out.append("======================================================================")
out.append(f" AUDIT PERIODIC INFORM - CPE BC9424-TCL (Fonte: {pcap})")
out.append("======================================================================")
out.append(f"Filtro utilizzato:     {tipo}")
out.append(f"Eventi trovati totali: {len(epochs)}")
out.append(f"Campioni analizzati:   {num}")
out.append(f"Intervallo atteso:     {expected} s (Tolleranza: ±{tol} s)")
out.append("----------------------------------------------------------------------")

if num < 2:
    out.append(f"[!] Rilevati solo {num} campioni nel file. Servono almeno 2 campioni per misurare il delta.")
    for i, ts in enumerate(timestamps):
        ts_format = ts.strftime("%Y-%m-%d %H:%M:%S")
        out.append(f"  Campione {i+1}: {ts_format}")
else:
    out.append(f"{'N°':<4} | {'Timestamp':<22} | {'Delta (s)':<12} | {'Esito':<15}")
    out.append("-" * 60)
    ok_all = True
    for i in range(num):
        ts_str = timestamps[i].strftime("%Y-%m-%d %H:%M:%S")
        if i == 0:
            out.append(f"{i+1:<4} | {ts_str:<22} | {'[Rif]':<12} | -")
        else:
            delta = (timestamps[i] - timestamps[i-1]).total_seconds()
            diff = abs(delta - expected)
            is_ok = diff <= tol
            if not is_ok:
                ok_all = False
            esito = "CONFORME" if is_ok else f"FAIL (diff: {diff:.1f}s)"
            out.append(f"{i+1:<4} | {ts_str:<22} | {f'{delta:.1f} s':<12} | {esito:<15}")
    out.append("----------------------------------------------------------------------")
    verdetto = "CONFORME AI REQUISITI TR-069" if ok_all else "NON CONFORME"
    out.append(f"VERDETTO: {verdetto}")

out.append("======================================================================\n")

text = "\n".join(out)
print(text)

with open(log_file, "w", encoding="utf-8") as f:
    f.write(text)
print(f"[+] Log salvato in: {log_file}")
