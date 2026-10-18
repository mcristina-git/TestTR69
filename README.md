# Test CPE TCL: TR-069, TR-181 e TOL

Unico programma: `tr069_validator_agent.py`. Usare l'ambiente virtuale esistente:

```bash
venv/bin/python tr069_validator_agent.py
```

L'esecuzione normale aggiorna il modello Device dall'ACS, esegue i GET della TOL,
rilegge il PeriodicInformInterval, acquisisce tre Inform 2 PERIODIC con relative
InformResponse HTTP 200, misura due intervalli e genera Excel, JSON e PowerPoint.
NTP resta in sola lettura; IMS CPE_0046 resta NON ESEGUITO.

```bash
venv/bin/python tr069_validator_agent.py --tol-active
venv/bin/python tr069_validator_agent.py --periodic-only
venv/bin/python tr069_validator_agent.py --report-only
venv/bin/python tr069_validator_agent.py --pcap /percorso/cattura.pcap
venv/bin/python tr069_validator_agent.py --self-test
venv/bin/python tr069_validator_agent.py --report-only --template report_template_tr069.pptx
```

- `--tol-active`: aggiunge le prove reversibili previste dalla TOL, registrando i
  valori iniziali, rileggendo gli esiti e ripristinando la configurazione. Nessun
  reset, firmware, APN o password viene modificato. Il preset ACS che impone
  30 secondi blocca la prova PI a 60 secondi senza modificare il preset.
- `--periodic-only`: nuova prova PCAP e PPTX, senza riscrivere la TOL.
- `--report-only`: esporta la cache ACS e rigenera il PPTX; non esegue nuovi test.
  Evidenze PCAP assenti, scadute o incoerenti producono WARN, non PASS.
- `--pcap`: usa una cattura fornita, estraendo i tre Inform periodici piu recenti
  del CPE con relative risposte. Verifica intervalli, retry, identita e data.
- `--self-test`: 22 controlli isolati, senza inviare task all'ACS/CPE.

Gli esiti validati dall'utente nella TOL esistente sono conservati, distinti
nell'ambito VALIDAZIONE UTENTE; le note riportano l'esito automatico corrente.
Questo non trasforma una prova automatica non eseguita in un test eseguito.

## File conservati

- `TOL _Huawei_H173-383.xlsx`: modello originale per mantenere invariate A-I.
- `TOL_TCL.xlsx`: TOL aggiornata; J valori/esiti, K note, Valori_TCL integrale.
- `test.xlsx`: valori attesi e audit sintetico. Se manca, lo script lo ricrea
  con i valori attesi incorporati; PeriodicInformInterval atteso 30 secondi.
- `report_audit_tr069_guida.pptx`: riferimento caricato dall?utente, conservato invariato.
- `report_template_tr069.pptx`: modello editoriale di 13 slide derivato dalla guida.
  Conserva stile, autrice e ordine. Le slide 9-13, comprese architettura e proposte
  Hyperway/rete, sono preparate editorialmente e copiate senza rigenerarle.
- `report_audit_tr069.pptx`: aggiorna data, riepilogo TOL, modello e prova PCAP
  nelle prime 8 slide. Dopo la slide 13 aggiunge tutti i parametri del data model
  come tabelle modificabili. Gli Excel restano nella cartella, senza oggetti OLE
  o file incorporati nel PowerPoint. Il riepilogo TOL riporta tutti i 78 casi
  presenti nel sorgente e distingue nelle note le validazioni utente.
- `cpe_datamodel_full.json`, `cpe_datamodel_flat.xlsx`: modello completo
  esportato con credenziali oscurate e distinzione fra vuoto, null e non letto.
- `periodic_inform_audit.log`: sintesi dei frame e degli intervalli verificati.
- `tr143_server/test_50MB.bin`: payload 50 MiB del server HTTP esistente.
- `lab.conf`, startup e Dockerfile: configurazione del laboratorio Kathara.
- `venv`, `requirements.txt`: interprete e dipendenze Python del programma.
- `logs`: risultati ed evidenze; ogni nuova prova PCAP conserva solo le sessioni
  dei tre Inform, il loro hash e la correlazione con il log ACS.

La cattura usa temporaneamente l'immagine Kathara gia installata, sulla rete
host e limitatamente alla porta CWMP ricavata dall'URL ACS del CPE. Il relativo
container viene arrestato e rimosso dal programma, senza fermare laboratori.
Sono richiesti Docker e tshark gia presenti sul server. Non vengono installati
pacchetti o modificati permessi/configurazioni di sistema. I PCAP sono privati
(mode 0600); il PPTX riporta solo timestamp, frame, eventi, risposte e misure.

I codici di uscita descrivono il flusso e l'evidenza periodica: 0 completato con
PCAP valido, 2 intervalli non conformi, 3 prerequisiti/evidenza indisponibili,
130 interruzione. Gli esiti individuali della TOL restano nel workbook.

La pulizia del 17/10/2026 e documentata in `logs/consolidation_20261017/cleanup_manifest.json`; il backup privato dei file rimossi e nella stessa cartella.

## Modello PowerPoint ed estensioni future

Il modello di 13 slide ? un ingresso necessario, come il TOL originale. Per
modificare testi e grafica editoriali usare `report_template_tr069.pptx`, non il
report gi? completo di appendice. `--template` permette di indicare una copia
dello stesso modello: conservare le 13 slide e gli oggetti aggiornati dal
validator nelle prime 8 slide. Le slide 10-13 non sono costruite dal codice.

CPE_0056_tr069 (riga 68) ? PASS su validazione utente del 18/10/2026. Il riuso
delle validazioni ? vincolato a identit? CPE, sorgente TOL e ID del caso. Le note
conservano l?esito automatico. Il PASS non sostituisce una nuova prova automatica.

Le matrici Hyperway IDU/ODU (due ordini di avvio e due ritardi: 5 e 300 secondi)
e gli scenari di interruzione/degrado rete sono proposte, non implementazioni
n? test eseguiti. Richiedono un adattatore di alimentazione/link, una topologia
controllabile e criteri/timeout concordati prima dell?esecuzione.
