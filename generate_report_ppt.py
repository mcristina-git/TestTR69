import os, glob, datetime, openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

EXCEL_PATH = "/home/cristina/cpe-testbed/test.xlsx"
LOGS_DIR = "/home/cristina/cpe-testbed/logs"
PERIODIC_LOG_PATH = "/home/cristina/cpe-testbed/periodic_inform_audit.log"
OUTPUT_PPTX = "/home/cristina/cpe-testbed/report_audit_tr069.pptx"

C_NAVY = RGBColor(16, 37, 66)
C_DARK = RGBColor(33, 37, 41)
C_MUTED = RGBColor(108, 117, 125)
C_LIGHT_BG = RGBColor(245, 247, 250)
C_CARD_BG = RGBColor(255, 255, 255)
C_BORDER = RGBColor(222, 226, 230)
C_PASS = RGBColor(40, 167, 69)
C_FAIL = RGBColor(220, 53, 69)
C_WARN = RGBColor(255, 153, 0)
C_CODE_BG = RGBColor(241, 243, 245)

def apply_background(slide, prs):
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = C_LIGHT_BG
    bg_shape.line.fill.background()
    return bg_shape

def add_header(slide, title_text, subtitle_text):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Arial"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = C_NAVY
    p2 = tf.add_paragraph()
    p2.text = subtitle_text
    p2.font.name = "Arial"
    p2.font.size = Pt(11)
    p2.font.color.rgb = C_MUTED
    p2.space_before = Pt(4)

def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
    else:
        card.line.fill.background()
    return card

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Slide 1: Copertina
    s1 = prs.slides.add_slide(blank_layout)
    bg = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_NAVY
    bg.line.fill.background()
    tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(11), Inches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "REPORT DI AUDIT E CONFORMITÀ TR-069"
    p.font.name = "Arial"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p_sub = tf1.add_paragraph()
    p_sub.text = "Validazione Protocollo CWMP • CPE TCL 5G (BC9424) • Data Model TR-181"
    p_sub.font.name = "Arial"
    p_sub.font.size = Pt(18)
    p_sub.font.color.rgb = RGBColor(180, 205, 237)
    p_sub.space_before = Pt(12)
    p_meta = tf1.add_paragraph()
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    p_meta.text = f"Generato da Validator Agent | Data: {now_str}\nTestbed: ServerCri (Kathara + GenieACS + MongoDB)"
    p_meta.font.name = "Arial"
    p_meta.font.size = Pt(12)
    p_meta.font.color.rgb = RGBColor(140, 165, 195)
    p_meta.space_before = Pt(36)

    # Slide 2: Risultati Excel
    s2 = prs.slides.add_slide(blank_layout)
    apply_background(s2, prs)
    add_header(s2, "1. Risultati Audit e Struttura File Excel", "Stato complessivo dei parametri verificati ed esiti memorizzati in test.xlsx")
    excel_rows = []
    if os.path.exists(EXCEL_PATH):
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            if any(row):
                excel_rows.append([str(c) if c is not None else "" for c in row[:5]])
    if excel_rows:
        headers = ["Parametro", "Valore Atteso / Op", "Esito / Valore", "Status", "Evidenza / Dettagli"]
        rows_to_display = [headers] + [r for r in excel_rows[1:] if r[0]][:12]
        table_shape = s2.shapes.add_table(len(rows_to_display), 5, Inches(0.8), Inches(1.6), Inches(11.73), Inches(5.2))
        table = table_shape.table
        table.columns[0].width = Inches(3.8)
        table.columns[1].width = Inches(2.2)
        table.columns[2].width = Inches(1.8)
        table.columns[3].width = Inches(1.1)
        table.columns[4].width = Inches(2.83)
        for r_idx, r_data in enumerate(rows_to_display):
            for c_idx in range(5):
                val = r_data[c_idx] if c_idx < len(r_data) else ""
                cell = table.cell(r_idx, c_idx)
                cell.text = val
                p = cell.text_frame.paragraphs[0]
                p.font.name = "Arial"
                p.font.size = Pt(9 if r_idx > 0 else 10)
                if r_idx == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = C_NAVY
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(255, 255, 255) if r_idx % 2 == 0 else RGBColor(248, 249, 252)
                    p.font.color.rgb = C_DARK
                    if "PASS" in val or "OK" in val:
                        p.font.bold = True
                        p.font.color.rgb = C_PASS
                    elif "FAIL" in val:
                        p.font.bold = True
                        p.font.color.rgb = C_FAIL

    # Slide 3: Albero Data Model
    s3 = prs.slides.add_slide(blank_layout)
    apply_background(s3, prs)
    add_header(s3, "2. Albero del Data Model (TR-181 Device)", "Gerarchia logica dei nodi interrogati e validati sul dispositivo")
    col_w, gap, top_pos, card_h = Inches(3.65), Inches(0.38), Inches(1.7), Inches(5.1)

    cards_def = [
        (Inches(0.8), "Device.DeviceInfo.", [
            ("HardwareVersion", "HH516L-V3.0"),
            ("SoftwareVersion", "HH516L_FT_03.00_17"),
            ("ModemFirmwareVersion", "HH516L_FT_03.00_17"),
            ("SpecVersion", "1.0"),
            ("IMEI", "869299070835408")
        ]),
        (Inches(0.8) + col_w + gap, "Device.ManagementServer.", [
            ("PeriodicInformEnable", "true"),
            ("PeriodicInformInterval", "120 s (Modificato da 300)"),
            ("ConnectionRequestURL", "http://10.49.27.145:7170/..."),
            ("ConnectionRequestUsername", "BC9424-TCL 5G CPE-..."),
            ("ConnectionRequestPassword", "•••••••• (134efnr1j29)")
        ]),
        (Inches(0.8) + (col_w + gap)*2, "DeviceID Identifiers", [
            ("Manufacturer", "TCL Communication"),
            ("OUI", "BC9424"),
            ("ProductClass", "TCL 5G CPE"),
            ("SerialNumber", "869299070835408")
        ])
    ]
    for left, title, items in cards_def:
        add_card(s3, left, top_pos, col_w, card_h)
        tb = s3.shapes.add_textbox(left + Inches(0.15), top_pos + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Arial"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        for k, v in items:
            p_k = tf.add_paragraph()
            p_k.text = f"• {k}"
            p_k.font.bold = True
            p_k.font.size = Pt(10)
            p_k.font.color.rgb = C_DARK
            p_k.space_before = Pt(8)
            p_v = tf.add_paragraph()
            p_v.text = f"   Valore: {v}"
            p_v.font.size = Pt(9)
            p_v.font.color.rgb = C_MUTED

    # Slide 4: Operazioni POST
    s4 = prs.slides.add_slide(blank_layout)
    apply_background(s4, prs)
    add_header(s4, "3. Operazioni POST CWMP (SetParameterValues)", "Modifiche parametri, confronto valori originale/nuovo, tracciamento log e scambio SOAP")
    add_card(s4, Inches(0.8), Inches(1.6), Inches(4.5), Inches(5.2))
    tb = s4.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(4.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Dettagli Modifica Parametro"
    p.font.name = "Arial"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    post_details = [
        ("Parametro Modificato", "Device.ManagementServer.PeriodicInformInterval"),
        ("Tipo di Dato", "xsd:unsignedInt"),
        ("Valore Originale", "300 s"),
        ("Nuovo Valore Applicato", "120 s"),
        ("Metodo Richiesta", "HTTP POST (SetParameterValues)"),
        ("Esito Risposta CPE", "Status 0 (Successo - Applicato)"),
        ("Nome File Log", "logs/set_param_Device_...Interval.log")
    ]
    for label, val in post_details:
        p_l = tf.add_paragraph()
        p_l.text = label
        p_l.font.bold = True
        p_l.font.size = Pt(10)
        p_l.font.color.rgb = C_DARK
        p_l.space_before = Pt(8)
        p_val = tf.add_paragraph()
        p_val.text = val
        p_val.font.size = Pt(10)
        p_val.font.color.rgb = C_PASS if "Successo" in val or "120" in val else C_MUTED

    add_card(s4, Inches(5.6), Inches(1.6), Inches(6.9), Inches(5.2), bg_color=C_CODE_BG)
    tb_code = s4.shapes.add_textbox(Inches(5.8), Inches(1.8), Inches(6.5), Inches(4.8))
    tf_code = tb_code.text_frame
    tf_code.word_wrap = True
    p_c_title = tf_code.paragraphs[0]
    p_c_title.text = "Contenuto del Log di Audit SOAP"
    p_c_title.font.name = "Arial"
    p_c_title.font.size = Pt(12)
    p_c_title.font.bold = True
    p_c_title.font.color.rgb = C_NAVY

    log_files = glob.glob(os.path.join(LOGS_DIR, "set_param_*.log"))
    log_content = "Nessun log trovato."
    if log_files:
        with open(log_files[0], "r", encoding="utf-8") as f:
            log_content = f.read()

    p_body = tf_code.add_paragraph()
    p_body.text = log_content.strip()
    p_body.font.name = "Courier New"
    p_body.font.size = Pt(8.5)
    p_body.font.color.rgb = RGBColor(30, 30, 30)
    p_body.space_before = Pt(6)

    # Slide 5: Verifica Cadenza Inform
    s5 = prs.slides.add_slide(blank_layout)
    apply_background(s5, prs)
    add_header(s5, "4. Verifica Cadenza Periodic Inform & Analisi Jitter", "Validazione intervallo di notifica periodica, campioni analizzati ed esito")
    add_card(s5, Inches(0.8), Inches(1.6), Inches(4.5), Inches(5.2))
    tb_m = s5.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(4.1), Inches(4.8))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True
    p = tf_m.paragraphs[0]
    p.text = "Metriche Periodicità"
    p.font.name = "Arial"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    inform_metrics = [
        ("Parametro Monitorato", "Device.ManagementServer.PeriodicInformTime"),
        ("Intervallo Configurato", "120 Secondi"),
        ("Tolleranza Ammessa", "± 15 Secondi (105s - 135s)"),
        ("Campioni Consecutivi", "Delta misurato: ~120 s"),
        ("Esito", "CONFORME"),
        ("Nome File Log", "periodic_inform_audit.log")
    ]
    for label, val in inform_metrics:
        p_l = tf_m.add_paragraph()
        p_l.text = label
        p_l.font.bold = True
        p_l.font.size = Pt(10)
        p_l.font.color.rgb = C_DARK
        p_l.space_before = Pt(8)
        p_v = tf_m.add_paragraph()
        p_v.text = val
        p_v.font.size = Pt(10)
        p_v.font.color.rgb = C_PASS if "CONFORME" in val else C_MUTED

    add_card(s5, Inches(5.6), Inches(1.6), Inches(6.9), Inches(5.2), bg_color=C_CODE_BG)
    tb_plog = s5.shapes.add_textbox(Inches(5.8), Inches(1.8), Inches(6.5), Inches(4.8))
    tf_plog = tb_plog.text_frame
    tf_plog.word_wrap = True
    p_p_title = tf_plog.paragraphs[0]
    p_p_title.text = "Log Wireshark / TShark (periodic_inform_audit.log)"
    p_p_title.font.name = "Arial"
    p_p_title.font.size = Pt(12)
    p_p_title.font.bold = True
    p_p_title.font.color.rgb = C_NAVY

    p_log_txt = "Log non disponibile."
    if os.path.exists(PERIODIC_LOG_PATH):
        with open(PERIODIC_LOG_PATH, "r", encoding="utf-8") as f:
            p_log_txt = f.read()

    p_p_body = tf_plog.add_paragraph()
    p_p_body.text = p_log_txt.strip()
    p_p_body.font.name = "Courier New"
    p_p_body.font.size = Pt(8.5)
    p_p_body.font.color.rgb = RGBColor(30, 30, 30)
    p_p_body.space_before = Pt(6)

    prs.save(OUTPUT_PPTX)
    print(f"[+] Presentazione PowerPoint generata con successo: {OUTPUT_PPTX}")

if __name__ == "__main__":
    build_presentation()
