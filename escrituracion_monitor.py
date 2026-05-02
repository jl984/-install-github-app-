#!/usr/bin/env python3
"""
Monitor de Escrituración — Tres Pascualas Apartments
─────────────────────────────────────────────────────
• Detecta cuando se registra FIRMA COMP. en la planilla Excel
• Envía email automático al equipo interno con datos de la operación
• Genera el Detalle de Escrituración (.xlsx) fiel al formato oficial

Requisitos:
    pip install openpyxl pandas

Configurar contraseña de aplicación Gmail:
    https://myaccount.google.com/apppasswords

Uso:
    python escrituracion_monitor.py              # Monitoreo continuo
    python escrituracion_monitor.py --generar 1  # Genera detalle op. N°1
    python escrituracion_monitor.py --test-email # Email de prueba
"""

import smtplib
import json
import os
import sys
import time
import argparse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ═══════════════════════════════════════════════════════
#  CONFIGURACIÓN — EDITAR ANTES DE USAR
# ═══════════════════════════════════════════════════════
GMAIL_USUARIO    = "jlastete@imodelo.cl"
GMAIL_CONTRASENA = "xxxx xxxx xxxx xxxx"       # ← Reemplaza con la contraseña real
EQUIPO_EMAILS    = [
    "jmunoz@latincapital.com",
    "carpide@latincapital.com",
    "rgamev@latincapital.com",
    "dgarces@ilatincapital.cl",
    "ventas@imodelo.cl",
    "fprieto@latincapital.com",
    "cespinoza@latincapital.com",
    "serviciocliente@imodelo.cl",
]
EXCEL_PATH    = "Tres_Pascualas_APARTMENTS_INTEGRADA__1_.xlsx"
ESTADO_PATH   = "estado_firmas.json"
INTERVALO_SEG = 300  # Revisar cada 5 minutos

# ═══════════════════════════════════════════════════════
#  ESTILOS PARA EL DETALLE DE ESCRITURACIÓN
# ═══════════════════════════════════════════════════════
BG_H1    = "FF2D2D2D"
BG_H2    = "FF3D3D3D"
BG_VAL   = "FFF5F5F5"
BG_WHITE = "FFFFFFFF"
BG_GRAY  = "FFEFEFEF"
BG_TH    = "FF5A5A5A"
BG_TOT   = "FF8DC63F"
WHITE    = "FFFFFFFF"
DARK     = "FF2D2D2D"

def _f(bold=False, sz=8, color=DARK):
    return Font(name="Arial", bold=bold, size=sz, color=color)

def _bg(rgb):
    return PatternFill("solid", fgColor=rgb)

def _al(h="left", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def sc(ws, coord, val, bold=False, sz=8, color=DARK, bg=BG_WHITE, h="left", wrap=False):
    c = ws[coord]
    c.value = val
    c.font  = _f(bold=bold, sz=sz, color=color)
    c.fill  = _bg(bg)
    c.alignment = _al(h=h, wrap=wrap)

def sec(ws, row, label):
    ws.merge_cells(f"C{row}:E{row}")
    sc(ws, f"C{row}", label, bold=True, sz=8, color=WHITE, bg=BG_H1)
    ws.row_dimensions[row].height = 15.95

def lv(ws, row, label, value, h=14.1):
    sc(ws, f"C{row}", label, bold=True, sz=8, color=WHITE, bg=BG_H2)
    sc(ws, f"D{row}", value, bold=False, sz=8, color=DARK, bg=BG_VAL)
    ws.row_dimensions[row].height = h

# ═══════════════════════════════════════════════════════
#  LECTURA DE PLANILLA
# ═══════════════════════════════════════════════════════
def leer_operaciones():
    df = pd.read_excel(EXCEL_PATH, sheet_name="Seguimiento", header=2)
    df.columns = [str(c).strip() for c in df.columns]
    op_col = df.columns[0]
    df = df[df[op_col].apply(lambda x: str(x).replace(".0","").isdigit())]
    df.index = df[op_col].astype(int)
    return df

def v(row, campo, default=""):
    val = row.get(campo, None)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    s = str(val).strip()
    return default if s in ("", "nan") else s

def num_int(row, campo, default=""):
    """Retorna número como entero si es válido (elimina .0)"""
    val = row.get(campo, None)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    try:
        return str(int(float(val)))
    except:
        return str(val).strip()

def uf(row, campo, default=0):
    try:
        val = float(row.get(campo, 0))
        return val if not pd.isna(val) and val != 0 else default
    except:
        return default

# ═══════════════════════════════════════════════════════
#  DETECCIÓN DE FIRMA COMPRADOR
# ═══════════════════════════════════════════════════════
def obtener_estado_firmas(df):
    estado = {}
    col_firma = "FIRMA COMP."
    for idx, row in df.iterrows():
        val = row.get(col_firma, None)
        estado[str(idx)] = "" if (val is None or (isinstance(val, float) and pd.isna(val))) else str(val).strip()
    return estado

def detectar_firmas_nuevas(estado_ant, estado_act, df):
    nuevas = []
    for idx_str, val_act in estado_act.items():
        val_ant = estado_ant.get(idx_str, "")
        if val_act and not val_ant:
            idx = int(idx_str)
            row = df.loc[idx]
            nuevas.append({"op_num": idx, "row": row, "fecha_firma": val_act})
    return nuevas

# ═══════════════════════════════════════════════════════
#  GENERADOR DETALLE DE ESCRITURACIÓN (.xlsx)
# ═══════════════════════════════════════════════════════
def generar_detalle_xlsx(op_num, df):
    row = df.loc[op_num]

    wb = Workbook()
    ws = wb.active
    ws.title = "Detalle Escrituración"

    ws.column_dimensions["B"].width = 1.57
    ws.column_dimensions["C"].width = 26.14
    ws.column_dimensions["D"].width = 33.57
    ws.column_dimensions["E"].width = 12.86
    ws.column_dimensions["F"].width = 1.43

    for r, h in {1: 5.1, 2: 3.95, 3: 28.5}.items():
        ws.row_dimensions[r].height = h

    # Título
    ws.merge_cells("C4:E4")
    sc(ws, "C4", "DETALLE DE ESCRITURACIÓN", bold=True, sz=10, color=WHITE, bg=BG_H1)
    ws.row_dimensions[4].height = 19.5
    ws.merge_cells("C5:E5")
    sc(ws, "C5", "Edificio TRES PASCUALAS APARTMENTS", bold=True, sz=8, color=WHITE, bg=BG_H1)
    ws.row_dimensions[5].height = 9.0
    ws.row_dimensions[6].height = 3.95

    # DATOS DEL COMPRADOR
    sec(ws, 7, "DATOS DEL COMPRADOR")
    lv(ws, 8,  "NOMBRE COMPRADOR",    v(row, "NOMBRE COMPLETO"))
    lv(ws, 9,  "CÉDULA DE IDENTIDAD", v(row, "RUT"))
    ws.row_dimensions[10].height = 5.1

    # INMOBILIARIA
    sec(ws, 11, "INMOBILIARIA VENDEDORA")
    lv(ws, 12, "RAZÓN SOCIAL", "INMOBILIARIA PAICAVÍ SpA")
    lv(ws, 13, "R.U.T.",       "77.070.568-1")
    ws.row_dimensions[14].height = 5.1

    # PROYECTO
    sec(ws, 15, "PROYECTO")
    lv(ws, 16, "NOMBRE DEL PROYECTO", "EDIFICIO TRES PASCUALAS APARTMENTS")
    lv(ws, 17, "ETAPA",               "ÚNICA")
    lv(ws, 18, "DIRECCIÓN",           "AVENIDA MANUEL RODRÍGUEZ N°1.049, CONCEPCIÓN")
    ws.row_dimensions[19].height = 5.1

    # FINANCIAMIENTO
    sec(ws, 20, "DESGLOSE DE FINANCIAMIENTO")
    lv(ws, 21, "HIPOTECARIO UF",    uf(row, "CRÉDITO HIPOT."))
    lv(ws, 22, "APORTE CONTADO UF", uf(row, "APORTE CONTADO"))
    ws.row_dimensions[23].height = 5.1

    # UNIDADES
    ws.merge_cells("C24:E24")
    sc(ws, "C24", "UNIDADES", bold=True, sz=8, color=WHITE, bg=BG_H1)
    ws.row_dimensions[24].height = 15.95

    for coord, lbl, al in [("C25", "TIPO / UNIDAD", "center"), ("D25", "ROL", "center"), ("E25", "PRECIO UF", "center")]:
        sc(ws, coord, lbl, bold=True, sz=8, color=WHITE, bg=BG_TH, h=al)
    ws.row_dimensions[25].height = 15.0

    # Departamento
    depto = num_int(row, "DEPARTAENTO")
    sc(ws, "C26", f"DEPARTAMENTO N° {depto}", bold=False, sz=8, bg=BG_WHITE)
    sc(ws, "D26", "", sz=8, bg=BG_WHITE, h="center")
    c = ws["E26"]; c.value = uf(row, "PRECIO DEPARTAMENTO")
    c.font = _f(bold=True, sz=8); c.fill = _bg(BG_WHITE); c.alignment = _al(h="right")
    ws.row_dimensions[26].height = 12.95

    # Estacionamiento
    est_num = num_int(row, "ESTACIONAMIENTO")
    est_uf  = uf(row, "PRECIO ESTACIONAMIENTO (ROL)")
    est_ug  = num_int(row, "ESTACIONAMIENTO USO Y GOCE")
    est_ug_uf = uf(row, "PRECIO USO Y GOCE. ESTACIONAMIENTO")
    if est_num:
        est_lbl = f"ESTACIONAMIENTO N° {est_num}"; est_p = est_uf
    elif est_ug:
        est_lbl = f"ESTACIONAMIENTO USO Y GOCE N° {est_ug}"; est_p = est_ug_uf
    else:
        est_lbl = "ESTACIONAMIENTO N°"; est_p = 0
    sc(ws, "C27", est_lbl, bold=False, sz=8, bg=BG_GRAY)
    sc(ws, "D27", "", sz=8, bg=BG_GRAY, h="center")
    c = ws["E27"]; c.value = est_p
    c.font = _f(bold=True, sz=8); c.fill = _bg(BG_GRAY); c.alignment = _al(h="right")
    ws.row_dimensions[27].height = 12.95

    # Bodega
    bod = num_int(row, "BODEGA")
    sc(ws, "C28", f"BODEGA N° {bod}" if bod else "BODEGA N°", bold=False, sz=8, bg=BG_WHITE)
    sc(ws, "D28", "", sz=8, bg=BG_WHITE, h="center")
    c = ws["E28"]; c.value = uf(row, "PRECIO BODEGA")
    c.font = _f(bold=True, sz=8); c.fill = _bg(BG_WHITE); c.alignment = _al(h="right")
    ws.row_dimensions[28].height = 12.95

    # Bicicletero
    bici = num_int(row, "BICICLETERO")
    sc(ws, "C29",
       f"ESTACIONAMIENTO DE BICICLETA N° {bici}" if bici else "ESTACIONAMIENTO DE BICICLETA N°",
       bold=False, sz=8, bg=BG_GRAY)
    sc(ws, "D29", "", sz=8, bg=BG_GRAY, h="center")
    c = ws["E29"]; c.value = uf(row, "PRECIO BICICLETERO")
    c.font = _f(bold=True, sz=8); c.fill = _bg(BG_GRAY); c.alignment = _al(h="right")
    ws.row_dimensions[29].height = 12.95

    # Total
    ws.merge_cells("C30:D30")
    sc(ws, "C30", "PRECIO TOTAL", bold=True, sz=9, color=WHITE, bg=BG_TOT)
    c = ws["E30"]; c.value = "=SUM(E26:E29)"
    c.font = _f(bold=True, sz=8, color=WHITE); c.fill = _bg(BG_TOT); c.alignment = _al(h="right")
    ws.row_dimensions[30].height = 18.0
    ws.row_dimensions[31].height = 5.1

    # CONDICIONES
    sec(ws, 32, "CONDICIONES DE LA OPERACIÓN")
    lv(ws, 33, "OPERACIÓN AFECTA A IVA", "SÍ")
    lv(ws, 34, "VIVIENDA SOCIAL", "")
    ws.row_dimensions[35].height = 5.1

    # DATOS BANCARIOS
    sec(ws, 36, "DATOS BANCARIOS Y GESTIÓN")
    lv(ws, 37, "BANCO ACREEDOR / ALZANTE", v(row, "BANCO", "BANCO SANTANDER"))
    lv(ws, 38, "ABOGADO DE ESCRITURACIÓN", "JOSÉ LUIS ASTETE GÓMEZ")
    lv(ws, 39, "MAIL ABOGADO",              "jlastete@imodelo.cl")
    lv(ws, 40, "GASTOS OPERACIONALES",      "PAGA COMPRADOR")
    lv(ws, 41, "CONTACTO TASACIÓN",
       "ROMINA VIDAL / serviciocliente@imodelo.cl / 932568032", h=25.5)
    lv(ws, 42, "NOTARÍA",           v(row, "NOTARÍA", ""))
    lv(ws, 43, "EJECUTIVA NOTARÍA", "Jeannette Figueroa — jfigueroa@notariasalgado.cl", h=24.0)
    ws.row_dimensions[44].height = 5.1

    ws.merge_cells("C45:E45")
    sc(ws, "C45", "INMOBILIARIA MODELO  |  www.imodelo.cl", bold=True, sz=8, color=WHITE, bg=BG_H1, h="center")
    ws.row_dimensions[45].height = 14.1

    apellido = v(row, "NOMBRE COMPLETO", "cliente").split()[-1]
    fname = f"Detalle_Op{op_num}_{apellido}.xlsx"
    wb.save(fname)
    return fname

# ═══════════════════════════════════════════════════════
#  EMAIL CON ADJUNTO
# ═══════════════════════════════════════════════════════
def enviar_email_firma(firma_info, adjunto_path):
    row       = firma_info["row"]
    op_num    = firma_info["op_num"]
    fecha     = firma_info["fecha_firma"]
    ahora     = datetime.now().strftime("%d/%m/%Y %H:%M")

    nombre    = v(row, "NOMBRE COMPLETO", "—")
    rut       = v(row, "RUT", "—")
    depto     = num_int(row, "DEPARTAENTO", "—")
    repertorio= v(row, "REPERTORIO", "—")
    notaria   = v(row, "NOTARÍA", "—")
    banco     = v(row, "BANCO", "—")
    total     = uf(row, "TOTAL ESCRITURA")
    proximo   = v(row, "PRÓXIMO HITO", "Por definir")

    # Unidades
    est_num  = num_int(row, "ESTACIONAMIENTO")
    est_ug   = num_int(row, "ESTACIONAMIENTO USO Y GOCE")
    bod_num  = num_int(row, "BODEGA")
    bici_num = num_int(row, "BICICLETERO")

    est_lbl = (f"Estacionamiento N° {est_num}" if est_num
               else (f"Est. Uso y Goce N° {est_ug}" if est_ug else "—"))

    unidades_rows = f"""
      <tr><td>Departamento</td><td><strong>N° {depto}</strong></td><td>UF {uf(row,"PRECIO DEPARTAMENTO"):.2f}</td></tr>
      <tr class="gray"><td>{est_lbl}</td><td></td>
        <td>UF {(uf(row,"PRECIO ESTACIONAMIENTO (ROL)") or uf(row,"PRECIO USO Y GOCE. ESTACIONAMIENTO")):.2f}</td></tr>
      {'<tr><td>Bodega</td><td>N° ' + bod_num + '</td><td>UF ' + f"{uf(row,'PRECIO BODEGA'):.2f}" + '</td></tr>' if bod_num else ""}
      {'<tr class="gray"><td>Bicicletero</td><td>N° ' + bici_num + '</td><td>UF ' + f"{uf(row,'PRECIO BICICLETERO'):.2f}" + '</td></tr>' if bici_num else ""}
    """

    asunto = (f"✅ Firma Comprador — Op.{op_num} | Dpto.{depto} | "
              f"{nombre.split()[0]} {nombre.split()[-1]} | Rep. {repertorio}")

    html = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8">
<style>
  body{{font-family:Arial,sans-serif;background:#f0f2f5;margin:0;padding:20px}}
  .wrap{{max-width:600px;margin:0 auto;background:#fff;border-radius:8px;
         box-shadow:0 2px 10px rgba(0,0,0,.1);overflow:hidden}}
  .hdr{{background:#1a3a5c;color:#fff;padding:20px 24px}}
  .hdr h1{{margin:0;font-size:17px;font-weight:700}}
  .hdr p{{margin:4px 0 0;font-size:12px;opacity:.7}}
  .badge{{background:#e8f5e9;border-left:4px solid #43a047;padding:14px 20px;margin:18px 22px 0;border-radius:0 6px 6px 0}}
  .badge h2{{margin:0;font-size:14px;color:#2e7d32}}
  .badge p{{margin:4px 0 0;font-size:12px;color:#555}}
  .sec{{padding:14px 22px 0}}
  .sec h3{{font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#888;margin:0 0 8px}}
  table{{width:100%;border-collapse:collapse;font-size:12px}}
  td{{padding:7px 8px;border-bottom:1px solid #f0f0f0;color:#333}}
  td:first-child{{color:#666;width:40%}}
  tr.gray td{{background:#fafafa}}
  .unid-hdr td{{background:#374151;color:#fff;font-weight:700;font-size:10px;text-transform:uppercase}}
  .total td{{background:#8dc63f;color:#fff;font-weight:700}}
  .next{{background:#fff8e1;border-left:3px solid #f9a825;padding:12px 20px;margin:14px 22px;
         border-radius:0 6px 6px 0;font-size:12px;color:#555}}
  .next strong{{color:#e65100}}
  .foot{{background:#1a3a5c;color:#fff;padding:12px 22px;font-size:10px;text-align:center;opacity:.9}}
</style></head>
<body><div class="wrap">
  <div class="hdr">
    <h1>📋 Tres Pascualas Apartments — Escrituración</h1>
    <p>{ahora}</p>
  </div>

  <div class="badge">
    <h2>✅ Firma del Comprador Registrada</h2>
    <p>Fecha de firma: <strong>{fecha}</strong></p>
  </div>

  <div class="sec"><h3>Operación</h3>
  <table>
    <tr><td>N° Operación</td><td><strong>Op. {op_num}</strong></td></tr>
    <tr class="gray"><td>Cliente</td><td><strong>{nombre}</strong></td></tr>
    <tr><td>RUT</td><td>{rut}</td></tr>
    <tr class="gray"><td>N° Repertorio</td><td><strong>{repertorio}</strong></td></tr>
    <tr><td>Notaría</td><td>{notaria}</td></tr>
    <tr class="gray"><td>Banco</td><td>{banco}</td></tr>
    <tr><td>Total Escritura</td><td><strong>UF {total:.2f}</strong></td></tr>
  </table></div>

  <div class="sec"><h3>Unidades Escrituradas</h3>
  <table>
    <tr class="unid-hdr"><td>Unidad</td><td>Identificación</td><td>Precio</td></tr>
    {unidades_rows}
    <tr class="total"><td colspan="2">PRECIO TOTAL ESCRITURA</td><td>UF {total:.2f}</td></tr>
  </table></div>

  <div class="next">📌 Próximo hito: <strong>{proximo}</strong></div>

  <div style="padding:10px 22px 16px;font-size:11px;color:#888">
    Se adjunta el Detalle de Escrituración completo en formato Excel.
  </div>

  <div class="foot">INMOBILIARIA MODELO | www.imodelo.cl — Sistema automático de escrituración</div>
</div></body></html>"""

    msg = MIMEMultipart("mixed")
    msg["Subject"] = asunto
    msg["From"]    = GMAIL_USUARIO
    msg["To"]      = ", ".join(EQUIPO_EMAILS)
    msg.attach(MIMEText(html, "html", "utf-8"))

    if adjunto_path and os.path.exists(adjunto_path):
        with open(adjunto_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(adjunto_path)}")
        msg.attach(part)

    # cPanel: servidor SMTP del dominio, puerto 465 con SSL.
    # Si falla con 465, reintenta con puerto 587 + STARTTLS.
    try:
        with smtplib.SMTP_SSL("mail.imodelo.cl", 465) as server:
            server.login(GMAIL_USUARIO, GMAIL_CONTRASENA)
            server.sendmail(GMAIL_USUARIO, EQUIPO_EMAILS, msg.as_string())
    except Exception:
        with smtplib.SMTP("mail.imodelo.cl", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(GMAIL_USUARIO, GMAIL_CONTRASENA)
            server.sendmail(GMAIL_USUARIO, EQUIPO_EMAILS, msg.as_string())

    print(f"  ✉️  Email enviado: Op.{op_num} | {nombre} | Rep. {repertorio}")

# ═══════════════════════════════════════════════════════
#  BUCLE PRINCIPAL
# ═══════════════════════════════════════════════════════
def monitorear():
    print(f"\n{'═'*55}")
    print("  🏢 Monitor de Escrituración — Tres Pascualas")
    print(f"  Hito monitoreado: FIRMA DEL COMPRADOR")
    print(f"  Revisando cada {INTERVALO_SEG//60} min | {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"{'═'*55}\n")

    if os.path.exists(ESTADO_PATH):
        with open(ESTADO_PATH, "r", encoding="utf-8") as f:
            estado_ant = json.load(f)
        print(f"  📂 Estado anterior: {len(estado_ant)} operaciones")
    else:
        estado_ant = {}
        print("  🆕 Primera ejecución — estableciendo estado base...")

    while True:
        ahora = datetime.now().strftime("%H:%M:%S")
        try:
            df = leer_operaciones()
            estado_act = obtener_estado_firmas(df)

            if estado_ant:
                nuevas = detectar_firmas_nuevas(estado_ant, estado_act, df)
                if nuevas:
                    print(f"\n  [{ahora}] 🖊️  {len(nuevas)} firma(s) nueva(s) detectada(s):")
                    for info in nuevas:
                        row = info["row"]
                        nombre = v(row, "NOMBRE COMPLETO", "?")
                        print(f"    Op.{info['op_num']} | {nombre[:35]:35} | {info['fecha_firma']}")
                        try:
                            adjunto = generar_detalle_xlsx(info["op_num"], df)
                            enviar_email_firma(info, adjunto)
                        except Exception as e:
                            print(f"    ⚠️  Error Op.{info['op_num']}: {e}")
                else:
                    print(f"  [{ahora}] Sin firmas nuevas — {len(df)} operaciones activas")
            else:
                print(f"  [{ahora}] Base establecida: {len(df)} operaciones. Monitoreando...")

            with open(ESTADO_PATH, "w", encoding="utf-8") as f:
                json.dump(estado_act, f, ensure_ascii=False, indent=2)
            estado_ant = estado_act

        except FileNotFoundError:
            print(f"  [{ahora}] ⚠️  No se encuentra: {EXCEL_PATH}")
        except Exception as e:
            print(f"  [{ahora}] ❌ Error: {e}")

        time.sleep(INTERVALO_SEG)

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Monitor Escrituración Tres Pascualas")
    parser.add_argument("--generar", type=int, metavar="N", help="Genera detalle xlsx de la op. N°")
    parser.add_argument("--test-email", action="store_true",  help="Envía email de prueba")
    args = parser.parse_args()

    if args.generar:
        df = leer_operaciones()
        if args.generar in df.index:
            out = generar_detalle_xlsx(args.generar, df)
            print(f"✅ Detalle generado: {out}")
        else:
            print(f"❌ Operación {args.generar} no encontrada")
    elif args.test_email:
        df = leer_operaciones()
        info = {"op_num": 1, "row": df.loc[1], "fecha_firma": datetime.now().strftime("%d/%m/%Y")}
        adjunto = generar_detalle_xlsx(1, df)
        enviar_email_firma(info, adjunto)
        print("✅ Email de prueba enviado.")
    else:
        monitorear()
