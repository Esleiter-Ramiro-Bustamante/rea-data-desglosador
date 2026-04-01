"""
main.py — ReaDesF1.9 — Validador Fiscal Adaptativo
===============================================================================

FLUJO COMPLETO v1.9 (100% GUI):

  INICIO
    ↓
  ASISTENTE GRÁFICO (UI)
    • Seleccionar archivo Excel fuente
    • Elegir Mes y Año del reporte
    • Aceptar / Rechazar Log de Auditoría y Privacidad
    ↓
  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  PASO 1 — SELECCIÓN DEL ARCHIVO
    • Validar que la carpeta y el archivo existen
    • Mostrar archivos disponibles si no se encuentra
  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    ↓
  🖥️  ANALIZAR HARDWARE    (RAM, CPU, SO)
    ↓
  📁 ANALIZAR ARCHIVO     (MB, filas, columnas)
    ↓
  🧠 SELECCIONAR MOTOR:
     ┌─────────────────────────────────────────────┐
     │ Filas < 5,000        → openpyxl   SEGURO    │
     │ 5k-30k + RAM < 4GB   → openpyxl   SEGURO    │
     │ 5k-30k + RAM ≥ 4GB   → pandas     TURBO     │
     │ >30k   + RAM < 4GB   → openpyxl   MÍNIMO    │
     │ >30k   + RAM ≥ 4GB   → chunks     CHUNKS    │
     │ RAM < 2GB (siempre)  → openpyxl   MÍNIMO    │
     └─────────────────────────────────────────────┘
    ↓
  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  PASO 2 — PROCESAR FACTURAS → guardar _validado.xlsx
    • Valida reglas fiscales
    • Escribe fórmulas auditables (sub1/sub0/sub2/iva_acred)
    • Aplica colores por régimen y deducibilidad
    • Si MemoryError → fallback automático a openpyxl
  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    ↓
  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  PASO 3 — GENERAR REPORTES desde _validado.xlsx
    • Lee todos los campos del CFDI (sin omisiones)
    • Detecta régimen dominante (612 / 626 / otros)
    • Detecta complementos CP01 y PPD pendientes
    • Genera _reporte.xlsx  — tabla con fórmulas auditables
    • Genera _reporte.html  — dashboard interactivo con filtros,
                              colores, búsqueda y estatus editables
  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    ↓
  Mostrar resumen completo
    ↓
  FIN

ARCHIVOS GENERADOS:
  NOMBRE_validado.xlsx   → Facturas validadas con fórmulas y colores
  NOMBRE_reporte.xlsx    → Reporte fiscal resumido (tabla auditable)
  NOMBRE_reporte.html    → Dashboard interactivo con filtros y búsqueda

ESTRUCTURA:
  ReaDesF1.8/
  ├── main.py                  ← EJECUTAR ESTE
  ├── generador_reporte.py     ← Paso 3 (se llama automáticamente)
  ├── analizador_sistema.py    ← RAM + CPU + archivo → motor
  ├── motor_openpyxl.py        ← Computadoras básicas
  ├── motor_pandas.py          ← Computadoras potentes (vectorizado)
  ├── motor_chunks.py          ← Archivos muy grandes (>30k filas)
  ├── validaciones_fiscales.py ← Reglas fiscales + fórmulas auditables
  └── seguridad.py             ← Privacidad y auditoría

===============================================================================
Versión: 1.9 | México | Marzo 2026
===============================================================================
"""

import os
import re
import time

from seguridad          import ConfiguracionSeguridad, LogAuditoria
from analizador_sistema import analizar_y_decidir

REGEX_XLSX = re.compile(r'\.xlsx$', re.IGNORECASE)

# ══════════════════════════════════════════════════════════════════
# ENCABEZADO
# ══════════════════════════════════════════════════════════════════

print("\n" + "=" * 70)
print("  ReaDesF1.9 — Validador Fiscal Adaptativo")
print("  3 Pasos: Selección → Validación → Reporte HTML + Excel")
print("=" * 70)

# Eliminamos advertencia inicial de CLI para usar GUI
log = LogAuditoria()

print("\n" + "─" * 70)
print("  PASO 1 — ASISTENTE GRÁFICO (Sinergia REA)")
print("─" * 70)

def mostrar_asistente_gui():
    import tkinter as tk
    from tkinter import filedialog, ttk, messagebox
    from datetime import datetime

    resultado = {'file_path': '', 'mes': '', 'log': True}

    # ── Paleta ────────────────────────────────────────────────────
    C_BG       = "#F4F4F8"
    C_CARD     = "#FFFFFF"
    C_HEADER   = "#1C1C28"
    C_BORDER   = "#E0E0E8"
    C_ROSA     = "#FF2D78"
    C_AMARILLO = "#FFD600"
    C_AZUL     = "#0057FF"
    C_VERDE    = "#16A34A"
    C_TX       = "#1A1A2E"
    C_GRIS_TX  = "#5A5A6E"

    root = tk.Tk()
    root.withdraw()                          # Ocultar mientras se construye
    root.title("Sinergia REA — ReaDesF 1.9")
    root.configure(bg=C_BG)
    root.resizable(False, False)

    W, H = 500, 570
    root.geometry(f"{W}x{H}")

    # ══════════════════════════════════════════════════════════════
    # HEADER — Canvas para texto multicolor sin separaciones raras
    # ══════════════════════════════════════════════════════════════
    HDR_H = 110
    canvas_hdr = tk.Canvas(root, width=W, height=HDR_H,
                           bg=C_HEADER, highlightthickness=0)
    canvas_hdr.pack(fill="x")

    def _draw_header(event=None):
        canvas_hdr.delete("all")
        cw = canvas_hdr.winfo_width() or W

        # ── SINERGIA REA — línea 1 ────────────────────────────────
        # Calcular posición centrada manualmente
        # Fuente grande: aproximamos anchos por caracter
        y1 = 34
        # Dibujar con create_text en posiciones relativas
        # Usamos anchor="w" y avanzamos x manualmente
        fnt_big = ("Segoe UI", 26, "bold")
        fnt_sm  = ("Segoe UI", 12, "bold")
        fnt_ver = ("Segoe UI",  9)
        fnt_bdg = ("Segoe UI",  7)

        # Medir texto creando temporales invisibles
        def _tw(txt, fnt):
            t = canvas_hdr.create_text(-999, -999, text=txt, font=fnt)
            bb = canvas_hdr.bbox(t)
            canvas_hdr.delete(t)
            return (bb[2]-bb[0]) if bb else 0

        # SINERGIA REA
        parts1 = [("SIN", C_ROSA), ("ERGIA ", "#E0E0DC"), ("REA", C_AMARILLO)]
        tw1    = sum(_tw(t, fnt_big) for t,_ in parts1)
        x1     = (cw - tw1) // 2
        for txt, col in parts1:
            canvas_hdr.create_text(x1, y1, text=txt, font=fnt_big,
                                   fill=col, anchor="nw")
            x1 += _tw(txt, fnt_big)

        # REA DES F 1.9
        y2     = y1 + 36
        parts2 = [("REA", C_ROSA), ("DES", "#E0E0DC"), ("F", C_AMARILLO)]
        tw2    = sum(_tw(t, fnt_sm) for t,_ in parts2) + _tw(" 1.9", fnt_ver) + 4
        x2     = (cw - tw2) // 2
        for txt, col in parts2:
            canvas_hdr.create_text(x2, y2, text=txt, font=fnt_sm,
                                   fill=col, anchor="nw")
            x2 += _tw(txt, fnt_sm)
        canvas_hdr.create_text(x2+4, y2+2, text="1.9", font=fnt_ver,
                               fill=C_AZUL, anchor="nw")

        # Badge
        canvas_hdr.create_text(cw//2, y2+26,
                               text="V A L I D A D O R   F I S C A L   A D A P T A T I V O",
                               font=fnt_bdg, fill="#4A4A5A", anchor="center")

        # Separador tricolor en la base del header
        t3 = cw // 3
        canvas_hdr.create_rectangle(0,    HDR_H-3, t3,    HDR_H, fill=C_ROSA,     outline="")
        canvas_hdr.create_rectangle(t3,   HDR_H-3, t3*2,  HDR_H, fill=C_AMARILLO, outline="")
        canvas_hdr.create_rectangle(t3*2, HDR_H-3, cw,    HDR_H, fill=C_AZUL,     outline="")

    canvas_hdr.bind("<Configure>", _draw_header)
    root.after(30, _draw_header)

    # ══════════════════════════════════════════════════════════════
    # CONTENIDO PRINCIPAL
    # ══════════════════════════════════════════════════════════════
    frame_main = tk.Frame(root, bg=C_BG)
    frame_main.pack(fill="both", expand=True, padx=32, pady=(14, 0))

    F_LABEL  = ("Segoe UI", 9,  "bold")
    F_BODY   = ("Segoe UI", 10)
    F_BTN    = ("Segoe UI", 10, "bold")
    F_BTN_LG = ("Segoe UI", 12, "bold")

    def _section_label(parent, numero, texto, color=C_ROSA):
        f = tk.Frame(parent, bg=C_BG)
        f.pack(fill="x", pady=(10, 5))
        tk.Label(f, text=f"{numero}.", font=F_LABEL,
                 bg=C_BG, fg=color).pack(side="left", padx=(0,5))
        tk.Label(f, text=texto, font=F_LABEL,
                 bg=C_BG, fg=color).pack(side="left")

    def _card(parent, border_color=C_BORDER):
        outer = tk.Frame(parent, bg=border_color)
        outer.pack(fill="x", ipady=0)
        # Barra lateral de color
        bar = tk.Frame(outer, bg=border_color, width=4)
        bar.pack(side="left", fill="y")
        inner = tk.Frame(outer, bg=C_CARD)
        inner.pack(side="left", fill="both", expand=True)
        return outer, bar, inner

    # ── 1. ARCHIVO ────────────────────────────────────────────────
    _section_label(frame_main, "1", "ARCHIVO EXCEL A PROCESAR", C_ROSA)

    card_f, bar_f, inner_f = _card(frame_main, C_BORDER)

    file_var = tk.StringVar(value="Ningún archivo seleccionado")
    lbl_file = tk.Label(inner_f, textvariable=file_var,
                        font=F_BODY, bg=C_CARD, fg=C_GRIS_TX,
                        anchor="w", padx=12, pady=10,
                        wraplength=360, justify="left")
    lbl_file.pack(fill="x", expand=True)

    def select_file():
        ruta = filedialog.askopenfilename(
            title="Seleccionar Excel",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if ruta:
            resultado['file_path'] = ruta
            file_var.set(f"📂  {os.path.basename(ruta)}")
            lbl_file.config(fg=C_VERDE)
            bar_f.config(bg=C_VERDE)
            card_f.config(bg=C_VERDE)

    tk.Button(frame_main, text="BUSCAR ARCHIVO",
              font=F_BTN, bg=C_ROSA, fg="#FFF",
              activebackground="#CC1F5F", activeforeground="#FFF",
              relief="flat", cursor="hand2", padx=14, pady=6,
              command=select_file).pack(anchor="e", pady=(5,0))

    # ── 2. PERIODO ────────────────────────────────────────────────
    _section_label(frame_main, "2", "PERIODO DEL REPORTE", C_AMARILLO)

    card_d, bar_d, inner_d = _card(frame_main, C_AMARILLO)
    bar_d.config(bg=C_AMARILLO)

    frame_sel = tk.Frame(inner_d, bg=C_CARD, padx=12, pady=10)
    frame_sel.pack(fill="x")

    meses = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
             "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
    cur_m = datetime.now().month
    cur_y = datetime.now().year

    # ── Selectores desplegables mejorados (Combobox) ────────────
    mes_var    = tk.StringVar(value=meses[cur_m - 1])
    year_var   = tk.StringVar(value=str(cur_y))
    preview_var = tk.StringVar(value=f"\U0001f4c5  {meses[cur_m-1]} {cur_y}")

    def _update_preview(*_):
        preview_var.set(f"\U0001f4c5  {mes_var.get()}  {year_var.get()}")

    mes_var.trace_add("write",  _update_preview)
    year_var.trace_add("write", _update_preview)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("REA.TCombobox",
                    fieldbackground=C_CARD,
                    background="#D0D0D8",
                    foreground=C_TX,
                    selectbackground=C_AMARILLO,
                    selectforeground="#0D0D1A",
                    padding=(10, 8),
                    relief="flat",
                    borderwidth=0,
                    arrowsize=14)
    style.map("REA.TCombobox",
              fieldbackground=[("readonly", C_CARD)],
              foreground=[("readonly", C_TX)],
              background=[("active", C_AMARILLO)])

    # Columna MES
    col_mes = tk.Frame(frame_sel, bg=C_CARD)
    col_mes.pack(side="left", padx=(0, 14))
    tk.Label(col_mes, text="MES", font=("Segoe UI", 7, "bold"),
             bg=C_CARD, fg=C_GRIS_TX).pack(anchor="w", padx=2)
    frm_mes = tk.Frame(col_mes, bg=C_AMARILLO, padx=2, pady=2)
    frm_mes.pack()
    cb_mes = ttk.Combobox(frm_mes, textvariable=mes_var,
                          values=meses, state="readonly",
                          style="REA.TCombobox", width=13,
                          font=("Segoe UI", 10, "bold"))
    cb_mes.pack()

    # Columna AÑO
    col_year = tk.Frame(frame_sel, bg=C_CARD)
    col_year.pack(side="left")
    tk.Label(col_year, text="A\u00d1O", font=("Segoe UI", 7, "bold"),
             bg=C_CARD, fg=C_GRIS_TX).pack(anchor="w", padx=2)
    frm_year = tk.Frame(col_year, bg=C_AMARILLO, padx=2, pady=2)
    frm_year.pack()
    años = [str(y) for y in range(2020, 2027)]
    cb_year = ttk.Combobox(frm_year, textvariable=year_var,
                           values=años, state="readonly",
                           style="REA.TCombobox", width=6,
                           font=("Segoe UI", 10, "bold"))
    cb_year.pack()

    # Preview del periodo seleccionado
    tk.Label(inner_d, textvariable=preview_var,
             font=("Segoe UI", 9, "bold"), bg=C_CARD,
             fg=C_AMARILLO, padx=12, pady=(0)).pack(anchor="w", padx=12, pady=(0,8))

    # ── 3. LOG DE AUDITORÍA ───────────────────────────────────────
    _section_label(frame_main, "3", "LOG DE AUDITORÍA", C_AZUL)

    card_l, bar_l, inner_l = _card(frame_main, C_AZUL)
    bar_l.config(bg=C_AZUL)

    frame_log = tk.Frame(inner_l, bg=C_CARD, padx=12, pady=10)
    frame_log.pack(fill="x")

    log_var = tk.BooleanVar(value=True)
    frm_chk = tk.Frame(frame_log, bg=C_CARD)
    frm_chk.pack(fill="x")
    tk.Checkbutton(frm_chk, text="  \U0001f510  Generar y guardar Log de Auditor\u00eda",
                   variable=log_var, bg=C_CARD, fg=C_TX,
                   selectcolor=C_AZUL, activebackground=C_CARD,
                   activeforeground=C_TX, font=("Segoe UI", 10, "bold"),
                   cursor="hand2").pack(anchor="w")
    tk.Label(frame_log,
             text="        Registro local de operaciones — sin env\u00edo de datos",
             font=("Segoe UI", 8), bg=C_CARD, fg=C_GRIS_TX).pack(anchor="w")

    # ── Diálogos custom Sinergia REA ─────────────────────────────
    def _rea_header_draw(canvas, W, HDR_H):
        canvas.delete("all")
        cw = canvas.winfo_width() or W
        C_ROSA2     = "#FF2D78"
        C_AMARILLO2 = "#FFD600"
        C_AZUL2     = "#0057FF"
        fnt_big = ("Segoe UI", 15, "bold")
        fnt_sm  = ("Segoe UI",  9, "bold")
        fnt_ver = ("Segoe UI",  7)
        fnt_bdg = ("Segoe UI",  6)

        def _tw(txt, fnt):
            t = canvas.create_text(-999, -999, text=txt, font=fnt)
            bb = canvas.bbox(t)
            canvas.delete(t)
            return (bb[2]-bb[0]) if bb else 0

        parts1 = [("SIN", C_ROSA2), ("ERGIA ", "#E0E0DC"), ("REA", C_AMARILLO2)]
        tw1    = sum(_tw(t, fnt_big) for t, _ in parts1)
        x1, y1 = (cw - tw1) // 2, 8
        for txt, col in parts1:
            canvas.create_text(x1, y1, text=txt, font=fnt_big, fill=col, anchor="nw")
            x1 += _tw(txt, fnt_big)

        y2     = y1 + 22
        parts2 = [("REA", C_ROSA2), ("DES", "#E0E0DC"), ("F", C_AMARILLO2)]
        tw2    = sum(_tw(t, fnt_sm) for t, _ in parts2) + _tw(" 1.9", fnt_ver) + 3
        x2     = (cw - tw2) // 2
        for txt, col in parts2:
            canvas.create_text(x2, y2, text=txt, font=fnt_sm, fill=col, anchor="nw")
            x2 += _tw(txt, fnt_sm)
        canvas.create_text(x2+3, y2+2, text="1.9", font=fnt_ver, fill=C_AZUL2, anchor="nw")

        canvas.create_text(cw//2, y2+18,
            text="V A L I D A D O R   F I S C A L   A D A P T A T I V O",
            font=fnt_bdg, fill="#4A4A5A", anchor="center")

        t3 = cw // 3
        canvas.create_rectangle(0,    HDR_H-3, t3,    HDR_H, fill=C_ROSA2,     outline="")
        canvas.create_rectangle(t3,   HDR_H-3, t3*2,  HDR_H, fill=C_AMARILLO2, outline="")
        canvas.create_rectangle(t3*2, HDR_H-3, cw,    HDR_H, fill=C_AZUL2,     outline="")

    def _dialogo_aviso(titulo, mensaje):
        """Aviso con diseño Sinergia REA — botón ACEPTAR."""
        dw = tk.Toplevel(root)
        dw.withdraw()
        dw.title(titulo)
        dw.configure(bg=C_BG)
        dw.resizable(False, False)
        dw.grab_set()
        DW, DH, HDR = 430, 290, 80
        dw.geometry(f"{DW}x{DH}")
        cv = tk.Canvas(dw, width=DW, height=HDR, bg=C_HEADER, highlightthickness=0)
        cv.pack(fill="x")
        cv.bind("<Configure>", lambda e: _rea_header_draw(cv, DW, HDR))
        dw.after(40, lambda: _rea_header_draw(cv, DW, HDR))
        fb = tk.Frame(dw, bg=C_BG)
        fb.pack(fill="both", expand=True, padx=22, pady=(14,0))
        fr = tk.Frame(fb, bg=C_BG)
        fr.pack(fill="x", pady=(0,8))
        tk.Label(fr, text="⚠️", font=("Segoe UI", 20),
                 bg=C_BG, fg=C_AMARILLO).pack(side="left", padx=(0,10))
        tk.Label(fr, text=titulo, font=("Segoe UI", 11, "bold"),
                 bg=C_BG, fg=C_AMARILLO).pack(side="left", anchor="s", pady=(4,0))
        tk.Frame(fb, bg=C_AMARILLO, height=2).pack(fill="x", pady=(0,10))
        card = tk.Frame(fb, bg=C_CARD, highlightbackground=C_BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True)
        tk.Label(card, text=mensaje, font=("Segoe UI", 9),
                 bg=C_CARD, fg=C_TX, justify="left", anchor="nw",
                 wraplength=360, padx=14, pady=12).pack(fill="both", expand=True)
        tk.Frame(dw, bg=C_BORDER, height=1).pack(fill="x", pady=(8,0))
        tk.Button(dw, text="  ACEPTAR  ",
                  font=("Segoe UI", 10, "bold"),
                  bg=C_AMARILLO, fg="#0D0D1A",
                  activebackground="#E6C200", activeforeground="#0D0D1A",
                  relief="flat", cursor="hand2", padx=22, pady=7,
                  command=dw.destroy).pack(pady=10)
        dw.update_idletasks()
        sw2 = dw.winfo_screenwidth(); sh2 = dw.winfo_screenheight()
        dw.geometry(f"{DW}x{DH}+{(sw2-DW)//2}+{(sh2-DH)//2}")
        dw.deiconify()
        dw.attributes("-topmost", True)
        dw.wait_window()

    def _dialogo_confirmar(titulo, mensaje):
        """Confirmación SÍ/NO con diseño Sinergia REA. Retorna True/False."""
        respuesta = [False]
        dw = tk.Toplevel(root)
        dw.withdraw()
        dw.title(titulo)
        dw.configure(bg=C_BG)
        dw.resizable(False, False)
        dw.grab_set()
        DW, DH, HDR = 430, 310, 80
        dw.geometry(f"{DW}x{DH}")
        cv = tk.Canvas(dw, width=DW, height=HDR, bg=C_HEADER, highlightthickness=0)
        cv.pack(fill="x")
        cv.bind("<Configure>", lambda e: _rea_header_draw(cv, DW, HDR))
        dw.after(40, lambda: _rea_header_draw(cv, DW, HDR))
        fb = tk.Frame(dw, bg=C_BG)
        fb.pack(fill="both", expand=True, padx=22, pady=(14,0))
        fr = tk.Frame(fb, bg=C_BG)
        fr.pack(fill="x", pady=(0,8))
        tk.Label(fr, text="🔐", font=("Segoe UI", 20),
                 bg=C_BG, fg=C_AZUL).pack(side="left", padx=(0,10))
        tk.Label(fr, text=titulo, font=("Segoe UI", 11, "bold"),
                 bg=C_BG, fg=C_AZUL).pack(side="left", anchor="s", pady=(4,0))
        tk.Frame(fb, bg=C_AZUL, height=2).pack(fill="x", pady=(0,10))
        card = tk.Frame(fb, bg=C_CARD, highlightbackground=C_BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True)
        tk.Label(card, text=mensaje, font=("Segoe UI", 9),
                 bg=C_CARD, fg=C_TX, justify="left", anchor="nw",
                 wraplength=360, padx=14, pady=12).pack(fill="both", expand=True)
        tk.Frame(dw, bg=C_BORDER, height=1).pack(fill="x", pady=(8,0))
        fr_btns = tk.Frame(dw, bg=C_BG)
        fr_btns.pack(pady=10)
        def _si():
            respuesta[0] = True; dw.destroy()
        def _no():
            respuesta[0] = False; dw.destroy()
        tk.Button(fr_btns, text="  SÍ  ",
                  font=("Segoe UI", 10, "bold"),
                  bg=C_VERDE, fg="#FFFFFF",
                  activebackground="#15803D", activeforeground="#FFF",
                  relief="flat", cursor="hand2", padx=20, pady=7,
                  command=_si).pack(side="left", padx=(0,10))
        tk.Button(fr_btns, text="  NO  ",
                  font=("Segoe UI", 10, "bold"),
                  bg=C_ROSA, fg="#FFFFFF",
                  activebackground="#CC1F5F", activeforeground="#FFF",
                  relief="flat", cursor="hand2", padx=20, pady=7,
                  command=_no).pack(side="left")
        dw.update_idletasks()
        sw2 = dw.winfo_screenwidth(); sh2 = dw.winfo_screenheight()
        dw.geometry(f"{DW}x{DH}+{(sw2-DW)//2}+{(sh2-DH)//2}")
        dw.deiconify()
        dw.attributes("-topmost", True)
        dw.wait_window()
        return respuesta[0]

    # ── Separador + Botón INICIAR ─────────────────────────────────
    tk.Frame(root, bg=C_BORDER, height=1).pack(fill="x", pady=(14,0))

    def on_iniciar():
        if not resultado['file_path']:
            _dialogo_aviso("Atención",
                           "Selecciona un archivo Excel primero.")
            return
        if _dialogo_confirmar("Confirmación LFPDPPP",
            "Sus datos se procesarán localmente\nsin uso de internet.\n\n"
            "¿Desea iniciar la validación fiscal del documento?"):
            resultado['mes'] = f"{cb_mes.get()} {cb_year.get()}"
            resultado['log'] = log_var.get()
            root.destroy()

    tk.Button(root, text="INICIAR PROCESO",
              font=F_BTN_LG, bg=C_AMARILLO, fg="#0D0D1A",
              activebackground="#E6C200", activeforeground="#0D0D1A",
              relief="flat", cursor="hand2", padx=32, pady=12,
              command=on_iniciar).pack(pady=14)

    tk.Label(root, text="Sinergia REA  ·  100% Local  ·  LFPDPPP",
             font=("Segoe UI", 7), bg=C_BG, fg=C_BORDER).pack(pady=(0,8))

    # ── Centrado real DESPUÉS de construir todo ─────────────────
    root.update_idletasks()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x  = (sw - W) // 2
    y  = (sh - H) // 2
    root.geometry(f"{W}x{H}+{x}+{y}")
    root.deiconify()                     # Mostrar ya centrada

    root.attributes("-topmost", True)
    root.focus_force()
    root.mainloop()
    return resultado

print("\n  Abriendo Asistente Interactivo...")
conf = mostrar_asistente_gui()

if not conf['file_path']:
    print("❌ Proceso cancelado.")
    raise SystemExit(1)

file_path = conf['file_path']
mes_reporte = conf['mes']
ConfiguracionSeguridad.CREAR_LOG_AUDITORIA = conf['log']
desktop_path = os.path.dirname(file_path)
file_name = os.path.basename(file_path)

print(f"\n  ✅ Archivo    : {file_path}")
print(f"  ✅ Mes        : {mes_reporte}")
print(f"  ✅ Guardar Log: {'SÍ' if conf['log'] else 'NO'}")

# ══════════════════════════════════════════════════════════════════
# PASO 2 — VALIDACIÓN Y GENERACIÓN DE _validado.xlsx
# ══════════════════════════════════════════════════════════════════

print("\n" + "─" * 70)
print("  PASO 2 — PROCESANDO FACTURAS")
print("─" * 70)

analisis = analizar_y_decidir(file_path)
log.registrar_inicio(file_path, analisis.motor)

base_name   = REGEX_XLSX.sub('', str(file_name))
suffix      = "_ANONIMIZADO_validado" if ConfiguracionSeguridad.MODO_ANONIMIZAR else "_validado"
output_path = os.path.join(str(desktop_path), f"{base_name}{suffix}.xlsx")

t_inicio = time.time()
stats    = {}

iconos = {'TURBO': '🚀', 'CHUNKS': '📦', 'SEGURO': '🔧', 'MÍNIMO': '🐢'}
print(f"\n  {iconos.get(analisis.modo, '⚙️')} Motor: "
      f"{analisis.motor.upper()} — {analisis.modo}\n")

def mostrar_mensaje_final(titulo, mensaje, tipo="info"):
    import tkinter as tk

    # ── Paleta idéntica al asistente principal ─────────────────────
    C_BG       = "#F4F4F8"
    C_CARD     = "#FFFFFF"
    C_HEADER   = "#1C1C28"
    C_BORDER   = "#E0E0E8"
    C_ROSA     = "#FF2D78"
    C_AMARILLO = "#FFD600"
    C_AZUL     = "#0057FF"
    C_VERDE    = "#16A34A"
    C_ROJO     = "#DC2626"
    C_TX       = "#1A1A2E"
    C_GRIS_TX  = "#5A5A6E"

    es_error   = (tipo == "error")
    COLOR_ACNT = C_ROJO if es_error else C_VERDE
    ICONO      = "\u274c" if es_error else "\u2705"

    win = tk.Tk()
    win.withdraw()
    win.title(titulo)
    win.configure(bg=C_BG)
    win.resizable(False, False)

    W, H = 440, 390
    win.geometry(f"{W}x{H}")

    # ── Header ────────────────────────────────────────────────────
    HDR_H = 88
    canvas_hdr = tk.Canvas(win, width=W, height=HDR_H,
                           bg=C_HEADER, highlightthickness=0)
    canvas_hdr.pack(fill="x")

    def _draw_hdr(event=None):
        canvas_hdr.delete("all")
        cw = canvas_hdr.winfo_width() or W
        fnt_big = ("Segoe UI", 16, "bold")
        fnt_sm  = ("Segoe UI", 10, "bold")
        fnt_ver = ("Segoe UI",  8)
        fnt_bdg = ("Segoe UI",  6)

        def _tw(txt, fnt):
            t = canvas_hdr.create_text(-999, -999, text=txt, font=fnt)
            bb = canvas_hdr.bbox(t)
            canvas_hdr.delete(t)
            return (bb[2]-bb[0]) if bb else 0

        # SINERGIA REA — línea 1
        parts1 = [("SIN", C_ROSA), ("ERGIA ", "#E0E0DC"), ("REA", C_AMARILLO)]
        tw1    = sum(_tw(t, fnt_big) for t, _ in parts1)
        x1, y1 = (cw - tw1) // 2, 10
        for txt, col in parts1:
            canvas_hdr.create_text(x1, y1, text=txt, font=fnt_big,
                                   fill=col, anchor="nw")
            x1 += _tw(txt, fnt_big)

        # READESF 1.9 — línea 2
        y2     = y1 + 26
        parts2 = [("REA", C_ROSA), ("DES", "#E0E0DC"), ("F", C_AMARILLO)]
        tw2    = sum(_tw(t, fnt_sm) for t, _ in parts2) + _tw(" 1.9", fnt_ver) + 4
        x2     = (cw - tw2) // 2
        for txt, col in parts2:
            canvas_hdr.create_text(x2, y2, text=txt, font=fnt_sm,
                                   fill=col, anchor="nw")
            x2 += _tw(txt, fnt_sm)
        canvas_hdr.create_text(x2 + 3, y2 + 2, text="1.9", font=fnt_ver,
                               fill=C_AZUL, anchor="nw")

        # Badge
        canvas_hdr.create_text(cw // 2, y2 + 20,
                               text="V A L I D A D O R   F I S C A L   A D A P T A T I V O",
                               font=fnt_bdg, fill="#4A4A5A", anchor="center")

        # Tricolor
        t3 = cw // 3
        canvas_hdr.create_rectangle(0,    HDR_H-3, t3,    HDR_H, fill=C_ROSA,     outline="")
        canvas_hdr.create_rectangle(t3,   HDR_H-3, t3*2,  HDR_H, fill=C_AMARILLO, outline="")
        canvas_hdr.create_rectangle(t3*2, HDR_H-3, cw,    HDR_H, fill=C_AZUL,     outline="")

    canvas_hdr.bind("<Configure>", _draw_hdr)
    win.after(30, _draw_hdr)

    # ── Cuerpo ────────────────────────────────────────────────────
    frame_body = tk.Frame(win, bg=C_BG)
    frame_body.pack(fill="both", expand=True, padx=28, pady=(16, 0))

    # Franja de color + ícono
    frm_top = tk.Frame(frame_body, bg=C_BG)
    frm_top.pack(fill="x", pady=(0, 10))

    tk.Label(frm_top, text=ICONO, font=("Segoe UI", 26),
             bg=C_BG, fg=COLOR_ACNT).pack(side="left", padx=(0, 12))

    tk.Label(frm_top, text=titulo,
             font=("Segoe UI", 13, "bold"),
             bg=C_BG, fg=COLOR_ACNT).pack(side="left", anchor="s", pady=(6,0))

    # Separador de acento
    tk.Frame(frame_body, bg=COLOR_ACNT, height=2).pack(fill="x", pady=(0, 12))

    # Card con mensaje
    card = tk.Frame(frame_body, bg=C_CARD,
                    highlightbackground=C_BORDER, highlightthickness=1)
    card.pack(fill="both", expand=True)

    tk.Label(card, text=mensaje,
             font=("Segoe UI", 9), bg=C_CARD, fg=C_TX,
             justify="left", anchor="nw",
             wraplength=370, padx=16, pady=14).pack(fill="both", expand=True)

    # ── Separador + Botón ─────────────────────────────────────────
    tk.Frame(win, bg=C_BORDER, height=1).pack(fill="x", pady=(10, 0))

    tk.Button(win, text="ACEPTAR",
              font=("Segoe UI", 10, "bold"),
              bg=C_AMARILLO, fg="#0D0D1A",
              activebackground="#E6C200", activeforeground="#0D0D1A",
              relief="flat", cursor="hand2", padx=28, pady=8,
              command=win.destroy).pack(pady=12)

    # ── Centrar y mostrar ─────────────────────────────────────────
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")
    win.deiconify()
    win.attributes("-topmost", True)
    win.focus_force()
    win.mainloop()

try:
    if analisis.motor == 'pandas':
        from motor_pandas import procesar_con_pandas
        stats = procesar_con_pandas(file_path, output_path, analisis.modo)

    elif analisis.motor == 'pandas_chunks':
        from motor_chunks import procesar_con_chunks
        stats = procesar_con_chunks(file_path, output_path, analisis.chunk_size)

    else:
        from motor_openpyxl import procesar_con_openpyxl
        stats = procesar_con_openpyxl(file_path, output_path, analisis.modo)

except MemoryError:
    print("\n⚠️  MEMORIA INSUFICIENTE — cambiando a openpyxl automáticamente...")
    from motor_openpyxl import procesar_con_openpyxl
    stats = procesar_con_openpyxl(file_path, output_path, 'SEGURO')
    analisis.motor = 'openpyxl (fallback)'
    analisis.modo  = 'SEGURO'

except Exception as e:
    print(f"\n❌ Error en Paso 2: {e}")
    log.registrar_error(e)
    log.guardar_log()
    mostrar_mensaje_final("Error en Proceso", f"Ocurrió un error en la validación:\n\n{e}", "error")
    raise SystemExit(1)

t_validacion = time.time() - t_inicio
velocidad    = analisis.filas_reales / t_validacion if t_validacion > 0 else 0
log.registrar_fin(output_path, analisis.filas_reales, t_validacion, analisis.motor)

print(f"\n  ✅ _validado.xlsx generado en {t_validacion:.2f}s")
print(f"  ⚡ {velocidad:,.0f} facturas/segundo")

# ── Resumen de validación ─────────────────────────────────────────
print(f"\n  📋 REGÍMENES:")
for reg, cnt in sorted(stats.get('regimenes', {}).items()):
    print(f"    • Régimen {reg}: {cnt:,} facturas")

print(f"\n  🍬 Dulces/Botanas:  Con IEPS: {stats.get('dulces_ieps8', 0):,}  |  Sin: {stats.get('dulces_sin', 0):,}")
print(f"\n  ⛽ Gasolina:")
print(f"    Con IEPS   : {stats.get('gas_ieps', 0):,}   Sin IEPS: {stats.get('gas_sin', 0):,}")
print(f"    626 efect. : {stats.get('gas_626', 0):,}   (agrup: {stats.get('gas_626_agrup', 0):,})")
print(f"    612 efect. : {stats.get('gas_612', 0):,}   (rechazadas)")
print(f"    Electrónica: {stats.get('gas_elec', 0):,}")
print(f"\n  🌱 Insumos Agrícolas:")
print(f"    ❌ Efectivo >$2,000  : {stats.get('ins_nd', 0):,}")
print(f"    ✅ Efectivo ≤$2,000  : {stats.get('ins_menor', 0):,}")
print(f"    ✅ Pago electrónico  : {stats.get('ins_elec', 0):,}")
print(f"\n  📋 General:")
print(f"    S01 (sin efectos)  : {stats.get('s01', 0):,}")
print(f"    Efectivo >$2,000   : {stats.get('ef_mayor', 0):,}")

# ══════════════════════════════════════════════════════════════════
# PASO 3 — GENERAR REPORTES HTML + EXCEL desde _validado.xlsx
# ══════════════════════════════════════════════════════════════════

print("\n" + "─" * 70)
print("  PASO 3 — GENERANDO REPORTES")
print("─" * 70)

try:
    from generador_reporte import generar_reporte
    generar_reporte(output_path, mes_reporte=mes_reporte)

except ImportError:
    msg = (
        "No se encontró el archivo generador_reporte.py.\n\n"
        "✅ El archivo _validado.xlsx SÍ fue generado correctamente.\n\n"
        "📋 Solución:\n"
        "   Asegúrate de que generador_reporte.py esté\n"
        "   en la misma carpeta que main.py."
    )
    print("\n⚠️  No se encontró generador_reporte.py")
    print(f"    _validado.xlsx generado: {output_path}")
    mostrar_mensaje_final("Paso 3 — Archivo faltante", msg, "error")

except PermissionError:
    msg = (
        "No se pudo guardar el reporte.\n\n"
        "✅ El archivo _validado.xlsx SÍ fue generado.\n\n"
        "📋 Posibles causas:\n"
        "   • El archivo _reporte.xlsx está abierto en Excel.\n"
        "   • No tienes permisos de escritura en esa carpeta.\n\n"
        "💡 Cierra Excel y vuelve a intentarlo."
    )
    print("\n⚠️  Error de permisos al guardar reporte")
    mostrar_mensaje_final("Paso 3 — Archivo en uso", msg, "error")

except MemoryError:
    msg = (
        "Memoria insuficiente al generar el reporte.\n\n"
        "✅ El archivo _validado.xlsx SÍ fue generado.\n\n"
        "📋 Solución:\n"
        "   • Cierra otras aplicaciones abiertas.\n"
        "   • Vuelve a ejecutar el programa."
    )
    print("\n⚠️  Memoria insuficiente en Paso 3")
    mostrar_mensaje_final("Paso 3 — Memoria insuficiente", msg, "error")

except Exception as e:
    import traceback
    detalle = traceback.format_exc()
    msg = (
        f"Error al generar los reportes:\n\n{e}\n\n"
        "✅ El archivo _validado.xlsx SÍ fue generado.\n\n"
        "📋 Puedes abrir el _validado.xlsx directamente.\n"
        "   El detalle técnico se guardó en el log."
    )
    print(f"\n⚠️  Error generando reportes: {e}")
    print(f"    _validado.xlsx generado: {output_path}")
    print(detalle)
    log.registrar_error(e)
    mostrar_mensaje_final("Paso 3 — Error en Reporte", msg, "error")

# ══════════════════════════════════════════════════════════════════
# RESUMEN FINAL COMPLETO
# ══════════════════════════════════════════════════════════════════

t_total_global = time.time() - t_inicio

print("\n" + "=" * 70)
print("  ✅ PROCESO COMPLETO — ReaDesF1.9")
print("=" * 70)
print(f"  📂 _validado.xlsx : {output_path}")
base_reporte = output_path.replace('_validado.xlsx', '_reporte')
print(f"  📊 _reporte.xlsx  : {base_reporte}.xlsx")
print(f"  🌐 _reporte.html  : {base_reporte}.html")
print(f"\n  📊 Filas procesadas : {analisis.filas_reales:,}")
print(f"  ⏱️  Tiempo total      : {t_total_global:.2f} segundos")
print(f"  🚗 Motor usado       : {analisis.motor.upper()} — {analisis.modo}")
print(f"  🖥️  RAM disponible    : {analisis.ram_disponible_gb:.1f} GB")

if analisis.motor == 'pandas_chunks':
    bloques = -(-analisis.filas_reales // analisis.chunk_size)
    print(f"  📦 Bloques proc.     : {bloques} × {analisis.chunk_size:,} filas")

print(f"\n  📐 FÓRMULAS AUDITABLES en _validado.xlsx:")
print(f"    sub1      = subtotal - descuento    → Base gravable real")
print(f"    sub0      = iva0 + iva_exento        → Total no gravado")
print(f"    sub2      = sub1 - sub0              → Base para IVA 16%")
print(f"    iva_acred = sub2 × 0.16              → IVA que debería ser")
print(f"    c_iva     = iva_acred - iva16        → Diferencia (0 = correcto)")
print(f"    comprob   = total_cfdi - t2          → Delta total")

print("\n  ⚖️  REGLAS APLICADAS:")
print("    Todos  : Art. 27 Fracc. III — efectivo >$2,000 NO deducible")
print("    626    : Gasolina ≤$2,000 o agrupada = deducible (facilidad)")
print("    612    : Art. 103 + Art. 27 Fracc. III LISR")
print("             Art. 74 AGAPES / Art. 147 = NO aplican a Régimen 612")

print("\n  🎨 COLORES _validado: 🟦 626  🟪 612  🟩 Deducible  🟥 NO ded.")
print("  🎨 COLORES _reporte : 🟩 16%  🟦 16Y0%  🟨 0%  🟥 No Ded  🟣 Complemento")

print("\n" + "=" * 70)
print("  🔐 Información fiscal CONFIDENCIAL — Guardar de forma segura")
print("=" * 70)

log.guardar_log()

msg_fin = (
    "✅ Reporte generado exitosamente.\n\n"
    f"Archivo procesado: {analisis.filas_reales:,} filas\n"
    f"Tiempo total: {t_total_global:.2f} segundos\n\n"
    f"Los archivos resultantes se guardaron en la misma carpeta del original."
)
mostrar_mensaje_final("Proceso Terminado", msg_fin, "info")
