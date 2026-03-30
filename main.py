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

    root = tk.Tk()
    root.title("Sinergia REA — Ajustes")
    root.geometry("500x480")
    root.configure(bg="#0A0A0A")
    try: root.eval('tk::PlaceWindow . center')
    except: pass

    COLOR_NEGRO = "#0A0A0A"
    COLOR_ROSA = "#FF2D78"
    COLOR_AMARILLO = "#FFD600"
    COLOR_BLANCO = "#F7F7F2"
    COLOR_GRIS = "#1A1A1A"

    tk.Label(root, text="ReaDesF 1.9", font=("Helvetica", 24, "bold"), bg=COLOR_NEGRO, fg=COLOR_BLANCO).pack(pady=(20, 0))
    tk.Label(root, text="Asistente de Configuración", font=("Courier", 10, "bold"), bg=COLOR_NEGRO, fg=COLOR_AMARILLO).pack(pady=(0, 20))

    frame_main = tk.Frame(root, bg=COLOR_NEGRO)
    frame_main.pack(fill="both", expand=True, padx=40)

    # 1. Archivo
    tk.Label(frame_main, text="1. ARCHIVO EXCEL A PROCESAR", font=("Century Gothic", 10, "bold"), bg=COLOR_NEGRO, fg=COLOR_ROSA, anchor="w").pack(fill="x", pady=(5, 5))
    file_var = tk.StringVar(value="[ Ningún archivo seleccionado ]")
    lbl_file = tk.Label(frame_main, textvariable=file_var, font=("Helvetica", 9), bg=COLOR_GRIS, fg="#AAAAAA", anchor="w", padx=10, pady=8)
    lbl_file.pack(fill="x")

    def select_file():
        ruta = filedialog.askopenfilename(title="Seleccionar Excel", filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if ruta:
            resultado['file_path'] = ruta
            file_var.set(f"📂 {os.path.basename(ruta)}")
            lbl_file.config(fg="#00FF00")

    tk.Button(frame_main, text="BUSCAR ARCHIVO", font=("Century Gothic", 9, "bold"), bg="#333", fg=COLOR_BLANCO, relief="flat", cursor="hand2", command=select_file).pack(anchor="e", pady=5)

    # 2. Mes del Reporte
    tk.Label(frame_main, text="2. PERIODO DEL REPORTE", font=("Century Gothic", 10, "bold"), bg=COLOR_NEGRO, fg=COLOR_ROSA, anchor="w").pack(fill="x", pady=(15, 5))
    frame_fecha = tk.Frame(frame_main, bg=COLOR_GRIS, padx=10, pady=10)
    frame_fecha.pack(fill="x")

    meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    cur_m = datetime.now().month
    cur_y = datetime.now().year

    cb_mes = ttk.Combobox(frame_fecha, values=meses, state="readonly", width=15, font=("Helvetica", 11))
    cb_mes.current(cur_m - 1)
    cb_mes.pack(side="left", padx=(0, 10))

    cb_year = ttk.Combobox(frame_fecha, values=[str(y) for y in range(2024, 2031)], state="readonly", width=10, font=("Helvetica", 11))
    cb_year.set(str(cur_y))
    cb_year.pack(side="left")

    # 3. Log de Auditoría
    tk.Label(frame_main, text="3. LOG DE AUDITORÍA", font=("Century Gothic", 10, "bold"), bg=COLOR_NEGRO, fg=COLOR_ROSA, anchor="w").pack(fill="x", pady=(20, 5))
    log_var = tk.BooleanVar(value=True)
    tk.Checkbutton(frame_main, text="Generar y guardar Log de Errores", variable=log_var, bg=COLOR_NEGRO, fg=COLOR_BLANCO, selectcolor=COLOR_GRIS, activebackground=COLOR_NEGRO, activeforeground=COLOR_BLANCO, font=("Helvetica", 10)).pack(anchor="w")

    def on_iniciar():
        if not resultado['file_path']:
            messagebox.showwarning("Atención", "Selecciona un archivo Excel primero.")
            return
        r = messagebox.askyesno("Confirmación LFPDPPP", "Sus datos se procesarán localmente sin uso de internet.\n\n¿Desea iniciar la validación fiscal del documento?")
        if r:
            resultado['mes'] = f"{cb_mes.get()} {cb_year.get()}"
            resultado['log'] = log_var.get()
            root.destroy()

    tk.Button(root, text="INICIAR PROCESO", font=("Century Gothic", 12, "bold"), bg=COLOR_AMARILLO, fg=COLOR_NEGRO, activebackground="#E6C200", activeforeground=COLOR_NEGRO, relief="flat", cursor="hand2", padx=24, pady=12, command=on_iniciar).pack(pady=15)
    
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
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    if tipo == "error":
        messagebox.showerror(titulo, mensaje, parent=root)
    else:
        messagebox.showinfo(titulo, mensaje, parent=root)
    root.destroy()

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
    print("\n⚠️  No se encontró generador_reporte.py en la misma carpeta.")
    print("    Asegúrate de que generador_reporte.py esté junto a main.py")
    print(f"    El archivo _validado.xlsx sí fue generado: {output_path}")

except Exception as e:
    print(f"\n⚠️  Error generando reportes: {e}")
    print(f"    El archivo _validado.xlsx sí fue generado: {output_path}")
    import traceback
    traceback.print_exc()

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
