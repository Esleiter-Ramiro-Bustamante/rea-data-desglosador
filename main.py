"""
main.py — ReaDesF1.8 — Validador Fiscal Adaptativo
===============================================================================

FLUJO COMPLETO v1.8:

  INICIO
    ↓
  Advertencia de privacidad
    ↓
  Solicitar nombre de archivo
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
  Procesar facturas
    ↓
  Si MemoryError → fallback automático a openpyxl
    ↓
  Guardar resultado con fórmulas auditables
    ↓
  Mostrar resumen
    ↓
  FIN

ESTRUCTURA:
  ReaDesF1.8/
  ├── main.py                  ← EJECUTAR ESTE
  ├── analizador_sistema.py    ← RAM + CPU + archivo → motor
  ├── motor_openpyxl.py        ← Computadoras básicas
  ├── motor_pandas.py          ← Computadoras potentes (vectorizado)
  ├── motor_chunks.py          ← Archivos muy grandes (>30k filas)
  ├── validaciones_fiscales.py ← Reglas fiscales + fórmulas auditables
  └── seguridad.py             ← Privacidad y auditoría

===============================================================================
Versión: 1.8 | México | Marzo 2026
===============================================================================
"""

import os
import re
import time

from seguridad          import ConfiguracionSeguridad, LogAuditoria
from analizador_sistema import analizar_y_decidir

REGEX_XLSX = re.compile(r'\.xlsx$', re.IGNORECASE)

desktop_path = os.path.join(
    os.path.expanduser('~'),
    'Desktop/GASTOS RESICO/2026/FEBRERO26'
)

# ══════════════════════════════════════════════════════════════════

print("\n" + "=" * 70)
print("  ReaDesF1.8 — Validador Fiscal Adaptativo")
print("  Adaptive Processing | 3 Motores | Fórmulas Auditables")
print("=" * 70)

ConfiguracionSeguridad.mostrar_advertencia_inicial()
log = LogAuditoria()

# Solicitar archivo
file_name = input("Nombre del archivo Excel (sin .xlsx): ").strip()
file_name = REGEX_XLSX.sub('', file_name) + '.xlsx'
file_path = os.path.join(desktop_path, file_name)

if not os.path.exists(desktop_path):
    print(f"\n❌ Carpeta no encontrada: {desktop_path}")
    input("ENTER para salir...")
    raise SystemExit(1)

if not os.path.exists(file_path):
    print(f"\n❌ Archivo no encontrado: {file_path}")
    excels = [f for f in os.listdir(desktop_path)
              if f.endswith(('.xlsx','.xls','.xlsm'))]
    if excels:
        print("\n📁 Archivos disponibles:")
        for i, f in enumerate(excels[:8], 1):
            print(f"  {i}. {f}")
    input("\nENTER para salir...")
    raise SystemExit(1)

# ── Análisis de recursos ─────────────────────────────────────────
analisis = analizar_y_decidir(file_path)
log.registrar_inicio(file_path, analisis.motor)

base_name   = REGEX_XLSX.sub('', file_name)
suffix      = "_ANONIMIZADO_validado" if ConfiguracionSeguridad.MODO_ANONIMIZAR else "_validado"
output_path = os.path.join(desktop_path, f"{base_name}{suffix}.xlsx")

t_inicio = time.time()
stats    = {}

iconos = {'TURBO':'🚀','CHUNKS':'📦','SEGURO':'🔧','MÍNIMO':'🐢'}
print(f"\n{iconos.get(analisis.modo,'⚙️')} Iniciando: "
      f"{analisis.motor.upper()} — {analisis.modo}\n")

# ── Ejecutar motor con fallback automático ───────────────────────
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
    print(f"\n❌ Error: {e}")
    log.registrar_error(e)
    log.guardar_log()
    input("ENTER para salir...")
    raise SystemExit(1)

t_total   = time.time() - t_inicio
velocidad = analisis.filas_reales / t_total if t_total > 0 else 0
log.registrar_fin(output_path, analisis.filas_reales, t_total, analisis.motor)

# ── Resumen final ────────────────────────────────────────────────
print("\n" + "=" * 70)
print("  ✅ PROCESO COMPLETADO — ReaDesF1.8")
print("=" * 70)
print(f"  📂 Archivo      : {output_path}")
print(f"  📊 Filas        : {analisis.filas_reales:,}")
print(f"  ⏱️  Tiempo        : {t_total:.2f} segundos")
print(f"  ⚡ Velocidad     : {velocidad:,.0f} facturas/segundo")
print(f"  🚗 Motor         : {analisis.motor.upper()} — {analisis.modo}")
print(f"  🖥️  RAM disponible: {analisis.ram_disponible_gb:.1f} GB")

if analisis.motor == 'pandas_chunks':
    bloques = -(-analisis.filas_reales // analisis.chunk_size)
    print(f"  📦 Bloques proc.: {bloques} × {analisis.chunk_size:,} filas")

print(f"\n  📐 FÓRMULAS AUDITABLES (escritas en todos los motores):")
print(f"    sub1      = subtotal - descuento    → Base gravable real")
print(f"    sub0      = iva0 + iva_exento        → Total no gravado")
print(f"    sub2      = sub1 - sub0              → Base para IVA 16%")
print(f"    iva_acred = sub2 × 0.16              → IVA que debería ser")
print(f"    c_iva     = iva_acred - iva16        → Diferencia (0 = correcto)")
print(f"    comprob   = total_cfdi - t2          → Delta total")

print(f"\n  📋 REGÍMENES:")
for reg, cnt in sorted(stats.get('regimenes',{}).items()):
    print(f"    • Régimen {reg}: {cnt:,} facturas")

print(f"\n  🍬 Dulces/Botanas:  Con IEPS: {stats.get('dulces_ieps8',0):,}  |  Sin: {stats.get('dulces_sin',0):,}")

print(f"\n  ⛽ Gasolina:")
print(f"    Con IEPS   : {stats.get('gas_ieps',0):,}   Sin IEPS: {stats.get('gas_sin',0):,}")
print(f"    626 efect. : {stats.get('gas_626',0):,}   (agrup: {stats.get('gas_626_agrup',0):,})")
print(f"    612 efect. : {stats.get('gas_612',0):,}   (rechazadas)")
print(f"    Electrónica: {stats.get('gas_elec',0):,}")

print(f"\n  🌱 Insumos Agrícolas:")
print(f"    ❌ Efectivo >$2,000  : {stats.get('ins_nd',0):,}")
print(f"    ✅ Efectivo ≤$2,000  : {stats.get('ins_menor',0):,}")
print(f"    ✅ Pago electrónico  : {stats.get('ins_elec',0):,}")

print(f"\n  📋 General:")
print(f"    S01 (sin efectos)  : {stats.get('s01',0):,}")
print(f"    Efectivo >$2,000   : {stats.get('ef_mayor',0):,}")

print("\n  ⚖️  REGLAS APLICADAS:")
print("    Todos  : Art. 27 Fracc. III — efectivo >$2,000 NO deducible")
print("    626    : Gasolina ≤$2,000 o agrupada = deducible (facilidad)")
print("    612    : Art. 103 + Art. 27 Fracc. III LISR")
print("             Art. 74 AGAPES / Art. 147 = NO aplican a Régimen 612")

print("\n  🎨 COLORES: 🟦 626  🟪 612  🟩 Deducible  🟥 NO ded.")
print("              🟨 Advertencia  🟢 Insumo elec.  🌸 Dulces IEPS")

print("\n" + "=" * 70)
print("  🔐 Información fiscal CONFIDENCIAL — Guardar de forma segura")
print("=" * 70)

log.guardar_log()
input("\n👉 ENTER para salir...")
