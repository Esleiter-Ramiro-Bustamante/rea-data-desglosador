# ReaDesF1.9 — Sinergia REA

> **Validador Fiscal Adaptativo** · Sistema inteligente de clasificación de gastos deducibles para contribuyentes mexicanos

![Version](https://img.shields.io/badge/versión-1.9-FF2D78?style=flat-square)
![Python](https://img.shields.io/badge/Python-3.8+-FFD600?style=flat-square&logo=python&logoColor=black)
![Status](https://img.shields.io/badge/estado-activo-00C851?style=flat-square)
![License](https://img.shields.io/badge/licencia-privada-0057FF?style=flat-square)

`RÉGIMEN 612` · `RÉGIMEN 626 RESICO` · `ART. 103 LISR` · `ART. 27 FRACC. III` · `IEPS 8%` · `INSUMOS AGRÍCOLAS` · `ADAPTIVE PROCESSING` · `FÓRMULAS AUDITABLES` · `DIOT` · `DASHBOARD HTML`

---

## ¿Qué es ReaDesF?

ReaDesF es un validador fiscal de gastos deducibles para **Régimen 612** (actividad empresarial) y **Régimen 626 RESICO**. Analiza archivos Excel con facturas CFDI, clasifica cada gasto según la ley, calcula fórmulas auditables y genera tres archivos de salida: el `_validado.xlsx` con colores y fórmulas, el `_reporte.xlsx` con tabla auditable resumida, y el `_reporte.html` con un dashboard interactivo completo.

Solo ejecutas `main.py`. El programa mide tu RAM y el tamaño del archivo, elige el motor óptimo automáticamente y procesa todas las facturas. **No necesitas configurar nada manualmente.**

---

## ¿Qué hay de nuevo en v1.9?

| Mejora | Detalle |
|---|---|
| Flujo de 3 pasos | PASO 1 selección → PASO 2 `_validado.xlsx` → PASO 3 `_reporte.xlsx` + `_reporte.html` |
| `generador_reporte.py` | Nuevo módulo separado — genera reporte Excel y dashboard HTML desde el `_validado` |
| Fórmulas cacheadas | Los 3 motores escriben fórmula + valor numérico en el XML del `.xlsx` — Python las lee sin depender de Excel |
| Fix SUB0% gasolina | Cuando `IVA_0% == IVA_Exento == IEPS` ya no se duplica el monto |
| Dashboard HTML | Filtros, buscador, estatus editable, contadores dinámicos, persistencia con `localStorage` |
| DIOT integrada | Informativa de Operaciones con Terceros — agrupada por RFC, actualización en tiempo real |
| Complementos CP01 | Cruce por UUID directo o por RFC+total cuando el UUID relacionado viene vacío |
| Forma 99 PPD | Se clasifica como PENDIENTE (no ERROR) — flujo correcto del SAT |

---

## 01 · Inicio Rápido

```bash
# 1. Instalar dependencias
pip install openpyxl pandas psutil

# 2. Ejecutar
python main.py

# 3. Seleccionar carpeta y archivo cuando se solicite
#    → El sistema elige el motor automáticamente

# 4. Se generan 3 archivos:
#    NOMBRE_validado.xlsx  → Facturas validadas con colores y fórmulas
#    NOMBRE_reporte.xlsx   → Tabla resumen auditable
#    NOMBRE_reporte.html   → Dashboard interactivo
```

| Paso | Acción |
|------|--------|
| 1 | Instalar dependencias |
| 2 | Ejecutar `python main.py` desde la terminal |
| 3 | El programa analiza tu **RAM** y elige el mejor motor automáticamente |
| 4 | Escribe el nombre del archivo **sin `.xlsx`** cuando se solicite |
| 5 | Indica el mes del reporte (ej: `FEBRERO 2026`) |
| 6 | Los 3 archivos de salida quedan en la misma carpeta del archivo fuente |

---

## 02 · Motores Adaptativos

ReaDesF1.9 incluye **3 motores** que se activan según los recursos de tu computadora y el tamaño del archivo.

### ⚙️ openpyxl — SEGURO / MÍNIMO
- RAM < 4 GB **o** archivo < 5,000 filas
- Estable en cualquier computadora
- Fórmulas auditables garantizadas + valores cacheados
- ⚡ ~5,000 facturas / min

### 🐼 pandas — TURBO
- RAM ≥ 4 GB + 5,000–30,000 filas
- Detección vectorizada completa
- Columnas a `category` (70% menos RAM)
- ⚡ ~20,000 facturas / min

### 🧩 pandas chunks — CHUNKS
- RAM ≥ 4 GB + > 30,000 filas
- Bloques de 5,000 filas, RAM constante
- Hasta 500,000 facturas posible
- ⚡ ~15,000 facturas / min

### Tabla de decisión automática

| Filas | RAM disponible | Motor | Modo | Velocidad est. |
|-------|---------------|-------|------|----------------|
| < 5,000 | Cualquiera | `openpyxl` | SEGURO 🔧 | ~5,000 / min |
| 5k – 30k | < 4 GB | `openpyxl` | SEGURO 🔧 | ~5,000 / min |
| 5k – 30k | ≥ 4 GB | `pandas` | TURBO 🚀 | ~20,000 / min |
| > 30k | < 4 GB | `openpyxl` | MÍNIMO 🐢 | ~2,000 / min |
| > 30k | ≥ 4 GB | `chunks` | CHUNKS 📦 | ~15,000 / min |
| Cualquiera | < 2 GB | `openpyxl` | MÍNIMO 🐢 | ~2,000 / min |

---

## 03 · Fórmulas Auditables

Las siguientes fórmulas se escriben como **fórmulas Excel vivas** en los 3 motores. Al pararte en cualquier celda se ve la operación completa. En v1.9, cada celda también lleva el **valor numérico cacheado en el XML** del archivo, por lo que `generador_reporte.py` las lee directamente sin necesitar que Excel las evalúe primero.

```
sub1      = subtotal - descuento    → Base gravable real
sub0      = iva0 + iva_exento       → Total no gravado (sin duplicar IEPS gasolina)
sub2      = sub1 - sub0             → Base para IVA 16%
iva_acred = sub2 × 0.16             → IVA que debería ser
c_iva     = iva_acred - iva16       → Diferencia (0 = correcto)
comprob   = total_cfdi - t2         → Delta total
```

| Columna | Fórmula | Descripción |
|---------|---------|-------------|
| `SUB1-16%` | `subtotal − descuento` | Base gravable real del gasto |
| `SUB0%` | `IVA 0% + IVA Exento` | Total no gravado — sin duplicar IEPS de gasolina |
| `SUB2-16%` | `sub1 − sub0` | Base real para IVA 16% |
| `A ACREDITABLE 16%` | `sub2 × 0.16` | IVA que debería ser según la base |
| `C IVA` | `iva_acred − IVA declarado` | Diferencia — si ≠ 0 hay discrepancia en el CFDI |
| `Comprobación T2` | `Total CFDI − T2` | Delta total de todos los componentes |

> ⚠️ Si `C IVA ≠ 0` hay discrepancia en el IVA del CFDI. Estas fórmulas permiten detectarla sin recalcular manualmente.

---

## 04 · Archivos de Salida

### `NOMBRE_validado.xlsx`
El archivo principal de validación. Contiene todas las columnas originales más las columnas calculadas con fórmulas auditables y colores por régimen y deducibilidad.

| Color | Significado |
|-------|-------------|
| 🟦 Azul | Régimen 626 RESICO |
| 🟪 Morado | Régimen 612 |
| 🟩 Verde | Deducible |
| 🟥 Rojo | No deducible |
| 🟧 Naranja | Gasolina con IEPS / régimen desconocido |
| 🟨 Amarillo | CP01 / Complemento de pago |

### `NOMBRE_reporte.xlsx`
Tabla resumen auditable con una fila por factura, columnas de estatus, montos desglosados y observaciones sobre complementos CP01.

### `NOMBRE_reporte.html`
Dashboard interactivo con:
- **Tarjetas de resumen** — DEDUCIBLES, EFECTIVO, NO DEDUCIBLES, PENDIENTES, TOTAL
- **Filtros por estatus** — DED · EFE · NO DED · PEND · EGRESO · CP01
- **Buscador en tiempo real** — por UUID, razón social o monto
- **Estatus editable** — cambia cualquier clasificación directamente en el reporte
- **Persistencia** — los cambios manuales se guardan con `localStorage` y sobreviven al recargar
- **DIOT integrada** — se despliega al hacer clic en la tarjeta DEDUCIBLES

---

## 05 · DIOT — Informativa de Operaciones con Terceros

En v1.9 el reporte HTML incluye la **DIOT mensual** generada automáticamente desde las facturas deducibles.

```
— Informativa de operaciones con terceros (DIOT)
```

| Columna | Fuente |
|---------|--------|
| RAZÓN SOCIAL | RFC emisor del `_validado` |
| RFC | RFC emisor del `_validado` |
| ACT. PAGADAS 16% | Suma de `SUB2-16%` por proveedor |
| IVA PAGADO | Suma de `IVA 16%` por proveedor |
| ACT. PAGADO TASA 0% | Suma de `SUB0%` por proveedor |
| TOTAL | Suma de `Total` por proveedor |

**Reglas:**
- ✅ Incluye facturas con estatus `DED` y `EFE`
- ❌ Excluye `COMPLEMENTO CP01` y `PENDIENTE PPD`
- Se actualiza en tiempo real cuando cambias el estatus de una factura

---

## 06 · Complementos CP01

El sistema detecta y cruza automáticamente complementos de pago con sus facturas PPD originales.

### Cruce Nivel 1 — UUID explícito
Cuando el `_validado` incluye el UUID relacionado en la columna `UUIDs relacionados`.

### Cruce Nivel 2 — RFC + Total
Cuando la columna `UUIDs relacionados` viene vacía, el sistema cruza por:
- Mismo RFC emisor
- Total idéntico entre el PPD y el CP01

```
OBS: "Complemento parcialidad 1 - factura 5D1EDA96AB6E saldo insoluto $0.00"
                                           ↑
                              Últimos 12 chars del UUID del PPD pagado
```

---

## 07 · Instalación

### ⚡ Mínimo requerido
```bash
pip install openpyxl
```

### 🚀 Recomendado (activa los 3 motores)
```bash
pip install openpyxl pandas psutil
```

### 📦 Un solo comando
```bash
pip install -r requirements.txt
```

---

## 08 · Estructura del Proyecto

```
ReaDesF1.9/
  │
  ├── main.py                   ← EJECUTAR ESTE
  │
  ├── analizador_sistema.py     ← RAM + CPU + archivo → motor
  │
  ├── motor_openpyxl.py         ← Computadoras básicas
  ├── motor_pandas.py           ← Computadoras potentes
  ├── motor_chunks.py           ← Archivos muy grandes +30k filas
  │
  ├── validaciones_fiscales.py  ← Reglas + fórmulas auditables
  ├── generador_reporte.py      ← Reporte Excel + Dashboard HTML + DIOT
  ├── seguridad.py              ← Privacidad y auditoría
  │
  ├── requirements.txt          ← pip install -r requirements.txt
  └── README.md                 ← Este archivo
```

---

## 09 · Historial de Versiones

| Versión | Nombre | Mejora principal |
|---------|--------|-----------------|
| v1.2 | Reglas diferenciadas | Validación separada para Régimen 626 RESICO y Régimen 612 |
| v1.3 | Gasolina agrupada RESICO | Detección de despachos agrupados separados por `\|` |
| v1.4 | Insumos agrícolas | 100+ palabras clave: fertilizantes, semillas, herbicidas |
| v1.4.1 | Corrección legal crítica | Art. 147 LISR eliminado — fundamento correcto: Art. 103 LISR |
| v1.5 | headers_map O(1) | Índice de columnas como dict comprehension |
| v1.6 | Optimizaciones Nivel 1-3 | `iter_rows`, sets de palabras clave, cache, regex precompilado |
| v1.7 | Adaptive Processing — 2 motores | Selección automática entre openpyxl y pandas |
| v1.8 | 3 Motores + Vectorización + Chunks | Motor chunks para +30k filas, masks vectorizadas, fórmulas auditables |
| **v1.9** ⭐ | **Reporte + Dashboard + DIOT** | `generador_reporte.py`, dashboard HTML interactivo, DIOT automática, fórmulas cacheadas, cruce CP01 por RFC+total |

---

## 10 · Hoja de Ruta

### 🟢 Completado — v1.9
- [x] Generador de reporte Excel + HTML
- [x] Dashboard interactivo con filtros y buscador
- [x] Estatus editable con persistencia `localStorage`
- [x] DIOT automática y actualización en tiempo real
- [x] Fórmulas auditables cacheadas en XML
- [x] Fix SUB0% gasolina con IEPS (no duplicar)
- [x] Cruce CP01 ↔ PPD por RFC + total
- [x] Forma 99 PPD → PENDIENTE (correcto SAT)

### 🟡 Corto plazo — v2.0
- [ ] `configuracion.py` externo — sin hardcodear rutas ni límites
- [ ] `actualizador_ppd.py` — cruce automático entre meses (enero→febrero)
- [ ] Interfaz gráfica tkinter
- [ ] Pruebas automatizadas

### 🔵 Mediano plazo — v2.1
- [ ] `comparador_meses.py` — tendencias y variaciones por proveedor
- [ ] Export PDF del reporte
- [ ] `validador_cfdi.py` — verificación de UUIDs contra SAT

### ⬜ Largo plazo — v3.0
- [ ] Instalador `.exe`
- [ ] Versión despachos
- [ ] Actualizaciones automáticas
- [ ] Multiusuario

---

## Notas legales y de seguridad

- Información fiscal **confidencial**
- Procesamiento **100% local** — ningún dato sale de tu computadora
- Cumple **LFPDPPP** (Ley Federal de Protección de Datos Personales en Posesión de Particulares)
- Fundamentos: **Art. 103 LISR** · **Art. 27 Fracc. III CFF** · **Régimen 612** · **Régimen 626 RESICO** · **IEPS 8%**

---

<div align="center">

**ReaDesF** · Sinergia REA · México 2026

*Validador Fiscal Adaptativo — procesamiento local, fórmulas auditables, cumplimiento fiscal*

</div>
