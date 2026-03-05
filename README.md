# ReaDesF1.8 — Sinergia REA

> **Validador Fiscal Adaptativo** · Sistema inteligente de clasificación de gastos deducibles para contribuyentes mexicanos

![Version](https://img.shields.io/badge/versión-1.8-FF2D78?style=flat-square)
![Python](https://img.shields.io/badge/Python-3.8+-FFD600?style=flat-square&logo=python&logoColor=black)
![Status](https://img.shields.io/badge/estado-activo-00C851?style=flat-square)
![License](https://img.shields.io/badge/licencia-privada-0057FF?style=flat-square)

`RÉGIMEN 612` · `RÉGIMEN 626 RESICO` · `ART. 103 LISR` · `ART. 27 FRACC. III` · `IEPS 8%` · `INSUMOS AGRÍCOLAS` · `ADAPTIVE PROCESSING` · `FÓRMULAS AUDITABLES`

---

## ¿Qué es ReaDesF?

ReaDesF es un validador fiscal de gastos deducibles para **Régimen 612** (actividad empresarial) y **Régimen 626 RESICO**. Analiza archivos Excel con facturas CFDI, clasifica cada gasto según la ley, calcula fórmulas auditables y genera un archivo de salida con colores, razones y columnas de validación contable.

Solo ejecutas `main.py`. El programa mide tu RAM y el tamaño del archivo, elige el motor óptimo automáticamente y procesa todas las facturas. **No necesitas configurar nada manualmente.**

---

## ¿Qué hay de nuevo en v1.8?

| Mejora | Detalle |
|---|---|
| 3 motores | `openpyxl` / `pandas` / `pandas_chunks` |
| Masks vectorizadas | Reemplaza `iterrows()` — 10-50x más rápido |
| Columnas `category` | Hasta 70% menos RAM en archivos grandes |
| Fórmulas auditables | sub1 / sub0 / sub2 / iva_acred en los 3 motores |
| Fallback automático | Si pandas falla por RAM → openpyxl automático |

---

## 01 · Inicio Rápido

```bash
# 1. Instalar dependencias
pip install openpyxl pandas psutil

# 2. Colocar tu archivo Excel en:
#    Desktop / GASTOS RESICO / 2026 / FEBRERO26

# 3. Ejecutar
python main.py

# 4. Escribir el nombre del archivo (sin .xlsx)
#    → GASTOS FEBRERO

# 5. El resultado se guarda como NOMBRE_validado.xlsx
```

| Paso | Acción |
|------|--------|
| 1 | Instalar dependencias en la carpeta `ReaDesF1.8/` |
| 2 | Colocar el archivo Excel en `Desktop/GASTOS RESICO/2026/FEBRERO26` |
| 3 | Ejecutar `python main.py` desde la terminal |
| 4 | El programa analiza tu **RAM** y elige el mejor motor automáticamente |
| 5 | Escribe el nombre del archivo **sin `.xlsx`** cuando se solicite |
| 6 | Resultado listo como `NOMBRE_validado.xlsx` con colores, fórmulas y razones |

---

## 02 · Motores Adaptativos

ReaDesF1.8 incluye **3 motores** que se activan según los recursos de tu computadora y el tamaño del archivo.

### ⚙️ openpyxl — SEGURO / MÍNIMO
- RAM < 4 GB **o** archivo < 5,000 filas
- Estable en cualquier computadora
- Fórmulas auditables garantizadas
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

### Capacidad máxima estimada

| Filas | Motor | Tiempo estimado |
|-------|-------|-----------------|
| 5,000 | openpyxl | ~1 min |
| 30,000 | pandas | ~2 min |
| 100,000 | chunks | ~7 min |
| 500,000 | chunks | ~35 min |

---

## 03 · Fórmulas Auditables

Las siguientes fórmulas se escriben como **fórmulas Excel vivas** en los 3 motores. Al pararte en cualquier celda se ve la operación completa.

```
sub1      = subtotal - descuento    → Base gravable real
sub0      = iva0 + iva_exento       → Total no gravado
sub2      = sub1 - sub0             → Base para IVA 16%
iva_acred = sub2 × 0.16             → IVA que debería ser
c_iva     = iva_acred - iva16       → Diferencia (0 = correcto)
comprob   = total_cfdi - t2         → Delta total
```

| Columna | Fórmula | Descripción |
|---------|---------|-------------|
| `sub1` | `subtotal − descuento` | Base gravable real del gasto. Punto de partida de toda la validación. |
| `sub0` | `IVA 0% + IVA Exento` | Total no gravado. Monto que NO genera IVA acreditable. |
| `sub2` | `sub1 − sub0` | Base real para IVA 16%. La porción que SÍ genera IVA acreditable. |
| `iva_acred` | `sub2 × 0.16` | IVA que *debería ser* según la base. Columna de validación clave. |
| `c_iva` | `iva_acred − IVA declarado` | Diferencia. Si ≠ 0 hay discrepancia en el CFDI. |
| `comprob` | `Total CFDI − T2` | Delta total. Verifica que todos los componentes sumen correctamente. |

> ⚠️ Si `c_iva ≠ 0` hay discrepancia en el IVA del CFDI. Estas fórmulas permiten detectarla sin necesidad de recalcular manualmente.

---

## 04 · Instalación

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

## 05 · Estructura del Proyecto

```
ReaDesF1.8/
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
  ├── seguridad.py              ← Privacidad y auditoría
  │
  ├── requirements.txt          ← pip install -r requirements.txt
  └── README.md                 ← Este archivo
```

---

## 06 · Historial de Versiones

| Versión | Nombre | Mejora principal |
|---------|--------|-----------------|
| v1.2 | Reglas diferenciadas | Validación separada para Régimen 626 RESICO y Régimen 612 actividad empresarial |
| v1.3 | Gasolina agrupada RESICO | Detección de despachos agrupados separados por `\|`. Facilidad RESICO aplicada correctamente |
| v1.4 | Insumos agrícolas | 100+ palabras clave: fertilizantes, semillas, herbicidas, fungicidas, enmiendas de suelo |
| v1.4.1 | Corrección legal crítica | Art. 147 LISR eliminado. Fundamento correcto: Art. 103 LISR |
| v1.5 | headers_map O(1) | Índice de columnas como dict comprehension. Búsqueda instantánea |
| v1.6 | Optimizaciones Nivel 1-3 | `iter_rows`, sets de palabras clave, cache, regex precompilado |
| v1.7 | Adaptive Processing — 2 motores | Selección automática entre openpyxl y pandas según RAM y archivo |
| **v1.8** ⭐ | **3 Motores + Vectorización + Chunks** | Motor chunks para +30k filas, masks vectorizadas, columnas `category`, fórmulas auditables en los 3 motores |

---

## 07 · Hoja de Ruta

### 🟡 Ahora — v1.8
- [x] Probar con datos reales
- [x] Validar 3 motores
- [x] Verificar fórmulas auditables

### 🔴 Corto plazo — v1.9
- [ ] Interfaz gráfica tkinter
- [ ] `config.json` externo
- [ ] Pruebas automatizadas
- [ ] Notificación al terminar

### 🔵 Mediano plazo — v2.0
- [ ] Reporte PDF automático
- [ ] Dashboard de resultados
- [ ] Detección de patrones
- [ ] Comparativo histórico

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
