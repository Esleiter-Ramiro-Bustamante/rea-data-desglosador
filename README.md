# 🧾 ReaDesF - Validador Fiscal Automatizado para México

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/tuusuario/readesf/graphs/commit-activity)

> Sistema profesional de validación y desglose de facturas CFDI para cumplimiento fiscal en México.

---

## 📋 Descripción

**ReaDesF** (Reconocedor y Desglosador Fiscal) es una herramienta automatizada que valida facturas electrónicas (CFDI) según las disposiciones fiscales mexicanas, específicamente:

- ✅ Artículo 27 de la Ley del ISR
- ✅ Validación de IEPS (8% y gasolina)
- ✅ Deducibilidad fiscal automática
- ✅ Detección de pagos en efectivo >$2,000
- ✅ Análisis de formas y métodos de pago
- ✅ Cumplimiento con LFPDPPP (Protección de Datos)

---

## ✨ Características Principales

### 🔍 Validación Fiscal
- Deducibilidad automática según Art. 27 LISR
- Validación de Uso CFDI (G01, G02, G03)
- Detección de métodos de pago válidos (PUE, PPD)
- Alertas de efectivo >$2,000
- Validación especial para gasolina

### 💰 Manejo de IEPS
- **IEPS 8%**: Productos dulces, botanas, pan
- **IEPS Gasolina**: Combustibles (va a IVA 0%)
- Detección automática por concepto
- Cálculos ajustados según tipo de IEPS

### 🔐 Seguridad y Privacidad
- Procesamiento 100% local
- Modo de anonimización para pruebas
- Log de auditoría con SHA256
- Cumplimiento LFPDPPP
- Sin conexión a internet

### ⚡ Optimización
- Pre-indexación de columnas IEPS
- Caché de referencias de columnas
- Velocidad: ~200-500 facturas/segundo
- Progreso en tiempo real

---

## 🚀 Instalación

### Requisitos
```bash
Python 3.8 o superior
```

### Instalar dependencias
```bash
pip install openpyxl
```

### Descargar
```bash
git clone https://github.com/tuusuario/readesf.git
cd readesf
```

---

## 📖 Uso Básico

### 1. Preparar archivo Excel
Coloca tu archivo Excel en:
```
Desktop/GASTOS RESICO/2026/ENERO26/
```

### 2. Ejecutar
```bash
python ReaDesF1.1.py
```

### 3. Ingresar nombre del archivo
```
Ingrese el nombre del archivo Excel (sin extensión .xlsx): facturas_enero
```

### 4. Resultado
```
✅ PROCESO COMPLETADO EXITOSAMENTE
📂 Archivo guardado como: facturas_enero_validado.xlsx
📊 Total de filas procesadas: 350
⏱️  Tiempo total: 12.45 segundos
⚡ Velocidad: 281 facturas/segundo
```

---

## ⚙️ Configuración

### Modos de Operación

#### 🏢 Modo Producción (Predeterminado)
```python
MODO_ANONIMIZAR = False
CREAR_LOG_AUDITORIA = True
```
Para trabajo diario con datos reales.

#### 🧪 Modo Desarrollo
```python
MODO_ANONIMIZAR = True
CREAR_LOG_AUDITORIA = False
```
Para pruebas y compartir archivos de forma segura.

#### 📋 Modo Auditoría
```python
MODO_ANONIMIZAR = False
CREAR_LOG_AUDITORIA = True
LOG_DIRECTORY = "auditoria_2026"
```
Para cumplimiento normativo y trazabilidad.

Ver [CONFIGURACION_RAPIDA_ReaDesF1.1.md](docs/CONFIGURACION_RAPIDA_ReaDesF1.1.md) para más detalles.

---

## 📊 Características Técnicas

### Columnas que Crea
- `SUB1-16%`: Subtotal con/sin IEPS
- `SUB0%`: Base exenta (IVA 0%)
- `SUB2-16%`: Base gravada al 16%
- `IVA ACREDITABLE 16%`: IVA calculado
- `C IVA`: Diferencia de IVA
- `T2`: Total recalculado
- `Comprobación T2`: Validación
- `Deducible`: SI/NO
- `Razón No Deducible`: Motivos de rechazo

### Código de Colores
- 🟦 **Azul**: Gasolina con IEPS / Régimen 626
- 🟩 **Verde**: Deducible / IVA correcto
- 🟥 **Rojo**: NO deducible
- 🟪 **Morado**: Régimen 612
- 🟧 **Naranja**: Gasolina sin IEPS
- 🟨 **Amarillo**: Advertencia / Anonimizado
- 🌸 **Rosa**: Dulces con IEPS 8%

---

## 📚 Documentación

- [Guía Rápida de Configuración](docs/CONFIGURACION_RAPIDA_ReaDesF1.1.md)
- [Guía Completa de Seguridad](docs/GUIA_SEGURIDAD_ReaDesF1.1.md)
- [Análisis de Rendimiento](docs/ANALISIS_RENDIMIENTO.md)

---

## 🔐 Seguridad y Privacidad

### Características
- ✅ Procesamiento 100% local
- ✅ Sin envío de datos a internet
- ✅ Modo de anonimización
- ✅ Log de auditoría con SHA256
- ✅ Cumplimiento LFPDPPP

### Advertencia
⚠️ Este software procesa información fiscal **CONFIDENCIAL**. 
- NO compartir archivos sin anonimizar
- Cifrar archivos de salida si es necesario
- Eliminar archivos temporales después de usar

---

## 🎯 Casos de Uso

### ✅ Contadores
Validación rápida de deducibilidad de facturas de clientes.

### ✅ Empresas RESICO
Cumplimiento del régimen 626 con facilidades especiales.

### ✅ Distribuidores
Validación masiva de facturas con IEPS (dulces, botanas).

### ✅ Gasolineras
Validación especial de IEPS en combustibles.

---

## 🛠️ Reglas Fiscales Implementadas

1. **Art. 27, Fracc. III LISR**: Efectivo >$2,000 NO deducible
2. **Gasolina**: NUNCA en efectivo (excepto RESICO ≤$2,000)
3. **IEPS 8%**: Se suma a SUB1
4. **IEPS Gasolina**: Va a IVA 0%
5. **Uso CFDI**: Solo G01, G02, G03 deducibles
6. **Método de pago**: Solo PUE, PPD válidos
7. **Forma de pago**: Según tipo de gasto

---

## 📈 Roadmap

### Versión 1.2 (Próximo)
- [ ] Validación UUID contra API del SAT
- [ ] Detección EFOS/EDOS
- [ ] Cifrado automático de archivos

### Versión 1.3 (Futuro)
- [ ] Complementos de pago
- [ ] Validación de retenciones
- [ ] Límites por tipo de gasto

### Versión 1.4 (Futuro)
- [ ] Interfaz gráfica (GUI)
- [ ] Integración con contabilidad electrónica
- [ ] API REST

---

## 🤝 Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crea una rama (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -am 'Agrega nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

### Reglas
- ✅ Mantener compatibilidad con Python 3.8+
- ✅ Documentar nuevas funcionalidades
- ✅ Agregar tests cuando sea posible
- ✅ Respetar privacidad de datos

---

## 📝 Licencia

Este proyecto está bajo la Licencia MIT. Ver [LICENSE](LICENSE) para más detalles.

---

## ⚠️ Disclaimer Legal

Este software es una **herramienta de apoyo** y NO sustituye:
- Asesoría fiscal profesional
- Validación oficial del SAT
- Responsabilidad fiscal del contribuyente

El usuario es responsable de:
- Verificar resultados antes de usar
- Cumplir con obligaciones fiscales
- Proteger datos confidenciales

---

## 👤 Autor

Desarrollado para la comunidad fiscal de México.

---

## 📞 Soporte

¿Preguntas? Abre un [Issue](https://github.com/tuusuario/readesf/issues)

---

## 🙏 Agradecimientos

- Esleiter Ramiro Bustamante Ataxca
- CP.Angelica Chagala Sixtega
- Comunidad fiscal mexicana
- Anthropic Claude por asistencia en desarrollo


---

**⭐ Si te fue útil, dale una estrella al repositorio!**
