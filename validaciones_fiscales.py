"""
validaciones_fiscales.py — Reglas fiscales + optimizaciones v1.8
ReaDesF1.8

NUEVO EN v1.8:
  ✓ Masks vectorizadas para pandas (reemplaza iterrows)
  ✓ Columnas a tipo 'category' para reducir RAM hasta 70%
  ✓ Fórmulas auditables Excel preservadas en los 3 motores
  ✓ evaluar_deducibilidad_vectorizado() — procesa DataFrame completo

FÓRMULAS AUDITABLES — NUNCA SE PIERDEN:
  ┌─────────────────────────────────────────────────────────────┐
  │ sub1     = subtotal - descuento   → Base gravable real      │
  │ sub0     = iva0 + iva_exento      → Total no gravado        │
  │ sub2     = sub1 - sub0            → Base para IVA 16%       │
  │ iva_acred= sub2 * 0.16            → IVA que debería ser     │
  │                                                             │
  │ Si IVA declarado ≠ iva_acred → DISCREPANCIA en el CFDI     │
  │ Al pararse en la celda en Excel se ve la fórmula VIVA       │
  └─────────────────────────────────────────────────────────────┘

FUNDAMENTO LEGAL — INSUMOS AGRÍCOLAS RÉGIMEN 612:
  ✅ Art. 27, Fracc. III LISR → Pagos >$2,000 deben ser electrónicos
  ✅ Art. 103 LISR            → Deducciones autorizadas Régimen 612
  ❌ Art. 147 LISR            → Deducciones personales. NO aplica aquí
  ❌ Art. 74 LISR (AGAPES)   → Exclusivo AGAPES. NO aplica a 612

  REGLA:
  • Insumo $1,500 efectivo      → ✅ Deducible
  • Insumo $3,000 efectivo      → ❌ NO Deducible
  • Insumo $3,000 transferencia → ✅ Deducible
"""

import re
import pandas as pd

# ══════════════════════════════════════════════════════════════════
# CONSTANTES FISCALES
# ══════════════════════════════════════════════════════════════════

USOS_DEDUCIBLES      = {'G01', 'G02', 'G03'}
METODOS_VALIDOS      = {'PUE', 'PPD'}
FORMAS_VALIDAS       = {'01', '02', '03', '04', '28'}
FORMAS_ELECTRONICAS  = {'02', '03', '04', '28'}
REGIMENES_TRABAJADOS = {'626', '612'}
USO_CFDI_VERDE       = "G02 - Devoluciones, descuentos o bonificaciones"
LIMITE_EFECTIVO      = 2000.0      # Art. 27 Fracc. III LISR

# ══════════════════════════════════════════════════════════════════
# SETS DE PALABRAS CLAVE — búsqueda O(1)
# ══════════════════════════════════════════════════════════════════

PALABRAS_GASOLINA: set = {
    'gasolina', 'combustible', 'magna', 'premium',
    'diesel', 'diésel', 'gasohol', 'gasoil',
    'nafta', 'petrol', 'gas', 'energético',
    'turbosina', 'jet fuel', 'bunker'
}

PALABRAS_DULCE: set = {
    'pan', 'roles', 'conchas', 'mantecadas', 'donas', 'panque',
    'gansito', 'pinguinos', 'submarinos', 'chocorol', 'principe',
    'pastisetas', 'canelitas', 'polvorones', 'triki', 'duo',
    'rebanada', 'colchones', 'cuernitos', 'medias noches',
    'pan blanco', 'pan integral', 'pan molido', 'empanizador',
    'tostada', 'tostaditas', 'tostado', 'tortillinas', 'salmas',
    'chocolate', 'bon o bon', 'hershey', 'reese', 'kit kat',
    'kremas', 'trident',
    'papas', 'chips', 'sabritas', 'doritos', 'cheetos', 'ruffles',
    'barcel', 'takis', 'hot nuts', 'cacahuates', 'kiyakis',
    'runners', 'churrumais', 'tostitos', 'fritos', 'big mix',
    'galletas', 'oreo', 'emperador', 'marías', 'animalitos',
    'chokis', 'sponch', 'barrita',
    'pepsi', 'coca', 'sprite', 'fanta', 'monster', 'gatorade',
    'ades', 'delsey'
}

PALABRAS_INSUMO: set = {
    # Fertilizantes
    'fertilizante', 'fertilizacion', 'fertilización',
    'abono', 'abonos', 'abonado',
    'urea', 'nitrato', 'nitrogeno', 'nitrógeno',
    'sulfato de amonio', 'amoniaco', 'sulfammo',
    'nitrofoska', 'nitrabor', 'kan-l', 'uan', 'nitramon', 'nitromax',
    'fosfato', 'fosforo', 'fósforo',
    'map', 'dap', 'superfosfato', 'triple super',
    'potasio', 'potásico', 'cloruro de potasio',
    'sulfato de potasio', 'sulfato potásico',
    'npk', 'haifa', 'basacote', 'nutri',
    'microelementos', 'humatos', 'calcio foliar',
    'quelatos', 'bioestimulante', 'aminoacidos',
    'acido humico', 'ácido húmico', 'fulvico',
    'yara', 'fertimex', 'compo', 'timac', 'mosaic', 'nutrimos', 'agrium',
    # Semillas
    'semilla', 'semillas', 'semillero',
    'esqueje', 'esquejes', 'plantula', 'plántula', 'plantulas',
    'variedad', 'hibrido', 'híbrido',
    'semilla de caña', 'punta de caña', 'trocito de caña',
    'cana semilla', 'caña semilla', 'propagacion', 'propagación',
    'maiz', 'maíz', 'sorgo', 'frijol',
    'garbanzo', 'ajonjoli', 'ajonjolí', 'girasol', 'cártamo', 'cartamo',
    # Herbicidas
    'herbicida', 'herbicidas', 'maleza', 'deshierbe', 'desherbe',
    'glifosato', 'roundup', 'atrazina', 'atrazine',
    '2-4d', '2,4-d', 'amine', 'pendimetalina', 'metribuzin',
    'diuron', 'diurón', 'hexazinona', 'ametrina', 'ametrine',
    'faena', 'gesapax', 'velpar', 'karmex', 'harness', 'dual',
    # Insecticidas
    'insecticida', 'insecticidas', 'plaguicida', 'plaguicidas',
    'pesticida', 'pesticidas',
    'clorpirifos', 'imidacloprid', 'lambda',
    'cipermetrina', 'deltametrina', 'malatión',
    'abamectina', 'spinosad', 'bifentrina', 'thiamethoxam', 'acetamiprid',
    'bacillus', 'beauveria', 'trichoderma', 'metarhizium',
    'control biologico', 'biol',
    'lorsban', 'confidor', 'regent', 'karate', 'decis', 'engeo',
    # Fungicidas
    'fungicida', 'fungicidas',
    'mancozeb', 'metalaxil', 'propiconazol', 'tebuconazol',
    'azoxistrobina', 'iprodiona', 'clorotalonil', 'captan', 'tiofanato',
    'dithane', 'ridomil', 'tilt', 'folicur', 'amistar', 'rovral',
    # Agronutrientes
    'agronutriente', 'agronutrientes', 'estimulante', 'estimulantes',
    'regulador de crecimiento', 'regulador',
    'citoquinina', 'auxina', 'giberelina',
    'ethephon', 'ethrel', 'madurante', 'maduracion', 'maduración',
    # Insumos de aplicación
    'coadyuvante', 'adherente', 'surfactante',
    'aceite agricola', 'aceite agrícola', 'dispersante', 'emulsificante',
    # Enmiendas
    'cal agricola', 'cal agrícola', 'calcita', 'dolomita',
    'yeso agricola', 'yeso agrícola',
    'azufre agricola', 'azufre agrícola', 'encalado',
    'enmienda', 'corrector de suelo', 'acondicionador de suelo',
}

# ══════════════════════════════════════════════════════════════════
# PATRONES REGEX COMPILADOS — para pandas .str.contains()
# Se compilan UNA SOLA VEZ al importar el módulo
# ══════════════════════════════════════════════════════════════════

PATRON_GASOLINA: re.Pattern = re.compile(
    '|'.join(re.escape(p) for p in PALABRAS_GASOLINA), re.IGNORECASE)

PATRON_DULCE: re.Pattern = re.compile(
    '|'.join(re.escape(p) for p in PALABRAS_DULCE), re.IGNORECASE)

PATRON_INSUMO: re.Pattern = re.compile(
    '|'.join(re.escape(p) for p in PALABRAS_INSUMO), re.IGNORECASE)

PALABRAS_TELECOM: set = {
    # Operadoras
    'telmex', 'telcel', 'att', 'izzi', 'axtel', 'megacable',
    'totalplay', 'telecomunicaciones', 'telecomunicacion',
    'telecomunicación', 'telecomunicaciones de mexico',
    'telefonos de mexico', 'teléfonos de méxico',
    # Servicios
    'telefonia', 'telefonía', 'servicio telefonico',
    'servicio telefónico', 'servicio de telecomunicaciones',
    'internet y telefonia', 'internet y telefonía',
    'plan de datos', 'renta mensual telefono',
    'servicio de internet', 'fibra optica', 'fibra óptica',
}

PATRON_TELECOM: re.Pattern = re.compile(
    '|'.join(re.escape(p) for p in PALABRAS_TELECOM), re.IGNORECASE)

# ══════════════════════════════════════════════════════════════════
# CACHE — evita re-evaluar conceptos repetidos (openpyxl motor)
# ══════════════════════════════════════════════════════════════════

_cache_conceptos: dict = {}

def detectar_tipo(concepto_lower: str) -> tuple:
    """
    Detecta tipo con cache. O(1) si ya fue evaluado.
    Retorna: (es_gasolina, es_dulce, es_insumo, es_telecom)
    """
    if concepto_lower in _cache_conceptos:
        return _cache_conceptos[concepto_lower]

    gas     = any(p in concepto_lower for p in PALABRAS_GASOLINA)
    dulce   = any(p in concepto_lower for p in PALABRAS_DULCE)
    insumo  = any(p in concepto_lower for p in PALABRAS_INSUMO)
    telecom = any(p in concepto_lower for p in PALABRAS_TELECOM)

    resultado = (gas, dulce, insumo, telecom)
    _cache_conceptos[concepto_lower] = resultado
    return resultado

def es_gasolina_agrupada(concepto_lower: str) -> bool:
    return '|' in concepto_lower

def extraer_codigo(value) -> str:
    if value and '-' in str(value):
        return str(value).split('-')[0].strip().upper()
    return str(value).strip().upper() if value else ''

# ══════════════════════════════════════════════════════════════════
# OPTIMIZACIÓN v1.8 — Columnas a tipo 'category'
# Reduce uso de RAM hasta 70% en columnas con valores repetidos
# (régimen, uso CFDI, forma pago, método pago)
# ══════════════════════════════════════════════════════════════════

COLUMNAS_CATEGORICAS = [
    'Regimen receptor',
    'Uso CFDI',
    'Forma pago',
    'Metodo pago',
]

def optimizar_tipos_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte columnas repetitivas a tipo 'category'.
    Reduce RAM hasta 70% en archivos grandes.
    Llamar después de pd.read_excel() y antes de procesar.
    """
    for col in COLUMNAS_CATEGORICAS:
        if col in df.columns:
            antes = df[col].memory_usage(deep=True)
            df[col] = df[col].astype('category')
            despues = df[col].memory_usage(deep=True)
            ahorro  = (1 - despues / antes) * 100 if antes > 0 else 0
    return df

# ══════════════════════════════════════════════════════════════════
# FÓRMULAS AUDITABLES — constantes de texto
# Se usan en los 3 motores para escribir fórmulas Excel vivas.
#
# CRÍTICO: estas fórmulas son la segunda validación contable.
# Al pararse en la celda en Excel se ve la operación completa.
# Permiten detectar discrepancias en IVA desglosado del CFDI.
#
#   sub1      = subtotal - descuento   → Base gravable real
#   sub0      = iva0 + iva_exento      → Total no gravado
#   sub2      = sub1 - sub0            → Base para IVA 16%
#   iva_acred = sub2 * 0.16            → IVA que debería ser
#
# Si IVA declarado ≠ iva_acred → hay discrepancia en el CFDI
# ══════════════════════════════════════════════════════════════════

def formulas_auditables(rn: int, CL: dict) -> dict:
    """
    Genera las fórmulas Excel auditables para una fila.
    rn  = número de fila en Excel
    CL  = dict de letras de columna {'ST': 'A', 'DC': 'B', ...}

    Retorna dict con las 7 fórmulas listas para escribir en Excel.
    """
    return {
        # ── Fórmula 1: SUB1 — Base gravable real ──────────────────
        # sub1 = subtotal - descuento
        # Refleja el monto real del gasto antes de impuestos
        'sub1': f"=({CL['ST']}{rn}-{CL['DC']}{rn})",

        # ── Fórmula 1b: SUB1 con IEPS 8% (dulces/botanas) ─────────
        # sub1_ieps8 = (subtotal - descuento) + IEPS8
        # El IEPS 8% forma parte de la base gravable
        'sub1_ieps8': f"=({CL['ST']}{rn}-{CL['DC']}{rn})+{CL.get('I8','')}{rn}",

        # ── Fórmula 1c: SUB1 con IEPS 3% (telefonía/telecomunicaciones) ──
        # sub1_ieps3 = (subtotal - descuento) + IEPS3
        # El IEPS 3% forma parte de la base gravable igual que IEPS 8%
        'sub1_ieps3': f"=({CL['ST']}{rn}-{CL['DC']}{rn})+{CL.get('I3','')}{rn}",

        # ── Fórmula 2: SUB0 — Total no gravado ────────────────────
        # sub0 = iva0 + iva_exento
        # Monto que NO genera IVA acreditable
        'sub0': f"={CL['I0']}{rn}+{CL['IE']}{rn}",

        # ── Fórmula 2b: SUB0 gasolina con IEPS ────────────────────
        # Cuando es gasolina con IEPS, el motor ya puso el IEPS en
        # IVA_0%. IVA_Exento tiene el mismo valor → no sumar (duplica)
        # sub0_gas = iva0   (solo uno, sin exento)
        'sub0_gas': f"={CL['I0']}{rn}",

        # ── Fórmula 3: SUB2 — Base para IVA 16% ───────────────────
        # sub2 = sub1 - sub0
        # La porción que SÍ genera IVA acreditable
        'sub2': f"={CL['S1']}{rn}-{CL['S0']}{rn}",

        # ── Fórmula 4: IVA ACREDITABLE — Validación contable ──────
        # iva_acred = sub2 * 0.16
        # Lo que el IVA debería ser según la base.
        # Si difiere del IVA declarado en el CFDI → DISCREPANCIA
        'iva_acred': f"={CL['S2']}{rn}*0.16",

        # ── Fórmula 5: C IVA — Diferencia de IVA ─────────────────
        # c_iva = iva_acred - iva16
        # Si ≠ 0 hay error en el CFDI. Celda de alerta.
        'c_iva': f"={CL['IA']}{rn}-{CL['I16']}{rn}",

        # ── Fórmula 6: T2 — Total reconstruido ────────────────────
        # t2 = sub2 + sub0 + iva16
        # Total calculado desde los componentes
        't2': f"={CL['S2']}{rn}+{CL['S0']}{rn}+{CL['I16']}{rn}",

        # ── Fórmula 7: Comprobación — Delta total ─────────────────
        # comprob = total_cfdi - t2
        # Si ≠ 0 hay discrepancia entre total y componentes
        'comprob': f"={CL['TOT']}{rn}-{CL['T2']}{rn}",
    }

# ══════════════════════════════════════════════════════════════════
# OPTIMIZACIÓN v1.8 — Evaluación vectorizada con MASKS pandas
# Reemplaza el iterrows() de v1.7 que era el cuello de botella
#
# En lugar de procesar fila por fila:
#   for _, r in df.iterrows():   ← LENTO
#       evaluar_deducibilidad(r)
#
# Se aplican masks sobre columnas completas:
#   mask = (df['_forma'] == '01') & (df['Total'] > 2000)  ← RÁPIDO
#   df.loc[mask, 'Deducible'] = 'NO'
# ══════════════════════════════════════════════════════════════════

def evaluar_deducibilidad_vectorizado(df: pd.DataFrame) -> pd.DataFrame:
    """
    Evalúa deducibilidad de TODO el DataFrame con masks vectorizadas.
    10-50x más rápido que iterrows().

    Requiere columnas previas:
      _regimen, _uso, _metodo, _forma, _es_gas, _es_insumo,
      _es_dulce, _agrupada, Total (numérico)

    Agrega columnas:
      _deducible ('SI'/'NO')
      _razon     (texto con fundamento)
    """

    # Inicializar — todos deducibles, sin razón
    df['_deducible'] = 'SI'
    df['_razon']     = ''

    total = df['Total'].fillna(0).astype(float)
    forma = df['_forma'].fillna('').astype(str)
    uso   = df['_uso'].fillna('').astype(str)
    met   = df['_metodo'].fillna('').astype(str)
    reg   = df['_regimen'].fillna('').astype(str)

    # ── MASK 1: Uso CFDI inválido ─────────────────────────────────
    mask_uso_inv = ~uso.isin(USOS_DEDUCIBLES)
    df.loc[mask_uso_inv, '_deducible'] = 'NO'
    df.loc[mask_uso_inv, '_razon'] += uso[mask_uso_inv].apply(
        lambda u: f"Uso CFDI {u} no deducible | ")

    # ── MASK 2: Método de pago inválido ──────────────────────────
    mask_met_inv = ~met.isin(METODOS_VALIDOS)
    df.loc[mask_met_inv, '_deducible'] = 'NO'
    df.loc[mask_met_inv, '_razon'] += met[mask_met_inv].apply(
        lambda m: f"Método {m} inválido | ")

    # ── MASK 3: GASOLINA EFECTIVO ─────────────────────────────────
    es_gas       = df['_es_gas'].fillna(False)
    mask_gas_ef  = es_gas & (forma == '01')

    # 3a. Gasolina 626 agrupada (facilidad RESICO) → Deducible
    mask_gas_626_agrup = mask_gas_ef & (reg == '626') & df['_agrupada'].fillna(False)
    df.loc[mask_gas_626_agrup, '_razon'] += \
        "RESICO (626): Gasolina agrupada efectivo — deducible (facilidad) | "

    # 3b. Gasolina 626 ≤$2,000 (facilidad RESICO) → Deducible
    mask_gas_626_menor = mask_gas_ef & (reg == '626') & ~df['_agrupada'].fillna(False) & (total <= LIMITE_EFECTIVO)
    df.loc[mask_gas_626_menor, '_razon'] += \
        "RESICO (626): Gasolina efectivo ≤$2,000 — deducible (facilidad) | "

    # 3c. Gasolina 626 >$2,000 individual → NO deducible
    mask_gas_626_mayor = mask_gas_ef & (reg == '626') & ~df['_agrupada'].fillna(False) & (total > LIMITE_EFECTIVO)
    df.loc[mask_gas_626_mayor, '_deducible'] = 'NO'
    df.loc[mask_gas_626_mayor, '_razon'] += \
        "RESICO (626): Gasolina individual efectivo >$2,000 — NO deducible | "

    # 3d. Gasolina 612 efectivo → SIEMPRE NO deducible
    mask_gas_612 = mask_gas_ef & (reg == '612')
    df.loc[mask_gas_612, '_deducible'] = 'NO'
    df.loc[mask_gas_612, '_razon'] += \
        "Régimen 612: Gasolina efectivo NO deducible. Art. 103 LISR + Art. 27 Fracc. III LISR | "

    # 3e. Gasolina otros regímenes efectivo → NO deducible
    mask_gas_otros = mask_gas_ef & ~reg.isin({'626', '612'})
    df.loc[mask_gas_otros, '_deducible'] = 'NO'
    df.loc[mask_gas_otros, '_razon'] += \
        "Gasolina efectivo NO deducible (Art. 27 Fracc. III LISR) | "

    # ── MASK 4: INSUMOS AGRÍCOLAS ─────────────────────────────────
    # Art. 103 LISR + Art. 27 Fracc. III LISR
    es_insumo = df['_es_insumo'].fillna(False)

    # 4a. Insumo efectivo >$2,000 → NO deducible
    mask_ins_nd = es_insumo & ~es_gas & (forma == '01') & (total > LIMITE_EFECTIVO)
    df.loc[mask_ins_nd, '_deducible'] = 'NO'
    df.loc[mask_ins_nd, '_razon'] += \
        "Insumo agrícola efectivo >$2,000 NO deducible. " \
        "Art. 103 LISR + Art. 27 Fracc. III LISR. " \
        "Facilidad AGAPES (Art. 74 LISR) NO aplica a Régimen 612 | "

    # 4b. Insumo efectivo ≤$2,000 → Deducible
    mask_ins_menor = es_insumo & ~es_gas & (forma == '01') & (total <= LIMITE_EFECTIVO)
    df.loc[mask_ins_menor, '_razon'] += \
        "Insumo agrícola efectivo ≤$2,000: deducible. Art. 27 Fracc. III LISR | "

    # 4c. Insumo pago electrónico → Deducible sin límite
    mask_ins_elec = es_insumo & ~es_gas & forma.isin(FORMAS_ELECTRONICAS)
    df.loc[mask_ins_elec, '_razon'] += \
        "Insumo agrícola pago electrónico: deducible. Art. 27 Fracc. III LISR cumplido | "

    # 4d. Insumo forma inválida
    mask_ins_inv = es_insumo & ~es_gas & ~(forma == '01') & ~forma.isin(FORMAS_ELECTRONICAS)
    df.loc[mask_ins_inv, '_deducible'] = 'NO'
    df.loc[mask_ins_inv, '_razon'] += forma[mask_ins_inv].apply(
        lambda f: f"Insumo agrícola: forma de pago {f} inválida | ")

    # ── MASK 5: GASTOS NORMALES efectivo >$2,000 ──────────────────
    mask_normal_nd = ~es_gas & ~es_insumo & (forma == '01') & (total > LIMITE_EFECTIVO)
    df.loc[mask_normal_nd, '_deducible'] = 'NO'
    df.loc[mask_normal_nd, '_razon'] += \
        "Efectivo >$2,000 NO deducible (Art. 27 Fracc. III LISR) | "

    # ── MASK 6: Forma de pago inválida (gastos normales) ──────────
    mask_forma_inv = ~es_gas & ~es_insumo & ~forma.isin(FORMAS_VALIDAS)
    df.loc[mask_forma_inv, '_deducible'] = 'NO'
    df.loc[mask_forma_inv, '_razon'] += forma[mask_forma_inv].apply(
        lambda f: f"Forma de pago {f} inválida | ")

    # ── MASK 7: Regímenes no trabajados — advertencia ─────────────
    mask_reg_ext = ~reg.isin(REGIMENES_TRABAJADOS)
    df.loc[mask_reg_ext, '_razon'] += reg[mask_reg_ext].apply(
        lambda r: f"⚠️ Régimen {r}: Verificar manualmente | ")

    # ── Limpiar razones vacías ────────────────────────────────────
    df['_razon'] = df['_razon'].str.rstrip(' |').str.strip()
    df.loc[df['_razon'] == '', '_razon'] = 'Cumple requisitos'

    return df


# ══════════════════════════════════════════════════════════════════
# Evaluación fila por fila — usada por motor_openpyxl.py
# Se mantiene para compatibilidad con el motor openpyxl
# ══════════════════════════════════════════════════════════════════

def evaluar_deducibilidad(
    uso_cfdi: str, metodo_pg: str, forma_pg: str,
    regimen: str, total: float,
    es_gas: bool, es_insumo: bool, concepto_lower: str
) -> tuple:
    """
    Evaluación fila por fila para motor openpyxl.
    Retorna: (es_deducible: bool, razones: list[str])
    """
    es_deducible = True
    razones      = []

    if uso_cfdi not in USOS_DEDUCIBLES:
        es_deducible = False
        razones.append(f"Uso CFDI {uso_cfdi} no deducible")

    if metodo_pg not in METODOS_VALIDOS:
        es_deducible = False
        razones.append(f"Método de pago {metodo_pg} inválido")

    if es_gas:
        if forma_pg == '01':
            if regimen == '626':
                if es_gasolina_agrupada(concepto_lower):
                    nd = concepto_lower.count('|') + 1
                    razones.append(
                        f"RESICO (626): {nd} despachos agrupados efectivo — deducible (facilidad)")
                elif total <= LIMITE_EFECTIVO:
                    razones.append("RESICO (626): Gasolina efectivo ≤$2,000 — deducible (facilidad)")
                else:
                    es_deducible = False
                    razones.append(
                        "RESICO (626): Gasolina individual efectivo >$2,000 — NO deducible")
            elif regimen == '612':
                es_deducible = False
                razones.append(
                    "Régimen 612: Gasolina efectivo NO deducible. "
                    "Art. 103 LISR + Art. 27 Fracc. III LISR.")
            else:
                es_deducible = False
                razones.append("Gasolina efectivo NO deducible (Art. 27 Fracc. III LISR)")
        else:
            if forma_pg not in FORMAS_ELECTRONICAS:
                es_deducible = False
                razones.append(f"Gasolina: forma de pago {forma_pg} inválida")

    elif es_insumo:
        if forma_pg == '01' and total > LIMITE_EFECTIVO:
            es_deducible = False
            razones.append(
                "Insumo agrícola efectivo >$2,000 NO deducible. "
                "Art. 103 LISR + Art. 27 Fracc. III LISR. "
                "Facilidad AGAPES (Art. 74 LISR) NO aplica a Régimen 612.")
        elif forma_pg == '01' and total <= LIMITE_EFECTIVO:
            razones.append(
                "Insumo agrícola efectivo ≤$2,000: deducible. Art. 27 Fracc. III LISR.")
        elif forma_pg in FORMAS_ELECTRONICAS:
            razones.append(
                "Insumo agrícola pago electrónico: deducible. Art. 27 Fracc. III LISR cumplido.")
        else:
            es_deducible = False
            razones.append(f"Insumo agrícola: forma de pago {forma_pg} inválida.")
    else:
        if forma_pg == '01' and total > LIMITE_EFECTIVO:
            es_deducible = False
            razones.append("Efectivo >$2,000 NO deducible (Art. 27 Fracc. III LISR)")
        elif forma_pg not in FORMAS_VALIDAS:
            es_deducible = False
            razones.append(f"Forma de pago {forma_pg} inválida")

    if regimen not in REGIMENES_TRABAJADOS:
        razones.append(f"⚠️ Régimen {regimen}: Verificar manualmente")

    return es_deducible, razones
