import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
import re

# Configuración inicial
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop/GASTOS RESICO/2026/DICIEMBRE25')
file_name = input("Ingrese el nombre del archivo Excel (sin extensión .xlsx): ")

# Limpiar y formatear el nombre del archivo
file_name = file_name.strip()
file_name = re.sub(r'\.xlsx$', '', file_name, flags=re.IGNORECASE) + '.xlsx'
file_path = os.path.join(desktop_path, file_name)

# Estilos
blue_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
purple_fill = PatternFill(start_color='800080', end_color='800080', fill_type='solid')
orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
result_style = Font(bold=True)
center_align = Alignment(horizontal='center')

# ==============================
# Funciones auxiliares
# ==============================

def extract_code(value):
    if value and '-' in str(value):
        return str(value).split('-')[0].strip()
    return str(value).strip() if value else ''

def find_column(sheet, column_name):
    """Busca una columna por nombre (case insensitive)"""
    for cell in sheet[1]:
        if cell.value and str(cell.value).strip().lower() == column_name.strip().lower():
            return cell.column
    return None

def create_column_if_missing(sheet, column_name, fill_color=blue_fill):
    """Crea una columna si no existe y devuelve su posición"""
    col = find_column(sheet, column_name)
    if col is None:
        last_col = sheet.max_column + 1
        sheet.cell(row=1, column=last_col, value=column_name)
        sheet.cell(row=1, column=last_col).fill = fill_color
        sheet.cell(row=1, column=last_col).alignment = center_align
        print(f"✅ Columna creada: '{column_name}' en posición {get_column_letter(last_col)}")
        return last_col
    return col

def es_gasolina(concepto):
    """Determina si un concepto es de gasolina"""
    palabras_gasolina = [
        'gasolina', 'combustible', 'magna', 'premium', 'diesel', 'diésel',
        'gasohol', 'gasoil', 'nafta', 'petrol', 'gas', 'energético'
    ]
    concepto = str(concepto).lower()
    for palabra in palabras_gasolina:
        if palabra in concepto:
            return True
    return False

# ==============================
# Reglas de deducibilidad
# ==============================

USOS_DEDUCIBLES = ['G01', 'G02', 'G03']
METODOS_VALIDOS = ['PUE', 'PPD']
FORMAS_VALIDAS = ['01', '02', '03', '04', '28']
USO_CFDI_VERDE = "G02 - Devoluciones, descuentos o bonificaciones"

# ==============================
# Abrir archivo Excel
# ==============================

try:
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
except Exception as e:
    print(f"Error al abrir el archivo: {e}")
    raise SystemExit(1)
# ==============================
# Columnas requeridas
# ==============================

required_columns = {
    'SubTotal': 'SubTotal',
    'Descuento': 'Descuento',
    'IVA Trasladado 0%': 'IVA Trasladado 0%',
    'IVA Exento': 'IVA Exento',
    'IVA Trasladado 16%': 'IVA Trasladado 16%',
    'Total': 'Total',
    'Uso CFDI': 'Uso CFDI',
    'Metodo pago': 'Metodo pago',
    'Forma pago': 'Forma pago',
    'Regimen receptor': 'Regimen receptor',
    'Razon emisor': 'Razon emisor',
    'Conceptos': 'Conceptos'
}

print("🔍 Buscando columnas requeridas...")
columns = {}

# Crear columnas requeridas si no existen
for key, col_name in required_columns.items():
    columns[key] = create_column_if_missing(sheet, col_name)
    print(f"   {col_name}: Columna {get_column_letter(columns[key])}")

# ==============================
# Columnas IEPS (opcionales)
# ==============================

ieps_encontrado = False
ieps_no_desglosado_encontrado = False

# IEPS Trasladado
ieps_col = find_column(sheet, 'IEPS Trasladado')
if ieps_col:
    columns['IEPS Trasladado'] = ieps_col
    ieps_encontrado = True
    print("✅ Columna IEPS Trasladado encontrada")
else:
    print("⚠️  Advertencia: No se encontró la columna 'IEPS Trasladado'")

# IEPS Trasladado No Desglosado
ieps_no_desglosado_col = find_column(sheet, 'IEPS Trasladado No Desglosado')
if ieps_no_desglosado_col:
    columns['IEPS Trasladado No Desglosado'] = ieps_no_desglosado_col
    ieps_no_desglosado_encontrado = True
    print("✅ Columna IEPS Trasladado No Desglosado encontrada")
else:
    print("⚠️  Advertencia: No se encontró la columna 'IEPS Trasladado No Desglosado'")

# ==============================
# Columna Efecto (si existe)
# ==============================

efecto_col = find_column(sheet, 'Efecto')
if efecto_col:
    print(f"✅ Columna Efecto encontrada: {get_column_letter(efecto_col)}")

# ==============================
# Inicializar columnas numéricas nuevas
# ==============================

print("📝 Inicializando columnas nuevas...")
for row in range(2, sheet.max_row + 1):
    for col_key in ['IVA Trasladado 16%', 'IVA Trasladado 0%', 'IVA Exento', 'Descuento']:
        col_idx = columns[col_key]
        cell = sheet.cell(row=row, column=col_idx)
        if cell.value is None:
            cell.value = 0
            cell.number_format = "0.00"

# ==============================
# Crear columnas de cálculo
# ==============================

last_column = sheet.max_column
headers = [
    'SUB1-16%',
    'SUB0%',
    'SUB2-16%',
    'IVA ACREDITABLE 16%',
    'C IVA',
    'T2',
    'Comprobación T2',
    'Deducible'
]

print("🧮 Creando columnas de cálculo...")
for i, header in enumerate(headers, start=1):
    hcell = sheet.cell(row=1, column=last_column + i, value=header)
    hcell.fill = blue_fill
    hcell.alignment = center_align
    print(f"   {header}: Columna {get_column_letter(last_column + i)}")

sub1_col = last_column + 1
sub0_col = last_column + 2
sub2_col = last_column + 3
iva_acred_col = last_column + 4
c_iva_col = last_column + 5
t2_col = last_column + 6
comprob_col = last_column + 7
deducible_col = last_column + 8

# ==============================
# Variables de control
# ==============================

ieps_procesados = 0
ieps_no_desglosado_procesados = 0
gasolina_con_ieps = 0
gasolina_sin_ieps = 0
uso_s01_count = 0

# 🔴 CONTADOR BUG EFECTIVO > 2000
efectivo_mayor_2000 = 0

print("🔧 Procesando filas...")
# ==============================
# Procesar filas
# ==============================

for row in range(2, sheet.max_row + 1):

    concepto = str(sheet.cell(row=row, column=columns['Conceptos']).value or '')
    iva0_val = sheet.cell(row=row, column=columns['IVA Trasladado 0%']).value or 0
    iva_exento_val = sheet.cell(row=row, column=columns['IVA Exento']).value or 0
    iva16_val = sheet.cell(row=row, column=columns['IVA Trasladado 16%']).value or 0
    total_val = sheet.cell(row=row, column=columns['Total']).value or 0
    uso_cfdi_r = sheet.cell(row=row, column=columns['Uso CFDI']).value

    # ------------------------------
    # Detección gasolina
    # ------------------------------
    es_gasolina_concepto = es_gasolina(concepto)

    tiene_ieps = False
    tiene_ieps_no_desglosado = False

    if ieps_encontrado:
        ieps_val = sheet.cell(row=row, column=columns['IEPS Trasladado']).value
        if ieps_val not in (None, 0, ''):
            tiene_ieps = True

    if ieps_no_desglosado_encontrado:
        ieps_nd_val = sheet.cell(row=row, column=columns['IEPS Trasladado No Desglosado']).value
        if ieps_nd_val not in (None, 0, ''):
            tiene_ieps_no_desglosado = True

    if es_gasolina_concepto:
        concepto_cell = sheet.cell(row=row, column=columns['Conceptos'])
        if tiene_ieps or tiene_ieps_no_desglosado:
            concepto_cell.fill = blue_fill
            gasolina_con_ieps += 1
        else:
            concepto_cell.fill = orange_fill
            gasolina_sin_ieps += 1

    # ------------------------------
    # SUB1-16%
    # ------------------------------
    sheet.cell(row=row, column=sub1_col).value = f"=({get_column_letter(columns['SubTotal'])}{row}-{get_column_letter(columns['Descuento'])}{row})"
    sheet.cell(row=row, column=sub1_col).number_format = "0.00"
    sheet.cell(row=row, column=sub1_col).font = result_style

    # ------------------------------
    # IEPS -> IVA 0% (SOLO GASOLINA)
    # ------------------------------
    if es_gasolina_concepto and (ieps_encontrado or ieps_no_desglosado_encontrado):
        for col in range(1, sheet.max_column + 1):
            header = sheet.cell(row=1, column=col).value
            if header and "IEPS" in str(header).upper():
                ieps_val = sheet.cell(row=row, column=col).value
                if ieps_val not in (None, 0, ''):
                    sheet.cell(row=row, column=columns['IVA Trasladado 0%']).value = ieps_val
                    sheet.cell(row=row, column=columns['IVA Trasladado 0%']).fill = orange_fill
                    ieps_procesados += 1

    # ------------------------------
    # SUB0%
    # ------------------------------
    sheet.cell(row=row, column=sub0_col).value = f"={get_column_letter(columns['IVA Trasladado 0%'])}{row}"
    sheet.cell(row=row, column=sub0_col).number_format = "0.00"
    sheet.cell(row=row, column=sub0_col).font = result_style

    # ------------------------------
    # SUB2-16%
    # ------------------------------
    sheet.cell(row=row, column=sub2_col).value = f"={get_column_letter(sub1_col)}{row}-{get_column_letter(sub0_col)}{row}"
    sheet.cell(row=row, column=sub2_col).number_format = "0.00"
    sheet.cell(row=row, column=sub2_col).font = result_style

    # ------------------------------
    # IVA acreditable
    # ------------------------------
    sheet.cell(row=row, column=iva_acred_col).value = f"={get_column_letter(sub2_col)}{row}*0.16"
    sheet.cell(row=row, column=iva_acred_col).number_format = "0.00"
    sheet.cell(row=row, column=iva_acred_col).font = result_style
    
    # ==============================
    # 🔍 COMPROBACIÓN: IVA TRASLADADO 16% == IVA ACREDITABLE 16%
    # (Cálculo en Python, NO lectura de fórmula)
    # ==============================
    try:
        subtotal_val = float(sheet.cell(row=row, column=columns['SubTotal']).value or 0)
        descuento_val = float(sheet.cell(row=row, column=columns['Descuento']).value or 0)
        iva0_val_num = float(sheet.cell(row=row, column=columns['IVA Trasladado 0%']).value or 0)
        iva16_tras_val = float(sheet.cell(row=row, column=columns['IVA Trasladado 16%']).value or 0)

        # Recalcular base 16%
        sub1_val = subtotal_val - descuento_val
        sub2_val = sub1_val - iva0_val_num

        iva_acred_calc = round(sub2_val * 0.16, 2)
        iva16_tras_val = round(iva16_tras_val, 2)

        # Comparar con tolerancia
        if abs(iva_acred_calc - iva16_tras_val) < 0.01:
            sheet.cell(row=row, column=columns['IVA Trasladado 16%']).fill = green_fill
            sheet.cell(row=row, column=iva_acred_col).fill = green_fill

    except Exception:
        pass



    # ------------------------------
    # C IVA
    # ------------------------------
    sheet.cell(row=row, column=c_iva_col).value = f"={get_column_letter(iva_acred_col)}{row}-{get_column_letter(columns['IVA Trasladado 16%'])}{row}"
    sheet.cell(row=row, column=c_iva_col).number_format = "0.00"
    sheet.cell(row=row, column=c_iva_col).font = result_style

    # ------------------------------
    # T2
    # ------------------------------
    sheet.cell(row=row, column=t2_col).value = f"={get_column_letter(sub2_col)}{row}+{get_column_letter(sub0_col)}{row}+{get_column_letter(columns['IVA Trasladado 16%'])}{row}"
    sheet.cell(row=row, column=t2_col).number_format = "0.00"
    sheet.cell(row=row, column=t2_col).font = result_style

    # ------------------------------
    # Comprobación T2
    # ------------------------------
    sheet.cell(row=row, column=comprob_col).value = f"={get_column_letter(columns['Total'])}{row}-{get_column_letter(t2_col)}{row}"
    sheet.cell(row=row, column=comprob_col).number_format = "0.00"
    sheet.cell(row=row, column=comprob_col).font = result_style

    # ------------------------------
    # Régimen receptor
    # ------------------------------
    regimen = extract_code(sheet.cell(row=row, column=columns['Regimen receptor']).value)
    if regimen == '626':
        sheet.cell(row=row, column=columns['Regimen receptor']).fill = blue_fill
    elif regimen == '612':
        sheet.cell(row=row, column=columns['Regimen receptor']).fill = purple_fill
    else:
        sheet.cell(row=row, column=columns['Regimen receptor']).fill = orange_fill
        sheet.cell(row=row, column=columns['Razon emisor']).fill = orange_fill

    # ------------------------------
    # Uso CFDI
    # ------------------------------
    uso_cfdi_codigo = extract_code(uso_cfdi_r).upper() if uso_cfdi_r else ''
    if uso_cfdi_r and str(uso_cfdi_r).strip() == USO_CFDI_VERDE:
        sheet.cell(row=row, column=columns['Uso CFDI']).fill = green_fill
    if uso_cfdi_codigo == 'S01':
        sheet.cell(row=row, column=columns['Uso CFDI']).fill = red_fill
        uso_s01_count += 1

    # ------------------------------
    # Efecto
    # ------------------------------
    es_egreso = False
    if efecto_col:
        efecto_val = sheet.cell(row=row, column=efecto_col).value
        if efecto_val and str(efecto_val).strip().upper() in ['EGRESO', 'E']:
            es_egreso = True

    # ------------------------------
    # BUG FIX: EFECTIVO > $2,000
    # ------------------------------
    metodo_pg = extract_code(sheet.cell(row=row, column=columns['Metodo pago']).value).upper()
    forma_pg = extract_code(sheet.cell(row=row, column=columns['Forma pago']).value).strip().upper()

    efectivo_mayor_2000_flag = False
    if forma_pg == '01' and total_val > 2000:
        efectivo_mayor_2000_flag = True
        efectivo_mayor_2000 += 1

    # ------------------------------
    # Deducibilidad FINAL
    # ------------------------------
    uso_cfdi = extract_code(uso_cfdi_r).upper() if uso_cfdi_r else ''

    es_deducible = (
        uso_cfdi in USOS_DEDUCIBLES and
        metodo_pg in METODOS_VALIDOS and
        forma_pg in FORMAS_VALIDAS and
        not efectivo_mayor_2000_flag
    )

    ded_cell = sheet.cell(row=row, column=deducible_col, value="SI" if es_deducible else "NO")
    if es_deducible:
        ded_cell.fill = blue_fill if es_egreso else green_fill
    else:
        ded_cell.fill = red_fill
    ded_cell.font = Font(bold=True, color='FFFFFF')
    ded_cell.alignment = center_align

    if efectivo_mayor_2000_flag:
        ded_cell.fill = red_fill
        if efecto_col:
            sheet.cell(row=row, column=efecto_col).fill = red_fill

# ==============================
# Ajustar ancho columnas
# ==============================

for col in range(last_column + 1, last_column + len(headers) + 1):
    sheet.column_dimensions[get_column_letter(col)].width = 15

# ==============================
# Guardar archivo
# ==============================

try:
    base_name = re.sub(r'\.xlsx$', '', file_name, flags=re.IGNORECASE)
    output_name = os.path.join(desktop_path, f"{base_name}_validado.xlsx")
    workbook.save(output_name)

    print("\n" + "=" * 80)
    print("✅ PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 80)
    print(f"📂 Archivo guardado como: {output_name}")
    print(f"📊 Total de filas procesadas: {sheet.max_row - 1}")

    print("\n🎨 FORMATOS APLICADOS:")
    print("  • Régimen 626: AZUL")
    print("  • Régimen 612: MORADO")
    print("  • Otros regímenes: NARANJA")
    print("  • Uso CFDI G02: VERDE")
    print("  • Uso CFDI S01: ROJO")
    print("  • Deducible egreso: AZUL")
    print("  • Deducible otros: VERDE")
    print("  • No deducible / Efectivo > $2,000: ROJO")

    print(f"  • IEPS Trasladado procesados: {ieps_procesados}")
    print(f"  • IEPS No Desglosado procesados: {ieps_no_desglosado_procesados}")
    print(f"  • Gasolina con IEPS: {gasolina_con_ieps}")
    print(f"  • Gasolina sin IEPS: {gasolina_sin_ieps}")
    print(f"  • Usos S01: {uso_s01_count}")
    print(f"  • Efectivo > $2,000: {efectivo_mayor_2000}")

    print("=" * 80)

except Exception as e:
    print(f"\n❌ Error al guardar el archivo: {e}")

