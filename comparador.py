# ═══════════════════════════════════════════════════════════════════
# VERSIÓN v4 — 2ª hoja "Prima Total" con totales grandes por plan
# Si esta línea está en tu Celda 2 = tienes la versión correcta.
# ═══════════════════════════════════════════════════════════════════

# -*- coding: utf-8 -*-
"""
COMPARATIVO DE COTIZACIONES DE SEGUROS DE AUTO
Olivas & Asociados - Seguros y Fianzas
Procesa cotizaciones (HDI, CHUBB, QUÁLITAS, GNP, INBURSA, ANA) y genera PDF comparativo.

CORRECCIONES v2:
  1. detectar_aseguradora: GNP se evalúa AL FINAL (su marca aparece a veces
     dentro de PDFs de otras aseguradoras y las "robaba").
  2. parse_gnp: la prima ya no puede confundirse con los descuentos
     (Fidelidad GNP / Campañas). Funciona para pago de Contado, Semestral,
     Trimestral y Mensual.
  3. procesar_cotizaciones: avisa si dos PDFs se detectan como la MISMA
     aseguradora (antes uno pisaba al otro en silencio).
"""

import pdfplumber
import re
import os
from reportlab.lib.pagesizes import letter, landscape, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT


def extraer_texto(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return '\n'.join([(page.extract_text() or '') for page in pdf.pages])


def detectar_aseguradora(texto):
    """Detecta la aseguradora. IMPORTANTE: GNP va al final porque su check es
    el más 'suelto' ('GNP' + 'SEGURO') y puede aparecer mencionada dentro de
    cotizaciones de otras compañías (p. ej. descuento 'Fidelidad GNP' citado
    en un comparativo, o pies de página)."""
    t = texto.upper()
    if 'HDI SEGUROS' in t or 'HDI.COM.MX' in t:
        return 'HDI'
    if 'CHUBB' in t:
        return 'CHUBB'
    qualitas_marks = (
        'QUALITAS', 'QUÁLITAS',
        'COTIZACION DE SEGURO DE AUTOMOVILES',
        'TASA FIN.P.F.',
        'GTOS. EXPEDICION POL.',
        'NO COTIZACION INTERNO',
        'TARIFA APLICADA',
    )
    if any(m in t for m in qualitas_marks):
        return 'QUALITAS'
    # ANA e INBURSA se evalúan ANTES que GNP (fix del bug de clasificación)
    if ('ANA SEGUROS' in t) or ('A.N.A. COMPA' in t) or ('ANASEGUROS' in t) or \
       ('ANA ASISTENCIA' in t and 'COTIZACIÓN DE SEGUROS AUTOMÓVILES' in t) or \
       ('ANA ASISTENCIA VIAL' in t) or ('DESBIELAMIENTO POR AGUA' in t):
        return 'ANA'
    if 'INBURSA' in t:
        return 'INBURSA'
    # GNP siempre al último
    if ('COTIZACIÓN SEGURO DE AUTOS GNP' in t) or ('COTIZACION SEGURO DE AUTOS GNP' in t):
        return 'GNP'
    if 'GNP' in t and ('SEGURO' in t or 'COTIZACIÓN' in t or 'COTIZACION' in t):
        return 'GNP'
    return None


def fmt_monto(valor):
    if valor is None:
        return 'NO AMPARADA'
    s = str(valor).replace(',', '').strip()
    try:
        n = float(s)
        return f"$ {n:,.2f}"
    except ValueError:
        return str(valor)


def fmt_suma(valor):
    if valor is None:
        return 'NO AMPARADA'
    s = str(valor).strip()
    if s.replace(',', '').replace('.', '').isdigit():
        return fmt_monto(s)
    return s


def normalizar_valor(valor):
    if valor is None:
        return 'NO AMPARADA'
    s = str(valor).strip()
    if s.upper() in ('AMPARADA', 'AMPARADO'):
        return 'AMPARADA'
    if s.upper() == 'VALOR COMERCIAL':
        return 'VALOR COMERCIAL'
    if s.upper() in ('NO AMPARADA', 'NO AMPARADO'):
        return 'NO AMPARADA'
    s = re.sub(r'\s*POR EVENTO\s*', '', s, flags=re.IGNORECASE).strip()
    m = re.match(r'^(\d+)(?:\.0+)?\s*%$', s)
    if m:
        return f"{m.group(1)}%"
    m = re.match(r'^(\d+)\s*UMAS?$', s, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return f"{n} UMA" if n <= 1 else f"{n} UMAS"
    if re.match(r'^TIPO CARGA', s, re.IGNORECASE):
        return s.upper()
    return s


def normalizar_coberturas(datos):
    cob = datos.get('coberturas', {})
    for k in cob:
        suma, ded = cob[k]
        cob[k] = (normalizar_valor(suma), normalizar_valor(ded))
    return datos


def buscar(patron, texto, grupo=1, default=None, flags=re.IGNORECASE):
    m = re.search(patron, texto, flags)
    return m.group(grupo).strip() if m else default


# ------------------------------------------------------------------
# Extracción de planes de pago (Contado / Semestral / Trimestral)
# ------------------------------------------------------------------
# Devuelve dict con 5 llaves:
#   contado, sem_pi (pago inicial), sem_ps (pago subsecuente),
#   tri_pi (pago inicial trimestral), tri_ps (pago subsecuente trimestral)
# Si un plan no está en el PDF, sus llaves quedan en None y en el
# comparativo se mostrará un guion "—".

def _planes_vacio():
    return {'contado': None, 'sem_pi': None, 'sem_ps': None,
            'tri_pi': None, 'tri_ps': None}


def _planes_qualitas(texto):
    """Qualitas muestra los 3 planes juntos en la sección 'FORMA DE PAGO'.
    Formato: CONTADO/SEMESTRAL/TRIMESTRAL <pago_inicial> [pago_subsecuente]"""
    planes = _planes_vacio()
    seccion = re.search(r'FORMA\s+DE\s+PAGO[:\s]([\s\S]{0,600})', texto, re.IGNORECASE)
    s = seccion.group(1) if seccion else texto
    m = re.search(r'\bCONTADO\s+([\d,]+\.\d{2})', s)
    if m: planes['contado'] = m.group(1)
    m = re.search(r'\bSEMESTRAL\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})', s)
    if m:
        planes['sem_pi'] = m.group(1)
        planes['sem_ps'] = m.group(2)
    m = re.search(r'\bTRIMESTRAL\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})', s)
    if m:
        planes['tri_pi'] = m.group(1)
        planes['tri_ps'] = m.group(2)
    return planes


def _planes_gnp(texto):
    """GNP normalmente muestra UN plan por PDF ('Forma de pago X').
    Extraemos ese plan; los otros quedan como None (mostrarán '—').
    Si futuras cotizaciones GNP muestran los 3 planes juntos, agregar
    aquí patrones adicionales."""
    planes = _planes_vacio()
    m_forma = re.search(r'Forma\s+de\s+pago\s+(Contado|Semestral|Trimestral|Mensual)',
                        texto, re.IGNORECASE)
    if not m_forma:
        return planes
    forma = m_forma.group(1).upper()
    # Detectar el pago único de contado ('1 pago de $ X') o el 1er pago
    # fraccionado ('1er pago de $ X') + los subsecuentes ('N pagos de $ Y')
    m1_contado = re.search(r'\b1\s+pago\s+de\s+\$?\s*([\d,]+(?:\.\d{2})?)', texto, re.IGNORECASE)
    m1_frac    = re.search(r'\b1er\s+pago\s+de\s+\$?\s*([\d,]+(?:\.\d{2})?)', texto, re.IGNORECASE)
    m_subs     = re.search(r'\b(\d+)\s+pagos?\s+de\s+\$?\s*([\d,]+(?:\.\d{2})?)', texto, re.IGNORECASE)
    if forma == 'CONTADO' and m1_contado:
        planes['contado'] = m1_contado.group(1)
    elif forma == 'SEMESTRAL':
        if m1_frac: planes['sem_pi'] = m1_frac.group(1)
        if m_subs:  planes['sem_ps'] = m_subs.group(2)
    elif forma == 'TRIMESTRAL':
        if m1_frac: planes['tri_pi'] = m1_frac.group(1)
        if m_subs:  planes['tri_ps'] = m_subs.group(2)

    # La 'Prima total anual' de GNP es el equivalente al pago de CONTADO.
    # Aunque el PDF esté cotizado a Trimestral, ese monto sirve para
    # comparar contra el Contado de las demás aseguradoras.
    if forma == 'CONTADO' and not planes['contado']:
        prima_anual = extraer_prima_gnp(texto)
        if prima_anual:
            planes['contado'] = prima_anual
    return planes


def _planes_chubb(texto):
    """Primero intenta el formato COTIZACIÓN (tabla 'Prima total' con los 3
    planes); si no existe, cae al formato PÓLIZA."""
    planes = _planes_chubb_cotizacion(texto)
    if any(planes.values()):
        return planes
    return _planes_chubb_poliza(texto)


def _num_espaciado(s):
    """CHUBB a veces extrae '2 4 , 6 4 2 .7 5' -> '24,642.75'."""
    return re.sub(r'\s+', '', s)


def _planes_chubb_cotizacion(texto):
    planes = _planes_vacio()
    m = re.search(r'Contado\s*\(Recargos[^)]*\)\s*Semestral\s*\(Recargos[^)]*\)\s*'
                  r'Trimestral\s*\(Recargos[^)]*\)\s*\$\s*([\d,]+\.\d{2})\s*\$\s*([\d,]+\.\d{2})'
                  r'\s*\$\s*([\d,]+\.\d{2})', texto, re.IGNORECASE)
    if not m:
        return planes
    planes['contado'] = m.group(1)
    resto = texto[m.end():m.end() + 400]
    num = r'\$\s*([\d\s,]+\.\s*[\d\s]{1,3}\d)'
    primeros = re.findall(r'1er\s+Pago\s+de\s*' + num, resto, re.IGNORECASE)
    m2 = re.search(r'2do\s+Pago\s+de\s*' + num, resto, re.IGNORECASE)
    m3 = re.search(r'3\s+Pagos\s+de\s*' + num, resto, re.IGNORECASE)
    if len(primeros) >= 1: planes['sem_pi'] = _num_espaciado(primeros[0])
    if m2: planes['sem_ps'] = _num_espaciado(m2.group(1))
    if len(primeros) >= 2: planes['tri_pi'] = _num_espaciado(primeros[1])
    if m3: planes['tri_ps'] = _num_espaciado(m3.group(1))
    return planes


def _planes_chubb_poliza(texto):
    """CHUBB emite PÓLIZAS (no cotizaciones comparativas), así que solo trae
    el plan YA contratado. Extraemos lo que sí existe:
      - 'Forma de pago: TRIMESTRAL' → qué plan es
      - 'Prima total  15,561.76'    → total del plan (ya con financiamiento)
      - 'Total a pagar: 5,197.43'   → primer pago (del aviso de cobro)
      - 'Serie del aviso 1/4'       → número de pagos
    El pago subsecuente se calcula: (total - primer_pago) / (n_pagos - 1).
    """
    planes = _planes_vacio()
    m_forma = re.search(r'Forma\s+de\s+pago:?\s*(CONTADO|SEMESTRAL|TRIMESTRAL|MENSUAL|ANUAL)',
                        texto, re.IGNORECASE)
    if not m_forma:
        return planes
    forma = m_forma.group(1).upper()

    m_total = re.search(r'Prima\s+total\s+([\d,]+\.\d{2})', texto, re.IGNORECASE)
    total = float(m_total.group(1).replace(',', '')) if m_total else None

    m_pi = re.search(r'Total\s+a\s+pagar:?\s*\$?\s*([\d,]+\.\d{2})', texto, re.IGNORECASE)
    primer = float(m_pi.group(1).replace(',', '')) if m_pi else None

    m_serie = re.search(r'Serie\s+del\s+aviso\s*(\d+)\s*/\s*(\d+)', texto, re.IGNORECASE)
    n_pagos = int(m_serie.group(2)) if m_serie else None

    if forma in ('CONTADO', 'ANUAL') and total:
        planes['contado'] = f"{total:,.2f}"
        return planes

    # Fraccionado: calcular el subsecuente si tenemos total, primer pago y n
    subsec = None
    if total and primer and n_pagos and n_pagos > 1:
        subsec = (total - primer) / (n_pagos - 1)

    if forma == 'SEMESTRAL':
        if primer: planes['sem_pi'] = f"{primer:,.2f}"
        if subsec: planes['sem_ps'] = f"{subsec:,.2f}"
    elif forma == 'TRIMESTRAL':
        if primer: planes['tri_pi'] = f"{primer:,.2f}"
        if subsec: planes['tri_ps'] = f"{subsec:,.2f}"
    return planes


def _planes_generico(texto):
    """Intento genérico para HDI/CHUBB/ANA/INBURSA. Prueba varios patrones
    típicos. Si no encuentra, retorna Nones (mostrarán '—')."""
    planes = _planes_vacio()
    # Patrón tipo tabla: "CONTADO $ X", "SEMESTRAL $ X $ Y", etc.
    m = re.search(r'\bCONTADO[\s:$]+([\d,]+\.\d{2})', texto, re.IGNORECASE)
    if m: planes['contado'] = m.group(1)
    m = re.search(r'\bSEMESTRAL[\s:$]+([\d,]+\.\d{2})[\s$]+([\d,]+\.\d{2})',
                  texto, re.IGNORECASE)
    if m:
        planes['sem_pi'] = m.group(1)
        planes['sem_ps'] = m.group(2)
    m = re.search(r'\bTRIMESTRAL[\s:$]+([\d,]+\.\d{2})[\s$]+([\d,]+\.\d{2})',
                  texto, re.IGNORECASE)
    if m:
        planes['tri_pi'] = m.group(1)
        planes['tri_ps'] = m.group(2)
    return planes


def extraer_planes_pago(texto, aseguradora):
    """Router de extracción por aseguradora."""
    aseg = (aseguradora or '').upper()
    if 'QUALITAS' in aseg or 'QUÁLITAS' in aseg:
        return _planes_qualitas(texto)
    if aseg == 'GNP':
        return _planes_gnp(texto)
    if aseg == 'CHUBB':
        return _planes_chubb(texto)
    return _planes_generico(texto)


def parse_hdi(texto):
    d = {'aseguradora': 'HDI', 'coberturas': {}}
    d['cliente'] = buscar(r'OLIVAS Y ASOCIADOS.*?\n.*?Nombre:\s*([^\n]*)', texto, default='')
    # Fecha: acepta "15/01/2026" o "15 DE ENERO DE 2026"
    d['vigencia_ini'] = buscar(r'Desde las.*?del\s+(\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+DE\s+\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4})', texto)
    d['vigencia_fin'] = buscar(r'Hasta las.*?del\s+(\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+DE\s+\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4})', texto)
    d['vehiculo'] = buscar(r'(DODGE,?\s*RAM[^\n]+)', texto)
    d['version'] = buscar(r'Versión:\s*([^\n]+?)(?:Transmisión|$)', texto)
    d['cp'] = buscar(r'C\.P\.\s*(\d{5})', texto)
    d['serie'] = buscar(r'serie:?\s*([A-Z0-9]*\d[A-Z0-9]{3,24})', texto, default='')
    d['paquete'] = buscar(r'Paquete:\s*([^\n]+?)(?:Tipo Suma|$)', texto)
    cob = d['coberturas']
    m = re.search(r'Daños\s+Materiales\s+Amparada\s+(\d+)%', texto, re.IGNORECASE)
    if m:
        cob['DAÑOS MATERIALES'] = ('VALOR COMERCIAL', f'{m.group(1)} %')
    m = re.search(r'Robo\s+Total\s+Amparada\s+(\d+)%', texto, re.IGNORECASE)
    if m:
        cob['ROBO TOTAL'] = ('VALOR COMERCIAL', f'{m.group(1)} %')
    patrones_monto = {
        'RESPONSABILIDAD CIVIL': r'Responsabilidad Civil\s*\(Límite[^)]*\)\s*([\d,\.]+)',
        'RC CATASTRÓFICA / EXCESO': r'Responsabilidad Civil en Exceso[^\d]*([\d,\.]+)',
        'GASTOS MÉDICOS OCUPANTES': r'Gastos Médicos Ocupantes[^\d]*([\d,\.]+)',
        'RESPONSABILIDAD CIVIL POR DAÑOS A OCUPANTES': r'Responsabilidad Civil por daños a los Ocupantes\s*([\d,\.]+)',
        'RESPONSABILIDAD CIVIL FAMILIAR': r'Responsabilidad Civil Familiar\s*([\d,\.]+)',
        'ACCIDENTES PERSONALES': r'Accidentes Automovilísticos al Conductor\s*([\d,\.]+)',
    }
    for k, pat in patrones_monto.items():
        m = re.search(pat, texto, re.IGNORECASE)
        if m:
            cob[k] = (fmt_monto(m.group(1)), 'NO APLICA')
    if re.search(r'Asistencia\s+en\s+viajes\s+Amparada', texto, re.IGNORECASE):
        cob['ASISTENCIA VIAL'] = ('Amparada', 'NO APLICA')
    if re.search(r'Asistencia\s+Jurídica\s+Amparada', texto, re.IGNORECASE):
        cob['ASISTENCIA LEGAL'] = ('Amparada', 'NO APLICA')
    if re.search(r'Asistencia\s+Médica\s+Amparada', texto, re.IGNORECASE):
        cob['ASISTENCIA MÉDICA'] = ('Amparada', 'NO APLICA')
    m = re.search(r'(?:Rotura de\s+)?Cristales\s+Amparada(?:\s+(\d+)\s*%)?', texto, re.IGNORECASE)
    if m:
        ded = f'{m.group(1)}%' if m.group(1) else 'NO APLICA'
        cob['CRISTALES'] = ('Amparada', ded)
    m = re.search(r'Una exhibición de\s+([\d,\.]+)', texto, re.IGNORECASE)
    if m:
        d['prima_total'] = fmt_monto(m.group(1))
    else:
        m = re.search(r'Total a Pagar.*?\n[\d,\.\s]+?([\d,]+\.\d{2})\s*$', texto, re.MULTILINE)
        if m:
            d['prima_total'] = fmt_monto(m.group(1))
    return d


def parse_chubb(texto):
    d = {'aseguradora': 'CHUBB', 'coberturas': {}}
    d['cliente'] = buscar(r'Nombre:\s*([^\n]+?)(?:\n|Agente)', texto)
    d['vigencia_ini'] = buscar(r'Vigencia:Del\s+(\d{2}/[A-Z]{3}/\d{4})', texto)
    d['vigencia_fin'] = buscar(r'al\s+(\d{2}/[A-Z]{3}/\d{4})', texto)
    d['vehiculo'] = buscar(r'Descripción del vehículo\*:\s*([^\n]+)', texto)
    d['cp'] = ''
    d['serie'] = buscar(r'serie:?\s*([A-Z0-9]*\d[A-Z0-9]{3,24})', texto, default='')
    d['paquete'] = buscar(r'Paquete:\s*(\w+)', texto)
    cob = d['coberturas']
    m = re.search(r'DAÑOS MATERIALES\s+(?:VALOR COMERCIAL|E\)\s*[\d,\.]+)\s+A\)\s*([\d\.]+)\s*%', texto)
    if m:
        m2 = re.search(r'DAÑOS MATERIALES\s+E\)\s*([\d,\.]+)\s+A\)', texto)
        suma = fmt_monto(m2.group(1)) if m2 else 'VALOR COMERCIAL'
        cob['DAÑOS MATERIALES'] = (suma, f'{m.group(1)} %')
    m = re.search(r'ROBO TOTAL\s+(?:VALOR COMERCIAL|E\)\s*[\d,\.]+)\s+A\)\s*([\d\.]+)\s*%', texto)
    if m:
        m2 = re.search(r'ROBO TOTAL\s+E\)\s*([\d,\.]+)\s+A\)', texto)
        suma = fmt_monto(m2.group(1)) if m2 else 'VALOR COMERCIAL'
        cob['ROBO TOTAL'] = (suma, f'{m.group(1)} %')
    patrones = {
        'RESPONSABILIDAD CIVIL': r'RESPONSABILIDAD CIVIL POR DAÑOS A TERCEROS\s+([\d,\.]+)',
        'RC CATASTRÓFICA / EXCESO': r'RESPONSABILIDAD CIVIL POR FALLECIMIENTO\s+([\d,\.]+)',
        'GASTOS MÉDICOS OCUPANTES': r'GASTOS MÉDICOS OCUPANTES\s+([\d,\.]+)',
        'ACCIDENTES PERSONALES': r'ACCIDENTES AUTOMOVILÍSTICOS AL CONDUCTOR\s+([\d,\.]+)',
        'RESPONSABILIDAD CIVIL POR DAÑOS A OCUPANTES': r'RESPONSABILIDAD CIVIL POR DAÑOS A OCUPANTES\s+([\d,\.]+)',
        'RESPONSABILIDAD CIVIL FAMILIAR': r'RESPONSABILIDAD CIVIL FAMILIAR\s+([\d,\.]+)',
        'ASISTENCIA LEGAL': r'ASISTENCIA LEGAL CDS\s*\*?\s+([\d,\.]+)',
    }
    for k, pat in patrones.items():
        m = re.search(pat, texto, re.IGNORECASE)
        if m:
            cob[k] = (fmt_monto(m.group(1)), 'NO APLICA')
    # RC Carga — CHUBB puede decir "AMPARADA" o un monto real (no confundir con la prima)
    m = re.search(r'RESPONSABILIDAD CIVIL CARGA\s*\w?\s+AMPARADA', texto, re.IGNORECASE)
    if m:
        cob['RESPONSABILIDAD CIVIL POR DAÑOS OCASIONADOS POR LA CARGA'] = ('AMPARADA', 'NO APLICA')
    else:
        m = re.search(r'RESPONSABILIDAD CIVIL CARGA\s*\w?\s+([\d,]+\.00)\s+NO APLICA', texto, re.IGNORECASE)
        if m:
            cob['RESPONSABILIDAD CIVIL POR DAÑOS OCASIONADOS POR LA CARGA'] = (fmt_monto(m.group(1)), 'NO APLICA')
    if re.search(r'RESPONSABILIDAD CIVIL EN USA Y CANADÁ\s+AMPARADA', texto):
        cob['RESPONSABILIDAD CIVIL EXTRANJERO / USA Y CANADÁ'] = ('AMPARADA', 'NO APLICA')
    if re.search(r'RESPONSABILIDAD CIVIL ADAPTACIONES Y CONVERSIONES\s+AMPARADA', texto):
        cob['RESPONSABILIDAD CIVIL ADAPTACIONES Y CONVERSIONES'] = ('AMPARADA', 'NO APLICA')
    if re.search(r'ASISTENCIA EN VIAJE CDS\s*\*?\s+AMPARADA', texto):
        cob['ASISTENCIA VIAL'] = ('AMPARADA', 'NO APLICA')
    m = re.search(r'CRISTALES\s+AMPARADA(?:\s+A?\)?\s*([\d\.]+)\s*%)?', texto, re.IGNORECASE)
    if m:
        ded = f'{int(float(m.group(1)))}%' if m.group(1) else 'NO APLICA'
        cob['CRISTALES'] = ('AMPARADA', ded)
    m = re.search(r'Prima total\s+([\d,\.]+)', texto)
    if m:
        d['prima_total'] = fmt_monto(m.group(1))
    return d


def parse_qualitas(texto):
    d = {'aseguradora': 'QUÁLITAS', 'coberturas': {}}
    d['cliente'] = buscar(r'INFORMACION DEL ASEGURADO\s*\n\s*([^\n]+)', texto)
    # Fix "2 al 2": Quálitas escribe fechas como "DEL 2 DE JULIO DE 2026", no como "01/07/2026"
    d['vigencia_ini'] = buscar(r'DESDE LAS.*?DEL\s+(\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+DE\s+\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4})', texto)
    d['vigencia_fin'] = buscar(r'HASTA LAS.*?DEL\s+(\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+DE\s+\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4})', texto)
    d['vehiculo'] = buscar(r'\d{4,5}\s+([A-Z][^\n]+?(?:OCUP\.|PICK[- ]UP[^\n]+))', texto)
    if not d['vehiculo']:
        d['vehiculo'] = buscar(r'((?:EQ\s+)?KENWORTH[^\n]+)', texto, default='')
    d['cp'] = buscar(r'C\.P\.\s*:?\s*(\d{5})', texto)
    d['serie'] = buscar(r'serie:?\s*([A-Z0-9]*\d[A-Z0-9]{3,24})', texto, default='')
    d['paquete'] = buscar(r'PLAN\s+USO.*?\n.*?(AMPLIA|LIMITADA|RC)', texto, flags=re.IGNORECASE | re.DOTALL)
    cob = d['coberturas']
    m = re.search(r'Daños materiales\s+Valor Comercial\s+(\d+)%', texto, re.IGNORECASE)
    if m: cob['DAÑOS MATERIALES'] = ('Valor Comercial', f'{m.group(1)}%')
    else:
        m = re.search(r'DANO?S\s+MATERIALES\s+\$\s*([\d,]+)\s+([\d\.]+)', texto)
        if m: cob['DAÑOS MATERIALES'] = (fmt_monto(m.group(1)), f'{int(float(m.group(2)))}%')
        else:
            m = re.search(r'Daños materiales\s+\$\s*([\d,\.]+)\s+(\d+)%', texto, re.IGNORECASE)
            if m: cob['DAÑOS MATERIALES'] = (fmt_monto(m.group(1)), f'{m.group(2)}%')
    m = re.search(r'Robo total\s+Valor Comercial\s+(\d+)%', texto, re.IGNORECASE)
    if m: cob['ROBO TOTAL'] = ('Valor Comercial', f'{m.group(1)}%')
    else:
        m = re.search(r'ROBO\s+TOTAL\s+\$\s*([\d,]+)\s+([\d\.]+)', texto)
        if m: cob['ROBO TOTAL'] = (fmt_monto(m.group(1)), f'{int(float(m.group(2)))}%')
        else:
            m = re.search(r'Robo total\s+\$\s*([\d,\.]+)\s+(\d+)%', texto, re.IGNORECASE)
            if m: cob['ROBO TOTAL'] = (fmt_monto(m.group(1)), f'{m.group(2)}%')
    m = re.search(r'RESPONSABILIDAD CIVIL POR DAÑOS A TERCEROS\s+\$\s*([\d,\.]+)(?:\s*POR EVENTO)?\s*(\d+\s*umas?)?',
                  texto, re.IGNORECASE)
    if m:
        ded = m.group(2) if m.group(2) else 'NO APLICA'
        cob['RESPONSABILIDAD CIVIL'] = (fmt_monto(m.group(1)), ded)
    else:
        m = re.search(r'Responsabilidad Civil\s+\$\s*([\d,\.]+)', texto, re.IGNORECASE)
        if m: cob['RESPONSABILIDAD CIVIL'] = (fmt_monto(m.group(1)), 'NO APLICA')
    patrones = {
        'RC CATASTRÓFICA / EXCESO': r'RC Complementaria Personas\s+\$\s*([\d,\.]+)',
        'GASTOS MÉDICOS OCUPANTES': r'(?:Gastos Med?icos Ocupantes|GASTOS MED?ICOS OCUPANTES)\s+\$\s*([\d,\.]+)',
        'RESPONSABILIDAD CIVIL POR DAÑOS A OCUPANTES': r'(?:RC\.?\s*Da[nñ]os.*?Ocupantes|RC\s+DANOS\s+POR\s+DANOS\s+A\s+OCUPANTES)\s+\$\s*([\d,\.]+)',
        'ACCIDENTES PERSONALES': r'(?:Muerte del Conductor|MUERTE DEL CONDUCTOR)\s*(?:X AA)?\s+\$\s*([\d,\.]+)',
        'DAÑOS ECOLÓGICOS POR CARGA': r'RESPONSABILIDAD CIVIL ECOLOGICA\s+\$\s*([\d,\.]+)',
    }
    for k, pat in patrones.items():
        m = re.search(pat, texto, re.IGNORECASE)
        if m:
            cob[k] = (fmt_monto(m.group(1)), 'NO APLICA')
    m = re.search(r'RC POR DA[NÑ]OS OCASIONADOS POR LA CARGA\s+AMPARADO\s*(\d+\s*umas?)?', texto, re.IGNORECASE)
    if m:
        ded = m.group(1) if m.group(1) else 'NO APLICA'
        cob['RESPONSABILIDAD CIVIL POR DAÑOS OCASIONADOS POR LA CARGA'] = ('AMPARADA', ded)
    elif re.search(r'Daños por carga\s+AMPARADO', texto, re.IGNORECASE):
        cob['RESPONSABILIDAD CIVIL POR DAÑOS OCASIONADOS POR LA CARGA'] = ('AMPARADA', 'NO APLICA')
    m = re.search(r'ADAPTACIONES Y/O CONVERSIONES DM\s+\$\s*([\d,\.]+)\s+([\d\.]+)', texto, re.IGNORECASE)
    if m:
        cob['RESPONSABILIDAD CIVIL ADAPTACIONES Y CONVERSIONES'] = (
            fmt_monto(m.group(1)), f'{int(float(m.group(2)))}%')
    if re.search(r'Maniobra de Carga y Descarga\s+AMPARADO', texto, re.IGNORECASE):
        cob['MANIOBRAS DE CARGA Y DESCARGA'] = ('AMPARADA', 'NO APLICA')
    if re.search(r'(?:Asistencia Vial|ASISTENCIA VIAL)\s+(?:QUALITAS\s+)?(?:PLUS\s+)?AMPARADO', texto, re.IGNORECASE):
        cob['ASISTENCIA VIAL'] = ('AMPARADA', 'NO APLICA')
    if re.search(r'(?:Gastos Legales|GASTOS LEGALES)\s+AMPARADO', texto, re.IGNORECASE):
        cob['ASISTENCIA LEGAL'] = ('AMPARADA', 'NO APLICA')
    m = re.search(r'CRISTALES\s+AMPARADO(?:\s+([\d\.]+))?', texto, re.IGNORECASE)
    if m:
        ded = f'{int(float(m.group(1)))}%' if m.group(1) else 'NO APLICA'
        cob['CRISTALES'] = ('AMPARADA', ded)
    m = re.search(r'IMPORTE TOTAL\s+([\d,\.]+)', texto)
    if m:
        d['prima_total'] = fmt_monto(m.group(1))
    else:
        m = re.search(r'CONTADO\s+([\d,\.]+)', texto)
        if m:
            d['prima_total'] = fmt_monto(m.group(1))
    return d


def _extraer_prima_gnp_tabla(pdf_path):
    """Método PRIMARIO: lee la tabla estructurada 'DATOS DE LA PRIMA' con
    pdfplumber y devuelve el valor de la columna 'Prima total anual'.
    Es el más confiable porque no depende del orden en que se extraiga el texto.
    """
    try:
        import pdfplumber as _plumber
        with _plumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    if not table:
                        continue
                    for i, row in enumerate(table):
                        if not row:
                            continue
                        for j, cell in enumerate(row):
                            cell_str = str(cell or '')
                            if 'PRIMA TOTAL ANUAL' not in cell_str.upper():
                                continue
                            # Caso 1: el valor está en la MISMA celda, debajo del texto del header
                            resto = re.split(r'PRIMA\s+TOTAL\s+ANUAL', cell_str,
                                             flags=re.IGNORECASE, maxsplit=1)
                            if len(resto) > 1:
                                m = re.search(r'\$?\s*([\d,]+(?:\.\d{2})?)', resto[1])
                                if m:
                                    monto = m.group(1)
                                    try:
                                        if float(monto.replace(',', '')) >= 1000:
                                            return monto
                                    except ValueError:
                                        pass
                            # Caso 2: el valor está en filas siguientes, misma columna
                            for row2 in table[i+1:]:
                                if not row2 or j >= len(row2):
                                    continue
                                cell2 = str(row2[j] or '').strip()
                                m = re.search(r'\$?\s*([\d,]+(?:\.\d{2})?)', cell2)
                                if m:
                                    monto = m.group(1)
                                    try:
                                        if float(monto.replace(',', '')) >= 1000:
                                            return monto
                                    except ValueError:
                                        continue
    except Exception:
        pass
    return None


def extraer_prima_gnp(texto, pdf_path=None):
    """Extrae la PRIMA TOTAL ANUAL de una cotización GNP.

    Estrategia (en orden de confiabilidad):
      1. Leer la tabla estructurada 'DATOS DE LA PRIMA' con pdfplumber.
      2. Buscar en la sección 'DATOS DE LA PRIMA' del texto y tomar el
         monto MÁS GRANDE (la prima total siempre es mayor que descuentos
         o pagos parciales — así ni $495, ni $940, ni $3,397, ni $2,609
         pueden confundirla).
      3. Fallback: pago único de contado.
    """
    # 1) Tabla estructurada
    if pdf_path:
        monto = _extraer_prima_gnp_tabla(pdf_path)
        if monto:
            return monto

    # 2) Región de la prima — MAX de los montos.
    # Anclas en orden de preferencia. Nota: pdfplumber a veces pega las
    # palabras ("Primatotal anual"), por eso los \s* en vez de \s+.
    seccion = None
    for ancla in (r'DATOS\s*DE\s*LA\s*PRIMA',
                  r'Prima\s*total\s*anual',
                  r'Forma\s*de\s*pago'):
        seccion = re.search(ancla + r'([\s\S]{0,300})', texto, re.IGNORECASE)
        if seccion:
            break
    if seccion:
        montos = re.findall(r'\$\s*([\d,]+(?:\.\d{2})?)', seccion.group(1))
        candidatos = []
        for s in montos:
            try:
                v = float(s.replace(',', ''))
                # >= 1000 descarta descuentos ($495, $940).
                # < 1000000 descarta sumas aseguradas ($3,000,000) por si la
                # ventana alcanza a tocar la tabla de coberturas.
                if 1000 <= v < 1000000:
                    candidatos.append((v, s))
            except ValueError:
                continue
        if candidatos:
            return max(candidatos)[1]

    # 3) Contado explícito
    m = re.search(r'\b1\s*pago\s*de\s*\$?\s*([\d,]+(?:\.\d{2})?)', texto, re.IGNORECASE)
    if m:
        return m.group(1)
    return None


def parse_gnp(texto, pdf_path=None):
    d = {'aseguradora': 'GNP', 'coberturas': {}}
    d['cliente'] = buscar(r'(?:Nombre|Contratante)[:\s]+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]{5,50})', texto, default='')
    d['cp'] = buscar(r'(?:Código Postal|C\.?P\.?)\s*(\d{5})', texto, default='')
    d['vehiculo'] = buscar(r'(KENWORTH[^\n]*)', texto, default='')
    if not d['vehiculo']:
        # Respaldo genérico: línea después de "Datos del vehículo"
        d['vehiculo'] = buscar(r'Datos del veh[ií]culo\s*\n\s*([^\n]+)', texto, default='')
    d['vigencia_ini'] = buscar(r'Fecha de cotización:?\s*([^\n]+?)(?:\s+Lugar|\n)', texto, default='')
    d['vigencia_fin'] = ''
    d['serie'] = ''
    d['descripcion_vehiculo'] = buscar(r'Folio de cotización:?\s*([A-Z0-9]+)', texto, default='')
    d['paquete'] = buscar(r'DETALLE DE COBERTURAS\s*\n\s*(Responsabilidad\s+Civil|Amplia(?:\s+Plus)?|Limitada|Reducida)',
                          texto, default='')
    if not d['paquete']:
        d['paquete'] = buscar(r'Paquete[^\n]*\n\s*(Responsabilidad|Amplia(?:\s+Plus)?|Limitada|Reducida)',
                              texto, default='')
        if d['paquete'] == 'Responsabilidad':
            d['paquete'] = 'Responsabilidad Civil'
    if d['paquete']:
        d['paquete'] = d['paquete'].upper()
    cob = d['coberturas']
    # --- PRIMA (corregido v3: lee la tabla estructurada; ya no confunde con
    # descuentos ni con pagos parciales sin importar cómo salga el texto) ---
    prima = extraer_prima_gnp(texto, pdf_path=pdf_path)
    if prima:
        d['prima_contado'] = prima.replace(',', '')
        d['prima_total'] = fmt_monto(prima)
    d['prima_semestral'] = None
    d['prima_trimestral'] = None
    if pdf_path:
        try:
            import pdfplumber as plumber
            with plumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    for table in page.extract_tables():
                        for row in table:
                            if not row or not row[0]: continue
                            nombre = str(row[0]).replace('\n', ' ').strip().upper()
                            suma = str(row[1]).strip() if len(row) > 1 and row[1] else ''
                            ded = str(row[2]).strip() if len(row) > 2 and row[2] else ''
                            if 'DAÑOS MATERIALES' in nombre and 'TOTAL' in nombre:
                                cob['DAÑOS MATERIALES'] = (suma, ded)
                            elif 'ROBO TOTAL' in nombre:
                                cob['ROBO TOTAL'] = (suma, ded)
                            elif 'RESPONSABILIDAD CIVIL' in nombre and 'TERCEROS' in nombre:
                                cob['RESPONSABILIDAD CIVIL'] = (suma, ded)
                            elif 'GASTOS MÉDICOS' in nombre or 'GASTOS MEDICOS' in nombre:
                                cob['GASTOS MÉDICOS OCUPANTES'] = (suma, 'NO APLICA')
                            elif 'ACCIDENTES' in nombre and 'CONDUCTOR' in nombre:
                                cob['ACCIDENTES PERSONALES'] = (suma, 'NO APLICA')
                            elif 'FALLECIMIENTO' in nombre:
                                cob['RC CATASTRÓFICA / EXCESO'] = (suma, 'NO APLICA')
                            elif 'RESPONSABILIDAD CIVIL' in nombre and 'OCUPANTES' in nombre:
                                cob['RESPONSABILIDAD CIVIL POR DAÑOS A OCUPANTES'] = (suma, 'NO APLICA')
                            elif 'PROTECCIÓN LEGAL' in nombre or 'PROTECCION LEGAL' in nombre:
                                cob['ASISTENCIA LEGAL'] = (suma, 'NO APLICA')
                            elif 'ASISTENCIA VIAL' in nombre:
                                cob['ASISTENCIA VIAL'] = (suma, 'NO APLICA')
                            elif 'CRISTALES' in nombre:
                                cob['CRISTALES'] = (suma, ded)
                            elif 'DAÑOS POR LA CARGA' in nombre or 'DANOS POR LA CARGA' in nombre:
                                cob['RESPONSABILIDAD CIVIL POR DAÑOS OCASIONADOS POR LA CARGA'] = (suma, 'NO APLICA')
                        # Respaldo de prima: fila de la tabla "DATOS DE LA PRIMA"
                        # (última columna = "Prima total anual")
                        if not d.get('prima_total'):
                            for row in table:
                                if not row: continue
                                celdas = [str(c or '') for c in row]
                                fila_txt = ' '.join(celdas).upper()
                                m2 = re.search(r'\$\s*([\d,]+)\s*$', celdas[-1].strip())
                                if m2 and ('PAGO' in fila_txt or 'AMPLIA' in fila_txt
                                           or 'LIMITADA' in fila_txt or 'REDUCIDA' in fila_txt
                                           or 'RESPONSABILIDAD' in fila_txt):
                                    d['prima_contado'] = m2.group(1).replace(',', '')
                                    d['prima_total'] = fmt_monto(m2.group(1))
                                    break
        except Exception as e:
            print(f"  ⚠️ Error tablas GNP: {e}")
    return d


def parse_inbursa(texto):
    d = {'aseguradora': 'INBURSA', 'coberturas': {}}
    d['cliente'] = buscar(r'Contratante:\s*([^\n]+?)(?:\s+Correo|$)', texto, default='')
    d['cp'] = ''
    d['vehiculo'] = buscar(r'Descripción:\s*([^\n]+?)(?:\s+Equipo|$)', texto, default='')
    d['vigencia_ini'] = buscar(r'Fecha:\s*([\d/]+)', texto, default='')
    d['vigencia_fin'] = ''
    d['serie'] = buscar(r'serie:?\s*([A-Z0-9]*\d[A-Z0-9]{3,24})', texto, default='')
    d['descripcion_vehiculo'] = buscar(r'Clave vehicular:?\s*(\d+)', texto, default='')
    d['paquete'] = buscar(r'COBERTURA\s+(AMPLIA\s+PLUS|AMPLIA|LIMITADA|REDUCIDA|RESPONSABILIDAD CIVIL)',
                          texto, default='')
    cob = d['coberturas']
    m = re.search(r'DAÑOS MATERIALES\s+VALOR COMERCIAL\s+(\d+)%', texto, re.IGNORECASE)
    if m: cob['DAÑOS MATERIALES'] = ('VALOR COMERCIAL', f'{m.group(1)}%')
    m = re.search(r'ROBO TOTAL\s+VALOR COMERCIAL\s+(\d+)%', texto, re.IGNORECASE)
    if m: cob['ROBO TOTAL'] = ('VALOR COMERCIAL', f'{m.group(1)}%')
    m = re.search(r'RESPONSABILIDAD CIVIL\s+([\d,]+)\s+(\d+\s*UMA)', texto, re.IGNORECASE)
    if m: cob['RESPONSABILIDAD CIVIL'] = (fmt_monto(m.group(1)), m.group(2).upper())
    m = re.search(r'RESPONSABILIDAD CIVIL FALLECIMIENTO\s+([\d,]+)', texto, re.IGNORECASE)
    if m: cob['RC CATASTRÓFICA / EXCESO'] = (fmt_monto(m.group(1)), 'NO APLICA')
    m = re.search(r'GASTOS MÉDICOS A OCUPANTES\s+([\d,]+)', texto, re.IGNORECASE)
    if m: cob['GASTOS MÉDICOS OCUPANTES'] = (fmt_monto(m.group(1)), 'NO APLICA')
    m = re.search(r'ACCIDENTES AL CONDUCTOR\s+([\d,]+)', texto, re.IGNORECASE)
    if m: cob['ACCIDENTES PERSONALES'] = (fmt_monto(m.group(1)), 'NO APLICA')
    m = re.search(r'DAÑOS CARGA LUC\s+AMPARADO\s*(TIPO CARGA:\s*\w+)?', texto, re.IGNORECASE)
    if m:
        ded = m.group(1).upper() if m.group(1) else 'NO APLICA'
        cob['RESPONSABILIDAD CIVIL POR DAÑOS OCASIONADOS POR LA CARGA'] = ('AMPARADA', ded)
    m = re.search(r'DAÑOS ECOLOGICOS LUC\s+AMPARADO\s*(TIPO CARGA:\s*\w+)?', texto, re.IGNORECASE)
    if m:
        ded = m.group(1).upper() if m.group(1) else 'NO APLICA'
        cob['DAÑOS ECOLÓGICOS POR CARGA'] = ('AMPARADA', ded)
    if re.search(r'ASISTENCIA EN VIAJES\s*\n?.*?AMPARADO', texto, re.IGNORECASE) or \
       re.search(r'ASISTENCIA EN VIAJES', texto, re.IGNORECASE):
        cob['ASISTENCIA VIAL'] = ('AMPARADA', 'NO APLICA')
    m = re.search(r'ASISTENCIA JURÍDICA\s+([\d,]+)', texto, re.IGNORECASE)
    if m: cob['ASISTENCIA LEGAL'] = (fmt_monto(m.group(1)), 'NO APLICA')
    m = re.search(r'PAGO DE FIANZAS\s+([\d,]+)', texto, re.IGNORECASE)
    if m: cob['PAGO DE FIANZAS'] = (fmt_monto(m.group(1)), 'NO APLICA')
    m = re.search(r'CONTADO\s*\n?\s*([\d,]+)\s+([\d,]+)', texto)
    if m:
        d['prima_total'] = fmt_monto(m.group(1))
    else:
        m = re.search(r'Prima total.*?\n.*?([\d,]+)', texto)
        if m:
            d['prima_total'] = fmt_monto(m.group(1))
    return d


def parse_ana(texto):
    """Extrae datos de cotización ANA Seguros."""
    d = {'aseguradora': 'ANA', 'coberturas': {}}
    d['cliente'] = ''  # ANA no muestra nombre del asegurado en este formato
    d['cp'] = ''  # ANA no muestra C.P. del asegurado, solo el estado/tarifa
    d['vehiculo'] = buscar(r'VEHÍCULO:\s*([^\n]+)', texto, default='')
    d['vigencia_ini'] = buscar(r'Vigencia:\s*([\d/]+)\s*AL', texto, default='')
    d['vigencia_fin'] = buscar(r'Vigencia:.*?AL\s+([\d/]+)', texto, default='')
    d['serie'] = buscar(r'serie:?\s*([A-Z0-9]*\d[A-Z0-9]{3,24})', texto, default='')
    d['descripcion_vehiculo'] = buscar(r'CLAVE TARIFA\s*:\s*(\d+)', texto, default='')
    cob = d['coberturas']

    # Daños Materiales / Robo Total: "V. COMERCIAL + 10%   5%   $ 6,675.83"
    m = re.search(r'DAÑOS MATERIALES\s+(V\.\s*COMERCIAL(?:\s*\+\s*\d+%)?)\s+(\d+)%', texto, re.IGNORECASE)
    if m: cob['DAÑOS MATERIALES'] = ('VALOR COMERCIAL', f'{m.group(2)}%')
    m = re.search(r'ROBO TOTAL\s+(V\.\s*COMERCIAL(?:\s*\+\s*\d+%)?)\s+(\d+)%', texto, re.IGNORECASE)
    if m: cob['ROBO TOTAL'] = ('VALOR COMERCIAL', f'{m.group(2)}%')

    # Responsabilidad Civil: "3,000,000.00 LUC"
    m = re.search(r'RESPONSABILIDAD CIVIL\s+([\d,\.]+)\s*LUC', texto, re.IGNORECASE)
    if m: cob['RESPONSABILIDAD CIVIL'] = (fmt_monto(m.group(1)), 'NO APLICA')

    # RC Catastrófica por muerte -> RC CATASTRÓFICA / EXCESO
    m = re.search(r'RC CATASTROFICA POR MUERTE\s+([\d,\.]+)', texto, re.IGNORECASE)
    if m: cob['RC CATASTRÓFICA / EXCESO'] = (fmt_monto(m.group(1)), 'NO APLICA')

    # Gastos médicos ocupantes
    m = re.search(r'GASTOS MEDICOS OCUPANTES\s+([\d,\.]+)', texto, re.IGNORECASE)
    if m: cob['GASTOS MÉDICOS OCUPANTES'] = (fmt_monto(m.group(1)), 'NO APLICA')

    # Gastos por muerte accidental -> ACCIDENTES PERSONALES
    m = re.search(r'GASTOS POR MUERTE ACCIDENTAL\s+([\d,\.]+)', texto, re.IGNORECASE)
    if m: cob['ACCIDENTES PERSONALES'] = (fmt_monto(m.group(1)), 'NO APLICA')

    # RC Ocupantes -> RESPONSABILIDAD CIVIL POR DAÑOS A OCUPANTES
    m = re.search(r'RESPONSABILIDAD CIVIL OCUPANTES\s+([\d,\.]+)', texto, re.IGNORECASE)
    if m: cob['RESPONSABILIDAD CIVIL POR DAÑOS A OCUPANTES'] = (fmt_monto(m.group(1)), 'NO APLICA')

    # RC daños por la carga
    if re.search(r'RESPONSABILIDAD CIVIL DAÑOS POR LA CARGA\s+AMPARADO', texto, re.IGNORECASE):
        cob['RESPONSABILIDAD CIVIL POR DAÑOS OCASIONADOS POR LA CARGA'] = ('AMPARADA', 'NO APLICA')

    # Defensa jurídica y asist. legal -> ASISTENCIA LEGAL
    if re.search(r'DEFENSA JURIDICA Y ASIST\.?\s*LEGAL\s+AMPARADO', texto, re.IGNORECASE):
        cob['ASISTENCIA LEGAL'] = ('AMPARADA', 'NO APLICA')

    # ANA Asistencia vial viajes -> ASISTENCIA VIAL
    if re.search(r'ANA ASISTENCIA VIAL VIAJES', texto, re.IGNORECASE):
        cob['ASISTENCIA VIAL'] = ('AMPARADA', 'NO APLICA')

    # Cobertura exclusiva de ANA sin equivalente estándar
    if re.search(r'DESBIELAMIENTO POR AGUA\s+AMPARADO', texto, re.IGNORECASE):
        cob['DESBIELAMIENTO POR AGUA'] = ('AMPARADA', 'NO APLICA')

    # Prima total
    m = re.search(r'PRIMA TOTAL:\s*([\d,\.]+)', texto, re.IGNORECASE)
    if m: d['prima_total'] = fmt_monto(m.group(1))

    return d


PARSERS = {
    'HDI': parse_hdi,
    'CHUBB': parse_chubb,
    'QUALITAS': parse_qualitas,
    'GNP': parse_gnp,
    'INBURSA': parse_inbursa,
    'ANA': parse_ana,
}


COBERTURAS_ORDEN = [
    'DAÑOS MATERIALES',
    'ROBO TOTAL',
    'RESPONSABILIDAD CIVIL',
    'RC CATASTRÓFICA / EXCESO',
    'GASTOS MÉDICOS OCUPANTES',
    'ASISTENCIA VIAL',
    'ASISTENCIA LEGAL',
    'ACCIDENTES PERSONALES',
    'RESPONSABILIDAD CIVIL POR DAÑOS A OCUPANTES',
    'RESPONSABILIDAD CIVIL POR DAÑOS OCASIONADOS POR LA CARGA',
    'MANIOBRAS DE CARGA Y DESCARGA',
    'RESPONSABILIDAD CIVIL FAMILIAR',
    'RESPONSABILIDAD CIVIL EXTRANJERO / USA Y CANADÁ',
    'CRISTALES',
    'RESPONSABILIDAD CIVIL ADAPTACIONES Y CONVERSIONES',
    'ASISTENCIA MÉDICA',
    'DAÑOS ECOLÓGICOS POR CARGA',
    'PAGO DE FIANZAS',
    'DESBIELAMIENTO POR AGUA',
]


COLOR_HEADER = colors.HexColor('#1F2A44')
COLOR_HEADER_BG = colors.HexColor('#F4F5F8')
COLOR_AMPARADA = colors.HexColor('#1A8A52')
COLOR_NO_AMPARADA = colors.HexColor('#9AA0A6')
COLOR_PRIMA = colors.HexColor('#1F2A44')
COLOR_BORDE = colors.HexColor('#E1E4EA')

LOGOS_MAP = {
    'HDI': 'hdi',
    'CHUBB': 'chubb',
    'QUÁLITAS': 'qualitas',
    'QUALITAS': 'qualitas',
    'ZURICH': 'zurich',
    'ANA': 'ana',
    'GNP': 'gnp',
    'ALLIANZ': 'allianz',
    'OLIVAS': 'olivas',
}


def extraer_logo(pdf_path, destino_dir, nombre):
    """Saca el logo de la aseguradora directo de su PDF: toma la imagen más
    grande que esté en el 25% superior de la primera página (ahí van los
    logos en GNP, Quálitas y CHUBB). Devuelve la ruta del PNG o None."""
    try:
        import pypdfium2 as pdfium
        import pypdfium2.raw as pdfium_c
        # Posición confiable con pdfplumber: tamaños (px) de las imágenes
        # que están en el 25% superior de la página 1
        with pdfplumber.open(pdf_path) as pp:
            pg = pp.pages[0]
            arriba = {tuple(int(x) for x in im.get('srcsize', (0, 0)))
                      for im in pg.images if im['top'] < pg.height * 0.25}
        pdf = pdfium.PdfDocument(pdf_path)
        page = pdf[0]
        mejor, mejor_area = None, 0
        for obj in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_IMAGE]):
            if tuple(int(x) for x in obj.get_px_size()) not in arriba:
                continue
            try:
                img = obj.get_bitmap(render=True).to_pil().convert('RGBA')
                # Respetar volteos de la matriz (CHUBB viene invertido)
                ma, _, _, md, _, _ = obj.get_matrix().get()
                from PIL import ImageOps
                if md < 0:
                    img = ImageOps.flip(img)
                if ma < 0:
                    img = ImageOps.mirror(img)
            except Exception:
                continue
            if img.width < 40 or img.height < 20:
                continue
            area = img.width * img.height
            if area > mejor_area:
                mejor, mejor_area = img, area
        pdf.close()
        if mejor is None:
            return None
        # Recortar márgenes blancos/transparentes
        from PIL import Image, ImageChops
        fondo = Image.new('RGBA', mejor.size, (255, 255, 255, 255))
        plano = Image.alpha_composite(fondo, mejor).convert('RGB')
        bbox = ImageChops.difference(plano, Image.new('RGB', plano.size, 'white')).getbbox()
        if bbox:
            plano = plano.crop(bbox)
        # Fondo blanco -> transparente (para que no se vea un recuadro)
        plano = plano.convert('RGBA')
        plano.putdata([(r, g, b, 0) if r > 242 and g > 242 and b > 242 else (r, g, b, 255)
                       for (r, g, b, _) in plano.getdata()])
        os.makedirs(destino_dir, exist_ok=True)
        ruta = os.path.join(destino_dir, f'logo_{nombre.lower()}.png')
        plano.save(ruta)
        return ruta
    except Exception as e:
        print(f"  (sin logo para {nombre}: {e})")
        return None


def logo_celda(ruta, ancho_max, alto_max, respaldo):
    """Imagen escalada para caber en (ancho_max x alto_max); si no hay ruta,
    regresa el flowable de respaldo (el nombre en texto)."""
    if not ruta or not os.path.isfile(ruta):
        return respaldo
    try:
        from reportlab.lib.utils import ImageReader
        w, h = ImageReader(ruta).getSize()
        escala = min(ancho_max / w, alto_max / h)
        return RLImage(ruta, width=w * escala, height=h * escala, hAlign='CENTER', mask='auto')
    except Exception:
        return respaldo


def buscar_logo(nombre_aseguradora, logos_dir):
    if not logos_dir:
        return None
    clave = LOGOS_MAP.get(nombre_aseguradora.upper(), nombre_aseguradora.lower())
    for ext in ('.jpeg', '.jpg', '.png'):
        ruta = os.path.join(logos_dir, clave + ext)
        if os.path.isfile(ruta):
            return ruta
    return None


def construir_pdf(datos_por_aseguradora, salida='comparativo.pdf', meta=None, logos_dir=None, watermark=None):
    aseguradoras = list(datos_por_aseguradora.keys())
    meta = meta or {}
    doc = SimpleDocTemplate(
        salida,
        pagesize=landscape(letter),
        leftMargin=1.0 * cm, rightMargin=1.0 * cm,
        topMargin=0.8 * cm, bottomMargin=0.8 * cm,
    )
    styles = getSampleStyleSheet()
    story = []
    estilo_valor = ParagraphStyle('valor', parent=styles['Normal'],
                                  fontName='Helvetica', fontSize=8,
                                  textColor=colors.HexColor('#333333'))
    estilo_marca = ParagraphStyle('marca', parent=styles['Normal'],
                                  fontName='Helvetica-Bold', fontSize=14,
                                  textColor=COLOR_HEADER, alignment=TA_RIGHT)
    cliente = meta.get('cliente', '')
    cp = meta.get('cp', '')
    vehiculo = meta.get('vehiculo', '')
    serie = meta.get('serie', '')
    vigencia = meta.get('vigencia', '')
    prima_anterior = meta.get('prima_anterior', '')
    paquetes = [(datos_por_aseguradora[a].get('paquete') or 'AMPLIA').upper() for a in aseguradoras]
    if len(set(paquetes)) == 1:
        paquete_header = paquetes[0]
    else:
        paquete_header = 'VER PAQUETE POR ASEGURADORA'
    logo_olivas_path = buscar_logo('OLIVAS', logos_dir) if logos_dir else None
    if logo_olivas_path and os.path.isfile(logo_olivas_path):
        celda_olivas = RLImage(logo_olivas_path, width=3.5 * cm, height=1.5 * cm,
                               kind='proportional', hAlign='RIGHT')
    else:
        celda_olivas = Paragraph("Olivas &amp; Asociados<br/><font size=7>SEGUROS Y FIANZAS</font>", estilo_marca)
    serie_txt = f"<br/>Serie: {serie}" if serie else ""
    header_data = [
        [
            Paragraph(f"<b>{cliente}</b><br/>C.P. {cp}", estilo_valor),
            Paragraph(f"<b>Paquete:</b> {paquete_header}<br/><b>Forma de Pago:</b> CONTADO", estilo_valor),
            Paragraph(f"<b>Vigencia:</b> {vigencia}<br/><b>Tipo:</b> {meta.get('tipo', 'RENOVACIÓN')}", estilo_valor),
            celda_olivas,
        ],
        [
            Paragraph(f"{vehiculo}{serie_txt}", estilo_valor),
            '', '',
            Paragraph(f"<font size=7><b>Prima Anterior:</b> {prima_anterior}</font>" if prima_anterior else '', estilo_valor),
        ],
    ]
    header_tbl = Table(header_data, colWidths=[6.5 * cm, 6 * cm, 7 * cm, 6 * cm])
    header_tbl.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('SPAN', (0, 1), (0, 1)),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(header_tbl)
    story.append(Spacer(1, 8))
    est_cob = ParagraphStyle('cob', fontName='Helvetica-Bold', fontSize=7,
                             textColor=COLOR_HEADER, leading=9, alignment=TA_LEFT)
    est_celda = ParagraphStyle('celda', fontName='Helvetica', fontSize=7,
                               textColor=colors.HexColor('#333333'), leading=9,
                               alignment=TA_CENTER)
    est_amparada = ParagraphStyle('amp', fontName='Helvetica-Bold', fontSize=7,
                                  textColor=COLOR_AMPARADA, leading=9,
                                  alignment=TA_CENTER)
    est_noamp = ParagraphStyle('noamp', fontName='Helvetica', fontSize=7,
                               textColor=COLOR_NO_AMPARADA, leading=9,
                               alignment=TA_CENTER)
    est_paquete = ParagraphStyle('paq', fontName='Helvetica-Bold', fontSize=7,
                                 textColor=COLOR_HEADER, leading=9,
                                 alignment=TA_CENTER)
    def celda_para(valor):
        if valor is None or valor == '':
            return Paragraph('', est_celda)
        v = str(valor).upper().strip()
        if v in ('AMPARADA', 'AMPARADO'):
            return Paragraph(valor, est_amparada)
        if v == 'NO AMPARADA':
            return Paragraph('NO AMPARADA', est_noamp)
        return Paragraph(valor, est_celda)
    n_aseg = len(aseguradoras)
    ancho_coberturas = 7.5 * cm
    ancho_par = (25.5 * cm - ancho_coberturas) / n_aseg

    def _logo(a):
        ruta = datos_por_aseguradora[a].get('logo_path')
        if not ruta and logos_dir:
            ruta = buscar_logo(a, logos_dir)
        return ruta

    tabla = [
        ['COBERTURAS'] + sum([[logo_celda(_logo(a), ancho_par * 0.6, 1.3 * cm,
                                          a), ''] for a in aseguradoras], []),
        [''] + sum([['SUMA ASEGURADA', 'DEDUCIBLE'] for _ in aseguradoras], []),
    ]
    fila_paq = [Paragraph('PAQUETE COTIZADO', est_cob)]
    for aseg in aseguradoras:
        paq = (datos_por_aseguradora[aseg].get('paquete') or 'AMPLIA').upper()
        fila_paq.extend([Paragraph(paq, est_paquete), Paragraph('', est_paquete)])
    tabla.append(fila_paq)
    FILA_PAQUETE_IDX = 2
    for cob in COBERTURAS_ORDEN:
        # Omitir la fila si NINGUNA aseguradora trae esta cobertura
        if not any(cob in datos_por_aseguradora[a].get('coberturas', {}) for a in aseguradoras):
            continue
        fila = [Paragraph(cob, est_cob)]
        for aseg in aseguradoras:
            d = datos_por_aseguradora[aseg]
            valor = d.get('coberturas', {}).get(cob)
            if valor is None:
                fila.extend([celda_para('NO AMPARADA'), Paragraph('', est_celda)])
            else:
                suma, ded = valor
                fila.extend([celda_para(suma), celda_para(ded)])
        tabla.append(fila)
    est_prima = ParagraphStyle('prima', fontName='Helvetica-Bold', fontSize=12,
                               textColor=colors.white, leading=14, alignment=TA_CENTER)
    est_prima_label = ParagraphStyle('plab', fontName='Helvetica-Bold', fontSize=10,
                                     textColor=colors.white, leading=14, alignment=TA_LEFT)
    fila_prima = [Paragraph('PRIMA TOTAL CONTADO', est_prima_label)]
    for aseg in aseguradoras:
        d = datos_por_aseguradora[aseg]
        pp = d.get('planes_pago') or {}
        if pp.get('contado'):
            prima = fmt_monto(pp['contado'])
        else:
            prima = d.get('prima_total', '')
            plan = ('SEMESTRAL' if pp.get('sem_pi') else
                    'TRIMESTRAL' if pp.get('tri_pi') else '')
            if prima and plan:
                prima += f"<br/><font size=7>(pago {plan.lower()}, no cotizó contado)</font>"
        fila_prima.extend([Paragraph(prima, est_prima), Paragraph('', est_prima)])
    tabla.append(fila_prima)
    n_aseg = len(aseguradoras)
    ancho_coberturas = 7.5 * cm
    ancho_total = 25.5 * cm - ancho_coberturas
    ancho_par = ancho_total / n_aseg
    col_widths = [ancho_coberturas] + [ancho_par / 2, ancho_par / 2] * n_aseg
    t = Table(tabla, colWidths=col_widths, repeatRows=2)
    estilo = [
        ('BACKGROUND', (0, 0), (-1, 1), COLOR_HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 1), COLOR_HEADER),
        ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 1), 8),
        ('ALIGN', (0, 0), (-1, 1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 1), 'MIDDLE'),
        ('SPAN', (0, 0), (0, 1)),
        ('ALIGN', (0, 0), (0, 1), 'LEFT'),
        ('LEFTPADDING', (0, 0), (0, 1), 8),
    ]
    for i, _ in enumerate(aseguradoras):
        col_ini = 1 + 2 * i
        col_fin = col_ini + 1
        estilo.append(('SPAN', (col_ini, 0), (col_fin, 0)))
        estilo.append(('SPAN', (col_ini, FILA_PAQUETE_IDX), (col_fin, FILA_PAQUETE_IDX)))
    estilo.append(('BACKGROUND', (0, FILA_PAQUETE_IDX), (-1, FILA_PAQUETE_IDX), colors.HexColor('#EDEFF4')))
    estilo += [
        ('VALIGN', (0, 2), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 2), (0, -2), 8),
        ('TOPPADDING', (0, 2), (-1, -2), 4),
        ('BOTTOMPADDING', (0, 2), (-1, -2), 4),
        ('GRID', (0, 0), (-1, -1), 0.4, COLOR_BORDE),
        ('LINEBELOW', (0, 1), (-1, 1), 0.8, COLOR_HEADER),
    ]
    for i in range(FILA_PAQUETE_IDX + 1, len(tabla) - 1):
        if i % 2 == 0:
            estilo.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#FAFBFC')))
    fila_prima_idx = len(tabla) - 1
    estilo += [
        ('BACKGROUND', (0, fila_prima_idx), (-1, fila_prima_idx), COLOR_PRIMA),
        ('TOPPADDING', (0, fila_prima_idx), (-1, fila_prima_idx), 8),
        ('BOTTOMPADDING', (0, fila_prima_idx), (-1, fila_prima_idx), 8),
        ('LEFTPADDING', (0, fila_prima_idx), (0, fila_prima_idx), 8),
    ]
    for i in range(n_aseg):
        col_ini = 1 + 2 * i
        col_fin = col_ini + 1
        estilo.append(('SPAN', (col_ini, fila_prima_idx), (col_fin, fila_prima_idx)))
    t.setStyle(TableStyle(estilo))
    story.append(t)
    story.append(Spacer(1, 10))
    nota = ParagraphStyle('nota', parent=styles['Normal'], fontSize=8,
                          textColor=colors.HexColor('#444444'), alignment=TA_LEFT,
                          backColor=colors.HexColor('#F4F5F8'), borderPadding=8)
    story.append(Paragraph(
        "<b>Nota importante:</b> Esta propuesta tiene una vigencia de "
        "<b>7 días naturales</b> a partir de su fecha de emisión. "
        "Los precios y condiciones pueden variar después de este periodo.",
        nota))
    story.append(Spacer(1, 12))
    cierre = ParagraphStyle('cierre', parent=styles['Normal'], fontSize=10,
                            textColor=COLOR_HEADER, alignment=TA_CENTER,
                            fontName='Helvetica-Oblique')
    story.append(Paragraph("¡Gracias por su preferencia!", cierre))

    # ==========================================================
    # SEGUNDA HOJA — OPCIONES DE PAGO
    # Layout: total del plan en grande, con el desglose (primer pago /
    # subsecuente) en chico debajo. Semestral total = 1er + 1 subs;
    # Trimestral total = 1er + 3 subs.
    # ==========================================================
    story.append(PageBreak())
    encabezado_p2 = ParagraphStyle('h2', parent=styles['Normal'], fontSize=14,
                                   textColor=COLOR_HEADER, alignment=TA_LEFT,
                                   fontName='Helvetica-Bold', spaceAfter=4)
    subhead_p2 = ParagraphStyle('sub2', parent=styles['Normal'], fontSize=9,
                                textColor=colors.HexColor('#555555'), alignment=TA_LEFT,
                                spaceAfter=14)
    story.append(Paragraph("Prima Total", encabezado_p2))
    subtitulo = f"{cliente}" + (f" — {vehiculo}" if vehiculo else "")
    story.append(Paragraph(subtitulo, subhead_p2))

    def _num(v):
        """Convierte string con comas/decimales a float, o None."""
        if v is None or str(v).strip() == '':
            return None
        try:
            return float(str(v).replace(',', '').strip())
        except ValueError:
            return None

    def _big(v):
        """Formato grande: '$ 12,061.67' o '—'."""
        return '—' if v is None else f"$ {v:,.2f}"

    def _small(v):
        """Formato chico: '$ 4,567.00' o '—'."""
        return '—' if v is None else f"$ {v:,.2f}"

    # Estilos
    est_label_plan = ParagraphStyle('lplan', fontName='Helvetica-Bold', fontSize=11,
                                    textColor=COLOR_HEADER, leading=13, alignment=TA_LEFT)
    est_label_sub = ParagraphStyle('lsub', fontName='Helvetica', fontSize=8,
                                   textColor=colors.HexColor('#666666'), leading=11,
                                   alignment=TA_LEFT)
    est_val_grande = ParagraphStyle('vg', fontName='Helvetica-Bold', fontSize=13,
                                    textColor=colors.HexColor('#1F2A44'), leading=15,
                                    alignment=TA_CENTER)
    est_val_grande_dim = ParagraphStyle('vgd', fontName='Helvetica', fontSize=13,
                                        textColor=colors.HexColor('#BBBBBB'), leading=15,
                                        alignment=TA_CENTER)
    est_val_chico = ParagraphStyle('vc', fontName='Helvetica', fontSize=8,
                                   textColor=colors.HexColor('#666666'), leading=11,
                                   alignment=TA_CENTER)
    est_aseg2 = ParagraphStyle('asg2', fontName='Helvetica-Bold', fontSize=10,
                               textColor=COLOR_HEADER, alignment=TA_CENTER, leading=12)

    def val_grande(v):
        return Paragraph(_big(v), est_val_grande if v is not None else est_val_grande_dim)

    def val_desglose(pi, ps):
        """Dos líneas: primer pago arriba, subsecuente abajo."""
        return Paragraph(f"{_small(pi)}<br/>{_small(ps)}", est_val_chico)

    # Precalcular totales por aseguradora
    resumen = {}
    for a in aseguradoras:
        pp = datos_por_aseguradora[a].get('planes_pago') or {}
        contado = _num(pp.get('contado'))
        sem_pi, sem_ps = _num(pp.get('sem_pi')), _num(pp.get('sem_ps'))
        tri_pi, tri_ps = _num(pp.get('tri_pi')), _num(pp.get('tri_ps'))
        sem_total = (sem_pi + sem_ps) if (sem_pi is not None and sem_ps is not None) else None
        tri_total = (tri_pi + 3 * tri_ps) if (tri_pi is not None and tri_ps is not None) else None
        resumen[a] = {
            'contado': contado,
            'sem_total': sem_total, 'sem_pi': sem_pi, 'sem_ps': sem_ps,
            'tri_total': tri_total, 'tri_pi': tri_pi, 'tri_ps': tri_ps,
        }

    # Construcción de tabla
    tabla2 = []
    # Fila 0: header con nombres de aseguradoras
    fila_head = ['']
    for a in aseguradoras:
        fila_head.append(logo_celda(_logo(a), (25.5 * cm - 5.0 * cm) / len(aseguradoras) * 0.5,
                                    1.4 * cm, Paragraph(a, est_aseg2)))
    tabla2.append(fila_head)

    # Fila 1: CONTADO
    fila = [Paragraph('Contado', est_label_plan)]
    for a in aseguradoras:
        fila.append(val_grande(resumen[a]['contado']))
    tabla2.append(fila)

    # Fila 2: SEMESTRAL (total grande)
    fila = [Paragraph('Semestral', est_label_plan)]
    for a in aseguradoras:
        fila.append(val_grande(resumen[a]['sem_total']))
    tabla2.append(fila)

    # Fila 3: Semestral - desglose (chico)
    fila = [Paragraph('Primer pago<br/>Subsecuente', est_label_sub)]
    for a in aseguradoras:
        fila.append(val_desglose(resumen[a]['sem_pi'], resumen[a]['sem_ps']))
    tabla2.append(fila)

    # Fila 4: TRIMESTRAL (total grande)
    fila = [Paragraph('Trimestral', est_label_plan)]
    for a in aseguradoras:
        fila.append(val_grande(resumen[a]['tri_total']))
    tabla2.append(fila)

    # Fila 5: Trimestral - desglose (chico)
    fila = [Paragraph('Primer pago<br/>(3) Subsecuentes', est_label_sub)]
    for a in aseguradoras:
        fila.append(val_desglose(resumen[a]['tri_pi'], resumen[a]['tri_ps']))
    tabla2.append(fila)

    n2 = len(aseguradoras)
    ancho_label2 = 5.0 * cm
    ancho_total2 = 25.5 * cm - ancho_label2
    col_widths2 = [ancho_label2] + [ancho_total2 / n2] * n2
    t2 = Table(tabla2, colWidths=col_widths2, repeatRows=1)
    estilo2 = [
        # Encabezado (aseguradoras)
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_HEADER_BG),
        ('ALIGN', (1, 0), (-1, 0), 'CENTER'),
        ('LINEBELOW', (0, 0), (-1, 0), 1, COLOR_HEADER),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        # Layout general
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (0, -1), 10),
        # Fila Contado (fila 1) — resaltada
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#FAFBFC')),
        ('TOPPADDING', (0, 1), (-1, 1), 12),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 12),
        # Fila Semestral total (2)
        ('TOPPADDING', (0, 2), (-1, 2), 14),
        ('BOTTOMPADDING', (0, 2), (-1, 2), 4),
        ('LINEABOVE', (0, 2), (-1, 2), 0.5, colors.HexColor('#DDDDDD')),
        # Fila Semestral desglose (3)
        ('TOPPADDING', (0, 3), (-1, 3), 0),
        ('BOTTOMPADDING', (0, 3), (-1, 3), 14),
        # Fila Trimestral total (4)
        ('TOPPADDING', (0, 4), (-1, 4), 14),
        ('BOTTOMPADDING', (0, 4), (-1, 4), 4),
        ('LINEABOVE', (0, 4), (-1, 4), 0.5, colors.HexColor('#DDDDDD')),
        # Fila Trimestral desglose (5)
        ('TOPPADDING', (0, 5), (-1, 5), 0),
        ('BOTTOMPADDING', (0, 5), (-1, 5), 14),
        # Sin bordes internos entre filas de un plan (más limpio)
        ('LINEBELOW', (0, -1), (-1, -1), 0.5, colors.HexColor('#DDDDDD')),
    ]
    t2.setStyle(TableStyle(estilo2))
    story.append(t2)

    story.append(Spacer(1, 14))
    nota2 = ParagraphStyle('nota2', parent=styles['Normal'], fontSize=7.5,
                           textColor=colors.HexColor('#666666'), alignment=TA_LEFT,
                           leading=10)
    story.append(Paragraph(
        "El total de Semestral y Trimestral incluye el cargo por financiamiento de la aseguradora. "
        "El guion \"—\" indica que la aseguradora no incluyó ese plan en el PDF cotizado.",
        nota2))

    doc.build(story)
    return salida


def procesar_cotizaciones(rutas_pdf, salida='comparativo.pdf', meta_override=None, logos_dir=None, watermark=None):
    datos = {}
    for ruta in rutas_pdf:
        texto = extraer_texto(ruta)
        aseguradora = detectar_aseguradora(texto)
        if not aseguradora:
            print(f"⚠️  No pude detectar la aseguradora en: {ruta}")
            continue
        parser = PARSERS[aseguradora]
        if aseguradora == 'GNP':
            d = parser(texto, pdf_path=ruta)
        else:
            d = parser(texto)
        normalizar_coberturas(d)
        d['planes_pago'] = extraer_planes_pago(texto, aseguradora)
        nombre_columna = d['aseguradora']
        # AVISO si dos PDFs se detectan como la misma aseguradora
        # (antes uno pisaba al otro sin decir nada)
        if nombre_columna in datos:
            print(f"⚠️  OJO: '{os.path.basename(ruta)}' también se detectó como {nombre_columna}. "
                  f"Está REEMPLAZANDO al PDF anterior de esa columna. "
                  f"Si son aseguradoras distintas, revisa los PDFs.")
        d['logo_path'] = extraer_logo(ruta, os.path.join(os.path.dirname(os.path.abspath(salida)), '_logos'),
                                      aseguradora)
        datos[nombre_columna] = d
        print(f"✓ {nombre_columna}: {len(d['coberturas'])} coberturas, "
              f"paquete {d.get('paquete') or 'N/D'}, prima {d.get('prima_total', 'N/D')}")
    def primero_no_vacio(campo, default=''):
        for d in datos.values():
            v = d.get(campo)
            if v and str(v).strip():
                return v
        return default
    vigencia = ''
    for d in datos.values():
        ini = d.get('vigencia_ini')
        fin = d.get('vigencia_fin')
        if ini and str(ini).strip() and fin and str(fin).strip():
            vigencia = f"{str(ini).strip()} al {str(fin).strip()}"
            break
    if not vigencia:
        vigencia = f"{primero_no_vacio('vigencia_ini')} al {primero_no_vacio('vigencia_fin')}"
    meta = {
        'cliente': primero_no_vacio('cliente'),
        'cp': primero_no_vacio('cp'),
        'vehiculo': primero_no_vacio('vehiculo'),
        'serie': primero_no_vacio('serie'),
        'vigencia': vigencia,
    }
    if meta_override:
        meta.update(meta_override)
    construir_pdf(datos, salida=salida, meta=meta, logos_dir=logos_dir, watermark=watermark)
    return salida