"""
Comparador de Cotizaciones — Olivas & Asociados
Interfaz web (Streamlit) sobre las funciones de comparador.py (antes Celda 2).
"""
import os
import tempfile
from datetime import datetime

import streamlit as st

import comparador as cmp

st.set_page_config(page_title="Comparador de Cotizaciones", page_icon="📄", layout="centered")

# ── Contraseña opcional ───────────────────────────────────────────
# Si en los "Secrets" del hosting defines  APP_PASSWORD = "algo",
# la app pedirá esa contraseña. Si no la defines, entra cualquiera.
def _password_ok():
    try:
        pwd = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        pwd = ""
    if not pwd:
        return True
    if st.session_state.get("auth"):
        return True
    intento = st.text_input("Contraseña", type="password")
    if intento and intento == pwd:
        st.session_state["auth"] = True
        st.rerun()
    elif intento:
        st.error("Contraseña incorrecta")
    return False


if not _password_ok():
    st.stop()

LOGOS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logos")

st.title("Comparador de Cotizaciones")
st.caption("Olivas & Asociados · Seguros y Fianzas — HDI, CHUBB, QUÁLITAS, GNP, INBURSA, ANA")

archivos = st.file_uploader("Sube los PDFs de las cotizaciones", type=["pdf"],
                            accept_multiple_files=True)

with st.expander("Datos opcionales del comparativo"):
    col1, col2 = st.columns(2)
    prima_anterior = col1.text_input("Prima anterior", placeholder="$ 0.00")
    tipo = col2.selectbox("Tipo", ["RENOVACIÓN", "NUEVO"])
    cliente = st.text_input("Cliente (si lo dejas vacío se toma del PDF)")
    vehiculo = st.text_input("Vehículo (si lo dejas vacío se toma del PDF)")
    cp = st.text_input("C.P. (si lo dejas vacío se toma del PDF)")
    nombre_salida = st.text_input("Nombre del archivo", value="COMPARATIVO")

if not archivos:
    st.info("Sube 2 o más cotizaciones en PDF para empezar.")
    st.stop()

# ── Guardar PDFs subidos y detectar aseguradora ───────────────────
tmp = tempfile.mkdtemp(prefix="cotiz_")
rutas, filas, detectadas = [], [], {}
for f in archivos:
    ruta = os.path.join(tmp, f.name)
    with open(ruta, "wb") as out:
        out.write(f.getbuffer())
    try:
        aseg = cmp.detectar_aseguradora(cmp.extraer_texto(ruta))
    except Exception as e:
        aseg, err = None, str(e)
    else:
        err = ""
    filas.append({"Archivo": f.name, "Aseguradora": aseg or "❌ no detectada", "Nota": err})
    if aseg:
        rutas.append(ruta)
        detectadas.setdefault(aseg, []).append(f.name)

st.subheader("Archivos detectados")
st.dataframe(filas, hide_index=True, use_container_width=True)

for aseg, nombres in detectadas.items():
    if len(nombres) > 1:
        st.warning(f"{len(nombres)} PDFs se detectaron como **{aseg}** ({', '.join(nombres)}). "
                   f"Solo se usará el último.")

if not rutas:
    st.error("No se reconoció ninguna aseguradora en los PDFs.")
    st.stop()

if st.button("Generar comparativo", type="primary", use_container_width=True):
    override = {"prima_anterior": prima_anterior, "tipo": tipo}
    for k, v in (("cliente", cliente), ("vehiculo", vehiculo), ("cp", cp)):
        if v.strip():
            override[k] = v.strip()
    salida = os.path.join(tmp, f"{(nombre_salida or 'COMPARATIVO').strip()}.pdf")
    with st.spinner("Leyendo cotizaciones y armando el PDF…"):
        try:
            cmp.procesar_cotizaciones(rutas, salida=salida, meta_override=override,
                                      logos_dir=LOGOS_DIR if os.path.isdir(LOGOS_DIR) else None)
        except Exception as e:
            st.exception(e)
            st.stop()
    with open(salida, "rb") as fh:
        st.session_state["pdf"] = fh.read()
        st.session_state["pdf_name"] = os.path.basename(salida)
    st.success("Comparativo listo")

if st.session_state.get("pdf"):
    st.download_button("⬇️ Descargar comparativo", st.session_state["pdf"],
                       file_name=st.session_state["pdf_name"], mime="application/pdf",
                       use_container_width=True)
