#!/usr/bin/env python3
"""Aplicación Streamlit para consultar documentos en el SVE ZOFRI.

Pensada para desplegarse en Streamlit Cloud (streamlit.app). No incluye
servidor HTTP propio ni scripts de Windows ni credenciales en disco: el
usuario y la contraseña se configuran como secretos/variables de entorno.
"""

from __future__ import annotations

import http.cookiejar
import io
import json
import os
import re
import shutil
import threading
import urllib.parse
import urllib.request
from pathlib import Path

import pymupdf as fitz
import pytesseract
import streamlit as st
from PIL import Image

st.set_page_config(page_title="Consulta de documentos — SVE ZOFRI", layout="wide")

st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
      html, body, [class*="css"] { font-family: "Inter", system-ui, sans-serif !important; }
      .stApp { background: #050505 !important; color: #F2F2F2 !important; }
      section[data-testid="stSidebar"] { display: none !important; }
      .block-container { padding: 32px !important; max-width: 1100px; }
      h1, h2, h3 { color: #F2F2F2 !important; font-weight: 600 !important; }
      p, label, .stCaption { color: #929292 !important; }
      .stTextInput input, .stTextArea textarea, .stFileUploader {
        background: #0f0f0f !important;
        border: 1px solid #1a1a1a !important;
        border-radius: 8px !important;
        color: #F2F2F2 !important;
      }
      .stButton > button {
        background: #202020 !important;
        color: #F2F2F2 !important;
        border: 1px solid #1a1a1a !important;
        border-radius: 6px !important;
        font-weight: 500 !important;
      }
      .stButton > button:hover { background: #141414 !important; }
      .stButton > button:disabled { opacity: 0.5 !important; }
      div[data-testid="stMetric"] {
        background: #090909 !important;
        border: 1px solid #1a1a1a !important;
        border-radius: 12px !important;
        padding: 16px !important;
      }
      div[data-testid="stMetricLabel"] { color: #929292 !important; }
      div[data-testid="stMetricValue"] { color: #F2F2F2 !important; }
      .stDataFrame, .stTable { border: 1px solid #1a1a1a !important; border-radius: 12px !important; }
      hr { border-color: #1a1a1a !important; }
      .stAlert { border-radius: 8px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE = "https://svepp.zofri.cl"
PANTALLA = BASE + "/index.php?sys&method=view"
DOC_RE = re.compile(r"^\d{3}-\d{2}-\d{6}$")
ROOT = Path(__file__).resolve().parent

MAX_PDF_BYTES = 20 * 1024 * 1024

# Credenciales del servidor: nunca se envían al navegador.
def _secreto(clave: str) -> str:
    valor = os.environ.get(clave, "")
    if valor:
        return valor
    try:
        return str(st.secrets.get(clave, "") or "")
    except Exception:  # noqa: BLE001
        return ""

SVE_USER = _secreto("SVE_USER")
SVE_PASS = _secreto("SVE_PASS")

# OCR para PDFs escaneados.
ruta_tesseract = os.environ.get("TESSERACT_CMD") or shutil.which("tesseract")
if ruta_tesseract:
    pytesseract.pytesseract.tesseract_cmd = ruta_tesseract

# Reexpedición: 20126 + 4 dígitos (9 en total) o 20126 + 6 (11 en total).
REEXPEDICION_RE = re.compile(r"(?<!\d)20126(\d{4,6})(?!\d)")

@st.cache_resource
def _estado_http():
    jar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
    return {"jar": jar, "opener": opener, "session": {"activa": False, "usuario": ""}, "lock": threading.Lock()}

_ESTADO = _estado_http()
jar = _ESTADO["jar"]
opener = _ESTADO["opener"]
SESSION = _ESTADO["session"]
BLOQUEO = _ESTADO["lock"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "es-CL,es;q=0.9",
}


def http_get(url: str, extra: dict | None = None, timeout: int = 20) -> str:
    cabeceras = dict(HEADERS)
    if extra:
        cabeceras.update(extra)
    peticion = urllib.request.Request(url, headers=cabeceras)
    with opener.open(peticion, timeout=timeout) as respuesta:
        return respuesta.read().decode("utf-8", "replace")


def http_post(url: str, datos: dict, timeout: int = 20) -> str:
    cuerpo = urllib.parse.urlencode(datos).encode("utf-8")
    peticion = urllib.request.Request(url, data=cuerpo, headers=HEADERS)
    with opener.open(peticion, timeout=timeout) as respuesta:
        return respuesta.read().decode("utf-8", "replace")


def iniciar_sesion(usuario: str, clave: str) -> dict:
    with BLOQUEO:
        if not usuario or not clave:
            raise ValueError("Faltan credenciales del SVE.")

        jar.clear()
        http_get(BASE + "/index.php")
        http_post(
            BASE + "/index.php?login",
            {"method": "autenticar", "user": usuario, "passwd": clave},
        )

        texto = http_get(PANTALLA)
        autenticado = (
            "Salir del sistema" in texto
            and 'name="user"' not in texto
            and "No ha iniciado" not in texto
        )
        if not autenticado:
            SESSION.update({"activa": False, "usuario": ""})
            raise RuntimeError("No se pudo iniciar sesión. Revisá las credenciales.")

        SESSION.update({"activa": True, "usuario": usuario})
        return dict(SESSION)


def normalizar(lista: list[str]) -> list[str]:
    salida: list[str] = []
    for valor in lista:
        doc = str(valor).strip().replace(" ", "")
        if re.fullmatch(r"\d{11}", doc):
            doc = f"{doc[:3]}-{doc[3:5]}-{doc[5:]}"
        if doc and doc not in salida:
            salida.append(doc)
    return salida


def consultar(doc: str) -> dict:
    if not DOC_RE.match(doc):
        return {
            "documento": doc,
            "estado": "FORMATO INVÁLIDO",
            "mensaje": "Debe tener el formato NNN-NN-NNNNNN",
        }

    url = BASE + "/index.php?tran&method=buscar_dt&doc=" + urllib.parse.quote(doc) + "&mic="
    cuerpo = http_get(url, {"Referer": PANTALLA, "X-Requested-With": "XMLHttpRequest"})
    if "No ha iniciado" in cuerpo:
        raise RuntimeError("La sesión caducó. Volvé a iniciar sesión.")

    coincide = re.search(r"\$\.msgbox\(\s*'((?:[^'\\]|\\.)*)'", cuerpo)
    mensaje = coincide.group(1) if coincide else ""
    mensaje = re.sub(r"<br\s*/?>", " ", mensaje, flags=re.I)
    mensaje = re.sub(r"<[^>]+>", "", mensaje).replace("\\n", " ").strip()
    bajo = mensaje.lower()

    if "ya fue controlado" in bajo:
        estado = "CONTROLADO"
    elif "no existe" in bajo:
        estado = "NO EXISTE"
    elif "no es controlable" in bajo:
        estado = "NO CONTROLABLE"
    elif re.search(r"var\s+error\s*=\s*0", cuerpo) or not mensaje:
        estado = "NO CONTROLADO"
    else:
        estado = "OTRO"

    return {"documento": doc, "estado": estado, "mensaje": mensaje}


def extraer_reexpediciones(texto: str) -> list[str]:
    encontrados: list[str] = []
    variantes = [texto, re.sub(r"\s+", "", texto)]
    for variante in variantes:
        for coincidencia in REEXPEDICION_RE.finditer(variante):
            codigo = f"201-26-{coincidencia.group(1).zfill(6)}"
            if codigo not in encontrados:
                encontrados.append(codigo)
    return encontrados


def procesar_pdf(contenido: bytes, nombre: str) -> dict:
    documento = fitz.open(stream=contenido, filetype="pdf")
    codigos: list[str] = []
    paginas_ocr = 0

    for pagina in documento:
        texto = pagina.get_text("text") or ""
        encontrados = extraer_reexpediciones(texto)

        if not encontrados:
            pixeles = pagina.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            imagen = Image.open(io.BytesIO(pixeles.tobytes("png")))
            texto_ocr = pytesseract.image_to_string(
                imagen, lang="eng", config="--psm 6"
            )
            paginas_ocr += 1
            encontrados = extraer_reexpediciones(texto_ocr)

        for codigo in encontrados:
            if codigo not in codigos:
                codigos.append(codigo)

    total_paginas = len(documento)
    documento.close()
    return {
        "nombre": nombre,
        "paginas": total_paginas,
        "paginas_ocr": paginas_ocr,
        "codigos": codigos,
    }


# ---------- Interfaz Streamlit ----------

st.title("Consulta de documentos · SVE ZOFRI")
st.caption("Revisa en lote si un código xxx-xx-xxxxxx ya fue controlado en Control de Tránsito.")

if "session_activa" not in st.session_state:
    st.session_state.session_activa = bool(SESSION.get("activa", False))
if "session_usuario" not in st.session_state:
    st.session_state.session_usuario = SESSION.get("usuario", "")
if "resultados" not in st.session_state:
    st.session_state.resultados = []
if "codigos_area" not in st.session_state:
    st.session_state.codigos_area = ""

# 1. Sesión
st.subheader("1. Sesión")
col_u, col_p, col_b = st.columns([2, 2, 1])
with col_u:
    usuario_input = st.text_input("Usuario del SVE", value=SVE_USER if SVE_USER else "", key="usuario")
with col_p:
    clave_input = st.text_input("Contraseña", type="password", key="clave")
with col_b:
    st.write("")
    iniciar = st.button("Iniciar sesión")

if SVE_USER and SVE_PASS:
    st.info("El servidor ya tiene las credenciales configuradas: podés dejar la contraseña vacía y clic en Iniciar sesión.")

if iniciar:
    usuario = (usuario_input or "").strip() or SVE_USER
    clave = (clave_input or "") or SVE_PASS
    try:
        datos = iniciar_sesion(usuario, clave)
        st.session_state.session_activa = bool(datos.get("activa"))
        st.session_state.session_usuario = datos.get("usuario", "")
        st.success("Sesión iniciada correctamente.")
    except Exception as error:  # noqa: BLE001
        st.session_state.session_activa = False
        st.session_state.session_usuario = ""
        st.error(str(error))

if st.session_state.session_activa:
    st.success(f"Conectado: {st.session_state.session_usuario}")
else:
    st.warning("Sin sesión")

# 2. Códigos
st.subheader("2. Códigos a revisar")
st.text_area(
    "Códigos (uno por línea, separados por coma o sin guiones)",
    height=160,
    key="codigos_area",
)

col_btn1, col_btn2 = st.columns([1, 1])
consultar_btn = col_btn1.button("Consultar documentos", disabled=not st.session_state.session_activa)
importar_btn = col_btn2.button("Importar reexpediciones (PDF)")

uploaded = st.file_uploader("Archivos PDF", type=["pdf"], accept_multiple_files=True, key="pdfs")

if importar_btn:
    if not uploaded:
        st.error("Elegí al menos un archivo PDF.")
    else:
        nuevos_total = 0
        detalles = []
        for archivo in uploaded:
            contenido = archivo.read()
            if len(contenido) > MAX_PDF_BYTES:
                detalles.append(f"{archivo.name}: El PDF es demasiado grande (máx. 20 MB).")
                continue
            try:
                informe = procesar_pdf(contenido, archivo.name)
            except Exception as error:  # noqa: BLE001
                informe = {"nombre": archivo.name, "error": str(error), "codigos": []}
            if "error" in informe:
                detalles.append(f"{informe['nombre']}: {informe['error']}")
            else:
                detalles.append(f"{informe['nombre']}: {len(informe.get('codigos', []))} código(s)")
                actuales = [c for c in st.session_state.codigos_area.splitlines() if c.strip()]
                # normalizar lista actual
                nuevos = [c for c in informe.get("codigos", []) if c not in actuales]
                if nuevos:
                    if st.session_state.codigos_area.strip():
                        st.session_state.codigos_area = st.session_state.codigos_area.rstrip() + "\n" + "\n".join(nuevos)
                    else:
                        st.session_state.codigos_area = "\n".join(nuevos)
                    nuevos_total += len(nuevos)
        msg = f"{sum(1 for _ in uploaded)} archivo(s) procesado(s). {nuevos_total} agregado(s). " + " · ".join(detalles)
        if nuevos_total:
            st.success(msg)
        else:
            st.info(msg)
        st.rerun()

if consultar_btn:
    codigos_lista = [c for c in re.split(r"[\s,;]+", st.session_state.codigos_area) if c.strip()]
    if not codigos_lista:
        st.error("Pegá al menos un código.")
    else:
        with st.spinner(f"Consultando {len(codigos_lista)} código(s)..."):
            documentos = normalizar(codigos_lista)
            resultados = []
            for doc in documentos:
                try:
                    resultados.append(consultar(doc))
                except Exception as error:  # noqa: BLE001
                    resultados.append({"documento": doc, "estado": "ERROR", "mensaje": str(error)})
            st.session_state.resultados = resultados
        st.success("Consulta terminada.")

# 3. Resultados
st.subheader("3. Resultados")
resultados = st.session_state.resultados
if resultados:
    conteo = {}
    for r in resultados:
        conteo[r["estado"]] = conteo.get(r["estado"], 0) + 1
    cols = st.columns(len(conteo) or 1)
    for i, (estado, n) in enumerate(conteo.items()):
        cols[i].metric(estado, n)

    def clase_estado(estado):
        if estado == "CONTROLADO":
            return "🔴"
        if estado == "NO CONTROLABLE":
            return "🟠"
        if estado == "NO CONTROLADO":
            return "🟢"
        return "⚪"

    filas = []
    for r in resultados:
        filas.append({
            "Documento": r.get("documento"),
            "Estado": f"{clase_estado(r.get('estado'))} {r.get('estado')}",
            "Mensaje": r.get("mensaje") or "—",
        })
    st.dataframe(filas, use_container_width=True)

    # CSV
    import csv
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter=";")
    writer.writerow(["documento", "estado", "mensaje"])
    for r in resultados:
        writer.writerow([r.get("documento"), r.get("estado"), r.get("mensaje")])
    st.download_button(
        "Descargar CSV",
        data="\ufeff" + buf.getvalue(),
        file_name="resultados_sve.csv",
        mime="text/csv",
    )
else:
    st.info("Aún no hay consultas.")
