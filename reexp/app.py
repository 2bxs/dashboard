#!/usr/bin/env python3
"""Aplicación web para consultar documentos en el SVE ZOFRI.

Pensada para desplegarse en un hosting con Docker (Render, Railway, Fly.io,
Hugging Face Spaces). No incluye scripts de Windows ni credenciales en disco:
el usuario y la contraseña se configuran como variables de entorno del servidor.
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
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pymupdf as fitz
import pytesseract
from PIL import Image

BASE = "https://svepp.zofri.cl"
PANTALLA = BASE + "/index.php?sys&method=view"
DOC_RE = re.compile(r"^\d{3}-\d{2}-\d{6}$")
ROOT = Path(__file__).resolve().parent

# El hosting define PORT; si no, usamos 7860 (compatible con HF Spaces).
PUERTO = int(os.environ.get("PORT") or 7860)
HOST = os.environ.get("HOST", "0.0.0.0")
MAX_PDF_BYTES = 20 * 1024 * 1024

# Credenciales del servidor: nunca se envían al navegador.
SVE_USER = os.environ.get("SVE_USER", "")
SVE_PASS = os.environ.get("SVE_PASS", "")

# OCR para PDFs escaneados (en Docker tesseract está en el PATH).
ruta_tesseract = os.environ.get("TESSERACT_CMD") or shutil.which("tesseract")
if ruta_tesseract:
    pytesseract.pytesseract.tesseract_cmd = ruta_tesseract

# Reexpedición: 20126 + 4 dígitos (9 en total) o 20126 + 6 (11 en total).
REEXPEDICION_RE = re.compile(r"(?<!\d)20126(\d{4,6})(?!\d)")

jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
SESSION = {"activa": False, "usuario": ""}
BLOQUEO = threading.Lock()

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
    """Devuelve códigos NNN-NN-NNNNNN desde reexpediciones del PDF.

    201264772 -> 201-26-004772 ; 20126477270 -> 201-26-477270.
    """
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


class Manejador(BaseHTTPRequestHandler):
    def log_message(self, formato: str, *args) -> None:
        print(f"[web] {self.address_string()} - {formato % args}", flush=True)

    def _json(self, objeto: dict, codigo: int = 200) -> None:
        cuerpo = json.dumps(objeto, ensure_ascii=False).encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(cuerpo)

    def _html(self) -> None:
        cuerpo = (ROOT / "index.html").read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(cuerpo)

    def _cuerpo_json(self) -> dict:
        largo = int(self.headers.get("Content-Length", 0))
        if not largo:
            return {}
        return json.loads(self.rfile.read(largo).decode("utf-8"))

    def _leer_multipart(self) -> list[tuple[dict[str, str], bytes]]:
        """Lee FormData sin depender de cgi (eliminado en Python 3.13)."""
        tipo = self.headers.get("Content-Type", "")
        limite = re.search(r'boundary=(?:"([^"]+)"|([^;]+))', tipo, re.I)
        if not limite:
            raise ValueError("El formulario no indicó un límite (boundary).")
        separador = b"--" + (limite.group(1) or limite.group(2)).strip().encode("ascii")
        crudo = self.rfile.read(int(self.headers.get("Content-Length", 0)))

        partes: list[tuple[dict[str, str], bytes]] = []
        for bloque in crudo.split(separador)[1:-1]:
            bloque = bloque.lstrip(b"\r\n")
            if bloque.endswith(b"\r\n"):
                bloque = bloque[:-2]
            if b"\r\n\r\n" not in bloque:
                continue
            crudo_headers, contenido = bloque.split(b"\r\n\r\n", 1)
            cabeceras: dict[str, str] = {}
            for linea in crudo_headers.decode("utf-8", "replace").split("\r\n"):
                if ":" in linea:
                    clave, valor = linea.split(":", 1)
                    cabeceras[clave.strip().lower()] = valor.strip()
            partes.append((cabeceras, contenido))
        return partes

    def _importar_pdf(self) -> None:
        largo = int(self.headers.get("Content-Length", 0))
        if largo > MAX_PDF_BYTES:
            self._json({"error": "El PDF es demasiado grande (máx. 20 MB)."}, 413)
            return

        archivos: list[tuple[str, bytes]] = []
        for cabeceras, contenido in self._leer_multipart():
            disposition = cabeceras.get("content-disposition", "")
            nombre = re.search(r'filename="([^"]*)"', disposition)
            if nombre and nombre.group(1) and contenido:
                archivos.append((nombre.group(1), contenido))

        if not archivos:
            self._json({"error": "Elegí al menos un archivo PDF."}, 400)
            return

        informes: list[dict] = []
        codigos: list[str] = []
        for nombre, contenido in archivos:
            try:
                informe = procesar_pdf(contenido, nombre)
            except Exception as error:  # noqa: BLE001
                informe = {"nombre": nombre, "error": str(error), "codigos": []}
            informes.append(informe)
            for codigo in informe.get("codigos", []):
                if codigo not in codigos:
                    codigos.append(codigo)

        self._json({"archivos": informes, "codigos": codigos})

    def do_GET(self) -> None:
        ruta = urllib.parse.urlparse(self.path).path
        if ruta in ("/", "/index.html"):
            self._html()
        elif ruta in ("/healthz", "/api/estado"):
            self._json(dict(SESSION) if ruta == "/api/estado" else {"ok": True})
        elif ruta == "/api/config":
            # Nunca se envía la contraseña al navegador.
            self._json(
                {
                    "usuario": SVE_USER,
                    "autologin": bool(SVE_USER and SVE_PASS),
                }
            )
        else:
            self._json({"error": "No encontrado"}, 404)

    def do_POST(self) -> None:
        ruta = urllib.parse.urlparse(self.path).path
        try:
            if ruta == "/api/importar-pdf":
                self._importar_pdf()
                return
            datos = self._cuerpo_json()
            if ruta == "/api/sesion":
                # Si el servidor tiene secretos configurados, se usan aunque el
                # navegador no envíe la contraseña.
                self._json(
                    iniciar_sesion(
                        datos.get("usuario", "") or SVE_USER,
                        datos.get("clave", "") or SVE_PASS,
                    )
                )
            elif ruta == "/api/consultar":
                if not SESSION["activa"]:
                    self._json({"error": "Primero iniciá sesión."}, 401)
                    return
                documentos = normalizar(datos.get("codigos", []))
                if not documentos:
                    self._json({"error": "No hay códigos para consultar."}, 400)
                    return
                resultados = []
                for doc in documentos:
                    try:
                        resultados.append(consultar(doc))
                    except Exception as error:  # noqa: BLE001
                        resultados.append(
                            {
                                "documento": doc,
                                "estado": "ERROR",
                                "mensaje": str(error),
                            }
                        )
                self._json({"resultados": resultados})
            else:
                self._json({"error": "No encontrado"}, 404)
        except Exception as error:  # noqa: BLE001
            self._json({"error": str(error)}, 500)


if __name__ == "__main__":
    servidor = ThreadingHTTPServer((HOST, PUERTO), Manejador)
    print(f"Consulta SVE ZOFRI -> http://{HOST}:{PUERTO}", flush=True)
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nDetenido.", flush=True)
    finally:
        servidor.server_close()
