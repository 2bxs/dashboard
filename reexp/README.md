# Consulta de documentos — SVE ZOFRI (versión web)

Aplicación web para revisar en lote si un código de documento **ya fue
controlado** en *Control de Tránsito* del SVE. Incluye tema oscuro e
importación de reexpediciones desde PDF (digitales o escaneados con OCR).

> Solo consulta: **no registra tránsitos ni modifica datos**.

---

## Despliegue desde GitHub

Esta aplicación necesita un proceso Python que mantenga la sesión del SVE, por
lo que **GitHub Pages no la puede ejecutar** (sirve solo HTML estático).
Subí este repo a GitHub y desplegalo en un hosting con Docker.

### Render (más simple)

1. Creá un repo en GitHub y subí **todo el contenido** de esta carpeta.
2. En Render: **New + → Blueprint** → elegí el repo (usa `render.yaml`).
3. Al crear el servicio, cargá `SVE_USER` y `SVE_PASS` como secretos.
4. Entrá a la URL que te entrega Render, por ejemplo
   `https://sve-zofri.onrender.com`.

### Railway / Fly.io / Hugging Face Spaces

Todos pueden usar el `Dockerfile` incluido. Configurá las variables
`SVE_USER` y `SVE_PASS` en el panel del proveedor. El contenedor escucha en
`0.0.0.0:$PORT` (por defecto `7860`, compatible con HF Spaces).

### Servidor propio con Docker

```bash
docker compose up -d --build
```

Queda disponible en `http://TU_SERVIDOR:8765`.

---

## Variables de entorno

| Variable | Uso | Obligatoria |
|---|---|---|
| `SVE_USER` | Usuario del SVE | No |
| `SVE_PASS` | Contraseña del SVE | No |
| `PORT` | Puerto que indica el hosting | No (default `7860`) |
| `HOST` | Interface de escucha | No (default `0.0.0.0`) |
| `TESSERACT_CMD` | Ruta a Tesseract | No (se detecta en el PATH) |

Si configurás `SVE_USER` y `SVE_PASS`, el usuario queda precargado y podés
iniciar sesión con un clic. **La contraseña nunca se envía al navegador**: se
usa dentro del servidor. Si no las configurás, escribí usuario y clave en el
formulario.

---

## Uso de la web

1. Clic en **Iniciar sesión** (o escribí tus credenciales si no están en el
   servidor).
2. Pegá códigos en el cuadro (uno por línea, con comas o sin guiones), o usá
   **importar reexpediciones** con uno o varios PDF.
3. Clic en **Consultar documentos**. Los controlados salen en **rojo**.
4. **Descargar CSV** exporta el resultado.

### Importar PDF

Se buscan códigos que empiezan con `20126`. Los de 9 dígitos se completan con
ceros a la izquierda del último bloque (`201264772` → `201-26-004772`) y se
agregan al cuadro sin repetir. Los PDF escaneados se procesan con Tesseract OCR.
Límite por archivo: 20 MB.

---

## Estados

| Estado | Significado |
|---|---|
| `CONTROLADO` | La web indicó *"ya fue controlado en avanzada"* (ventana roja). |
| `NO EXISTE` | El código no existe en el sistema. |
| `NO CONTROLABLE` | Existe, pero *"no es controlable en avanzada"*. |
| `NO CONTROLADO` | La web no devolvió mensaje de error. |
| `FORMATO INVÁLIDO` | El código no cumple `NNN-NN-NNNNNN`. |

---

## Cómo funciona

El servidor consulta:

```
GET https://svepp.zofri.cl/index.php?tran&method=buscar_dt&doc=201-26-477270&mic=
```

con la sesión iniciada, y clasifica la respuesta JavaScript que devuelve el SVE.

Endpoints: `GET /`, `GET /healthz`, `GET /api/estado`, `GET /api/config`,
`POST /api/sesion`, `POST /api/consultar`, `POST /api/importar-pdf`.

---

## Seguridad

- No subas credenciales al repo: `.gitignore` excluye `.env` y el ejemplo está
  vacío. Usá los secretos del proveedor.
- La sesión es **global** (una sola cuenta a la vez): la herramienta está
  pensada para uso personal.
- Cualquier persona con la URL podría usar esa sesión. Protegé el acceso con
  la autenticación de tu proveedor (Cloudflare Access, login del hosting,
  etc.) si no querés que sea público.
- El hosting debe poder llegar a `svepp.zofri.cl`. Probá una consulta antes de
  compartir la URL; los planes gratuitos suelen dormir el servicio.

---

## Estructura

```
app.py            servidor HTTP y endpoints
index.html        interfaz (HTML/CSS/JS, sin build)
Dockerfile        imagen con Python + Tesseract
docker-compose.yml  despliegue local en Docker
render.yaml       blueprint para Render
Procfile          para plataformas tipo Heroku
requirements.txt  dependencias Python
```
