"""
╔══════════════════════════════════════════════════════════════════╗
║     AUTOMATIZADOR DE JUICIOS DE EVALUACIÓN — SERVIDOR WEB        ║
║                      SENA Colombia                               ║
╚══════════════════════════════════════════════════════════════════╝

Ejecutar:
    pip install flask selenium webdriver-manager pandas openpyxl xlrd reportlab
    python app.py

Acceder en: http://localhost:5000
"""

from __future__ import annotations

import glob
import io
import json
import zipfile
import os
import queue
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Optional

from flask import Flask, Response, jsonify, render_template, request, send_file

# ── Import backend ────────────────────────────────────────────────
import sofia_backend as bk
from pathlib import Path as _Path
import platform as _platform

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit

# ══════════════════════════════════════════════════════════════════
#  GESTIÓN DE TAREAS
# ══════════════════════════════════════════════════════════════════

_tasks: dict[str, dict] = {}   # task_id → {queue, status, archivo, error}
_task_lock = threading.Lock()
_TASK_TTL = 3600               # segundos antes de limpiar tarea antigua


def _new_task(task_id: str) -> dict:
    task = {
        "id":      task_id,
        "queue":   queue.Queue(),
        "status":  "running",   # running | done | error
        "archivo": None,        # ruta local del PDF generado
        "created": time.time(),
    }
    with _task_lock:
        _tasks[task_id] = task
    return task


def _get_task(task_id: str) -> Optional[dict]:
    with _task_lock:
        return _tasks.get(task_id)


def _put(task: dict, type_: str, **kwargs) -> None:
    """Encola un evento SSE."""
    task["queue"].put({"type": type_, **kwargs})


def _cleanup_old_tasks():
    """Elimina tareas más antiguas de TTL del diccionario."""
    now = time.time()
    with _task_lock:
        dead = [tid for tid, t in _tasks.items() if now - t["created"] > _TASK_TTL]
        for tid in dead:
            del _tasks[tid]


# ══════════════════════════════════════════════════════════════════
#  CONTEXTO DE LOG → SSE
# ══════════════════════════════════════════════════════════════════

@contextmanager
def _log_context(task: dict):
    """
    Redirige todos los logs del backend (ok/err/info/titulo)
    al queue SSE de la tarea mediante contextvars.
    """
    def _cb(level: str, msg: str):
        # Limpiar ANSI antes de enviar al browser
        import re
        clean = re.sub(r"\x1b\[[0-9;]*m", "", msg)
        _put(task, "log", level=level, msg=clean)

    token = bk._log_ctx.set(_cb)
    try:
        yield
    finally:
        bk._log_ctx.reset(token)


# ══════════════════════════════════════════════════════════════════
#  HELPERS DE PASOS
# ══════════════════════════════════════════════════════════════════

def _paso(task: dict, index: int, estado: str):
    """Emite evento de actualización de paso (en_curso | listo | error)."""
    _put(task, "paso", index=index, estado=estado)


def _fin(task: dict, exito: bool, msg: str = ""):
    task["status"] = "done" if exito else "error"
    _put(task, "fin", exito=exito, msg=msg)
    # Señal especial para que el generador SSE cierre
    _put(task, "_cerrar")


# ══════════════════════════════════════════════════════════════════
#  LÓGICA DE TAREAS EN HILO
# ══════════════════════════════════════════════════════════════════

PASOS_DESCARGA = [
    "Abriendo navegador", "Iniciando sesión", "Seleccionando rol",
    "Navegando al reporte", "Buscando ficha", "Descargando archivo",
]
PASOS_CONSULTA = ["Leyendo reporte", "Filtrando aprendices", "Generando documento"]
PASOS_COMPLETO = [
    "Abriendo navegador", "Iniciando sesión", "Seleccionando rol",
    "Navegando al reporte", "Buscando ficha", "Descargando archivo",
    "Generando informe",
]


def _run_descargar(task_id: str, usuario: str, contrasena: str, numero_ficha: str):
    task   = _get_task(task_id)
    driver = None
    _t0    = time.time()
    with _log_context(task):
        try:
            _paso(task, 0, "en_curso")
            driver = bk.configurar_driver()
            _paso(task, 0, "listo")

            _paso(task, 1, "en_curso")
            bk.iniciar_sesion(driver, usuario, contrasena)
            _paso(task, 1, "listo")

            _paso(task, 2, "en_curso")
            bk.seleccionar_rol_instructor(driver)
            _paso(task, 2, "listo")

            _paso(task, 3, "en_curso")
            bk.navegar_reporte_juicios(driver)
            _paso(task, 3, "listo")

            _paso(task, 4, "en_curso")
            mtime_ref = [0.0]
            bk.buscar_ficha_y_aprendiz(driver, numero_ficha, mtime_ref)
            _paso(task, 4, "listo")

            _paso(task, 5, "en_curso")
            ruta = bk.descargar_excel(numero_ficha, mtime_ref[0],
                                      t_inicio=_t0,
                                      carpeta_tmp=getattr(driver, "_sofia_tmp_dir", ""))
            if ruta:
                task["archivo"] = ruta
                _paso(task, 5, "listo")
                _put(task, "archivo", nombre=os.path.basename(ruta))
                _fin(task, True, f"Reporte descargado: {os.path.basename(ruta)}")
            else:
                _paso(task, 5, "error")
                _fin(task, False, "No se encontró el archivo descargado.")

        except Exception as ex:
            _paso(task, -1, "error")
            _fin(task, False, _mensaje_error(ex))
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass


def _run_consultar(
    task_id: str, numero_ficha: str, nombre: str, apellido: str,
    tipo: str, solo_activos: bool,
):
    task = _get_task(task_id)
    with _log_context(task):
        try:
            _paso(task, 0, "en_curso")
            ruta_excel = bk.buscar_excel_ficha(numero_ficha)
            if not ruta_excel:
                _paso(task, 0, "error")
                _fin(task, False,
                     f"No hay reporte descargado para la ficha {numero_ficha}. "
                     "Descárgalo primero desde la pestaña 'Descargar'.")
                return
            _paso(task, 0, "listo")

            _paso(task, 1, "en_curso")
            # el filtrado ocurre dentro de analizar_*
            _paso(task, 1, "listo")

            _paso(task, 2, "en_curso")
            if tipo == "independiente":
                carpeta = bk.analizar_juicios_independientes(
                    ruta_excel, nombre, apellido, numero_ficha,
                    solo_activos=solo_activos,
                )
                task["archivo"] = carpeta
                _paso(task, 2, "listo")
                _put(task, "archivo", nombre=os.path.basename(carpeta) + ".zip", es_zip=True)
                _fin(task, True, "Informes independientes generados.")
            else:
                ruta_doc = bk.analizar_juicios(
                    ruta_excel, nombre, apellido, solo_activos=solo_activos
                )
                if ruta_doc:
                    task["archivo"] = ruta_doc
                    _paso(task, 2, "listo")
                    _put(task, "archivo", nombre=os.path.basename(ruta_doc))
                    _fin(task, True, f"Informe generado: {os.path.basename(ruta_doc)}")
                else:
                    _paso(task, 2, "error")
                    _fin(task, False, "No se generó el informe. Revisa el log.")

        except Exception as ex:
            _paso(task, -1, "error")
            _fin(task, False, _mensaje_error(ex))


def _run_completo(
    task_id: str, usuario: str, contrasena: str, numero_ficha: str,
    nombre: str, apellido: str, tipo: str, solo_activos: bool,
):
    task   = _get_task(task_id)
    driver = None
    _t0    = time.time()
    with _log_context(task):
        try:
            _paso(task, 0, "en_curso")
            driver = bk.configurar_driver()
            _paso(task, 0, "listo")

            _paso(task, 1, "en_curso")
            bk.iniciar_sesion(driver, usuario, contrasena)
            _paso(task, 1, "listo")

            _paso(task, 2, "en_curso")
            bk.seleccionar_rol_instructor(driver)
            _paso(task, 2, "listo")

            _paso(task, 3, "en_curso")
            bk.navegar_reporte_juicios(driver)
            _paso(task, 3, "listo")

            _paso(task, 4, "en_curso")
            mtime_ref = [0.0]
            bk.buscar_ficha_y_aprendiz(driver, numero_ficha, mtime_ref)
            _paso(task, 4, "listo")

            _paso(task, 5, "en_curso")
            ruta_excel = bk.descargar_excel(numero_ficha, mtime_ref[0],
                                             t_inicio=_t0,
                                             carpeta_tmp=getattr(driver, "_sofia_tmp_dir", ""))
            if not ruta_excel:
                _paso(task, 5, "error")
                _fin(task, False, "No se encontró el archivo descargado.")
                return
            _paso(task, 5, "listo")

            _paso(task, 6, "en_curso")
            if tipo == "independiente":
                carpeta = bk.analizar_juicios_independientes(
                    ruta_excel, nombre, apellido, numero_ficha,
                    solo_activos=solo_activos,
                )
                task["archivo"] = carpeta
                _put(task, "archivo", nombre=os.path.basename(carpeta) + ".zip", es_zip=True)
                _paso(task, 6, "listo")
                _fin(task, True, "Proceso completo — informes independientes generados.")
            else:
                ruta_doc = bk.analizar_juicios(
                    ruta_excel, nombre, apellido, solo_activos=solo_activos
                )
                if ruta_doc:
                    task["archivo"] = ruta_doc
                    _put(task, "archivo", nombre=os.path.basename(ruta_doc))
                    _paso(task, 6, "listo")
                    _fin(task, True, f"Proceso completo — {os.path.basename(ruta_doc)}")
                else:
                    _paso(task, 6, "error")
                    _fin(task, False, "No se generó el informe. Revisa el log.")

        except Exception as ex:
            _paso(task, -1, "error")
            _fin(task, False, _mensaje_error(ex))
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass


def _mensaje_error(ex: Exception) -> str:
    txt = str(ex).lower()
    if "portal_no_disponible" in txt:
        return (
            "⚠️ El portal Sofía Plus no se encuentra disponible en este momento. "
            "Por favor intenta de nuevo más tarde."
        )
    if "credenciales_invalidas" in txt:
        return "Usuario o contraseña incorrectos. Verifica tus credenciales de Sofía Plus."
    if "formulario_no_encontrado" in txt:
        return (
            "⚠️ El portal Sofía Plus no se encuentra disponible en este momento. "
            "Por favor intenta de nuevo más tarde."
        )
    if "navegador_no_encontrado" in txt:
        return ("No se encontró Chrome ni Edge instalado. "
                "Instala uno de los dos navegadores para continuar.")
    if "ficha_no_encontrada" in txt:
        return f"La ficha no fue encontrada en Sofía Plus. Verifica el número."
    if "popup_no_abierto" in txt:
        return "No se pudo abrir el popup de fichas. Intenta de nuevo."
    if "consultar_fallido" in txt:
        return "No se pudo hacer clic en 'Consultar'. Intenta de nuevo."
    if "read timed out" in txt or "readtimeouterror" in txt:
        print(f"[DEBUG timeout navegador] {str(ex)}")
        return (
            "⚠️ El portal Sofía Plus tardó demasiado en responder (más de 2 minutos) "
            "y el navegador se quedó esperando. Puede ser el portal, o la conexión "
            "del servidor hacia SENA. Intenta de nuevo en unos minutos."
        )
    if "chrome" in txt or "driver" in txt or "webdriver" in txt:
        print(f"[DEBUG navegador] {str(ex)}")
        return "Problema con el navegador. Asegúrate de tener Chrome o Edge instalado y actualizado."
    return f"Error inesperado: {str(ex)[:200]}"


# ══════════════════════════════════════════════════════════════════
#  RUTAS API
# ══════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/descargar", methods=["POST"])
def api_descargar():
    data         = request.get_json(force=True)
    usuario      = data.get("usuario", "").strip()
    contrasena   = data.get("contrasena", "").strip()
    numero_ficha = data.get("numero_ficha", "").strip()

    if not all([usuario, contrasena, numero_ficha]):
        return jsonify(error="Faltan campos obligatorios"), 400

    task_id = str(uuid.uuid4())
    _new_task(task_id)
    # Enviar definición de pasos al cliente junto con el task_id
    t = threading.Thread(
        target=_run_descargar,
        args=(task_id, usuario, contrasena, numero_ficha),
        daemon=True,
    )
    t.start()
    return jsonify(task_id=task_id, pasos=PASOS_DESCARGA)


@app.route("/api/consultar", methods=["POST"])
def api_consultar():
    data         = request.get_json(force=True)
    numero_ficha = data.get("numero_ficha", "").strip()
    nombre       = data.get("nombre", "").strip()
    apellido     = data.get("apellido", "").strip()
    tipo         = data.get("tipo", "consolidado")
    solo_activos = bool(data.get("solo_activos", False))

    if not numero_ficha:
        return jsonify(error="Número de ficha requerido"), 400

    task_id = str(uuid.uuid4())
    _new_task(task_id)
    t = threading.Thread(
        target=_run_consultar,
        args=(task_id, numero_ficha, nombre, apellido, tipo, solo_activos),
        daemon=True,
    )
    t.start()
    return jsonify(task_id=task_id, pasos=PASOS_CONSULTA)


@app.route("/api/completo", methods=["POST"])
def api_completo():
    data         = request.get_json(force=True)
    usuario      = data.get("usuario", "").strip()
    contrasena   = data.get("contrasena", "").strip()
    numero_ficha = data.get("numero_ficha", "").strip()
    nombre       = data.get("nombre", "").strip()
    apellido     = data.get("apellido", "").strip()
    tipo         = data.get("tipo", "consolidado")
    solo_activos = bool(data.get("solo_activos", False))

    if not all([usuario, contrasena, numero_ficha]):
        return jsonify(error="Faltan campos obligatorios"), 400

    task_id = str(uuid.uuid4())
    _new_task(task_id)
    t = threading.Thread(
        target=_run_completo,
        args=(task_id, usuario, contrasena, numero_ficha,
              nombre, apellido, tipo, solo_activos),
        daemon=True,
    )
    t.start()
    return jsonify(task_id=task_id, pasos=PASOS_COMPLETO)


@app.route("/api/stream/<task_id>")
def api_stream(task_id: str):
    """Server-Sent Events — transmite el progreso de la tarea en tiempo real."""
    task = _get_task(task_id)
    if not task:
        return Response("Task not found", status=404)

    def _generate():
        q = task["queue"]
        while True:
            try:
                event = q.get(timeout=30)
            except queue.Empty:
                yield "event: ping\ndata: {}\n\n"
                continue

            if event.get("type") == "_cerrar":
                break

            payload = json.dumps(event, ensure_ascii=False)
            yield f"data: {payload}\n\n"

        # Enviar evento final y cerrar stream
        yield f"data: {json.dumps({'type': '_eof'})}\n\n"

    return Response(
        _generate(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control":  "no-cache",
            "X-Accel-Buffering": "no",  # Nginx: deshabilitar buffering
        },
    )


@app.route("/api/descargar-archivo/<task_id>")
def api_descargar_archivo(task_id: str):
    """Sirve el PDF (o ZIP de carpeta) generado para descarga en el browser."""
    task = _get_task(task_id)
    if not task or not task.get("archivo"):
        return jsonify(error="Archivo no disponible"), 404

    ruta = task["archivo"]

    # ── Carpeta de independientes → empaquetar en ZIP en memoria ─
    if os.path.isdir(ruta):
        nombre_zip = os.path.basename(ruta) + ".zip"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for raiz, _, archivos in os.walk(ruta):
                for archivo in archivos:
                    ruta_abs = os.path.join(raiz, archivo)
                    ruta_rel = os.path.relpath(ruta_abs, ruta)
                    zf.write(ruta_abs, ruta_rel)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=nombre_zip,
            mimetype="application/zip",
        )

    # ── Archivo único (PDF) ──────────────────────────────────────
    if not os.path.isfile(ruta):
        return jsonify(error="Archivo no encontrado en disco"), 404

    ext = os.path.splitext(ruta)[1].lower()
    mimetype = (
        "application/pdf"
        if ext == ".pdf" else "application/octet-stream"
    )
    return send_file(
        ruta,
        as_attachment=True,
        download_name=os.path.basename(ruta),
        mimetype=mimetype,
    )


@app.route("/api/dashboard/<numero_ficha>")
def api_dashboard(numero_ficha: str):
    """Devuelve los datos calculados del dashboard para una ficha."""
    datos = bk.generar_datos_dashboard(numero_ficha.strip())
    if "error" in datos:
        return jsonify(datos), 404
    return jsonify(datos)


@app.route("/api/consulta_ra/<numero_ficha>")
def api_consulta_ra(numero_ficha: str):
    """Devuelve competencias y RAs con datos de evaluación para consulta interactiva."""
    solo_activos = request.args.get("solo_activos", "false").lower() == "true"
    datos = bk.generar_consulta_ra(numero_ficha.strip(), solo_activos=solo_activos)
    if "error" in datos:
        return jsonify(datos), 404
    return jsonify(datos)


@app.route("/api/estado")
def api_estado():
    """Ping para verificar que el servidor está activo."""
    _cleanup_old_tasks()
    return jsonify(ok=True, version="1.0.0")


@app.route("/api/carpeta")
def api_carpeta():
    return jsonify(carpeta=str(bk.cfg.carpeta_descarga))


@app.route("/api/carpeta", methods=["POST"])
def api_carpeta_set():
    """Actualiza la carpeta de trabajo y la persiste en disco."""
    data  = request.get_json(force=True)
    nueva = data.get("carpeta", "").strip()

    if not nueva:
        return jsonify(error="Ruta vacía"), 400

    # Intentar crear la carpeta si no existe
    try:
        os.makedirs(nueva, exist_ok=True)
    except Exception as ex:
        return jsonify(error=f"No se pudo crear la carpeta: {ex}"), 400

    if not os.path.isdir(nueva):
        return jsonify(error="La ruta no es una carpeta válida"), 400

    bk.cfg.carpeta_descarga = nueva
    bk.guardar_carpeta(nueva)
    return jsonify(ok=True, carpeta=nueva)


@app.route("/api/explorar")
def api_explorar():
    """
    Lista el contenido de una carpeta para el selector de directorios.
    Parámetro: ?ruta=C:/Users/...
    """
    ruta = request.args.get("ruta", "").strip()

    # Sin ruta → devolver discos/home según OS
    if not ruta:
        import platform
        if platform.system() == "Windows":
            import string
            drives = [f"{d}:/" for d in string.ascii_uppercase
                      if os.path.exists(f"{d}:/")]
            return jsonify(tipo="raiz", elementos=drives)
        else:
            ruta = str(str(_Path.home()))

    if not os.path.isdir(ruta):
        return jsonify(error="Ruta no válida"), 400

    try:
        elementos = []
        # Padre
        padre = str(_Path(ruta).parent)
        if padre != ruta:
            elementos.append({"nombre": ".. (subir)", "ruta": padre, "tipo": "padre"})

        # Subcarpetas (ordenadas)
        for entrada in sorted(os.scandir(ruta), key=lambda e: e.name.lower()):
            if entrada.is_dir(follow_symlinks=False) and not entrada.name.startswith("."):
                elementos.append({
                    "nombre": entrada.name,
                    "ruta":   entrada.path,
                    "tipo":   "carpeta",
                })

        return jsonify(tipo="carpeta", ruta_actual=ruta, elementos=elementos)
    except PermissionError:
        return jsonify(error="Sin permiso para leer esta carpeta"), 403


# ══════════════════════════════════════════════════════════════════
#  PUNTO DE ENTRADA
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = os.environ.get("HOST", "127.0.0.1")
    print(f"\n  🌿 Sofía Plus — Servidor web corriendo en http://{host}:{port}\n")
    # use_reloader=False importante: evita que el driver se inicie dos veces
    app.run(host=host, port=port, debug=False, use_reloader=False, threaded=True)
