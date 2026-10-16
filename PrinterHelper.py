"""
Denario Impresora (PrinterHelper)
=================================

Servidor local que permite a la web app de Denario (Flutter web) imprimir
comandas y facturas en impresoras térmicas conectadas a Windows.

Comportamiento al distribuirse como .exe (PyInstaller --onefile):
  * Primer doble clic (desde Descargas u otra carpeta): se auto-instala en
    %LOCALAPPDATA%\\Programs\\Denario Impresora, se registra para iniciar con
    Windows, crea acceso directo en el menú Inicio, registra el protocolo
    denario-printer:// y una entrada en "Agregar o quitar programas".
  * Siguientes inicios: arranca solo al iniciar sesión, sin ventanas.
  * Si ya está corriendo y se vuelve a abrir, abre la página de configuración.

Argumentos de línea de comandos:
  --autostart   Lanzado por Windows al iniciar sesión (sin notificaciones).
  --first-run   Primer arranque tras la instalación (abre la configuración).
  --portable    No auto-instalar; ejecutar desde la ubicación actual.
  --uninstall   Desinstalar (usado por "Agregar o quitar programas").
  --quiet       Con --uninstall: sin cuadros de diálogo.
  denario-printer://...  Lanzado desde la web app vía protocolo.

API HTTP (http://localhost:8000) — compatible con la versión 1.x:
  GET  /                 Página de configuración.
  GET  /status           Estado del helper.
  GET  /printers         Lista de impresoras.
  POST /print            Imprime datos RAW (ESC/POS). ?printer=NOMBRE opcional.
  POST /select_printer   Guarda impresora predeterminada (form o JSON).
  POST /test_print       Imprime una página de prueba.
  POST /settings/autostart  {"enabled": bool}
"""

import ctypes
import json
import logging
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
import winreg
from logging.handlers import RotatingFileHandler

import pywintypes
import win32api
import win32event
import win32print
import winerror
from flask import Flask, jsonify, redirect, render_template_string, request
from PIL import Image, ImageDraw
from pystray import Icon, Menu, MenuItem
from waitress import serve

# ----------------------------
# CONSTANTES
# ----------------------------
APP_NAME = "Denario Impresora"
APP_ID = "DenarioPrinter"
VERSION = "2.0.0"
PUBLISHER = "Denario"

HOST = "127.0.0.1"
PORT = 8000
SETTINGS_URL = f"http://localhost:{PORT}/"
PROTOCOL = "denario-printer"

# Orígenes web autorizados a usar el helper (CORS + verificación de Origin).
# Soporta todos los dominios de Denario (web.app, info, firebaseapp) y entornos locales.
ALLOWED_ORIGIN_PATTERNS = [
    re.compile(r"^https?://([a-zA-Z0-9-]+\.)*(denario\.(web\.app|info|firebaseapp\.com)|mi-denario\.web\.app)(:\d+)?$", re.I),
    re.compile(r"^https?://(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$", re.I),
]
# Protección contra DNS rebinding: sólo aceptamos estos Host.
ALLOWED_HOSTS = {f"localhost:{PORT}", f"127.0.0.1:{PORT}"}

IS_FROZEN = getattr(sys, "frozen", False)

_APPDATA = os.environ.get("APPDATA") or os.path.expanduser("~")
_LOCALAPPDATA = os.environ.get("LOCALAPPDATA") or _APPDATA
DATA_DIR = os.path.join(_APPDATA, "Denario", "PrinterHelper")
CONFIG_FILE = os.path.join(DATA_DIR, "printer_config.json")
LOG_DIR = os.path.join(DATA_DIR, "logs")
LEGACY_CONFIG_NAME = "printer_config.json"

INSTALL_DIR = os.path.join(_LOCALAPPDATA, "Programs", APP_NAME)
INSTALLED_EXE = os.path.join(INSTALL_DIR, f"{APP_ID}.exe")
START_MENU_LNK = os.path.join(
    _APPDATA, "Microsoft", "Windows", "Start Menu", "Programs", f"{APP_NAME}.lnk"
)

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
STARTUP_APPROVED_KEY = r"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
UNINSTALL_KEY = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_ID}"
PROTOCOL_KEY = rf"Software\Classes\{PROTOCOL}"

MUTEX_NAME = "Local\\DenarioPrinterHelper.Instance"
QUIT_EVENT_NAME = "Local\\DenarioPrinterHelper.Quit"

# Win32
SYNCHRONIZE = 0x00100000
EVENT_MODIFY_STATE = 0x0002
DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200
CREATE_NO_WINDOW = 0x08000000
MB_OK, MB_YESNO, IDYES = 0x0, 0x4, 6
MB_ICONERROR, MB_ICONQUESTION, MB_ICONINFO = 0x10, 0x20, 0x40
MB_SETFOREGROUND = 0x10000

log = logging.getLogger("denario")
_tray_icon = None
_instance_mutex = None


# ----------------------------
# UTILIDADES GENERALES
# ----------------------------
def setup_logging():
    """Log rotativo en %APPDATA% (con console=False los print() se pierden)."""
    os.makedirs(LOG_DIR, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    file_handler = RotatingFileHandler(
        os.path.join(LOG_DIR, "helper.log"), maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(file_handler)
    if not IS_FROZEN:
        console = logging.StreamHandler()
        console.setFormatter(fmt)
        root.addHandler(console)


def resource_path(name):
    """Ruta a un recurso empaquetado (funciona en desarrollo y en el .exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def message_box(text, flags=MB_OK | MB_ICONINFO):
    return ctypes.windll.user32.MessageBoxW(None, text, APP_NAME, flags | MB_SETFOREGROUND)


def launch_detached(exe, *args):
    """Lanza un proceso independiente (sobrevive a la salida del actual)."""
    env = os.environ.copy()
    # Evita que un .exe onefile hijo reutilice la carpeta temporal del padre.
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    subprocess.Popen(
        [exe, *args],
        env=env,
        cwd=os.path.dirname(exe),
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
    )


# ----------------------------
# CONFIGURACIÓN
# ----------------------------
_config_lock = threading.RLock()
DEFAULT_CONFIG = {"default_printer": None, "autostart": True}


def load_config():
    with _config_lock:
        cfg = dict(DEFAULT_CONFIG)
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg.update(json.load(f))
        except Exception:
            log.exception("No se pudo leer la configuración")
        return cfg


def save_config(**updates):
    """Guarda la configuración de forma atómica (evita archivos corruptos)."""
    with _config_lock:
        cfg = load_config()
        cfg.update(updates)
        os.makedirs(DATA_DIR, exist_ok=True)
        tmp = CONFIG_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        os.replace(tmp, CONFIG_FILE)
        return cfg


def migrate_legacy_config():
    """Importa la impresora guardada por la v1 (printer_config.json relativo)."""
    if os.path.exists(CONFIG_FILE):
        return
    exe_dir = os.path.dirname(sys.executable if IS_FROZEN else os.path.abspath(__file__))
    for folder in (os.getcwd(), exe_dir, os.path.join(os.path.expanduser("~"), "Downloads")):
        path = os.path.join(folder, LEGACY_CONFIG_NAME)
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    printer = json.load(f).get("default_printer")
                if printer:
                    save_config(default_printer=printer)
                    log.info("Configuración v1 migrada desde %s (%s)", path, printer)
                    return
        except Exception:
            log.exception("No se pudo migrar %s", path)


def load_selected_printer():
    return load_config().get("default_printer")


def save_selected_printer(printer_name):
    save_config(default_printer=printer_name)


# ----------------------------
# IMPRESORAS
# ----------------------------
def list_windows_printers():
    """Impresoras locales + conexiones de red del usuario, sin duplicados."""
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        names = [p[2] for p in win32print.EnumPrinters(flags)]
        return list(dict.fromkeys(names))
    except Exception:
        log.exception("Error al listar impresoras")
        return []


def send_raw_to_printer(printer_name, data, job_name="Denario"):
    """Envía bytes RAW (ESC/POS). Siempre libera el handle, incluso si falla."""
    handle = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(handle, 1, (job_name, None, "RAW"))
        try:
            win32print.StartPagePrinter(handle)
            try:
                return win32print.WritePrinter(handle, data)
            finally:
                win32print.EndPagePrinter(handle)
        finally:
            win32print.EndDocPrinter(handle)
    finally:
        win32print.ClosePrinter(handle)


def build_test_ticket(printer_name):
    esc, gs = b"\x1b", b"\x1d"

    def txt(s):
        return s.encode("cp437", "replace")

    return b"".join([
        esc + b"@",                        # reset
        esc + b"a\x01",                    # centrado
        esc + b"!\x30", b"DENARIO\n",      # doble alto/ancho
        esc + b"!\x00", b"Pagina de prueba\n",
        b"-" * 32 + b"\n",
        esc + b"a\x00",                    # izquierda
        txt(f"Impresora: {printer_name}\n"),
        txt(f"Fecha: {time.strftime('%d/%m/%Y %H:%M')}\n"),
        txt(f"Version: {VERSION}\n"),
        b"-" * 32 + b"\n",
        esc + b"a\x01",
        b"Si podes leer esto,\nla impresora funciona!\n",
        b"\n\n\n\n",
        gs + b"V\x42\x00",                 # corte parcial
    ])


# ----------------------------
# REGISTRO DE WINDOWS (HKCU, sin admin)
# ----------------------------
def _reg_set(path, name, value, kind=winreg.REG_SZ):
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, name, 0, kind, value)


def _reg_get(path, name):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            return winreg.QueryValueEx(key, name)[0]
    except OSError:
        return None


def _reg_delete_value(path, name):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, name)
    except OSError:
        pass


def _reg_delete_tree(path):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_ALL_ACCESS) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                _reg_delete_tree(f"{path}\\{child}")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except OSError:
        pass


def set_autostart(enabled, exe=None):
    exe = exe or (sys.executable if IS_FROZEN else None)
    if not exe:
        log.info("Autoarranque omitido (modo desarrollo)")
        return
    if enabled:
        _reg_set(RUN_KEY, APP_ID, f'"{exe}" --autostart')
    else:
        _reg_delete_value(RUN_KEY, APP_ID)


def _disabled_in_task_manager():
    """True si el usuario lo deshabilitó en Administrador de tareas > Inicio."""
    data = _reg_get(STARTUP_APPROVED_KEY, APP_ID)
    return bool(data) and data[0] % 2 == 1


def is_autostart_enabled():
    return bool(_reg_get(RUN_KEY, APP_ID)) and not _disabled_in_task_manager()


def toggle_autostart(enabled):
    save_config(autostart=enabled)
    set_autostart(enabled)
    if enabled:
        # Si el usuario lo había deshabilitado en el Administrador de tareas, lo reactivamos.
        _reg_delete_value(STARTUP_APPROVED_KEY, APP_ID)
    log.info("Autoarranque %s", "activado" if enabled else "desactivado")


def register_protocol(exe):
    """denario-printer:// permite que la web app lance el helper si está cerrado."""
    _reg_set(PROTOCOL_KEY, "", f"URL:{APP_NAME}")
    _reg_set(PROTOCOL_KEY, "URL Protocol", "")
    _reg_set(PROTOCOL_KEY + r"\DefaultIcon", "", f'"{exe}",0')
    _reg_set(PROTOCOL_KEY + r"\shell\open\command", "", f'"{exe}" "%1"')


def register_uninstall_entry(exe):
    _reg_set(UNINSTALL_KEY, "DisplayName", APP_NAME)
    _reg_set(UNINSTALL_KEY, "DisplayVersion", VERSION)
    _reg_set(UNINSTALL_KEY, "Publisher", PUBLISHER)
    _reg_set(UNINSTALL_KEY, "DisplayIcon", f'"{exe}",0')
    _reg_set(UNINSTALL_KEY, "InstallLocation", INSTALL_DIR)
    _reg_set(UNINSTALL_KEY, "UninstallString", f'"{exe}" --uninstall')
    _reg_set(UNINSTALL_KEY, "QuietUninstallString", f'"{exe}" --uninstall --quiet')
    _reg_set(UNINSTALL_KEY, "NoModify", 1, winreg.REG_DWORD)
    _reg_set(UNINSTALL_KEY, "NoRepair", 1, winreg.REG_DWORD)
    try:
        size_kb = os.path.getsize(exe) // 1024
        _reg_set(UNINSTALL_KEY, "EstimatedSize", size_kb, winreg.REG_DWORD)
    except OSError:
        pass


def create_start_menu_shortcut(exe):
    """Acceso directo en el menú Inicio para que aparezca al buscar en Windows."""
    try:
        import win32com.client

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(START_MENU_LNK)
        shortcut.TargetPath = exe
        shortcut.WorkingDirectory = os.path.dirname(exe)
        shortcut.IconLocation = f"{exe},0"
        shortcut.Description = "Imprimir comandas y facturas desde Denario"
        shortcut.save()
    except Exception:
        log.exception("No se pudo crear el acceso directo del menú Inicio")


# ----------------------------
# INSTANCIA ÚNICA
# ----------------------------
def acquire_single_instance():
    global _instance_mutex
    _instance_mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
    return win32api.GetLastError() != winerror.ERROR_ALREADY_EXISTS


def is_instance_running():
    try:
        handle = win32event.OpenMutex(SYNCHRONIZE, False, MUTEX_NAME)
        win32api.CloseHandle(handle)
        return True
    except pywintypes.error:
        return False


def listen_for_quit_signal():
    """Permite que una nueva versión (o el desinstalador) cierre esta instancia."""
    event = win32event.CreateEvent(None, True, False, QUIT_EVENT_NAME)
    win32event.WaitForSingleObject(event, win32event.INFINITE)
    log.info("Señal de cierre recibida")
    shutdown()


def stop_running_instance(timeout=8.0):
    """Cierra la instancia v2 en ejecución (vía evento) y, si hace falta, una v1."""
    try:
        event = win32event.OpenEvent(EVENT_MODIFY_STATE, False, QUIT_EVENT_NAME)
        win32event.SetEvent(event)
        win32api.CloseHandle(event)
    except pywintypes.error:
        pass

    deadline = time.time() + timeout
    while is_instance_running() and time.time() < deadline:
        time.sleep(0.2)

    if is_helper_listening():
        kill_legacy_helper()


def is_helper_listening():
    try:
        with urllib.request.urlopen(f"http://{HOST}:{PORT}/status", timeout=1.5) as r:
            return json.load(r).get("status") == "running"
    except Exception:
        return False


def _pid_listening_on_port():
    try:
        out = subprocess.run(
            ["netstat", "-ano", "-p", "TCP"], capture_output=True, text=True,
            creationflags=CREATE_NO_WINDOW, timeout=10,
        ).stdout
    except Exception:
        return None
    for line in out.splitlines():
        parts = line.split()
        if len(parts) >= 5 and parts[1].endswith(f":{PORT}") and parts[3].upper() in ("LISTENING", "ESCUCHANDO"):
            return int(parts[4])
    return None


def kill_legacy_helper():
    """La v1 no escucha la señal de cierre: la identificamos por el puerto."""
    pid = _pid_listening_on_port()
    if not pid or pid in (os.getpid(), os.getppid()):
        return
    try:
        info = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"], capture_output=True,
            text=True, creationflags=CREATE_NO_WINDOW, timeout=10,
        ).stdout.lower()
        if "denario" not in info and "printerhelper" not in info:
            log.warning("El puerto %s lo usa otro programa (PID %s); no se cierra", PORT, pid)
            return
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True,
                       creationflags=CREATE_NO_WINDOW, timeout=10)
        log.info("Versión anterior cerrada (PID %s)", pid)
        time.sleep(1.0)
    except Exception:
        log.exception("No se pudo cerrar la versión anterior")


# ----------------------------
# INSTALACIÓN / DESINSTALACIÓN
# ----------------------------
def is_running_installed_copy():
    return os.path.normcase(os.path.abspath(sys.executable)) == os.path.normcase(INSTALLED_EXE)


def self_install():
    """Copia el .exe a una ubicación fija y registra autoarranque, protocolo, etc."""
    log.info("Instalando desde %s", sys.executable)
    migrate_legacy_config()  # antes de cerrar la v1, mientras sabemos dónde estaba
    stop_running_instance()

    os.makedirs(INSTALL_DIR, exist_ok=True)
    for attempt in range(30):  # el exe anterior puede tardar en liberar el archivo
        try:
            shutil.copy2(sys.executable, INSTALLED_EXE)
            break
        except OSError:
            if attempt == 29:
                raise
            time.sleep(0.5)

    register_protocol(INSTALLED_EXE)
    register_uninstall_entry(INSTALLED_EXE)
    create_start_menu_shortcut(INSTALLED_EXE)
    if load_config().get("autostart", True):
        set_autostart(True, INSTALLED_EXE)

    launch_detached(INSTALLED_EXE, "--first-run")
    log.info("Instalación completa en %s", INSTALLED_EXE)


def uninstall(quiet=False):
    if not quiet:
        answer = message_box(
            f"¿Desinstalar {APP_NAME}?\n\nDenario ya no podrá imprimir en esta computadora.",
            MB_YESNO | MB_ICONQUESTION,
        )
        if answer != IDYES:
            return

    stop_running_instance()
    _reg_delete_value(RUN_KEY, APP_ID)
    _reg_delete_value(STARTUP_APPROVED_KEY, APP_ID)
    _reg_delete_tree(PROTOCOL_KEY)
    _reg_delete_tree(UNINSTALL_KEY)
    try:
        os.remove(START_MENU_LNK)
    except OSError:
        pass

    # El exe en ejecución no puede borrarse a sí mismo: lo hace cmd tras unos segundos.
    if os.path.isdir(INSTALL_DIR):
        subprocess.Popen(
            f'cmd /c ping 127.0.0.1 -n 4 >nul & rmdir /s /q "{INSTALL_DIR}"',
            cwd=os.environ.get("TEMP", _LOCALAPPDATA),
            creationflags=CREATE_NO_WINDOW,
        )
    log.info("Desinstalado (la configuración se conserva en %s)", DATA_DIR)
    if not quiet:
        message_box(f"{APP_NAME} fue desinstalada.")


# ----------------------------
# SERVIDOR HTTP
# ----------------------------
app = Flask(__name__)


def is_allowed_origin(origin):
    if not origin:
        return True
    return any(p.match(origin) for p in ALLOWED_ORIGIN_PATTERNS)


@app.before_request
def security_guard():
    if (request.host or "").lower() not in ALLOWED_HOSTS:
        return jsonify({"error": "Host no permitido"}), 403
    # CORS sólo impide LEER respuestas; un POST "simple" de cualquier web igual llegaría.
    # Por eso rechazamos explícitamente orígenes desconocidos.
    origin = request.headers.get("Origin")
    if origin and not is_allowed_origin(origin):
        log.warning("Solicitud bloqueada desde origen no autorizado: %s", origin)
        return jsonify({"error": "Origen no autorizado"}), 403
    if request.method == "OPTIONS":
        return app.make_default_options_response()


@app.after_request
def add_cors_headers(response):
    origin = request.headers.get("Origin")
    if origin and is_allowed_origin(origin):
        response.headers["Access-Control-Allow-Origin"] = origin
        response.headers.add("Vary", "Origin")
        response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = (
            request.headers.get("Access-Control-Request-Headers") or "Content-Type"
        )
        response.headers["Access-Control-Max-Age"] = "600"
        # Chrome/Edge: acceso de una web pública a la red local (localhost).
        if request.headers.get("Access-Control-Request-Private-Network") == "true":
            response.headers["Access-Control-Allow-Private-Network"] = "true"
    return response


def _json_body():
    return request.get_json(silent=True) or {}


VIRTUAL_PRINTER_RE = re.compile(r"onenote|pdf|xps|fax|anydesk|send to|enviar a", re.I)


@app.route("/")
def index():
    printers = list_windows_printers()
    current = load_selected_printer()
    # Sin impresora guardada: sugerimos la primera que no sea virtual (PDF, OneNote...).
    suggested = current or next((p for p in printers if not VIRTUAL_PRINTER_RE.search(p)), None)
    return render_template_string(
        SETTINGS_HTML,
        printers=printers,
        current=current,
        selected=suggested,
        current_available=current in printers,
        autostart=is_autostart_enabled(),
        can_autostart=IS_FROZEN,
        version=VERSION,
        saved=request.args.get("saved") == "1",
    )


@app.route("/select_printer", methods=["POST"])
def select_printer():
    is_form = bool(request.form)
    selected = request.form.get("printer_name") or _json_body().get("printer_name")
    if not selected:
        return jsonify({"error": "No se seleccionó ninguna impresora"}), 400
    printers = list_windows_printers()
    if printers and selected not in printers:
        return jsonify({"error": f"La impresora '{selected}' no existe en este equipo"}), 400
    save_selected_printer(selected)
    log.info("Impresora predeterminada: %s", selected)
    if is_form:
        return redirect("/?saved=1", code=303)
    return jsonify({"status": "ok", "default_printer": selected})


@app.route("/printers", methods=["GET"])
def list_printers_api():
    return jsonify({"printers": list_windows_printers(), "default_printer": load_selected_printer()})


@app.route("/print", methods=["POST"])
def print_data():
    """Imprime datos RAW. ?printer=NOMBRE sobrescribe la impresora predeterminada."""
    printer_name = request.args.get("printer") or load_selected_printer()
    if not printer_name:
        return jsonify({
            "error": "No se ha especificado ninguna impresora.",
            "details": "Configurá una impresora predeterminada o agregá ?printer=NOMBRE al URL.",
        }), 400

    data = request.get_data()
    if not data:
        return jsonify({"error": "No se han recibido datos para imprimir."}), 400

    job_name = request.args.get("job") or "Denario"
    try:
        send_raw_to_printer(printer_name, data, job_name)
        log.info("Impreso en '%s' (%d bytes)", printer_name, len(data))
        return jsonify({
            "status": "Printed successfully",
            "target_printer": printer_name,
            "data_size_bytes": len(data),
        })
    except Exception as e:
        log.exception("Falló la impresión en '%s'", printer_name)
        return jsonify({"error": f"No se pudo imprimir con '{printer_name}'. Detalles: {e}"}), 500


@app.route("/test_print", methods=["POST"])
def test_print():
    printer_name = (
        request.args.get("printer") or _json_body().get("printer") or load_selected_printer()
    )
    if not printer_name:
        return jsonify({"error": "Primero elegí una impresora."}), 400
    try:
        send_raw_to_printer(printer_name, build_test_ticket(printer_name), "Denario - Prueba")
        log.info("Página de prueba enviada a '%s'", printer_name)
        return jsonify({"status": "ok", "target_printer": printer_name})
    except Exception as e:
        log.exception("Falló la página de prueba en '%s'", printer_name)
        return jsonify({"error": f"No se pudo imprimir con '{printer_name}'. Detalles: {e}"}), 500


@app.route("/status", methods=["GET"])
def status():
    printer_name = load_selected_printer()
    return jsonify({
        "status": "running",
        "version": VERSION,
        "default_printer": printer_name,
        "printer_available": bool(printer_name) and printer_name in list_windows_printers(),
        "autostart": is_autostart_enabled(),
    })


@app.route("/settings/autostart", methods=["POST"])
def settings_autostart():
    if not IS_FROZEN:
        return jsonify({"error": "Disponible sólo en la versión instalada"}), 400
    enabled = bool(_json_body().get("enabled"))
    toggle_autostart(enabled)
    return jsonify({"status": "ok", "autostart": is_autostart_enabled()})


def port_is_free():
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        sock.bind((HOST, PORT))
        return True
    except OSError:
        return False
    finally:
        sock.close()


def run_server_forever():
    """Servidor de producción; si se cae, se reinicia solo."""
    while True:
        try:
            log.info("Servidor escuchando en http://%s:%s", HOST, PORT)
            serve(app, host=HOST, port=PORT, threads=6, _quiet=True)
        except Exception:
            log.exception("El servidor se detuvo; reiniciando en 3 s")
        time.sleep(3)


# ----------------------------
# ÍCONO DE BANDEJA
# ----------------------------
def load_tray_image():
    try:
        return Image.open(resource_path("icon.ico"))
    except Exception:
        image = Image.new("RGB", (64, 64), (76, 175, 80))
        ImageDraw.Draw(image).rectangle((16, 16, 48, 48), fill=(255, 255, 255))
        return image


def open_settings(icon=None, item=None):
    webbrowser.open(SETTINGS_URL)


def open_logs(icon, item):
    os.startfile(LOG_DIR)


def on_test_print(icon, item):
    printer = load_selected_printer()
    if not printer:
        icon.notify("Primero elegí una impresora.", APP_NAME)
        open_settings()
        return
    try:
        send_raw_to_printer(printer, build_test_ticket(printer), "Denario - Prueba")
        icon.notify(f"Página de prueba enviada a {printer}.", APP_NAME)
    except Exception as e:
        log.exception("Falló la página de prueba")
        icon.notify(f"No se pudo imprimir: {e}", APP_NAME)


def on_toggle_autostart(icon, item):
    toggle_autostart(not is_autostart_enabled())


def shutdown():
    try:
        if _tray_icon:
            _tray_icon.stop()
    except Exception:
        pass
    os._exit(0)


def quit_app(icon, item):
    log.info("Cerrado por el usuario")
    shutdown()


def run_tray(notify_message=None, open_browser=False):
    global _tray_icon

    def setup(icon):
        icon.visible = True
        if notify_message:
            try:
                icon.notify(notify_message, f"{APP_NAME} está activa")
            except Exception:
                log.exception("No se pudo mostrar la notificación")
        if open_browser:
            open_settings()

    _tray_icon = Icon(
        APP_ID,
        load_tray_image(),
        APP_NAME,
        menu=Menu(
            MenuItem(f"{APP_NAME} v{VERSION}", None, enabled=False),
            MenuItem(
                lambda item: f"Impresora: {load_selected_printer() or 'sin configurar'}",
                None,
                enabled=False,
            ),
            Menu.SEPARATOR,
            MenuItem("Configurar impresora", open_settings, default=True),
            MenuItem("Imprimir página de prueba", on_test_print),
            MenuItem(
                "Iniciar con Windows",
                on_toggle_autostart,
                checked=lambda item: is_autostart_enabled(),
                enabled=IS_FROZEN,
            ),
            Menu.SEPARATOR,
            MenuItem("Abrir carpeta de registros", open_logs),
            MenuItem("Salir", quit_app),
        ),
    )
    _tray_icon.run(setup=setup)


# ----------------------------
# PUNTO DE ENTRADA
# ----------------------------
def main():
    setup_logging()
    args = sys.argv[1:]
    log.info("Iniciando %s v%s | exe=%s | args=%s", APP_NAME, VERSION, sys.executable, args)

    if "--uninstall" in args:
        uninstall(quiet="--quiet" in args)
        return

    if IS_FROZEN and not is_running_installed_copy() and "--portable" not in args:
        try:
            self_install()
            return
        except Exception:
            log.exception("La auto-instalación falló; se ejecuta en modo portable")

    from_protocol = any(a.lower().startswith(PROTOCOL + ":") for a in args)
    from_autostart = "--autostart" in args
    first_run = "--first-run" in args

    if not acquire_single_instance():
        log.info("Ya hay una instancia en ejecución")
        if not (from_protocol or from_autostart):
            open_settings()
        return

    migrate_legacy_config()
    if IS_FROZEN:
        # Re-afirmamos el registro en cada inicio (ruta correcta aunque se haya movido).
        try:
            if load_config().get("autostart", True) and not _disabled_in_task_manager():
                set_autostart(True)
            register_protocol(sys.executable)
        except Exception:
            log.exception("No se pudo actualizar el registro")

    if not port_is_free():
        log.error("El puerto %s está ocupado", PORT)
        message_box(
            f"No se pudo iniciar {APP_NAME}.\n\n"
            f"Otro programa está usando el puerto {PORT}. "
            "Cerralo o reiniciá la computadora e intentá de nuevo.",
            MB_OK | MB_ICONERROR,
        )
        return

    threading.Thread(target=listen_for_quit_signal, daemon=True).start()
    threading.Thread(target=run_server_forever, daemon=True).start()

    if first_run:
        message = "Listo. Se iniciará automáticamente con Windows."
    elif from_autostart or from_protocol:
        message = None
    else:
        message = "Denario ya puede imprimir en esta computadora."
    run_tray(notify_message=message, open_browser=first_run)


# ----------------------------
# PÁGINA DE CONFIGURACIÓN
# ----------------------------
SETTINGS_HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Denario Impresora · Configuración</title>
  <meta name="description" content="Configuración del asistente de impresión de Denario en Windows.">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --border: #e2e8f0;
      --text: #0f172a;
      --text-muted: #64748b;
      --primary: #10b981;
      --primary-hover: #059669;
      --primary-subtle: #ecfdf5;
      --primary-border: #a7f3d0;
      --radius: 12px;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      min-height: 100vh;
      font-family: 'Inter', system-ui, -apple-system, sans-serif;
      background-color: var(--bg);
      color: var(--text);
      display: flex;
      justify-content: center;
      align-items: center;
      padding: 30px 16px;
    }
    main {
      width: 100%;
      max-width: 480px;
      display: flex;
      flex-direction: column;
      gap: 16px;
    }
    header {
      display: flex;
      align-items: center;
      gap: 14px;
      padding-bottom: 4px;
    }
    header img {
      width: 46px;
      height: 46px;
      object-fit: contain;
      border-radius: 10px;
      background: #ffffff;
      border: 1px solid var(--border);
      padding: 6px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .header-text h1 {
      font-size: 1.25rem;
      font-weight: 700;
      color: var(--text);
      letter-spacing: -0.01em;
    }
    .header-text p {
      font-size: 0.82rem;
      color: var(--text-muted);
      margin-top: 2px;
    }
    .badge {
      margin-left: auto;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 10px;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 600;
      color: #065f46;
      background: var(--primary-subtle);
      border: 1px solid var(--primary-border);
      white-space: nowrap;
    }
    .dot {
      width: 7px;
      height: 7px;
      border-radius: 50%;
      background: var(--primary);
    }
    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 22px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04), 0 4px 12px rgba(0, 0, 0, 0.02);
    }
    .card h2 {
      font-size: 0.98rem;
      font-weight: 600;
      color: var(--text);
      margin-bottom: 4px;
    }
    .card .desc {
      font-size: 0.82rem;
      color: var(--text-muted);
      line-height: 1.45;
    }
    label.field {
      display: block;
      font-size: 0.8rem;
      font-weight: 600;
      color: #334155;
      margin: 16px 0 6px;
    }
    .select-wrap {
      position: relative;
    }
    .select-wrap::after {
      content: "";
      position: absolute;
      right: 14px;
      top: 50%;
      width: 7px;
      height: 7px;
      border-right: 2px solid #64748b;
      border-bottom: 2px solid #64748b;
      transform: translateY(-65%) rotate(45deg);
      pointer-events: none;
    }
    select {
      width: 100%;
      appearance: none;
      padding: 10px 36px 10px 14px;
      border-radius: 8px;
      font-size: 0.9rem;
      font-family: inherit;
      color: var(--text);
      background-color: #ffffff;
      border: 1px solid #cbd5e1;
      cursor: pointer;
      transition: border-color 0.15s, box-shadow 0.15s;
    }
    select:focus {
      outline: none;
      border-color: var(--primary);
      box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.15);
    }
    .actions {
      display: flex;
      gap: 10px;
      margin-top: 16px;
    }
    button {
      flex: 1;
      padding: 10px 14px;
      border-radius: 8px;
      font-size: 0.88rem;
      font-family: inherit;
      font-weight: 600;
      cursor: pointer;
      border: 1px solid transparent;
      transition: all 0.15s ease;
    }
    button:active {
      transform: scale(0.99);
    }
    button:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }
    .btn-primary {
      background-color: var(--primary);
      color: #ffffff;
    }
    .btn-primary:hover:not(:disabled) {
      background-color: var(--primary-hover);
    }
    .btn-secondary {
      background-color: #ffffff;
      color: #334155;
      border-color: #cbd5e1;
    }
    .btn-secondary:hover:not(:disabled) {
      background-color: #f1f5f9;
      border-color: #94a3b8;
    }
    .status-box {
      margin-top: 14px;
      padding: 10px 12px;
      border-radius: 8px;
      font-size: 0.82rem;
      line-height: 1.4;
    }
    .status-box.ok {
      background-color: #f0fdf4;
      border: 1px solid #bbf7d0;
      color: #166534;
    }
    .status-box.warn {
      background-color: #fffbeb;
      border: 1px solid #fde68a;
      color: #92400e;
    }
    .status-box.err {
      background-color: #fef2f2;
      border: 1px solid #fecaca;
      color: #991b1b;
    }
    .status-box strong {
      font-weight: 600;
    }
    .row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }
    .switch {
      position: relative;
      width: 44px;
      height: 24px;
      flex-shrink: 0;
    }
    .switch input {
      opacity: 0;
      width: 0;
      height: 0;
    }
    .slider {
      position: absolute;
      inset: 0;
      border-radius: 9999px;
      cursor: pointer;
      background-color: #cbd5e1;
      transition: background-color 0.2s;
    }
    .slider::before {
      content: "";
      position: absolute;
      width: 18px;
      height: 18px;
      left: 3px;
      top: 3px;
      border-radius: 50%;
      background: white;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.2);
      transition: transform 0.2s;
    }
    .switch input:checked + .slider {
      background-color: var(--primary);
    }
    .switch input:checked + .slider::before {
      transform: translateX(20px);
    }
    .switch input:disabled + .slider {
      opacity: 0.4;
      cursor: not-allowed;
    }
    .tip {
      display: flex;
      align-items: center;
      gap: 10px;
      background: #f8fafc;
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 12px 14px;
      font-size: 0.8rem;
      color: var(--text-muted);
      line-height: 1.45;
    }
    .tip svg {
      flex-shrink: 0;
      color: #64748b;
    }
    footer {
      text-align: center;
      font-size: 0.72rem;
      color: #94a3b8;
      padding-top: 4px;
    }
    .toast {
      position: fixed;
      left: 50%;
      bottom: 24px;
      transform: translate(-50%, 15px);
      opacity: 0;
      padding: 9px 16px;
      border-radius: 9999px;
      font-size: 0.82rem;
      font-weight: 500;
      color: #ffffff;
      background-color: #0f172a;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      pointer-events: none;
      transition: opacity 0.2s, transform 0.2s;
    }
    .toast.show {
      opacity: 1;
      transform: translate(-50%, 0);
    }
    .toast.err {
      background-color: #dc2626;
    }
  </style>
</head>
<body>
<main>
  <header>
    <img src="https://firebasestorage.googleapis.com/v0/b/cafe-galia.appspot.com/o/Denario%2FLogo%20denario.png?alt=media&token=094e6157-6a33-478e-b701-cb3e8ebe5f7a" alt="Logo Denario">
    <div class="header-text">
      <h1>Denario Impresora</h1>
      <p>Asistente de impresión para comandas y facturas</p>
    </div>
    <span class="badge"><span class="dot"></span>Activa · v{{ version }}</span>
  </header>

  <section class="card" aria-labelledby="printer-title">
    <h2 id="printer-title">Impresora predeterminada</h2>
    <p class="desc">Se utiliza cuando la aplicación web no especifica una impresora particular.</p>

    {% if printers %}
    <label class="field" for="printer_name">Seleccioná tu impresora:</label>
    <div class="select-wrap">
      <select id="printer_name" name="printer_name">
        {% for printer in printers %}
        <option value="{{ printer }}" {% if printer == selected %}selected{% endif %}>{{ printer }}</option>
        {% endfor %}
      </select>
    </div>
    <div class="actions">
      <button id="save-btn" class="btn-primary" type="button">Guardar</button>
      <button id="test-btn" class="btn-secondary" type="button">Imprimir prueba</button>
    </div>
    {% else %}
    <div class="status-box err">No se encontraron impresoras instaladas en Windows. Conectá tu impresora e instalá sus controladores.</div>
    {% endif %}

    <div id="current-status">
      {% if current and current_available %}
      <div class="status-box ok">Impresora configurada:&nbsp;<strong id="current-name">{{ current }}</strong></div>
      {% elif current %}
      <div class="status-box warn">La impresora guardada «{{ current }}» no está conectada o está apagada.</div>
      {% else %}
      <div class="status-box warn">Aún no se ha seleccionado ninguna impresora predeterminada.</div>
      {% endif %}
    </div>
  </section>

  <section class="card" aria-labelledby="startup-title">
    <div class="row">
      <div>
        <h2 id="startup-title">Iniciar con Windows</h2>
        <p class="desc">{% if can_autostart %}Inicia automáticamente en segundo plano al encender el equipo.{% else %}Disponible en la versión instalada.{% endif %}</p>
      </div>
      <label class="switch" aria-label="Iniciar con Windows">
        <input id="autostart-toggle" type="checkbox" {% if autostart %}checked{% endif %} {% if not can_autostart %}disabled{% endif %}>
        <span class="slider"></span>
      </label>
    </div>
  </section>

  <div class="tip">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></svg>
    <span>Podés cerrar esta ventana. Denario seguirá funcionando en segundo plano junto al reloj de Windows.</span>
  </div>

  <footer>Denario · Asistente de impresión v{{ version }}</footer>
</main>
<div id="toast" class="toast" role="status" aria-live="polite"></div>

<script>
  const $ = (s) => document.querySelector(s);
  let toastTimer;
  function toast(msg, kind = "ok") {
    const t = $("#toast");
    t.textContent = msg;
    t.className = `toast show ${kind === "err" ? "err" : ""}`;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), 2800);
  }
  async function api(path, body) {
    const r = await fetch(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body || {}) });
    const data = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(data.error || "Error inesperado");
    return data;
  }
  async function withBusy(btn, fn) {
    btn.disabled = true;
    try { await fn(); } catch (e) { toast(e.message, "err"); } finally { btn.disabled = false; }
  }
  const saveBtn = $("#save-btn"), testBtn = $("#test-btn"), select = $("#printer_name");
  saveBtn && saveBtn.addEventListener("click", () => withBusy(saveBtn, async () => {
    const d = await api("/select_printer", { printer_name: select.value });
    $("#current-status").innerHTML = '<div class="status-box ok">Impresora configurada:&nbsp;<strong></strong></div>';
    $("#current-status strong").textContent = d.default_printer;
    toast("Impresora guardada correctamente");
  }));
  testBtn && testBtn.addEventListener("click", () => withBusy(testBtn, async () => {
    await api("/test_print", { printer: select.value });
    toast("Página de prueba enviada a la impresora");
  }));
  const toggle = $("#autostart-toggle");
  toggle.addEventListener("change", async () => {
    try {
      const d = await api("/settings/autostart", { enabled: toggle.checked });
      toggle.checked = d.autostart;
      toast(d.autostart ? "Se iniciará con Windows" : "Inicio automático desactivado");
    } catch (e) {
      toggle.checked = !toggle.checked;
      toast(e.message, "err");
    }
  });
  {% if saved %}toast("Impresora guardada correctamente");{% endif %}
</script>
</body>
</html>
"""


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log.exception("Error fatal")
        raise