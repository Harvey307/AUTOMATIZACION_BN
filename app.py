from __future__ import annotations

import csv
import json
import queue
import re
import sys
import threading
import time
import traceback
from dataclasses import dataclass
from datetime import datetime, time as dt_time
from pathlib import Path
from typing import Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "AUTOMATIZACION_BN"
APP_VERSION = "1.1.0"
DISPLAY_COLS = ["codigo", "nombre", "incidente", "fecha", "hora", "estado", "detalle"]
COL_TITLES = {
    "codigo": "Código",
    "nombre": "Nombre y apellidos",
    "incidente": "Incidente",
    "fecha": "Fecha",
    "hora": "Hora",
    "estado": "Estado",
    "detalle": "Detalle",
}

# Nombres admitidos para las columnas de los Excel/TXT.
ALIASES = {
    "codigo": [
        "cod. empleado",
        "cod empleado",
        "cod. emplea",
        "cod emplea",
        "codigo",
        "código",
        "codigo del trabajador",
        "código del trabajador",
        "cod_emplea",
    ],
    "nombre": ["nombre y apellidos", "nombre", "trabajador", "apellidos y nombres"],
    "incidente": ["incidente", "tipo incidente"],
    "fecha_inicio": ["fecha inicio", "fecha", "fecha_inicio"],
    "hora_inicio": ["hora inicio", "hora_inicio", "ingreso"],
    "hora_fin": ["hora fin", "hora_fin", "salida"],
    "estado_origen": ["estado"],
}

# Valores obtenidos del mapa real de UI Automation de SYGNUS.
SYGNUS = {
    "main_title_regex": r"^Maestr(?:a|ía) de Asistencia$",
    "main_class": "FNWND3105",
    "process_name": "tar_base",

    # Ventana Buscar / Find
    "find_title": "Find",
    "find_class": "FNWNS3105",
    "find_where_combo_id": 1002,
    "find_text_id": 1008,
    "search_direction_combo_id": 1009,
    "find_next_id": 1006,
    "find_cancel_id": 1007,

    # Lista principal de trabajadores
    "workers_window_title": "Lista de Trabajadores",
    "workers_window_id": 200,
    "workers_window_class": "FNWND3105",
    "workers_grid_id": 1000,
    "workers_grid_class": "pbdw105",

    # Ver Movimientos Trabajador
    "movements_title": "Ver Movimientos Trabajador",
    "movements_id": 201,
    "movements_class": "FNWND3105",
    "tabs_id": 1001,
    "tabs_class": "PBTabControl32_100",
    "marks_pane_title": "Marcaciones",
    "marks_pane_id": 1004,
    "marks_pane_class": "FNUDO3105",
    "new_mark_button_id": 1002,  # botón izquierdo de la barra de Marcaciones

    # Nuevo Marcación
    "new_mark_title": "Nuevo Marcación",
    "new_mark_class": "FNWNS3105",
    "new_mark_datawindow_id": 1002,
    "new_mark_datawindow_class": "pbdw105",
    "save_button_id": 1000,
    "close_button_id": 1001,
}

DEFAULT_DELAYS = {
    "short": 0.25,
    "medium": 0.65,
    "after_search": 0.45,
    "after_open_worker": 0.65,
    "after_open_mark": 0.50,
    "after_save": 0.85,
}

# Fallbacks geométricos. Se usan sólo para controles antiguos que el mapa no expone.
# El botón Buscar no apareció como control individual en UI Automation.
FALLBACK = {
    # Dentro de FNFIXEDBAR105, el botón Buscar está aprox. a 335 px desde la izquierda.
    "toolbar_buscar_offset": (335, 18),
    # Dentro del PBTabControl32_100, cabecera de la pestaña Marcaciones.
    "marcaciones_tab_points": [(0.355, 0.030), (0.330, 0.030), (0.385, 0.030)],
    # Dentro del pbdw105 de "Nuevo Marcación".
    # El DataWindow no expone Fecha/Hora como AutomationElement separados.
    "fecha_point": (0.28, 0.17),
    "hora_point": (0.28, 0.36),
}


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def load_delays() -> dict:
    """Lee los delays del profile si existe, sin depender obligatoriamente del JSON."""
    delays = dict(DEFAULT_DELAYS)
    profile = resource_path("automation_profile.json")
    try:
        if profile.exists():
            with open(profile, "r", encoding="utf-8") as f:
                data = json.load(f)
            delays.update(data.get("delays", {}))
    except Exception:
        pass
    return delays


def normalize_col(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def find_column(columns, aliases):
    normalized = {normalize_col(c): c for c in columns}
    for alias in aliases:
        n_alias = normalize_col(alias)
        if n_alias in normalized:
            return normalized[n_alias]
    for n, original in normalized.items():
        for alias in aliases:
            a = normalize_col(alias)
            if a in n or n in a:
                return original
    return None


def normalize_code(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    digits = re.sub(r"\D", "", s)
    if digits:
        return digits.zfill(7)
    return s


def normalize_date(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%d/%m/%Y")
    s = str(value).strip()
    if not s:
        return ""
    for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%d/%m/%y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    try:
        return pd.to_datetime(value, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return s


def normalize_time(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, dt_time):
        return value.strftime("%H:%M")
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%H:%M")
    if isinstance(value, (float, int)) and 0 <= float(value) < 1:
        total = round(float(value) * 24 * 60)
        return f"{total // 60:02d}:{total % 60:02d}"
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p"):
        try:
            return datetime.strptime(s, fmt).strftime("%H:%M")
        except ValueError:
            pass
    m = re.match(r"^(\d{1,2}):(\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}:{int(m.group(2)):02d}"
    return s


def read_input(path: str) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, dtype=object)
    if ext in (".txt", ".csv"):
        try:
            return pd.read_csv(path, sep=None, engine="python", dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(path, sep=None, engine="python", dtype=str, encoding="latin1")
    raise ValueError("Formato no admitido. Usa Excel, TXT o CSV.")


def transform_data(raw: pd.DataFrame) -> pd.DataFrame:
    code_col = find_column(raw.columns, ALIASES["codigo"])
    incident_col = find_column(raw.columns, ALIASES["incidente"])
    date_col = find_column(raw.columns, ALIASES["fecha_inicio"])
    hi_col = find_column(raw.columns, ALIASES["hora_inicio"])
    hf_col = find_column(raw.columns, ALIASES["hora_fin"])
    name_col = find_column(raw.columns, ALIASES["nombre"])

    missing = []
    if not code_col:
        missing.append("Cod. Empleado")
    if not incident_col:
        missing.append("Incidente")
    if not date_col:
        missing.append("Fecha inicio")
    if missing:
        raise ValueError("No encontré estas columnas obligatorias: " + ", ".join(missing))

    rows = []
    for _, r in raw.iterrows():
        incident = "" if pd.isna(r[incident_col]) else str(r[incident_col]).strip()
        upper = incident.upper()
        if "INGRESO" in upper:
            source_time = r[hi_col] if hi_col else ""
        elif "SALIDA" in upper:
            source_time = r[hf_col] if hf_col else ""
        else:
            source_time = r[hi_col] if hi_col and not pd.isna(r[hi_col]) else (r[hf_col] if hf_col else "")
        rows.append(
            {
                "codigo": normalize_code(r[code_col]),
                "nombre": "" if not name_col or pd.isna(r[name_col]) else str(r[name_col]).strip(),
                "incidente": incident,
                "fecha": normalize_date(r[date_col]),
                "hora": normalize_time(source_time),
                "estado": "PENDIENTE",
                "detalle": "",
            }
        )
    return pd.DataFrame(rows, columns=DISPLAY_COLS)


class RowEditor(tk.Toplevel):
    def __init__(self, parent, title, values=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.vars = {}
        values = values or {}
        fields = ["codigo", "nombre", "incidente", "fecha", "hora"]
        for i, key in enumerate(fields):
            ttk.Label(self, text=COL_TITLES[key]).grid(row=i, column=0, padx=10, pady=6, sticky="w")
            v = tk.StringVar(value=values.get(key, ""))
            self.vars[key] = v
            ttk.Entry(self, textvariable=v, width=46).grid(row=i, column=1, padx=10, pady=6)
        btns = ttk.Frame(self)
        btns.grid(row=len(fields), column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="Guardar", command=self.save).pack(side="left", padx=5)
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="left", padx=5)
        self.transient(parent)
        self.grab_set()

    def save(self):
        data = {k: v.get().strip() for k, v in self.vars.items()}
        data["codigo"] = normalize_code(data["codigo"])
        data["fecha"] = normalize_date(data["fecha"])
        data["hora"] = normalize_time(data["hora"])
        if not data["codigo"] or not data["fecha"] or not data["hora"]:
            messagebox.showwarning("Datos incompletos", "Código, fecha y hora son obligatorios.", parent=self)
            return
        data["estado"] = "PENDIENTE"
        data["detalle"] = ""
        self.result = data
        self.destroy()


@dataclass
class AutomationResult:
    ok: bool
    detail: str


class AttendanceAutomator:
    """
    Automatizador específico para la UI real de Maestra de Asistencia (tar_base).

    Prioridad:
      1) IDs/clases reales obtenidos del mapa UI Automation.
      2) teclado / menú.
      3) coordenadas relativas sólo donde PowerBuilder no expone el control.
    """

    def __init__(self, log_callback, cancel_event: threading.Event):
        self.log = log_callback
        self.cancel_event = cancel_event
        self.delays = load_delays()
        self.main = None
        self.pid = None

    def _check_cancel(self):
        if self.cancel_event.is_set():
            raise InterruptedError("Proceso cancelado por el usuario")

    def _sleep(self, key: str):
        self._check_cancel()
        time.sleep(float(self.delays.get(key, DEFAULT_DELAYS.get(key, 0.4))))
        self._check_cancel()

    @staticmethod
    def _spec_exists(spec, timeout=0.25) -> bool:
        try:
            return bool(spec.exists(timeout=timeout))
        except Exception:
            return False

    def connect(self):
        from pywinauto import Desktop

        self.log("Buscando 'Maestra de Asistencia'...")
        desktop = Desktop(backend="win32")

        # Primero: título + clase reales del mapa.
        main = desktop.window(
            title_re=SYGNUS["main_title_regex"],
            class_name=SYGNUS["main_class"],
        )

        try:
            main.wait("exists visible", timeout=10)
        except Exception as exc:
            # Diagnóstico útil: muestra las ventanas visibles que sí encontró pywinauto.
            visible = []
            try:
                for w in desktop.windows():
                    title = (w.window_text() or "").strip()
                    if title:
                        visible.append(title)
            except Exception:
                pass
            sample = ", ".join(visible[:12]) if visible else "(ninguna)"
            raise RuntimeError(
                "No se encontró la ventana 'Maestra de Asistencia'. "
                "Verifica que SYGNUS esté abierto en Asistencia de trabajador y que ambos programas "
                "se ejecuten con el mismo nivel de permisos. "
                f"Ventanas visibles: {sample}"
            ) from exc

        self.main = main.wrapper_object()
        try:
            self.pid = self.main.process_id()
        except Exception:
            self.pid = None

        try:
            self.main.set_focus()
        except Exception:
            pass

        # Verifica además que la Lista de Trabajadores exista.
        workers = self._workers_window_spec()
        if not self._spec_exists(workers, timeout=1.0):
            raise RuntimeError(
                "Se encontró Maestra de Asistencia, pero no la ventana 'Lista de Trabajadores'. "
                "Entra primero a la opción Asistencia de trabajador."
            )

        pid_text = f" PID={self.pid}" if self.pid else ""
        self.log(f"Conectado a Maestra de Asistencia.{pid_text}")
        return self.pid

    def _workers_window_spec(self):
        return self.main.child_window(
            title=SYGNUS["workers_window_title"],
            control_id=SYGNUS["workers_window_id"],
            class_name=SYGNUS["workers_window_class"],
        )

    def _workers_grid(self):
        workers = self._workers_window_spec()
        workers.wait("exists", timeout=3)
        grid = workers.child_window(
            control_id=SYGNUS["workers_grid_id"],
            class_name=SYGNUS["workers_grid_class"],
        )
        grid.wait("exists", timeout=3)
        return grid.wrapper_object()

    def _find_spec(self):
        return self.main.child_window(
            title=SYGNUS["find_title"],
            class_name=SYGNUS["find_class"],
        )

    def _open_find(self):
        from pywinauto import mouse

        self._check_cancel()
        existing = self._find_spec()
        if self._spec_exists(existing, timeout=0.2):
            return existing.wrapper_object()

        self.log("Abriendo Buscar...")

        # 1) Intenta usar el menú nativo: evita depender del mouse.
        opened = False
        for menu_path in ("Acciones->Buscar", "Acciones->Buscar...", "Buscar"):
            try:
                self.main.menu_select(menu_path)
                self._sleep("short")
                if self._spec_exists(self._find_spec(), timeout=0.6):
                    opened = True
                    break
            except Exception:
                pass

        # 2) Fallback al botón de la barra. El mapa expone la barra FNFIXEDBAR105,
        #    pero no el botón Buscar como nodo independiente.
        if not opened:
            try:
                bar = self.main.child_window(class_name="FNFIXEDBAR105").wrapper_object()
                rect = bar.rectangle()
                dx, dy = FALLBACK["toolbar_buscar_offset"]
                mouse.click(coords=(rect.left + dx, rect.top + dy))
            except Exception as exc:
                raise RuntimeError("No se pudo accionar el botón Buscar.") from exc

        find_spec = self._find_spec()
        find_spec.wait("exists visible", timeout=5)
        return find_spec.wrapper_object()

    @staticmethod
    def _set_combo(combo, text: str):
        from pywinauto.keyboard import send_keys

        try:
            combo.select(text)
            return
        except Exception:
            pass
        try:
            combo.click_input()
            send_keys("^a")
            send_keys(text, with_spaces=True)
            send_keys("{ENTER}")
        except Exception as exc:
            raise RuntimeError(f"No se pudo seleccionar '{text}' en el combo.") from exc

    def _detect_not_found_dialog(self) -> bool:
        """Detecta, si existe, un diálogo de 'no encontrado' del mismo proceso."""
        from pywinauto import Desktop
        from pywinauto.keyboard import send_keys

        if not self.pid:
            return False
        try:
            for w in Desktop(backend="win32").windows(process=self.pid):
                title = (w.window_text() or "").strip()
                cls = w.class_name()
                # No confundir las ventanas conocidas con un posible MessageBox.
                if title in {
                    "Maestra de Asistencia",
                    "Find",
                    "Ver Movimientos Trabajador",
                    "Nuevo Marcación",
                    "Lista de Trabajadores",
                }:
                    continue
                texts = " ".join(t for t in w.texts() if t).lower()
                candidate = f"{title} {texts}".lower()
                if cls == "#32770" and any(k in candidate for k in ("no encontr", "not found", "no existe")):
                    try:
                        w.set_focus()
                        send_keys("{ENTER}")
                    except Exception:
                        pass
                    return True
        except Exception:
            pass
        return False

    def _search_once(self, code: str, direction: str) -> bool:
        """Busca un código con Up o Down usando IDs reales del diálogo Find."""
        self._check_cancel()
        win = self._open_find()

        try:
            where_combo = win.child_window(
                control_id=SYGNUS["find_where_combo_id"], class_name="ComboBox"
            ).wrapper_object()
            text_edit = win.child_window(
                control_id=SYGNUS["find_text_id"], class_name="Edit"
            ).wrapper_object()
            direction_combo = win.child_window(
                control_id=SYGNUS["search_direction_combo_id"], class_name="ComboBox"
            ).wrapper_object()
            find_next = win.child_window(
                control_id=SYGNUS["find_next_id"], class_name="Button"
            ).wrapper_object()
        except Exception as exc:
            raise RuntimeError("No se pudieron detectar los controles internos de Find.") from exc

        self._set_combo(where_combo, "Código Del Trabajador")
        try:
            text_edit.set_edit_text(code)
        except Exception:
            text_edit.click_input()
            from pywinauto.keyboard import send_keys
            send_keys("^a{BACKSPACE}")
            send_keys(code)
        self._set_combo(direction_combo, direction)

        self.log(f"Buscando trabajador {code} ({direction})...")
        find_next.click_input()
        self._sleep("after_search")

        if self._detect_not_found_dialog():
            self.log(f"Sin coincidencia con Search={direction}.")
            return False
        return True

    def _close_find(self):
        from pywinauto.keyboard import send_keys

        spec = self._find_spec()
        if not self._spec_exists(spec, timeout=0.1):
            return
        win = spec.wrapper_object()
        try:
            cancel = win.child_window(control_id=SYGNUS["find_cancel_id"], class_name="Button")
            if cancel.exists(timeout=0.3):
                cancel.click_input()
            else:
                send_keys("{ESC}")
        except Exception:
            try:
                send_keys("{ESC}")
            except Exception:
                pass
        self._sleep("short")

    def _try_open_worker(self, timeout=2.5):
        """
        Abre el trabajador seleccionado. PowerBuilder no expone las filas de la grilla;
        se intenta actuar sobre la fila actual mediante teclado y, como respaldo, sobre
        la primera fila visible de la DataWindow.
        """
        from pywinauto.keyboard import send_keys
        from pywinauto import mouse

        self._check_cancel()
        mov_spec = self.main.child_window(
            title=SYGNUS["movements_title"],
            control_id=SYGNUS["movements_id"],
            class_name=SYGNUS["movements_class"],
        )

        # Si ya está abierto, reutilizarlo.
        if self._spec_exists(mov_spec, timeout=0.1):
            return mov_spec.wrapper_object()

        grid = self._workers_grid()

        # Intento 1: Enter sobre la selección actual que dejó Find Next.
        try:
            grid.set_focus()
        except Exception:
            pass
        try:
            send_keys("{ENTER}")
            if mov_spec.exists(timeout=1.0):
                mov_spec.wait("visible", timeout=1.0)
                return mov_spec.wrapper_object()
        except Exception:
            pass

        # Intento 2: clic en la primera fila visible; Find suele desplazar la coincidencia.
        try:
            rect = grid.rectangle()
            mouse.click(coords=(rect.left + 120, rect.top + 18))
            if mov_spec.exists(timeout=0.8):
                mov_spec.wait("visible", timeout=1.0)
                return mov_spec.wrapper_object()
        except Exception:
            pass

        # Intento 3: doble clic en esa misma fila.
        try:
            rect = grid.rectangle()
            mouse.double_click(coords=(rect.left + 120, rect.top + 18), interval=0.12)
            mov_spec.wait("exists visible", timeout=timeout)
            return mov_spec.wrapper_object()
        except Exception:
            return None

    def _find_and_open_worker(self, code: str):
        """
        Ejecuta la regla de negocio indicada por el usuario:
        buscar Up primero y, si no permite abrir el trabajador, repetir con Down.
        """
        last_error = None
        for direction in ("Up", "Down"):
            self._check_cancel()
            try:
                searched = self._search_once(code, direction)
                self._close_find()
                if not searched:
                    continue
                mov = self._try_open_worker()
                if mov is not None:
                    self.log(f"Trabajador {code} abierto correctamente.")
                    return mov
                last_error = f"No se pudo abrir el trabajador después de buscar {direction}."
                self.log(last_error)
            except Exception as exc:
                last_error = str(exc)
                self._close_find()
                self.log(f"Intento {direction} falló: {exc}")

        raise RuntimeError(last_error or f"No se pudo localizar al trabajador {code}.")

    def _marks_pane_spec(self, mov):
        return mov.child_window(
            title=SYGNUS["marks_pane_title"],
            control_id=SYGNUS["marks_pane_id"],
            class_name=SYGNUS["marks_pane_class"],
        )

    def _marks_active(self, mov) -> bool:
        pane = self._marks_pane_spec(mov)
        try:
            if not pane.exists(timeout=0.15):
                return False
            w = pane.wrapper_object()
            return w.is_visible() and w.is_enabled()
        except Exception:
            return False

    def _activate_marks_tab(self, mov):
        from pywinauto import mouse
        from pywinauto.keyboard import send_keys

        self._check_cancel()
        if self._marks_active(mov):
            return self._marks_pane_spec(mov).wrapper_object()

        tabs = mov.child_window(
            control_id=SYGNUS["tabs_id"], class_name=SYGNUS["tabs_class"]
        )
        tabs.wait("exists", timeout=3)
        tab = tabs.wrapper_object()

        # PowerBuilder no expone los TabItem. Se pulsa la zona de la cabecera "Marcaciones"
        # y se valida que aparezca el FNUDO3105 con ID 1004.
        rect = tab.rectangle()
        for xr, yr in FALLBACK["marcaciones_tab_points"]:
            self._check_cancel()
            mouse.click(coords=(rect.left + int(rect.width() * xr), rect.top + int(rect.height() * yr)))
            self._sleep("short")
            if self._marks_active(mov):
                return self._marks_pane_spec(mov).wrapper_object()

        # Último respaldo: ciclar pestañas con Ctrl+Tab.
        try:
            tab.set_focus()
        except Exception:
            pass
        for _ in range(6):
            send_keys("^{TAB}")
            self._sleep("short")
            if self._marks_active(mov):
                return self._marks_pane_spec(mov).wrapper_object()

        raise RuntimeError("No se pudo activar la pestaña Marcaciones.")

    def _open_new_mark(self, mov):
        self._check_cancel()
        marks = self._activate_marks_tab(mov)
        self.log("Marcaciones activa. Abriendo Nueva Marcación...")

        try:
            new_button = marks.child_window(
                control_id=SYGNUS["new_mark_button_id"], class_name="Button"
            )
            new_button.wait("exists enabled", timeout=3)
            new_button.click_input()
        except Exception as exc:
            raise RuntimeError("No se pudo accionar el botón Nueva Marcación (ID 1002).") from exc

        self._sleep("after_open_mark")
        new_spec = self.main.child_window(
            title=SYGNUS["new_mark_title"], class_name=SYGNUS["new_mark_class"]
        )
        new_spec.wait("exists visible", timeout=5)
        return new_spec.wrapper_object()

    @staticmethod
    def _click_ratio(control, point):
        from pywinauto import mouse

        xr, yr = point
        rect = control.rectangle()
        mouse.click(coords=(rect.left + int(rect.width() * xr), rect.top + int(rect.height() * yr)))

    def _type_into_datawindow(self, datawindow, point, value: str):
        from pywinauto.keyboard import send_keys

        self._click_ratio(datawindow, point)
        time.sleep(0.12)
        send_keys("^a{BACKSPACE}")
        send_keys(value, with_spaces=True)
        time.sleep(0.12)

    def _fill_and_save(self, win, fecha: str, hora: str):
        """
        El mapa de UI Automation muestra Fecha/Hora dentro de un único pbdw105 (ID 1002),
        no como Edit separados. Por eso se escriben por posición relativa dentro del DataWindow.
        Guardar sí está expuesto de forma estable como Button ID 1000.
        """
        self._check_cancel()
        try:
            datawindow = win.child_window(
                control_id=SYGNUS["new_mark_datawindow_id"],
                class_name=SYGNUS["new_mark_datawindow_class"],
            ).wrapper_object()
        except Exception as exc:
            raise RuntimeError("No se detectó el formulario interno de Nuevo Marcación (pbdw105 ID 1002).") from exc

        self.log(f"Ingresando Fecha={fecha} Hora={hora}...")
        self._type_into_datawindow(datawindow, FALLBACK["fecha_point"], fecha)
        self._type_into_datawindow(datawindow, FALLBACK["hora_point"], hora)

        try:
            save = win.child_window(control_id=SYGNUS["save_button_id"], class_name="Button")
            save.wait("exists enabled", timeout=3)
            save.click_input()
        except Exception as exc:
            raise RuntimeError("No se pudo accionar Guardar (ID 1000).") from exc

        self._sleep("after_save")

        # Si el formulario sigue visible, puede haber rechazado los datos o estar esperando algo.
        try:
            if win.exists(timeout=0.2) and win.is_visible():
                raise RuntimeError(
                    "Nuevo Marcación sigue abierto después de Guardar. "
                    "Revisa la Fecha/Hora o algún mensaje de validación del sistema."
                )
        except RuntimeError:
            raise
        except Exception:
            # Si la ventana ya no existe, el guardado/close fue correcto.
            pass

    def _close_movements(self, mov):
        try:
            mov.close()
            self._sleep("short")
        except Exception:
            try:
                from pywinauto.keyboard import send_keys
                mov.set_focus()
                send_keys("%{F4}")
                self._sleep("short")
            except Exception:
                pass

    def process_row(self, code: str, fecha: str, hora: str) -> AutomationResult:
        mov = None
        try:
            self._check_cancel()
            try:
                self.main.set_focus()
            except Exception:
                pass

            mov = self._find_and_open_worker(code)
            self._sleep("after_open_worker")
            new = self._open_new_mark(mov)
            self._fill_and_save(new, fecha, hora)
            self._close_movements(mov)
            return AutomationResult(True, "Marcación creada correctamente")
        except InterruptedError:
            raise
        except Exception as exc:
            # Intentar volver a un estado utilizable para continuar con la siguiente fila.
            try:
                self._close_find()
            except Exception:
                pass
            try:
                if mov is not None:
                    self._close_movements(mov)
            except Exception:
                pass
            return AutomationResult(False, f"{type(exc).__name__}: {exc}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} - v{APP_VERSION}")
        self.geometry("1210x740")
        self.minsize(980, 620)
        self.df = pd.DataFrame(columns=DISPLAY_COLS)
        self.file_path = None
        self.worker = None
        self.cancel_event = threading.Event()
        self.events = queue.Queue()
        self._build_ui()
        self.after(150, self._drain_events)

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Button(top, text="Abrir Excel/TXT", command=self.load_file).pack(side="left", padx=4)
        ttk.Button(top, text="Agregar", command=self.add_row).pack(side="left", padx=4)
        ttk.Button(top, text="Modificar", command=self.edit_row).pack(side="left", padx=4)
        ttk.Button(top, text="Eliminar", command=self.delete_row).pack(side="left", padx=4)
        ttk.Button(top, text="Exportar resultado", command=self.export_result).pack(side="left", padx=4)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=8)
        self.test_btn = ttk.Button(top, text="Probar conexión", command=self.test_connection)
        self.test_btn.pack(side="left", padx=4)

        self.start_btn = ttk.Button(top, text="Iniciar automatización", command=self.start_automation)
        self.start_btn.pack(side="right", padx=4)
        self.cancel_btn = ttk.Button(top, text="Cancelar", command=self.cancel_automation, state="disabled")
        self.cancel_btn.pack(side="right", padx=4)

        info = ttk.Frame(self, padding=(12, 0, 12, 4))
        info.pack(fill="x")
        self.file_label = ttk.Label(info, text="Ningún archivo cargado")
        self.file_label.pack(side="left", fill="x", expand=True)
        self.connection_label = ttk.Label(info, text="Sistema: no probado")
        self.connection_label.pack(side="right")

        table_frame = ttk.Frame(self, padding=10)
        table_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table_frame, columns=DISPLAY_COLS, show="headings", selectmode="browse")
        widths = {
            "codigo": 90,
            "nombre": 250,
            "incidente": 160,
            "fecha": 100,
            "hora": 75,
            "estado": 110,
            "detalle": 330,
        }
        for c in DISPLAY_COLS:
            self.tree.heading(c, text=COL_TITLES[c])
            self.tree.column(c, width=widths[c], minwidth=60, anchor="w")
        y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", lambda _e: self.edit_row())

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x")
        self.status = ttk.Label(bottom, text="Listo")
        self.status.pack(anchor="w", pady=(5, 0))

        log_frame = ttk.LabelFrame(self, text="Registro de ejecución", padding=5)
        log_frame.pack(fill="both", padx=10, pady=(0, 10))
        self.log_text = tk.Text(log_frame, height=8, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)

    def log(self, text):
        self.events.put(("log", text))

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, row in self.df.iterrows():
            values = [row.get(c, "") for c in DISPLAY_COLS]
            self.tree.insert("", "end", iid=str(i), values=values)

    def load_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Excel o TXT",
            filetypes=[("Excel/TXT", "*.xlsx *.xlsm *.xls *.txt *.csv"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            raw = read_input(path)
            self.df = transform_data(raw)
            self.df.index = range(len(self.df))
            self.file_path = path
            self.file_label.config(text=path)
            self.refresh_tree()
            self.status.config(text=f"{len(self.df)} registros cargados")
        except Exception as exc:
            messagebox.showerror("Error al cargar", str(exc))

    def selected_index(self) -> Optional[int]:
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def add_row(self):
        dlg = RowEditor(self, "Agregar registro")
        self.wait_window(dlg)
        if dlg.result:
            self.df.loc[len(self.df)] = dlg.result
            self.refresh_tree()

    def edit_row(self):
        idx = self.selected_index()
        if idx is None:
            messagebox.showinfo("Modificar", "Selecciona una fila.")
            return
        dlg = RowEditor(self, "Modificar registro", self.df.loc[idx].to_dict())
        self.wait_window(dlg)
        if dlg.result:
            for k, v in dlg.result.items():
                self.df.at[idx, k] = v
            self.refresh_tree()

    def delete_row(self):
        idx = self.selected_index()
        if idx is None:
            messagebox.showinfo("Eliminar", "Selecciona una fila.")
            return
        if messagebox.askyesno("Eliminar", "¿Eliminar el registro seleccionado?"):
            self.df = self.df.drop(index=idx).reset_index(drop=True)
            self.refresh_tree()

    def export_result(self):
        if self.df.empty:
            messagebox.showinfo("Exportar", "No hay datos para exportar.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.df.to_excel(path, index=False)
            messagebox.showinfo("Exportar", "Archivo guardado correctamente.")

    def validate_rows(self, indices=None):
        errors = []
        indices = list(self.df.index if indices is None else indices)
        for i in indices:
            r = self.df.loc[i]
            if not str(r["codigo"]).strip():
                errors.append(f"Fila {i + 1}: código vacío")
            if not str(r["fecha"]).strip():
                errors.append(f"Fila {i + 1}: fecha vacía")
            if not str(r["hora"]).strip():
                errors.append(f"Fila {i + 1}: hora vacía")
        return errors

    def _pending_indices(self):
        # Seguridad contra duplicados durante la misma sesión: no reprocesar REGISTRADO.
        return [
            idx
            for idx in self.df.index
            if str(self.df.at[idx, "estado"]).strip().upper() != "REGISTRADO"
        ]

    def test_connection(self):
        if self.worker and self.worker.is_alive():
            return
        self.test_btn.config(state="disabled")
        self.connection_label.config(text="Sistema: comprobando...")

        def worker():
            try:
                temp_cancel = threading.Event()
                auto = AttendanceAutomator(self.log, temp_cancel)
                pid = auto.connect()
                self.events.put(("connection", True, f"Conectado - PID {pid}" if pid else "Conectado"))
            except Exception as exc:
                self.events.put(("connection", False, str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def start_automation(self):
        if self.worker and self.worker.is_alive():
            return
        if self.df.empty:
            messagebox.showwarning("Sin datos", "Carga primero un Excel/TXT.")
            return

        pending = self._pending_indices()
        if not pending:
            messagebox.showinfo("Sin pendientes", "Todos los registros ya figuran como REGISTRADO.")
            return

        errors = self.validate_rows(pending)
        if errors:
            messagebox.showerror("Datos incompletos", "\n".join(errors[:15]))
            return

        if not messagebox.askyesno(
            "Confirmar automatización",
            "Antes de continuar:\n"
            "1. Abre SYGNUS.\n"
            "2. Entra a Asistencia de trabajador.\n"
            "3. Deja visible la Lista de Trabajadores.\n\n"
            f"Se procesarán {len(pending)} registros pendientes.\n\n"
            "¿Deseas iniciar?",
        ):
            return

        self.cancel_event.clear()
        self.start_btn.config(state="disabled")
        self.test_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.progress["maximum"] = len(pending)
        self.progress["value"] = 0
        self.worker = threading.Thread(target=self._automation_worker, args=(pending,), daemon=True)
        self.worker.start()

    def cancel_automation(self):
        self.cancel_event.set()
        self.status.config(text="Cancelando...")
        self.log("Se solicitó cancelar el proceso.")

    def _automation_worker(self, indices):
        try:
            automator = AttendanceAutomator(self.log, self.cancel_event)
            pid = automator.connect()
            self.events.put(("connection", True, f"Conectado - PID {pid}" if pid else "Conectado"))

            processed = 0
            for idx in indices:
                if self.cancel_event.is_set():
                    raise InterruptedError()

                code = str(self.df.at[idx, "codigo"])
                fecha = str(self.df.at[idx, "fecha"])
                hora = str(self.df.at[idx, "hora"])
                self.events.put(("row", idx, "PROCESANDO", "Buscando trabajador"))
                self.log(f"[{processed + 1}/{len(indices)}] {code} - {fecha} {hora}")

                result = automator.process_row(code, fecha, hora)
                state = "REGISTRADO" if result.ok else "ERROR"
                self.events.put(("row", idx, state, result.detail))
                processed += 1
                self.events.put(("progress", processed))

            self.events.put(("done", "Proceso finalizado."))
        except InterruptedError:
            self.events.put(("done", "Proceso cancelado por el usuario."))
        except Exception as exc:
            self.log(traceback.format_exc())
            self.events.put(("connection", False, str(exc)))
            self.events.put(("done", f"Error general: {exc}"))

    def _drain_events(self):
        try:
            while True:
                evt = self.events.get_nowait()
                if evt[0] == "log":
                    self.log_text.config(state="normal")
                    self.log_text.insert("end", evt[1] + "\n")
                    self.log_text.see("end")
                    self.log_text.config(state="disabled")
                elif evt[0] == "row":
                    _, idx, state, detail = evt
                    self.df.at[idx, "estado"] = state
                    self.df.at[idx, "detalle"] = detail
                    if self.tree.exists(str(idx)):
                        vals = [self.df.at[idx, c] for c in DISPLAY_COLS]
                        self.tree.item(str(idx), values=vals)
                        self.tree.see(str(idx))
                        self.tree.selection_set(str(idx))
                    self.status.config(text=f"{self.df.at[idx, 'codigo']}: {state}")
                elif evt[0] == "progress":
                    self.progress["value"] = evt[1]
                elif evt[0] == "connection":
                    _, ok, detail = evt
                    self.connection_label.config(text=f"Sistema: {'CONECTADO' if ok else 'NO CONECTADO'}")
                    if ok:
                        self.status.config(text=detail)
                    else:
                        self.status.config(text=detail)
                        self.log(f"Conexión: {detail}")
                    self.test_btn.config(state="normal")
                elif evt[0] == "done":
                    self.start_btn.config(state="normal")
                    self.test_btn.config(state="normal")
                    self.cancel_btn.config(state="disabled")
                    self.status.config(text=evt[1])
                    self.log(evt[1])
        except queue.Empty:
            pass
        self.after(150, self._drain_events)


if __name__ == "__main__":
    App().mainloop()
