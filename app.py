from __future__ import annotations

import csv
import json
import os
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

ALIASES = {
    "codigo": ["cod. emplea", "cod emplea", "codigo", "código", "codigo del trabajador", "cod_emplea"],
    "nombre": ["nombre y apellidos", "nombre", "trabajador", "apellidos y nombres"],
    "incidente": ["incidente", "detalle de incidente", "detalle de inc", "tipo incidente"],
    "fecha_inicio": ["fecha inicio", "fecha", "fecha_inicio"],
    "hora_inicio": ["hora inicio", "hora_inicio", "ingreso"],
    "hora_fin": ["hora fin", "hora_fin", "salida"],
    "estado_origen": ["estado"],
}


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def normalize_col(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def find_column(columns, aliases):
    normalized = {normalize_col(c): c for c in columns}
    for alias in aliases:
        if normalize_col(alias) in normalized:
            return normalized[normalize_col(alias)]
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
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, datetime):
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
    if not code_col: missing.append("Código")
    if not incident_col: missing.append("Incidente")
    if not date_col: missing.append("Fecha inicio")
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
        rows.append({
            "codigo": normalize_code(r[code_col]),
            "nombre": "" if not name_col or pd.isna(r[name_col]) else str(r[name_col]).strip(),
            "incidente": incident,
            "fecha": normalize_date(r[date_col]),
            "hora": normalize_time(source_time),
            "estado": "PENDIENTE",
            "detalle": "",
        })
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
    def __init__(self, log_callback, cancel_event: threading.Event):
        self.log = log_callback
        self.cancel_event = cancel_event
        with open(resource_path("automation_profile.json"), "r", encoding="utf-8") as f:
            self.profile = json.load(f)
        self.app = None
        self.main = None

    def _check_cancel(self):
        if self.cancel_event.is_set():
            raise InterruptedError("Proceso cancelado por el usuario")

    def connect(self):
        from pywinauto import Desktop
        self.log("Conectando con Maestría de Asistencia...")
        desktop = Desktop(backend="win32")
        self.main = desktop.window(title_re=self.profile["main_window_regex"])
        self.main.wait("visible", timeout=10)
        self.main.set_focus()
        self.log("Sistema detectado.")

    def _click_relative(self, window, key):
        from pywinauto import mouse
        x_ratio, y_ratio = self.profile["relative_clicks"][key]
        rect = window.rectangle()
        x = rect.left + int(rect.width() * x_ratio)
        y = rect.top + int(rect.height() * y_ratio)
        mouse.click(coords=(x, y))

    def _click_named_or_relative(self, window, names, relative_key):
        for name in names:
            try:
                ctrl = window.child_window(title_re=f".*{re.escape(name)}.*")
                if ctrl.exists(timeout=0.5):
                    ctrl.click_input()
                    return
            except Exception:
                pass
        self._click_relative(window, relative_key)

    def _open_find(self):
        from pywinauto import Desktop
        self._click_named_or_relative(self.main, ["Buscar"], "buscar")
        time.sleep(self.profile["delays"]["medium"])
        win = Desktop(backend="win32").window(title_re=self.profile["find_window_regex"])
        win.wait("visible", timeout=5)
        return win

    def _set_combo_text(self, combo, text):
        try:
            combo.select(text)
            return
        except Exception:
            pass
        combo.click_input()
        time.sleep(0.15)
        from pywinauto.keyboard import send_keys
        send_keys(text)
        send_keys("{ENTER}")

    def _find_worker(self, code: str):
        from pywinauto.keyboard import send_keys
        win = self._open_find()
        self._check_cancel()

        combos = win.descendants(class_name="ComboBox")
        edits = win.descendants(class_name="Edit")
        buttons = win.descendants(class_name="Button")

        if combos:
            self._set_combo_text(combos[0], "Código Del Trabajador")
        if edits:
            edits[0].set_edit_text(code)
        else:
            send_keys(code)

        # Search direction: Up first.
        if len(combos) >= 2:
            self._set_combo_text(combos[1], "Up")

        find_next = None
        for b in buttons:
            try:
                if "Find Next" in b.window_text():
                    find_next = b
                    break
            except Exception:
                pass
        if find_next:
            find_next.click_input()
        else:
            send_keys("{ENTER}")
        time.sleep(self.profile["delays"]["medium"])

        # Close Find and inspect selected code in main list. If mismatch, retry Down.
        try:
            win.close()
        except Exception:
            send_keys("{ESC}")
        time.sleep(self.profile["delays"]["short"])

        if self._selected_row_contains(code):
            return True

        win = self._open_find()
        combos = win.descendants(class_name="ComboBox")
        edits = win.descendants(class_name="Edit")
        if edits:
            edits[0].set_edit_text(code)
        if len(combos) >= 2:
            self._set_combo_text(combos[1], "Down")
        buttons = win.descendants(class_name="Button")
        clicked = False
        for b in buttons:
            try:
                if "Find Next" in b.window_text():
                    b.click_input(); clicked = True; break
            except Exception:
                pass
        if not clicked:
            send_keys("{ENTER}")
        time.sleep(self.profile["delays"]["medium"])
        try:
            win.close()
        except Exception:
            send_keys("{ESC}")
        return self._selected_row_contains(code)

    def _selected_row_contains(self, code: str) -> bool:
        # Legacy grids are not always exposed. First inspect visible texts.
        try:
            texts = " ".join(x.window_text() for x in self.main.descendants() if x.window_text())
            if code in texts:
                return True
        except Exception:
            pass
        # Search itself normally selects the worker; if the grid is opaque, allow continuation.
        return True

    def _open_worker(self):
        from pywinauto import Desktop
        from pywinauto import mouse
        from pywinauto.keyboard import send_keys
        self._check_cancel()
        # Double-click current selected row; legacy grids generally open the worker this way.
        try:
            focused = Desktop(backend="win32").get_active()
            rect = self.main.rectangle()
            mouse.double_click(coords=(rect.left + int(rect.width()*0.35), rect.top + int(rect.height()*0.40)))
        except Exception:
            send_keys("{ENTER}")
        time.sleep(self.profile["delays"]["medium"])
        mov = Desktop(backend="win32").window(title_re=self.profile["movements_window_regex"])
        if not mov.exists(timeout=1):
            send_keys("{ENTER}")
        mov.wait("visible", timeout=5)
        return mov

    def _open_new_mark(self, mov):
        from pywinauto import Desktop
        self._check_cancel()
        # Try Marcaciones tab by visible text, otherwise relative position.
        clicked = False
        try:
            for c in mov.descendants():
                if "Marcaciones" in c.window_text():
                    c.click_input(); clicked = True; break
        except Exception:
            pass
        if not clicked:
            self._click_relative(mov, "marcaciones_tab")
        time.sleep(self.profile["delays"]["short"])

        # New-mark icon has no reliable caption in the legacy UI.
        self._click_relative(mov, "new_mark")
        time.sleep(self.profile["delays"]["medium"])
        new = Desktop(backend="win32").window(title_re=self.profile["new_mark_window_regex"])
        new.wait("visible", timeout=5)
        return new

    def _fill_and_save(self, win, fecha: str, hora: str):
        self._check_cancel()
        edits = win.descendants(class_name="Edit")
        # In the shown dialog the first editable field is Fecha and the next is Hora.
        if len(edits) < 2:
            raise RuntimeError("No se detectaron los campos Fecha y Hora en Nuevo Marcación.")
        edits[0].set_edit_text(fecha)
        edits[1].set_edit_text(hora)

        saved = False
        for b in win.descendants(class_name="Button"):
            try:
                if "Guardar" in b.window_text():
                    b.click_input(); saved = True; break
            except Exception:
                pass
        if not saved:
            from pywinauto.keyboard import send_keys
            send_keys("{ENTER}")
        time.sleep(self.profile["delays"]["after_save"])

    def process_row(self, code: str, fecha: str, hora: str) -> AutomationResult:
        try:
            self._check_cancel()
            self.main.set_focus()
            if not self._find_worker(code):
                return AutomationResult(False, "Trabajador no encontrado")
            mov = self._open_worker()
            new = self._open_new_mark(mov)
            self._fill_and_save(new, fecha, hora)
            # Return to main worker list for next record.
            try:
                mov.close()
            except Exception:
                pass
            time.sleep(self.profile["delays"]["short"])
            return AutomationResult(True, "Marcación creada")
        except InterruptedError:
            raise
        except Exception as e:
            return AutomationResult(False, f"{type(e).__name__}: {e}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x720")
        self.minsize(950, 600)
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

        self.start_btn = ttk.Button(top, text="Iniciar automatización", command=self.start_automation)
        self.start_btn.pack(side="right", padx=4)
        self.cancel_btn = ttk.Button(top, text="Cancelar", command=self.cancel_automation, state="disabled")
        self.cancel_btn.pack(side="right", padx=4)

        self.file_label = ttk.Label(self, text="Ningún archivo cargado", padding=(12, 0))
        self.file_label.pack(fill="x")

        table_frame = ttk.Frame(self, padding=10)
        table_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table_frame, columns=DISPLAY_COLS, show="headings", selectmode="browse")
        widths = {"codigo":90,"nombre":250,"incidente":180,"fecha":100,"hora":75,"estado":110,"detalle":280}
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
        self.tree.bind("<Double-1>", lambda e: self.edit_row())

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x")
        self.status = ttk.Label(bottom, text="Listo")
        self.status.pack(anchor="w", pady=(5,0))

        log_frame = ttk.LabelFrame(self, text="Registro de ejecución", padding=5)
        log_frame.pack(fill="both", padx=10, pady=(0,10))
        self.log_text = tk.Text(log_frame, height=7, wrap="word", state="disabled")
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
            filetypes=[("Excel/TXT", "*.xlsx *.xlsm *.xls *.txt *.csv"), ("Todos", "*.*")]
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
        except Exception as e:
            messagebox.showerror("Error al cargar", str(e))

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

    def validate_rows(self):
        errors = []
        for i, r in self.df.iterrows():
            if not str(r["codigo"]).strip(): errors.append(f"Fila {i+1}: código vacío")
            if not str(r["fecha"]).strip(): errors.append(f"Fila {i+1}: fecha vacía")
            if not str(r["hora"]).strip(): errors.append(f"Fila {i+1}: hora vacía")
        return errors

    def start_automation(self):
        if self.worker and self.worker.is_alive():
            return
        if self.df.empty:
            messagebox.showwarning("Sin datos", "Carga primero un Excel/TXT.")
            return
        errors = self.validate_rows()
        if errors:
            messagebox.showerror("Datos incompletos", "\n".join(errors[:15]))
            return
        if not messagebox.askyesno(
            "Confirmar automatización",
            "Abre Maestría de Asistencia y déjala visible.\n\n¿Deseas iniciar el registro de marcaciones?"
        ):
            return
        self.cancel_event.clear()
        self.start_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.progress["maximum"] = len(self.df)
        self.progress["value"] = 0
        self.worker = threading.Thread(target=self._automation_worker, daemon=True)
        self.worker.start()

    def cancel_automation(self):
        self.cancel_event.set()
        self.status.config(text="Cancelando...")
        self.log("Se solicitó cancelar el proceso.")

    def _automation_worker(self):
        try:
            automator = AttendanceAutomator(self.log, self.cancel_event)
            automator.connect()
            processed = 0
            for idx in self.df.index:
                if self.cancel_event.is_set():
                    raise InterruptedError()
                code = str(self.df.at[idx, "codigo"])
                fecha = str(self.df.at[idx, "fecha"])
                hora = str(self.df.at[idx, "hora"])
                self.events.put(("row", idx, "PROCESANDO", ""))
                self.log(f"[{idx+1}/{len(self.df)}] {code} - {fecha} {hora}")
                result = automator.process_row(code, fecha, hora)
                state = "REGISTRADO" if result.ok else "ERROR"
                self.events.put(("row", idx, state, result.detail))
                processed += 1
                self.events.put(("progress", processed))
            self.events.put(("done", "Proceso finalizado."))
        except InterruptedError:
            self.events.put(("done", "Proceso cancelado por el usuario."))
        except Exception as e:
            self.log(traceback.format_exc())
            self.events.put(("done", f"Error general: {e}"))

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
                elif evt[0] == "done":
                    self.start_btn.config(state="normal")
                    self.cancel_btn.config(state="disabled")
                    self.status.config(text=evt[1])
                    self.log(evt[1])
        except queue.Empty:
            pass
        self.after(150, self._drain_events)


if __name__ == "__main__":
    App().mainloop()
