# -*- coding: utf-8 -*-
"""
mod_descuentos_ecommerce.py
Subproceso MAIN: "Cálculo de Descuentos Ecommerce"

Versión actualizada:
1) Prioridad de reglas editable por usuario mediante arrastre.
2) Se elimina módulo Checklist de %.
3) Listado Prod. 1P ahora detecta valores dinámicos: 1P, BTS 1P, 1P CYBER, etc.
   El usuario define si cada valor permite descuento y qué porcentaje aplica.
4) Descuentos Retail mantiene lógica:
   - Hoja PARES: StyleColor columna C, Descuento columna I.
   - Hoja ACC:   StyleColor columna C, Descuento columna G.
5) Listado Prod. Futbol deja de usarse. Se reemplaza por Tech Sports:
   - Style-Color columna A.
   - DETALLE PARA DSCTO columna K.
   - El usuario define si cada valor permite descuento y qué porcentaje aplica.
6) Compras Retail-Ecomm-BTS ahora usa grupos dinámicos:
   - Style-Color columna B.
   - M:Y agrupadas por regla comercial.
   - El usuario define si cada grupo permite descuento y qué porcentaje aplica.
7) El orden de prioridad es el primer filtro: la primera regla que aplique resuelve el producto.
8) Tech Sports: si el valor es APTO PARA DSCTO, el % se configura por SEASON ACTUAL desde Tech Sports columna I.
"""

import os
from pathlib import Path
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import pandas as pd


# ------------------ Helpers ------------------

def _norm_sc(x) -> str:
    return (str(x).strip() if x is not None else "").strip()


def _norm_key(x) -> str:
    return _norm_sc(x).replace("-", "").replace(" ", "")


def _cell_is_value(v) -> bool:
    if v is None:
        return False
    s = str(v).strip()
    if not s:
        return False
    return s != "-"


def _to_float_discount(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s or s == "-":
        return None
    s = s.replace("%", "").replace(",", ".").strip()
    try:
        f = float(s)
    except Exception:
        return None
    if 0 < f <= 1:
        return round(f * 100, 4)
    return float(f)


def _safe_str(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() == "nan":
        return ""
    return s


KI_KEYS = ["hotshot", "cozy fit", "glide-step", "glide step"]

PRIORITY_RULES = [
    ("retail", "Descuento Retail"),
    ("onep", "Listado Prod. 1P"),
    ("techsports", "Tech Sports"),
    ("compras", "Compras Retail-Ecomm-BTS"),
    ("ki", "Key Initiative"),
]

COMPRA_GROUPS = [
    ("COMPRA RETAIL BTS-26", ["M"]),
    ("COMPRA RETAIL 2026-1", ["N", "O", "P"]),
    ("COMPRA ECOMM 2026-1", ["Q"]),
    ("COMPRA ECOMM BTS-26", ["R"]),
    ("COMPRA RETAIL 2025-2", ["S", "T", "U"]),
    ("COMPRA RETAIL 2026-2", ["V", "W", "X"]),
    ("COMPRA ECOMM 2026-2", ["Y"]),
]


# ------------------ Rutas por defecto ------------------

def _find_xlsx_by_basename(folder: Path, base_name: str) -> str:
    try:
        folder = Path(folder)
        if not folder.exists():
            return ""

        for ext in [".xlsx", ".xlsm", ".xls"]:
            exact = folder / f"{base_name}{ext}"
            if exact.exists():
                return str(exact)

        base_low = base_name.lower()
        cands = [p for p in folder.glob("*.xls*") if p.stem.lower().startswith(base_low)]
        if not cands:
            return ""
        cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return str(cands[0])
    except Exception:
        return ""


def _find_newest_open_file(open_folder: Path) -> str:
    try:
        open_folder = Path(open_folder)
        if not open_folder.exists():
            return ""
        xlsx = list(open_folder.glob("*.xls*"))
        if not xlsx:
            return ""
        filtered = [p for p in xlsx if "openordersreportbystyle" in p.name.lower()]
        cands = filtered if filtered else xlsx
        cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return str(cands[0])
    except Exception:
        return ""


def _resolve_default_paths() -> dict:
    home = Path.home()
    candidate_roots = [
        home / "Skechers USA",
        home / "OneDrive - Skechers USA",
    ]

    paths_out = {
        "compra": "",
        "techsports": "",
        "onep": "",
        "retail": "",
        "beta": "",
        "open": "",
        "disponible": "",
    }

    for root in candidate_roots:
        base_mod_candidates = [
            root / "Macarena Caballero Gonzalez - Team Ecommerce" / "Python" / "mod_descuentos_ecommerce",
            root / "Team Ecommerce" / "Python" / "mod_descuentos_ecommerce",
        ]
        beta_candidates = [
            root / "Macarena Caballero Gonzalez - Team Ecommerce" / "Beta.xlsx",
            root / "Team Ecommerce" / "Beta.xlsx",
        ]
        open_dir_candidates = [
            root / "Macarena Caballero Gonzalez - Team Ecommerce" / "Stock Files" / "OPEN" / "2026",
            root / "Stock Files" / "OPEN" / "2026",
            root / "Team Ecommerce" / "Stock Files" / "OPEN" / "2026",
        ]
        disponible_candidates = [
            root / "Macarena Caballero Gonzalez - Team Ecommerce" / "Stock Files" / "Disponible.xlsx",
            root / "Stock Files" / "Disponible.xlsx",
            root / "Team Ecommerce" / "Stock Files" / "Disponible.xlsx",
        ]

        base_mod = next((p for p in base_mod_candidates if p.exists()), None)
        beta_path = next((p for p in beta_candidates if p.exists()), None)
        open_dir = next((p for p in open_dir_candidates if p.exists()), None)
        disponible_path = next((p for p in disponible_candidates if p.exists()), None)

        if not (base_mod or beta_path or open_dir or disponible_path):
            continue

        if base_mod:
            if not paths_out["compra"]:
                paths_out["compra"] = _find_xlsx_by_basename(base_mod, "Compras Retail-Ecomm-BTS") or _find_xlsx_by_basename(base_mod, "Compra 2026-1 Retail-Ecomm-BTS + 25-2")
            if not paths_out["techsports"]:
                paths_out["techsports"] = _find_xlsx_by_basename(base_mod, "Tech Sports")
            if not paths_out["onep"]:
                paths_out["onep"] = _find_xlsx_by_basename(base_mod, "Listado Prod. 1P")
            if not paths_out["retail"]:
                paths_out["retail"] = _find_xlsx_by_basename(base_mod, "Descuentos Retail")

        if not paths_out["beta"]:
            paths_out["beta"] = str(beta_path) if beta_path and beta_path.exists() else ""

        if not paths_out["open"]:
            paths_out["open"] = _find_newest_open_file(open_dir) if open_dir else ""

        if not paths_out["disponible"]:
            paths_out["disponible"] = str(disponible_path) if disponible_path and disponible_path.exists() else ""

    return paths_out


# ------------------ Vista ------------------

class DescuentosEcommerceView(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.paths = {
            "style": tk.StringVar(),
            "retail": tk.StringVar(),
            "beta": tk.StringVar(),
            "onep": tk.StringVar(),
            "techsports": tk.StringVar(),
            "open": tk.StringVar(),
            "compra": tk.StringVar(),
            "disponible": tk.StringVar(),
        }

        self._apply_default_paths()

        self.priority_rule_ids = [rule_id for rule_id, _label in PRIORITY_RULES]
        self.priority_labels = {rule_id: label for rule_id, label in PRIORITY_RULES}
        self.priority_listbox = None
        self._drag_index = None

        # Configs dinámicas por run.
        # Formato por valor/grupo: {"enabled": bool, "pct": float}
        self.onep_config = {}
        self.tech_config = {}
        self.compra_config = {}
        self.options_loaded = False

        self.onep_tree = None
        self.tech_tree = None
        self.compra_tree = None

        self._q = queue.Queue()
        self._worker = None
        self._total = 0
        self._done = 0
        self._out_path = None

        self._build_ui()
        self._poll_queue()

    # --------------- UI ---------------

    def _build_ui(self):
        header = skechers_header(
            self,
            "Cálculo de Descuentos Ecommerce",
            "Prioriza reglas y define descuentos dinámicos para 1P / Tech Sports / Compras / Retail."
        )
        header.pack(fill="x", pady=(0, 12))

        box = ttk.LabelFrame(self, text="Inputs", padding=10)
        box.pack(fill="x", pady=(0, 10))

        ttk.Label(box, text="Archivo Excel (Style-Color)", width=24).grid(row=0, column=0, sticky="w", pady=3)
        ent = ttk.Entry(box, textvariable=self.paths["style"], width=92)
        ent.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(box, text="Seleccionar…", command=lambda: self._pick_file("style")).grid(row=0, column=2, padx=6)

        ttk.Label(box, text="OPEN (auto)", width=24).grid(row=1, column=0, sticky="w", pady=3)
        ttk.Label(box, text="(se carga en background)").grid(row=1, column=1, sticky="w", padx=6)
        ttk.Button(box, text="Refrescar OPEN", command=self.refresh_open).grid(row=1, column=2, padx=6)

        box.columnconfigure(1, weight=1)

        main_area = ttk.Frame(self)
        main_area.pack(fill="both", expand=True, pady=(0, 10))

        left = ttk.LabelFrame(main_area, text="Prioridad de reglas", padding=10)
        left.pack(side="left", fill="y", padx=(0, 8))

        ttk.Label(left, text="Arrastra para ordenar. La primera regla que aplique define el resultado.", wraplength=300).pack(anchor="w", pady=(0, 8))
        self.priority_listbox = tk.Listbox(left, height=len(PRIORITY_RULES), exportselection=False, activestyle="none")
        self.priority_listbox.pack(fill="x")
        for i, rule_id in enumerate(self.priority_rule_ids, start=1):
            self.priority_listbox.insert("end", f"{i}. {self.priority_labels[rule_id]}")
        self.priority_listbox.bind("<Button-1>", self._on_priority_click)
        self.priority_listbox.bind("<B1-Motion>", self._on_priority_drag)
        self.priority_listbox.bind("<ButtonRelease-1>", self._renumber_priority_listbox)

        ttk.Button(left, text="Restablecer orden", command=self._reset_priority).pack(fill="x", pady=(8, 0))
        ttk.Label(left, text="Ejemplo: si Retail queda primero y existe descuento, se aplica antes de evaluar Compras/1P/Tech.", wraplength=300).pack(anchor="w", pady=(10, 0))

        right = ttk.LabelFrame(main_area, text="Configuración dinámica de reglas", padding=10)
        right.pack(side="left", fill="both", expand=True)

        top_cfg = ttk.Frame(right)
        top_cfg.pack(fill="x", pady=(0, 8))
        ttk.Button(top_cfg, text="Cargar / actualizar opciones desde archivos", command=self.load_dynamic_options).pack(side="left")
        ttk.Label(top_cfg, text="Tip: doble clic en 'Permite descuento' para alternar. Doble clic en '%' para editar.").pack(side="left", padx=10)

        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill="both", expand=True)

        self.onep_tree = self._make_config_tab("Listado Prod. 1P")
        self.tech_tree = self._make_config_tab("Tech Sports")
        self.compra_tree = self._make_config_tab("Compras Retail-Ecomm-BTS")

        ctl = ttk.Frame(self)
        ctl.pack(fill="x", pady=(0, 8))

        self.btn_run = ttk.Button(ctl, text="Ejecutar cálculo", style="Primary.TButton", command=self.run)
        self.btn_run.pack(side="left")

        self.btn_open = ttk.Button(ctl, text="Abrir output", command=self.open_output, state="disabled")
        self.btn_open.pack(side="left", padx=8)

        self.progress_label = ttk.Label(ctl, text="Progreso: 0%")
        self.progress_label.pack(side="right")

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(0, 8))

        ttk.Label(self, text="Mensajes:", style="SectionTitle.TLabel").pack(anchor="w")
        self.log = tk.Text(self, height=9)
        self.log.pack(fill="both", expand=False)

    def _make_config_tab(self, title):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=title)
        cols = ("valor", "enabled", "pct")
        tree = ttk.Treeview(tab, columns=cols, show="headings", height=8)
        tree.heading("valor", text="Valor / Grupo detectado")
        tree.heading("enabled", text="Permite descuento")
        tree.heading("pct", text="%")
        tree.column("valor", width=360, anchor="w")
        tree.column("enabled", width=140, anchor="center")
        tree.column("pct", width=80, anchor="center")
        tree.pack(fill="both", expand=True)
        tree.bind("<Double-1>", self._on_config_double_click)
        return tree

    def _apply_default_paths(self):
        defaults = _resolve_default_paths()
        for k in ["retail", "beta", "onep", "techsports", "open", "compra", "disponible"]:
            self.paths[k].set(defaults.get(k, "") or "")

    def refresh_open(self):
        try:
            defaults = _resolve_default_paths()
            new_open = (defaults.get("open") or "").strip()
            if not new_open or not os.path.exists(new_open):
                self._log("OPEN no encontrado: revisa carpeta OPEN\\2026 (Skechers USA / OneDrive - Skechers USA).")
                return
            old_open = (self.paths["open"].get() or "").strip()
            self.paths["open"].set(new_open)
            if old_open and old_open != new_open:
                self._log(f"OPEN actualizado: {Path(new_open).name}")
            else:
                self._log(f"OPEN vigente: {Path(new_open).name}")
        except Exception as e:
            self._log(f"Error refrescando OPEN: {e}")

    def _pick_file(self, key):
        if key != "style":
            return
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xlsm;*.xls")])
        if path:
            self.paths[key].set(path)

    def _log(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    # -------- Prioridad drag & drop --------

    def _refresh_priority_listbox(self):
        self.priority_listbox.delete(0, "end")
        for i, rule_id in enumerate(self.priority_rule_ids, start=1):
            self.priority_listbox.insert("end", f"{i}. {self.priority_labels[rule_id]}")

    def _on_priority_click(self, event):
        self._drag_index = self.priority_listbox.nearest(event.y)

    def _on_priority_drag(self, event):
        if self._drag_index is None:
            return
        new_index = self.priority_listbox.nearest(event.y)
        if new_index == self._drag_index or new_index < 0 or new_index >= len(self.priority_rule_ids):
            return
        self.priority_rule_ids.insert(new_index, self.priority_rule_ids.pop(self._drag_index))
        self._drag_index = new_index
        self._refresh_priority_listbox()
        self.priority_listbox.selection_set(new_index)

    def _renumber_priority_listbox(self, _event=None):
        self._drag_index = None
        self._refresh_priority_listbox()

    def _reset_priority(self):
        self.priority_rule_ids = [rule_id for rule_id, _label in PRIORITY_RULES]
        self._refresh_priority_listbox()

    def _get_priority_order(self):
        return list(self.priority_rule_ids)

    # -------- Config UI --------

    def _bool_txt(self, enabled):
        return "Sí" if enabled else "No"

    def _tree_to_config(self, tree):
        cfg = {}
        for iid in tree.get_children():
            valor, enabled_txt, pct_txt = tree.item(iid, "values")
            existing_item = {}
            for src in (self.onep_config, self.tech_config, self.compra_config):
                if str(valor) in src:
                    existing_item = dict(src.get(str(valor), {}))
                    break
            existing_item["enabled"] = str(enabled_txt).strip().lower() in ("sí", "si", "s", "true", "1")
            parsed_pct = _to_float_discount(pct_txt)
            if parsed_pct is not None:
                existing_item["pct"] = float(parsed_pct)
            else:
                existing_item.setdefault("pct", 0.0)
            cfg[str(valor)] = existing_item
        return cfg

    def _populate_tree(self, tree, config):
        tree.delete(*tree.get_children())
        for valor in sorted(config.keys(), key=lambda x: str(x).upper()):
            item = config[valor]
            pct_value = item.get("pct", 0.0)
            if str(valor).upper() == "APTO PARA DSCTO" and item.get("season_pct"):
                vals_unique = sorted(set(float(v or 0.0) for v in item.get("season_pct", {}).values()))
                pct_value = vals_unique[0] if len(vals_unique) == 1 else "Por season"
            tree.insert("", "end", values=(valor, self._bool_txt(item.get("enabled", False)), pct_value))

    def _on_config_double_click(self, event):
        tree = event.widget
        iid = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not iid:
            return
        vals = list(tree.item(iid, "values"))
        if col == "#2":
            vals[1] = "No" if str(vals[1]).strip().lower() in ("sí", "si") else "Sí"
            tree.item(iid, values=vals)
        elif col == "#3":
            valor = str(vals[0]).strip()

            # Caso especial Tech Sports: APTO PARA DSCTO se configura por SEASON FINAL
            # detectada desde OPEN columna BB, cruzada por Style-Color columna D.
            if tree is self.tech_tree and valor.upper() == "APTO PARA DSCTO":
                self._edit_tech_apto_by_season(valor, iid)
                return

            current = vals[2]
            new_pct = simpledialog.askfloat("Editar porcentaje", f"% para: {vals[0]}", initialvalue=float(_to_float_discount(current) or 0.0), minvalue=0, maxvalue=100)
            if new_pct is not None:
                vals[2] = float(new_pct)
                tree.item(iid, values=vals)

    def _edit_tech_apto_by_season(self, valor, iid):
        """Popup especial para APTO PARA DSCTO: permite definir % por SEASON FINAL."""
        cfg = self.tech_config.setdefault(valor, {"enabled": True, "pct": 0.0, "season_pct": {}})
        seasons = list(cfg.get("seasons", []))
        season_pct = dict(cfg.get("season_pct", {}))

        if not seasons:
            messagebox.showwarning(
                "Sin seasons detectadas",
                "No encontré SEASON FINAL en OPEN para productos Tech Sports con APTO PARA DSCTO.\n\n"
                "Presiona 'Cargar / actualizar opciones desde archivos' y revisa que OPEN tenga Style-Color en D y SEASON FINAL en BB."
            )
            return

        win = tk.Toplevel(self)
        win.title("Editar % por Season - Tech Sports")
        win.transient(self.winfo_toplevel())
        win.grab_set()
        win.resizable(False, False)

        frm = ttk.Frame(win, padding=14)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="% para: APTO PARA DSCTO", font=("TkDefaultFont", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Label(frm, text="Define el porcentaje por SEASON ACTUAL detectada en Tech Sports.").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        vars_by_season = {}
        for r, season in enumerate(seasons, start=2):
            ttk.Label(frm, text=f"{season}:", width=22).grid(row=r, column=0, sticky="w", pady=3)
            v = tk.DoubleVar(value=float(season_pct.get(season, cfg.get("pct", 0.0)) or 0.0))
            vars_by_season[season] = v
            sp = ttk.Spinbox(frm, from_=0, to=100, increment=0.5, textvariable=v, width=10, format="%.1f")
            sp.grid(row=r, column=1, sticky="w", pady=3)
            ttk.Label(frm, text="%").grid(row=r, column=2, sticky="w", padx=(4, 0))

        btns = ttk.Frame(frm)
        btns.grid(row=2 + len(seasons), column=0, columnspan=3, sticky="e", pady=(12, 0))

        def save():
            new_map = {season: float(var.get() or 0.0) for season, var in vars_by_season.items()}
            cfg["enabled"] = True
            cfg["season_pct"] = new_map
            # El % visible queda como resumen. Si todos son iguales, muestra ese valor; si no, indica "Por season".
            unique_vals = sorted(set(new_map.values()))
            visible_pct = unique_vals[0] if len(unique_vals) == 1 else "Por season"
            vals = list(self.tech_tree.item(iid, "values"))
            vals[1] = "Sí"
            vals[2] = visible_pct
            self.tech_tree.item(iid, values=vals)
            win.destroy()

        ttk.Button(btns, text="Guardar", command=save).pack(side="left", padx=(0, 6))
        ttk.Button(btns, text="Cancelar", command=win.destroy).pack(side="left")

        win.update_idletasks()
        x = self.winfo_toplevel().winfo_x() + (self.winfo_toplevel().winfo_width() // 2) - (win.winfo_width() // 2)
        y = self.winfo_toplevel().winfo_y() + (self.winfo_toplevel().winfo_height() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{x}+{y}")

    def load_dynamic_options(self):
        self._apply_default_paths()
        try:
            missing = []
            for k in ["onep", "techsports", "compra"]:
                p = self.paths[k].get().strip()
                if not p or not os.path.exists(p):
                    missing.append(f"- {k}: {p or '(no encontrado)'}")
            if missing:
                messagebox.showerror("Archivos no encontrados", "No pude cargar opciones dinámicas:\n\n" + "\n".join(missing))
                return False

            onep_idx = self._load_1p(self.paths["onep"].get())
            tech_idx = self._load_techsports(self.paths["techsports"].get())
            open_idx = self._load_open(self.paths["open"].get())
            compra_idx = self._load_compra(self.paths["compra"].get())

            self.onep_config = self._detect_onep_options(onep_idx, keep_existing=True)
            self.tech_config = self._detect_tech_options(tech_idx, keep_existing=True)
            self.compra_config = self._detect_compra_options(compra_idx, keep_existing=True)

            self._populate_tree(self.onep_tree, self.onep_config)
            self._populate_tree(self.tech_tree, self.tech_config)
            self._populate_tree(self.compra_tree, self.compra_config)

            self.options_loaded = True
            self._log("Opciones dinámicas cargadas correctamente.")
            return True
        except Exception as e:
            messagebox.showerror("Error cargando opciones", str(e))
            return False

    # --------------- Progreso ---------------

    def _set_progress(self, done, total, status=None):
        self._done = done
        self._total = max(total, 1)
        pct = int((self._done / self._total) * 100)
        self.progress["value"] = pct
        self.progress_label.config(text=f"Progreso: {pct}%")
        if status:
            self._log(status)

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    self._set_progress(payload["done"], payload["total"], payload.get("status"))
                elif kind == "done":
                    self._out_path = payload
                    self.btn_open.config(state="normal" if payload else "disabled")
                    self.btn_run.config(state="normal")
                    if payload:
                        messagebox.showinfo("Descuentos Ecommerce", f"✅ Output generado:\n{payload}")
                elif kind == "error":
                    self.btn_run.config(state="normal")
                    messagebox.showerror("Error", str(payload))
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    # --------------- Core ---------------

    def run(self):
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("En progreso", "Ya hay un proceso corriendo.")
            return

        self._apply_default_paths()

        if not self.paths["style"].get().strip():
            messagebox.showwarning("Falta input", "Debes seleccionar el archivo Excel (Style-Color).")
            return

        auto_keys = ["retail", "beta", "onep", "techsports", "open", "compra", "disponible"]
        auto_missing = [k for k in auto_keys if not self.paths[k].get().strip() or not os.path.exists(self.paths[k].get().strip())]
        if auto_missing:
            detalle = "\n".join([f"- {k}: {self.paths[k].get().strip() or '(no encontrado)'}" for k in auto_missing])
            messagebox.showerror(
                "No se encontraron archivos internos",
                "No pude resolver estos inputs automáticamente. Revisa rutas/nombres:\n\n" + detalle
            )
            return

        if not self.options_loaded:
            ok = self.load_dynamic_options()
            if not ok:
                return

        self.onep_config = self._tree_to_config(self.onep_tree)
        self.tech_config = self._tree_to_config(self.tech_tree)
        self.compra_config = self._tree_to_config(self.compra_tree)

        self.log.delete("1.0", "end")
        self.progress["value"] = 0
        self.progress_label.config(text="Progreso: 0%")
        self.btn_open.config(state="disabled")
        self.btn_run.config(state="disabled")
        self._out_path = None

        self._worker = threading.Thread(target=self._worker_run, daemon=True)
        self._worker.start()

    def open_output(self):
        if not self._out_path or not os.path.exists(self._out_path):
            return
        try:
            os.startfile(self._out_path)
        except Exception:
            pass

    # --------- Cargas ---------

    def _load_style_colors(self, path):
        df = pd.read_excel(path, header=0, usecols=[0])
        vals = [_norm_sc(v) for v in df.iloc[:, 0].tolist()]
        return [v for v in vals if v]

    def _load_beta(self, path):
        xl = pd.ExcelFile(path)
        sheet = "All (2)" if "All (2)" in xl.sheet_names else ("All" if "All" in xl.sheet_names else xl.sheet_names[0])
        df = pd.read_excel(path, sheet_name=sheet, header=0, usecols="B,E,F,G,J")
        df.columns = ["Style-Color", "Name", "Genero", "Clase", "DIV"]
        df["Style-Color"] = df["Style-Color"].map(_norm_sc)
        return df.set_index("Style-Color", drop=True)

    def _load_open(self, path):
        xl = pd.ExcelFile(path)
        sheet = "OPEN" if "OPEN" in xl.sheet_names else xl.sheet_names[0]
        # D = Style-Color para cruce, BB = SEASON FINAL para lógica Tech Sports por season.
        # AU se mantiene como SEASON original para compatibilidad/comentario 2026-1.
        df = pd.read_excel(path, sheet_name=sheet, header=0, usecols="D,R,AS,AU,BB,BC")
        df.columns = ["Style-Color", "DIV", "Genero", "SEASON", "SEASON_FINAL", "Name"]
        df["Style-Color"] = df["Style-Color"].map(_norm_sc)
        df["SEASON_FINAL"] = df["SEASON_FINAL"].fillna("").astype(str).str.strip()
        return df.set_index("Style-Color", drop=True)

    def _load_1p(self, path):
        df = pd.read_excel(path, header=0, usecols="A,D,E,F,G,H")
        df.columns = ["Style-Color", "Falabella", "Meli", "Paris", "Ripley", "Ecommerce"]
        df["Style-Color"] = df["Style-Color"].map(_norm_sc)
        for c in ["Falabella", "Meli", "Paris", "Ripley", "Ecommerce"]:
            df[c] = df[c].fillna("-").astype(str).str.strip()
        return df.set_index("Style-Color", drop=True)

    def _load_techsports(self, path):
        # Columna A: Style-Color / Skechers.
        # Columna I: SEASON ACTUAL.
        # Columna K: Detalle para descuento.
        df = pd.read_excel(path, header=0, usecols="A,I,K")
        df.columns = ["Style-Color", "SEASON_ACTUAL", "Detalle"]
        df["Style-Color"] = df["Style-Color"].map(_norm_sc)
        df["SEASON_ACTUAL"] = df["SEASON_ACTUAL"].fillna("").astype(str).str.strip()
        df["Detalle"] = df["Detalle"].fillna("-").astype(str).str.strip()
        return df.set_index("Style-Color", drop=True)

    def _load_compra(self, path):
        df = pd.read_excel(path, header=0, usecols="B,M,N,O,P,Q,R,S,T,U,V,W,X,Y")
        df.columns = ["Style-Color", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"]
        df["Style-Color"] = df["Style-Color"].map(_norm_sc)
        return df.set_index("Style-Color", drop=True)

    def _load_disponible(self, path):
        """Carga QTY y Q desde Disponible.xlsx, hoja ALL.

        Cruce solicitado:
          - Style-Color / SKECHERS: columna A
          - QTY: columna AH
          - Q: columna AI

        Se usa posición fija porque Disponible tiene encabezados repetidos (P/Q por tallas)
        y buscar por nombre puede tomar columnas incorrectas.
        """
        raw = pd.read_excel(path, sheet_name="ALL", header=None)

        key_pos = 0     # A = SKECHERS
        qty_pos = 33    # AH = QTY
        q_pos = 34      # AI = Q

        # Detectar la fila real de encabezados: donde A sea SKECHERS y AH sea QTY.
        header_row = None
        for idx, row in raw.iterrows():
            a = str(row.iloc[key_pos]).strip().upper() if len(row) > key_pos else ""
            ah = str(row.iloc[qty_pos]).strip().upper() if len(row) > qty_pos else ""
            ai = str(row.iloc[q_pos]).strip().upper() if len(row) > q_pos else ""
            if a == "SKECHERS" and ah == "QTY" and ai == "Q":
                header_row = idx
                break

        if header_row is None:
            raise ValueError("No encontré encabezados esperados en Disponible.xlsx / ALL: A=SKECHERS, AH=QTY, AI=Q.")

        out = raw.iloc[header_row + 1:, [key_pos, qty_pos, q_pos]].copy()
        out.columns = ["Style-Color", "QTY", "Q"]
        out["Style-Color"] = out["Style-Color"].map(_norm_sc)
        out = out[out["Style-Color"] != ""]
        return out.drop_duplicates(subset=["Style-Color"], keep="first").set_index("Style-Color", drop=True)

    def _get_disponible_values(self, sc, disponible_idx):
        if sc in disponible_idx.index:
            row = disponible_idx.loc[sc]
            return row.get("QTY", ""), row.get("Q", "")
        return "", ""

    def _load_retail(self, path):
        lookup = {}
        xl = pd.ExcelFile(path)
        pares = "PARES" if "PARES" in xl.sheet_names else xl.sheet_names[0]
        if "ACC" in xl.sheet_names:
            acc = "ACC"
        elif len(xl.sheet_names) > 1:
            acc = xl.sheet_names[1]
        else:
            acc = xl.sheet_names[0]

        df_pares = pd.read_excel(path, sheet_name=pares, header=0, usecols="C,I")
        df_pares.columns = ["StyleColor", "Descuento"]

        df_acc = pd.read_excel(path, sheet_name=acc, header=0, usecols="C,G")
        df_acc.columns = ["StyleColor", "Descuento"]

        for df in (df_pares, df_acc):
            df["StyleColor"] = df["StyleColor"].fillna("").astype(str).str.strip()
            df["Descuento"] = df["Descuento"].apply(_to_float_discount)
            for _, r in df.iterrows():
                k = _norm_key(r["StyleColor"])
                if not k:
                    continue
                d = r["Descuento"]
                if d is None:
                    continue
                prev = lookup.get(k)
                if prev is None or d > prev:
                    lookup[k] = d
        return lookup

    # --------- Detección de opciones dinámicas ---------

    def _detect_onep_options(self, onep_idx, keep_existing=False):
        existing = self.onep_config if keep_existing else {}
        cfg = dict(existing)
        cols = ["Falabella", "Meli", "Paris", "Ripley", "Ecommerce"]
        for col in cols:
            if col not in onep_idx.columns:
                continue
            for val in onep_idx[col].dropna().astype(str).str.strip().unique():
                if not val or val == "-" or val.lower() == "nan":
                    continue
                cfg.setdefault(val, {"enabled": False, "pct": 0.0})
        return cfg

    def _detect_tech_options(self, tech_idx, open_idx=None, keep_existing=False):
        existing = self.tech_config if keep_existing else {}
        cfg = dict(existing)
        if "Detalle" in tech_idx.columns:
            for val in tech_idx["Detalle"].dropna().astype(str).str.strip().unique():
                if not val or val == "-" or val.lower() == "nan":
                    continue
                default_enabled = val.strip().upper() == "APTO PARA DSCTO"
                cfg.setdefault(val, {"enabled": default_enabled, "pct": 0.0})

        # Para APTO PARA DSCTO, detectar las SEASON ACTUAL existentes en el mismo
        # archivo Tech Sports, columna I, solo para productos APTO PARA DSCTO.
        apto_key = next((k for k in cfg.keys() if str(k).strip().upper() == "APTO PARA DSCTO"), None)
        if apto_key:
            seasons = set()
            try:
                for _sc, row in tech_idx.iterrows():
                    if isinstance(row, pd.DataFrame):
                        row = row.iloc[0]
                    detalle = _safe_str(row.get("Detalle", ""))
                    if detalle.strip().upper() == "APTO PARA DSCTO":
                        season_actual = _safe_str(row.get("SEASON_ACTUAL", ""))
                        if season_actual and season_actual != "-":
                            seasons.add(season_actual)
            except Exception:
                seasons = set()

            item = cfg.setdefault(apto_key, {"enabled": True, "pct": 0.0})
            prev_map = dict(item.get("season_pct", {}))
            item["seasons"] = sorted(seasons, key=lambda x: str(x))
            item["season_pct"] = {season: float(prev_map.get(season, item.get("pct", 0.0)) or 0.0) for season in item["seasons"]}
            item["enabled"] = True

        return cfg

    def _detect_compra_options(self, compra_idx, keep_existing=False):
        existing = self.compra_config if keep_existing else {}
        cfg = dict(existing)
        for group_name, cols in COMPRA_GROUPS:
            present = False
            for col in cols:
                if col in compra_idx.columns and compra_idx[col].apply(_cell_is_value).any():
                    present = True
                    break
            if present:
                cfg.setdefault(group_name, {"enabled": False, "pct": 0.0})
        return cfg

    # --------- Evaluadores ---------

    def _pick_from_beta_open(self, sc, beta_idx, open_idx):
        out = {"Name": None, "Genero": None, "Clase": None, "DIV": None, "SEASON": None, "SEASON_FINAL": None}
        if sc in beta_idx.index:
            row = beta_idx.loc[sc]
            out["Name"] = _safe_str(row.get("Name", ""))
            out["Genero"] = _safe_str(row.get("Genero", ""))
            out["Clase"] = _safe_str(row.get("Clase", ""))
            out["DIV"] = _safe_str(row.get("DIV", ""))

        if sc in open_idx.index:
            row = open_idx.loc[sc]
            out["SEASON"] = _safe_str(row.get("SEASON", ""))
            out["SEASON_FINAL"] = _safe_str(row.get("SEASON_FINAL", ""))
            if not out["Name"]:
                out["Name"] = _safe_str(row.get("Name", ""))
            if not out["Genero"]:
                out["Genero"] = _safe_str(row.get("Genero", ""))
            if not out["DIV"]:
                out["DIV"] = _safe_str(row.get("DIV", ""))
        return out

    def _eval_ki(self, name: str):
        name_l = (name or "").lower()
        for k in KI_KEYS:
            if k in name_l:
                if k in ["glide step", "glide-step"]:
                    return "Glide-Step"
                if k == "cozy fit":
                    return "Cozy Fit"
                if k == "hotshot":
                    return "Hotshot"
        return None

    def _get_retail_discount(self, sc, retail_lookup):
        key = _norm_key(sc)
        return retail_lookup.get(key) if key else None

    def _onep_matches(self, sc, onep_idx):
        matches = []
        if sc not in onep_idx.index:
            return matches
        row = onep_idx.loc[sc]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        for col, canal in [
            ("Falabella", "Falabella"),
            ("Meli", "Meli 26-1"),
            ("Paris", "Paris 26-1"),
            ("Ripley", "Ripley 26-1"),
            ("Ecommerce", "Ecommerce"),
        ]:
            val = _safe_str(row.get(col, "-"))
            if val and val != "-":
                matches.append((val, canal))
        return matches

    def _tech_match(self, sc, tech_idx):
        if sc not in tech_idx.index:
            return None
        row = tech_idx.loc[sc]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        val = _safe_str(row.get("Detalle", "-"))
        season_actual = _safe_str(row.get("SEASON_ACTUAL", ""))
        if not val or val == "-":
            return None
        return val, season_actual

    def _compra_matches(self, sc, compra_idx):
        matches = []
        if sc not in compra_idx.index:
            return matches
        row = compra_idx.loc[sc]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        for group_name, cols in COMPRA_GROUPS:
            if any(_cell_is_value(row.get(c)) for c in cols):
                matches.append(group_name)
        return matches

    def _eval_configured_matches(self, rule_name, matches, config, detail_builder=None):
        if not matches:
            return None

        enabled = []
        disabled = []
        for raw in matches:
            key = raw[0] if isinstance(raw, tuple) else raw
            cfg = config.get(key, {"enabled": False, "pct": 0.0})
            if cfg.get("enabled", False):
                enabled.append((raw, float(cfg.get("pct", 0.0))))
            else:
                disabled.append(raw)

        if enabled:
            max_pct = max(p for _raw, p in enabled)
            selected = [raw for raw, p in enabled if p == max_pct]
            if detail_builder:
                comment = detail_builder(selected)
            else:
                comment = ", ".join([str(x[0] if isinstance(x, tuple) else x) for x in selected])
            return {
                "habilitado": True,
                "pct": max_pct,
                "comentario": f"{rule_name} - {comment} - {max_pct:g}%",
            }

        if disabled:
            if detail_builder:
                comment = detail_builder(disabled)
            else:
                comment = ", ".join([str(x[0] if isinstance(x, tuple) else x) for x in disabled])
            return {
                "habilitado": False,
                "pct": 0.0,
                "comentario": f"{rule_name} - {comment}",
            }
        return None

    def _eval_priority_rule(self, rule_id, sc, name, season_final, onep_idx, tech_idx, compra_idx, retail_lookup):
        if rule_id == "retail":
            retail_discount = self._get_retail_discount(sc, retail_lookup)
            if retail_discount is not None:
                return {
                    "habilitado": True,
                    "pct": float(retail_discount),
                    "comentario": f"Descuento Retail - {float(retail_discount):g}%",
                }
            return None

        if rule_id == "onep":
            matches = self._onep_matches(sc, onep_idx)
            return self._eval_configured_matches(
                "Listado Prod. 1P",
                matches,
                self.onep_config,
                detail_builder=lambda items: ", ".join([f"{v} {canal}" for v, canal in items])
            )

        if rule_id == "techsports":
            match = self._tech_match(sc, tech_idx)
            if not match:
                return None
            val, tech_season_actual = match
            cfg = self.tech_config.get(val, {"enabled": False, "pct": 0.0})

            # APTO PARA DSCTO puede tener % distinto por SEASON ACTUAL
            # del mismo archivo Tech Sports, columna I.
            if val.strip().upper() == "APTO PARA DSCTO" and cfg.get("season_pct"):
                if cfg.get("enabled", False):
                    pct = float(cfg.get("season_pct", {}).get(tech_season_actual, cfg.get("pct", 0.0)) or 0.0)
                    season_txt = f" {tech_season_actual}" if tech_season_actual else ""
                    return {
                        "habilitado": True,
                        "pct": pct,
                        "comentario": f"Tech Sports - {val}{season_txt} - {pct:g}%",
                    }
                return {
                    "habilitado": False,
                    "pct": 0.0,
                    "comentario": f"Tech Sports - {val}",
                }

            return self._eval_configured_matches("Tech Sports", [val], self.tech_config)

        if rule_id == "compras":
            matches = self._compra_matches(sc, compra_idx)
            return self._eval_configured_matches("Compras", matches, self.compra_config)

        if rule_id == "ki":
            ki = self._eval_ki(name)
            if ki:
                return {
                    "habilitado": False,
                    "pct": 0.0,
                    "comentario": f"Key Initiative - {ki}",
                }
            return None

        return None

    def _eval_by_priority(self, sc, name, season, season_final, onep_idx, tech_idx, compra_idx, retail_lookup, priority_order):
        for rule_id in priority_order:
            result = self._eval_priority_rule(rule_id, sc, name, season_final, onep_idx, tech_idx, compra_idx, retail_lookup)
            if result is not None:
                comments = [result["comentario"]]
                if season == "2026-1":
                    comments.append("2026-1")
                return float(result["pct"]), bool(result["habilitado"]), comments

        comments = []
        if season == "2026-1":
            comments.append("2026-1")
        if not comments:
            comments = ["Habilitados a tener descuento"]
        return 0.0, True, comments

    def _worker_run(self):
        try:
            self._q.put(("log", "Leyendo Style-Color..."))
            style_colors = self._load_style_colors(self.paths["style"].get())
            self._q.put(("log", f"Style-Color encontrados: {len(style_colors)}"))

            self._q.put(("log", "Cargando Beta..."))
            beta_idx = self._load_beta(self.paths["beta"].get())

            self._q.put(("log", "Cargando OPEN..."))
            open_idx = self._load_open(self.paths["open"].get())

            self._q.put(("log", "Cargando Listado Prod. 1P..."))
            onep_idx = self._load_1p(self.paths["onep"].get())

            self._q.put(("log", "Cargando Tech Sports..."))
            tech_idx = self._load_techsports(self.paths["techsports"].get())

            self._q.put(("log", "Cargando Compras Retail-Ecomm-BTS..."))
            compra_idx = self._load_compra(self.paths["compra"].get())

            retail_path = self.paths["retail"].get()
            self._q.put(("log", f"Cargando Descuentos Retail: {Path(retail_path).name}"))
            retail_lookup = self._load_retail(retail_path)

            disponible_path = self.paths["disponible"].get()
            self._q.put(("log", f"Cargando Disponible: {Path(disponible_path).name}"))
            disponible_idx = self._load_disponible(disponible_path)

            priority_order = self._get_priority_order()
            priority_txt = " > ".join(self.priority_labels.get(rule_id, rule_id) for rule_id in priority_order)
            self._q.put(("log", f"Prioridad seleccionada: {priority_txt}"))

            total = len(style_colors)
            rows_out = []

            for i, sc in enumerate(style_colors, start=1):
                info = self._pick_from_beta_open(sc, beta_idx, open_idx)
                name = (info.get("Name") or "").strip()
                season = (info.get("SEASON") or "").strip()
                season_final = (info.get("SEASON_FINAL") or "").strip()

                final_pct, habilitado, comments = self._eval_by_priority(
                    sc=sc,
                    name=name,
                    season=season,
                    season_final=season_final,
                    onep_idx=onep_idx,
                    tech_idx=tech_idx,
                    compra_idx=compra_idx,
                    retail_lookup=retail_lookup,
                    priority_order=priority_order,
                )

                habil_txt = "Habilitado para tener Descuento" if habilitado else "No Habilitado para tener Descuento"
                disponible_qty, disponible_q = self._get_disponible_values(sc, disponible_idx)

                rows_out.append({
                    "Style-Color": sc,
                    "%": float(final_pct),
                    "Habilitado/No Habilitado": habil_txt,
                    "Comentario": ", ".join(dict.fromkeys(comments)),
                    "QTY": disponible_qty,
                    "Q": disponible_q,
                })

                self._q.put(("progress", {
                    "done": i,
                    "total": total,
                    "status": f"Procesado {i}/{total}: {sc}"
                }))

            out_df = pd.DataFrame(rows_out, columns=["Style-Color", "%", "Habilitado/No Habilitado", "Comentario", "QTY", "Q"])
            base_dir = os.path.dirname(self.paths["style"].get())
            out_path = os.path.join(base_dir, "output_descuentos_ecommerce.xlsx")
            out_df.to_excel(out_path, index=False)
            self._q.put(("done", out_path))

        except Exception as e:
            self._q.put(("error", e))
