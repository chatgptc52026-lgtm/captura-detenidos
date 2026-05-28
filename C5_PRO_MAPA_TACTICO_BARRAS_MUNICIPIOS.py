# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import json
import random
import re
import threading
import traceback
import webbrowser
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import folium
    from folium import FeatureGroup
    from folium.plugins import MarkerCluster, HeatMap
except Exception as e:
    raise SystemExit(
        "Faltan dependencias.\n\nInstala con:\n"
        "pip install folium pandas openpyxl"
    ) from e


APP_TITLE = "C5 PRO - MAPA TÁCTICO INTELIGENTE"
APP_SIZE = "1520x960"
OUTPUT_DIR_NAME = "salidas_mapa_tactico_c5"

LAT_CANDIDATES = ["lat", "latitude", "latitud", "coordenada_y", "coord_y", "y"]
LON_CANDIDATES = ["lon", "lng", "long", "longitude", "longitud", "coordenada_x", "coord_x", "x"]
DATE_CANDIDATES = ["fecha", "fecha_hecho", "fecha del hecho", "date", "fecha_reporte"]
CATEGORY_CANDIDATES = ["tipo", "categoria", "incidente", "delito", "id_homologado", "evento", "clasificacion"]
MUNICIPIO_CANDIDATES = ["municipio"]
COLONIA_CANDIDATES = ["colonia"]
SECTOR_CANDIDATES = ["sector"]
CUADRANTE_CANDIDATES = ["cuadrante"]
ANIO_CANDIDATES = ["anio", "año", "year"]
MES_CANDIDATES = ["mes", "month"]

SMALL_DATASET_LIMIT = 8000
MEDIUM_DATASET_LIMIT = 50000
LARGE_DATASET_LIMIT = 250000

# Rango aproximado Quintana Roo
LAT_MIN, LAT_MAX = 17.5, 21.8
LON_MIN, LON_MAX = -89.8, -86.3


def normalize_text(value: str) -> str:
    text = str(value).strip().lower()
    for a, b in {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "ä": "a", "ë": "e", "ï": "i", "ö": "o", "ü": "u",
        "ñ": "n"
    }.items():
        text = text.replace(a, b)
    return re.sub(r"\s+", " ", text).strip()


def normalize_key(value: str) -> str:
    text = normalize_text(value)
    text = re.sub(r"\bregion\b", "reg", text)
    text = re.sub(r"\bsupermanzana\b", "sm", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def safe_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def suggest_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    normalized = {col: normalize_text(col) for col in columns}
    for col, norm in normalized.items():
        if norm in candidates:
            return col
    for candidate in candidates:
        for col, norm in normalized.items():
            if candidate in norm:
                return col
    return None


def random_color_from_text(text: str) -> str:
    seed = sum(ord(c) for c in str(text))
    palette = [
        "#ef4444", "#f97316", "#f59e0b", "#22c55e", "#06b6d4",
        "#3b82f6", "#8b5cf6", "#d946ef", "#ec4899", "#84cc16"
    ]
    return palette[seed % len(palette)]


def month_name_to_number(text: str) -> Optional[int]:
    value = normalize_text(text)
    mapping = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
        "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
        "noviembre": 11, "diciembre": 12
    }
    if value.isdigit():
        n = int(value)
        if 1 <= n <= 12:
            return n
    return mapping.get(value)


def detect_csv_separator(path: Path) -> str:
    try:
        with open(path, "r", encoding="utf-8-sig", errors="ignore", newline="") as f:
            sample = f.read(5000)
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except Exception:
        return ","


def parse_kml_coordinates(coord_text: str) -> List[Tuple[float, float]]:
    coords = []
    for p in coord_text.strip().split():
        vals = p.split(",")
        if len(vals) >= 2:
            try:
                lon = float(vals[0])
                lat = float(vals[1])
                coords.append((lat, lon))
            except Exception:
                pass
    return coords


def extract_kml_text_from_file(path: Path) -> str:
    if path.suffix.lower() == ".kml":
        return path.read_text(encoding="utf-8", errors="ignore")
    if path.suffix.lower() == ".kmz":
        with zipfile.ZipFile(path, "r") as zf:
            names = [n for n in zf.namelist() if n.lower().endswith(".kml")]
            if not names:
                raise ValueError("El KMZ no contiene KML interno.")
            with zf.open(names[0]) as f:
                return f.read().decode("utf-8", errors="ignore")
    raise ValueError("Formato no soportado.")


def load_kml_features(path: Path) -> List[Dict]:
    text = extract_kml_text_from_file(path)
    root = ET.fromstring(text)
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    placemarks = []

    for pm in root.findall(".//kml:Placemark", ns):
        name = pm.findtext("kml:name", default="", namespaces=ns)
        desc = pm.findtext("kml:description", default="", namespaces=ns)
        point = pm.find(".//kml:Point/kml:coordinates", ns)
        line = pm.find(".//kml:LineString/kml:coordinates", ns)
        poly = pm.find(".//kml:Polygon//kml:coordinates", ns)

        if point is not None and point.text:
            coords = parse_kml_coordinates(point.text)
            if coords:
                placemarks.append({"type": "Point", "name": safe_text(name), "description": safe_text(desc), "coordinates": coords})
        elif line is not None and line.text:
            coords = parse_kml_coordinates(line.text)
            if coords:
                placemarks.append({"type": "LineString", "name": safe_text(name), "description": safe_text(desc), "coordinates": coords})
        elif poly is not None and poly.text:
            coords = parse_kml_coordinates(poly.text)
            if coords:
                placemarks.append({"type": "Polygon", "name": safe_text(name), "description": safe_text(desc), "coordinates": coords})

    return placemarks


class TacticalMapApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(APP_SIZE)
        self.root.minsize(1320, 820)

        self.data_path: Optional[Path] = None
        self.sheet_name: Optional[str] = None
        self.df_work: Optional[pd.DataFrame] = None

        self.catalog_path: Optional[Path] = None
        self.catalog_sheet_name: Optional[str] = None
        self.df_catalog: Optional[pd.DataFrame] = None

        self.kml_paths: List[Path] = []
        self.worker_running = False

        self._build_ui()

    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Sub.TLabel", font=("Segoe UI", 10))

        # Contenedor con scroll vertical para que se vea completo en pantallas pequeñas.
        container = ttk.Frame(self.root)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, highlightthickness=0)
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)

        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        outer = ttk.Frame(canvas, padding=12)
        window_id = canvas.create_window((0, 0), window=outer, anchor="nw")

        def _on_frame_configure(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            # Mantiene el contenido al ancho de la ventana cuando sea posible.
            canvas.itemconfigure(window_id, width=max(event.width, 1320))

        def _on_mousewheel(event):
            # Windows / Mac
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _on_shift_mousewheel(event):
            canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

        outer.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)

        # Soporte Linux para rueda del mouse
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-3, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(3, "units"))

        ttk.Label(outer, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Carga Excel/CSV con coordenadas o cruza por municipio-colonia con un catálogo geográfico.",
            style="Sub.TLabel"
        ).pack(anchor="w", pady=(0, 10))

        main = ttk.Frame(outer)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        right = ttk.Frame(main)
        right.pack(side="left", fill="both", expand=True)

        self._build_left(left)
        self._build_right(right)

    def _build_left(self, parent):
        file_frame = ttk.LabelFrame(parent, text="1. Archivo principal")
        file_frame.pack(fill="x", pady=(0, 8))

        row = ttk.Frame(file_frame)
        row.pack(fill="x", padx=8, pady=8)
        self.path_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.path_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="Seleccionar Excel / CSV", command=self.select_file).pack(side="left")

        row2 = ttk.Frame(file_frame)
        row2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(row2, text="Hoja:").pack(side="left")
        self.sheet_combo = ttk.Combobox(row2, state="readonly", width=46)
        self.sheet_combo.pack(side="left", padx=(8, 8))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.load_selected_data())
        ttk.Button(row2, text="Cargar", command=self.load_selected_data).pack(side="left")

        cat_frame = ttk.LabelFrame(parent, text="2. Catálogo opcional de colonias")
        cat_frame.pack(fill="x", pady=(0, 8))

        rowc = ttk.Frame(cat_frame)
        rowc.pack(fill="x", padx=8, pady=8)
        self.catalog_path_var = tk.StringVar()
        ttk.Entry(rowc, textvariable=self.catalog_path_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(rowc, text="Seleccionar catálogo", command=self.select_catalog_file).pack(side="left")

        rowc2 = ttk.Frame(cat_frame)
        rowc2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(rowc2, text="Hoja catálogo:").pack(side="left")
        self.catalog_sheet_combo = ttk.Combobox(rowc2, state="readonly", width=36)
        self.catalog_sheet_combo.pack(side="left", padx=(8, 8))
        self.catalog_sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.load_catalog_data())
        ttk.Button(rowc2, text="Cargar catálogo", command=self.load_catalog_data).pack(side="left")

        cols_frame = ttk.LabelFrame(parent, text="3. Columnas del archivo principal")
        cols_frame.pack(fill="x", pady=(0, 8))
        grid = ttk.Frame(cols_frame)
        grid.pack(fill="x", padx=8, pady=8)

        specs = [
            ("Latitud", "lat_combo", 0, 0), ("Longitud", "lon_combo", 0, 2),
            ("Categoría", "category_combo", 1, 0), ("Popup principal", "popup_combo", 1, 2),
            ("Fecha", "date_combo", 2, 0), ("Municipio", "municipio_combo", 2, 2),
            ("Colonia", "colonia_combo", 3, 0), ("Sector", "sector_combo", 3, 2),
            ("Cuadrante", "cuadrante_combo", 4, 0), ("Año", "anio_combo", 4, 2),
            ("Mes", "mes_combo", 5, 0),
        ]
        for label, attr, r, c in specs:
            ttk.Label(grid, text=label).grid(row=r, column=c, sticky="w", padx=(0, 8), pady=5)
            combo = ttk.Combobox(grid, state="readonly", width=34)
            setattr(self, attr, combo)
            combo.grid(row=r, column=c + 1, sticky="ew", pady=5)
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

        cat_cols_frame = ttk.LabelFrame(parent, text="4. Columnas del catálogo")
        cat_cols_frame.pack(fill="x", pady=(0, 8))
        gridc = ttk.Frame(cat_cols_frame)
        gridc.pack(fill="x", padx=8, pady=8)

        cat_specs = [
            ("Municipio catálogo", "cat_municipio_combo", 0, 0),
            ("Colonia catálogo", "cat_colonia_combo", 0, 2),
            ("Latitud catálogo", "cat_lat_combo", 1, 0),
            ("Longitud catálogo", "cat_lon_combo", 1, 2),
            ("Sector catálogo", "cat_sector_combo", 2, 0),
            ("Cuadrante catálogo", "cat_cuadrante_combo", 2, 2),
        ]
        for label, attr, r, c in cat_specs:
            ttk.Label(gridc, text=label).grid(row=r, column=c, sticky="w", padx=(0, 8), pady=5)
            combo = ttk.Combobox(gridc, state="readonly", width=34)
            setattr(self, attr, combo)
            combo.grid(row=r, column=c + 1, sticky="ew", pady=5)
        gridc.columnconfigure(1, weight=1)
        gridc.columnconfigure(3, weight=1)

        opt_frame = ttk.LabelFrame(parent, text="5. Opciones tácticas")
        opt_frame.pack(fill="x", pady=(0, 8))
        grid2 = ttk.Frame(opt_frame)
        grid2.pack(fill="x", padx=8, pady=8)

        ttk.Label(grid2, text="Modo").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.map_mode_var = tk.StringVar(value="Automático")
        self.map_mode_combo = ttk.Combobox(
            grid2, state="readonly", textvariable=self.map_mode_var,
            values=["Automático", "Marcadores", "Cluster", "Calor", "Marcadores + Cluster", "Todo"],
            width=34
        )
        self.map_mode_combo.grid(row=0, column=1, sticky="ew", pady=5)

        ttk.Label(grid2, text="Zoom").grid(row=0, column=2, sticky="w", padx=(15, 8), pady=5)
        self.zoom_var = tk.StringVar(value="11")
        ttk.Entry(grid2, textvariable=self.zoom_var, width=12).grid(row=0, column=3, sticky="w", pady=5)

        ttk.Label(grid2, text="Radio calor").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.heat_radius_var = tk.StringVar(value="22")
        ttk.Entry(grid2, textvariable=self.heat_radius_var, width=12).grid(row=1, column=1, sticky="w", pady=5)

        ttk.Label(grid2, text="Límite tabla").grid(row=1, column=2, sticky="w", padx=(15, 8), pady=5)
        self.table_limit_var = tk.StringVar(value="150")
        ttk.Entry(grid2, textvariable=self.table_limit_var, width=12).grid(row=1, column=3, sticky="w", pady=5)
        grid2.columnconfigure(1, weight=1)

        self.open_browser_var = tk.BooleanVar(value=True)
        self.include_popup_details_var = tk.BooleanVar(value=False)
        self.color_by_category_var = tk.BooleanVar(value=True)
        self.enable_sampling_var = tk.BooleanVar(value=True)
        self.enable_geocode_catalog_var = tk.BooleanVar(value=True)

        for text, var in [
            ("Abrir HTML automáticamente", self.open_browser_var),
            ("Popup detallado", self.include_popup_details_var),
            ("Colorear por categoría", self.color_by_category_var),
            ("Muestreo inteligente para bases masivas", self.enable_sampling_var),
            ("Usar catálogo si faltan coordenadas", self.enable_geocode_catalog_var),
        ]:
            ttk.Checkbutton(opt_frame, text=text, variable=var).pack(anchor="w", padx=8)

        kml_frame = ttk.LabelFrame(parent, text="6. Capas KML / KMZ")
        kml_frame.pack(fill="x", pady=(0, 8))
        kml_btns = ttk.Frame(kml_frame)
        kml_btns.pack(fill="x", padx=8, pady=8)
        ttk.Button(kml_btns, text="Agregar KML/KMZ", command=self.add_kml_files).pack(side="left")
        ttk.Button(kml_btns, text="Limpiar capas", command=self.clear_kml_files).pack(side="left", padx=8)
        self.kml_list = tk.Listbox(kml_frame, height=3)
        self.kml_list.pack(fill="x", padx=8, pady=(0, 8))

        action_frame = ttk.Frame(parent)
        action_frame.pack(fill="x", pady=(4, 0))
        self.preview_btn = ttk.Button(action_frame, text="Vista previa", command=self.preview_dataset)
        self.preview_btn.pack(side="left")
        self.generate_btn = ttk.Button(action_frame, text="Generar mapa táctico", command=self.start_generate_map)
        self.generate_btn.pack(side="left", padx=8)
        self.reset_btn = ttk.Button(action_frame, text="Limpiar todo", command=self.reset_form)
        self.reset_btn.pack(side="left")

        prog_frame = ttk.LabelFrame(parent, text="7. Progreso")
        prog_frame.pack(fill="x", pady=(8, 0))
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=8, pady=(8, 6))
        self.progress_label_var = tk.StringVar(value="Esperando proceso...")
        ttk.Label(prog_frame, textvariable=self.progress_label_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=8, pady=(0, 8))

    def _build_right(self, parent):
        status_frame = ttk.LabelFrame(parent, text="Resumen")
        status_frame.pack(fill="x", pady=(0, 8))
        self.status_var = tk.StringVar(value="Carga un Excel o CSV para comenzar.")
        ttk.Label(status_frame, textvariable=self.status_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=8, pady=8)

        columns_frame = ttk.LabelFrame(parent, text="Columnas detectadas")
        columns_frame.pack(fill="both", expand=True, pady=(0, 8))
        self.columns_list = tk.Listbox(columns_frame, height=18)
        self.columns_list.pack(fill="both", expand=True, padx=8, pady=8)

        preview_frame = ttk.LabelFrame(parent, text="Vista previa")
        preview_frame.pack(fill="both", expand=True)
        self.preview_text = tk.Text(preview_frame, wrap="none", height=24)
        self.preview_text.pack(fill="both", expand=True, padx=8, pady=8)

    def set_busy(self, is_busy: bool):
        state = "disabled" if is_busy else "normal"
        for w in [self.preview_btn, self.generate_btn, self.reset_btn, self.sheet_combo, self.catalog_sheet_combo]:
            try:
                w.configure(state=state)
            except Exception:
                pass
        self.worker_running = is_busy

    def update_progress(self, percent: float, text: str):
        self.progress_var.set(max(0, min(100, percent)))
        self.progress_label_var.set(text)
        self.root.update_idletasks()

    def update_progress_async(self, percent: float, text: str):
        self.root.after(0, lambda: self.update_progress(percent, text))

    def read_data(self, path: Path, sheet_name: Optional[str]) -> pd.DataFrame:
        if path.suffix.lower() == ".csv":
            sep = detect_csv_separator(path)
            for enc in ["utf-8-sig", "utf-8", "latin-1", "cp1252"]:
                try:
                    return pd.read_csv(path, sep=sep, encoding=enc, low_memory=False)
                except Exception:
                    pass
            return pd.read_csv(path, sep=sep, low_memory=False)
        return pd.read_excel(path, sheet_name=sheet_name)

    def select_file(self):
        if self.worker_running:
            return
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo principal",
            filetypes=[("Excel y CSV", "*.xlsx *.xls *.csv"), ("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Todos", "*.*")]
        )
        if not file_path:
            return
        self.data_path = Path(file_path)
        self.path_var.set(str(self.data_path))

        if self.data_path.suffix.lower() == ".csv":
            self.sheet_combo["values"] = ["CSV"]
            self.sheet_combo.set("CSV")
            self.load_selected_data()
            return

        xl = pd.ExcelFile(self.data_path)
        self.sheet_combo["values"] = xl.sheet_names
        self.sheet_combo.set(xl.sheet_names[0])
        self.load_selected_data()

    def load_selected_data(self):
        if not self.data_path:
            messagebox.showwarning("Aviso", "Primero selecciona un archivo.")
            return
        sheet = self.sheet_combo.get().strip() or None
        try:
            df = self.read_data(self.data_path, None if self.data_path.suffix.lower() == ".csv" else sheet)
            if df.empty:
                messagebox.showwarning("Aviso", "La hoja está vacía.")
                return
            self.df_work = df
            self.sheet_name = "CSV" if self.data_path.suffix.lower() == ".csv" else sheet
            self._fill_main_controls([str(c) for c in df.columns])
            self._update_preview()
            self.status_var.set(f"Archivo cargado | Registros: {len(df):,} | Columnas: {len(df.columns)} | Fuente: {self.sheet_name}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo.\n\n{e}")

    def select_catalog_file(self):
        if self.worker_running:
            return
        file_path = filedialog.askopenfilename(
            title="Seleccionar catálogo",
            filetypes=[("Excel y CSV", "*.xlsx *.xls *.csv"), ("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Todos", "*.*")]
        )
        if not file_path:
            return
        self.catalog_path = Path(file_path)
        self.catalog_path_var.set(str(self.catalog_path))

        if self.catalog_path.suffix.lower() == ".csv":
            self.catalog_sheet_combo["values"] = ["CSV"]
            self.catalog_sheet_combo.set("CSV")
            self.load_catalog_data()
            return

        xl = pd.ExcelFile(self.catalog_path)
        self.catalog_sheet_combo["values"] = xl.sheet_names
        self.catalog_sheet_combo.set(xl.sheet_names[0])
        self.load_catalog_data()

    def load_catalog_data(self):
        if not self.catalog_path:
            messagebox.showwarning("Aviso", "Primero selecciona un catálogo.")
            return
        sheet = self.catalog_sheet_combo.get().strip() or None
        try:
            df = self.read_data(self.catalog_path, None if self.catalog_path.suffix.lower() == ".csv" else sheet)
            if df.empty:
                messagebox.showwarning("Aviso", "El catálogo está vacío.")
                return
            self.df_catalog = df
            self.catalog_sheet_name = "CSV" if self.catalog_path.suffix.lower() == ".csv" else sheet
            self._fill_catalog_controls([str(c) for c in df.columns])
            messagebox.showinfo("Catálogo cargado", f"Catálogo cargado.\n\nRegistros: {len(df):,}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el catálogo.\n\n{e}")

    def _fill_main_controls(self, columns: List[str]):
        values = [""] + columns
        combos = [
            self.lat_combo, self.lon_combo, self.category_combo, self.popup_combo,
            self.date_combo, self.municipio_combo, self.colonia_combo, self.sector_combo,
            self.cuadrante_combo, self.anio_combo, self.mes_combo
        ]
        for combo in combos:
            combo["values"] = values
            combo.set("")

        self.lat_combo.set(suggest_column(columns, LAT_CANDIDATES) or "")
        self.lon_combo.set(suggest_column(columns, LON_CANDIDATES) or "")
        self.category_combo.set(suggest_column(columns, CATEGORY_CANDIDATES) or "")
        self.popup_combo.set(suggest_column(columns, CATEGORY_CANDIDATES) or "")
        self.date_combo.set(suggest_column(columns, DATE_CANDIDATES) or "")
        self.municipio_combo.set(suggest_column(columns, MUNICIPIO_CANDIDATES) or "")
        self.colonia_combo.set(suggest_column(columns, COLONIA_CANDIDATES) or "")
        self.sector_combo.set(suggest_column(columns, SECTOR_CANDIDATES) or "")
        self.cuadrante_combo.set(suggest_column(columns, CUADRANTE_CANDIDATES) or "")
        self.anio_combo.set(suggest_column(columns, ANIO_CANDIDATES) or "")
        self.mes_combo.set(suggest_column(columns, MES_CANDIDATES) or "")

        self.columns_list.delete(0, tk.END)
        for col in columns:
            self.columns_list.insert(tk.END, col)

    def _fill_catalog_controls(self, columns: List[str]):
        values = [""] + columns
        combos = [
            self.cat_municipio_combo, self.cat_colonia_combo, self.cat_lat_combo,
            self.cat_lon_combo, self.cat_sector_combo, self.cat_cuadrante_combo
        ]
        for combo in combos:
            combo["values"] = values
            combo.set("")

        self.cat_municipio_combo.set(suggest_column(columns, MUNICIPIO_CANDIDATES) or "")
        self.cat_colonia_combo.set(suggest_column(columns, COLONIA_CANDIDATES) or "")
        self.cat_lat_combo.set(suggest_column(columns, LAT_CANDIDATES) or "")
        self.cat_lon_combo.set(suggest_column(columns, LON_CANDIDATES) or "")
        self.cat_sector_combo.set(suggest_column(columns, SECTOR_CANDIDATES) or "")
        self.cat_cuadrante_combo.set(suggest_column(columns, CUADRANTE_CANDIDATES) or "")

    def _update_preview(self):
        self.preview_text.delete("1.0", tk.END)
        if self.df_work is None:
            return
        self.preview_text.insert("1.0", self.df_work.head(12).to_string(index=False))

    def add_kml_files(self):
        files = filedialog.askopenfilenames(title="Seleccionar KML/KMZ", filetypes=[("KML / KMZ", "*.kml *.kmz"), ("Todos", "*.*")])
        for f in files:
            p = Path(f)
            if p not in self.kml_paths:
                self.kml_paths.append(p)
        self.refresh_kml_list()

    def clear_kml_files(self):
        self.kml_paths = []
        self.refresh_kml_list()

    def refresh_kml_list(self):
        self.kml_list.delete(0, tk.END)
        for p in self.kml_paths:
            self.kml_list.insert(tk.END, p.name)

    def preview_dataset(self):
        if self.df_work is None:
            messagebox.showwarning("Aviso", "Primero carga un archivo.")
            return
        messagebox.showinfo(
            "Vista previa",
            f"Registros principales: {len(self.df_work):,}\n"
            f"Catálogo cargado: {'Sí' if self.df_catalog is not None else 'No'}\n"
            f"KML/KMZ: {len(self.kml_paths)}"
        )

    def reset_form(self):
        self.data_path = None
        self.sheet_name = None
        self.df_work = None
        self.catalog_path = None
        self.df_catalog = None
        self.kml_paths = []
        self.path_var.set("")
        self.catalog_path_var.set("")
        self.sheet_combo.set("")
        self.catalog_sheet_combo.set("")
        self.columns_list.delete(0, tk.END)
        self.preview_text.delete("1.0", tk.END)
        self.kml_list.delete(0, tk.END)
        self.status_var.set("Carga un Excel o CSV para comenzar.")
        self.update_progress(0, "Esperando proceso...")

    def catalog_geocode(self, work: pd.DataFrame) -> Tuple[pd.DataFrame, int, int]:
        if self.df_catalog is None or self.df_catalog.empty:
            return work, 0, 0

        municipio_col = self.municipio_combo.get().strip()
        colonia_col = self.colonia_combo.get().strip()
        cat_mun = self.cat_municipio_combo.get().strip()
        cat_col = self.cat_colonia_combo.get().strip()
        cat_lat = self.cat_lat_combo.get().strip()
        cat_lon = self.cat_lon_combo.get().strip()
        cat_sec = self.cat_sector_combo.get().strip()
        cat_cua = self.cat_cuadrante_combo.get().strip()

        if not municipio_col or not colonia_col or not cat_mun or not cat_col or not cat_lat or not cat_lon:
            return work, 0, 0

        cat = self.df_catalog.copy()
        cat["__KEY__"] = cat[cat_mun].map(normalize_key) + "|" + cat[cat_col].map(normalize_key)
        cat["__CAT_LAT__"] = pd.to_numeric(cat[cat_lat], errors="coerce")
        cat["__CAT_LON__"] = pd.to_numeric(cat[cat_lon], errors="coerce")
        cat = cat.dropna(subset=["__CAT_LAT__", "__CAT_LON__"]).drop_duplicates("__KEY__")

        keep = ["__KEY__", "__CAT_LAT__", "__CAT_LON__"]
        if cat_sec:
            cat["__CAT_SECTOR__"] = cat[cat_sec]
            keep.append("__CAT_SECTOR__")
        if cat_cua:
            cat["__CAT_CUADRANTE__"] = cat[cat_cua]
            keep.append("__CAT_CUADRANTE__")
        cat = cat[keep]

        work["__KEY__"] = work[municipio_col].map(normalize_key) + "|" + work[colonia_col].map(normalize_key)
        missing_before = work["__LAT__"].isna() | work["__LON__"].isna()
        merged = work.merge(cat, on="__KEY__", how="left")

        can_fill = missing_before & merged["__CAT_LAT__"].notna() & merged["__CAT_LON__"].notna()
        merged.loc[can_fill, "__LAT__"] = merged.loc[can_fill, "__CAT_LAT__"]
        merged.loc[can_fill, "__LON__"] = merged.loc[can_fill, "__CAT_LON__"]
        merged.loc[can_fill, "__GEO_STATUS__"] = "Sin coordenadas (geocodificado)"

        if "__CAT_SECTOR__" in merged.columns:
            mask = merged["__SECTOR__"].astype(str).str.strip().eq("")
            merged.loc[mask & merged["__CAT_SECTOR__"].notna(), "__SECTOR__"] = merged.loc[mask & merged["__CAT_SECTOR__"].notna(), "__CAT_SECTOR__"]

        if "__CAT_CUADRANTE__" in merged.columns:
            mask = merged["__CUADRANTE__"].astype(str).str.strip().eq("")
            merged.loc[mask & merged["__CAT_CUADRANTE__"].notna(), "__CUADRANTE__"] = merged.loc[mask & merged["__CAT_CUADRANTE__"].notna(), "__CAT_CUADRANTE__"]

        geocoded = int(can_fill.sum())
        not_found = int((missing_before & (merged["__LAT__"].isna() | merged["__LON__"].isna())).sum())
        merged = merged.drop(columns=[c for c in ["__KEY__", "__CAT_LAT__", "__CAT_LON__", "__CAT_SECTOR__", "__CAT_CUADRANTE__"] if c in merged.columns], errors="ignore")
        return merged, geocoded, not_found

    def prepare_dataframe(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict]:
        lat_col = self.lat_combo.get().strip()
        lon_col = self.lon_combo.get().strip()
        date_col = self.date_combo.get().strip()
        anio_col = self.anio_combo.get().strip()
        mes_col = self.mes_combo.get().strip()
        sector_col = self.sector_combo.get().strip()
        cuadrante_col = self.cuadrante_combo.get().strip()

        work = df.copy()
        work["__FILA_EXCEL__"] = work.index + 2

        if lat_col and lon_col:
            work["__LAT__"] = pd.to_numeric(work[lat_col], errors="coerce")
            work["__LON__"] = pd.to_numeric(work[lon_col], errors="coerce")
        else:
            work["__LAT__"] = pd.NA
            work["__LON__"] = pd.NA

        work["__GEO_STATUS__"] = "Con coordenadas"
        work.loc[work["__LAT__"].isna() | work["__LON__"].isna(), "__GEO_STATUS__"] = "Sin coordenadas"
        work["__SECTOR__"] = work[sector_col] if sector_col else ""
        work["__CUADRANTE__"] = work[cuadrante_col] if cuadrante_col else ""

        geocoded, not_found = 0, 0
        if self.enable_geocode_catalog_var.get():
            work, geocoded, not_found = self.catalog_geocode(work)

        if date_col:
            dt = pd.to_datetime(work[date_col], errors="coerce", dayfirst=True)
            work["__ANIO__"] = dt.dt.year
            work["__MES__"] = dt.dt.month
            work["__FECHA_FMT__"] = dt.dt.strftime("%d/%m/%Y %H:%M").fillna("")
            work["__FECHA_ISO__"] = dt.dt.strftime("%Y-%m-%d").fillna("")
        else:
            work["__FECHA_FMT__"] = ""
            work["__FECHA_ISO__"] = ""

        if "__ANIO__" not in work.columns and anio_col:
            work["__ANIO__"] = pd.to_numeric(work[anio_col], errors="coerce")
        if "__MES__" not in work.columns and mes_col:
            work["__MES__"] = work[mes_col].astype(str).map(month_name_to_number)

        valid = work["__LAT__"].notna() & work["__LON__"].notna()
        work["__FUERA_RANGO__"] = False
        work.loc[valid, "__FUERA_RANGO__"] = (
            (work.loc[valid, "__LAT__"] < LAT_MIN) |
            (work.loc[valid, "__LAT__"] > LAT_MAX) |
            (work.loc[valid, "__LON__"] < LON_MIN) |
            (work.loc[valid, "__LON__"] > LON_MAX)
        )
        work.loc[valid & work["__FUERA_RANGO__"], "__GEO_STATUS__"] = "Fuera de rango"

        stats = {
            "total": len(work),
            "con_coordenadas": int(valid.sum()),
            "sin_coordenadas": int((~valid).sum()),
            "fuera_rango": int(work["__FUERA_RANGO__"].sum()),
            "geocodificados": geocoded,
            "no_encontrados": not_found,
        }
        return work, stats

    def build_popup_html(self, row: Dict, popup_col: str, category_col: str, municipio_col: str, colonia_col: str) -> str:
        title = safe_text(row.get(popup_col, "")) if popup_col else ""
        cat = safe_text(row.get(category_col, "")) if category_col else ""
        municipio = safe_text(row.get(municipio_col, "")) if municipio_col else ""
        colonia = safe_text(row.get(colonia_col, "")) if colonia_col else ""
        status = safe_text(row.get("__GEO_STATUS__", ""))
        color = "#ef4444" if status == "Fuera de rango" else ("#f59e0b" if status.startswith("Sin coordenadas") else "#22c55e")

        html = f"""
        <div style="font-family:Segoe UI,Arial,sans-serif;min-width:260px;background:#020617;color:#e5e7eb;padding:6px;">
            <div style="font-weight:800;font-size:14px;color:{color};margin-bottom:8px">{escape(title or cat or "REGISTRO")}</div>
            <b>Fila Excel:</b> {escape(safe_text(row.get("__FILA_EXCEL__", "")))}<br>
            <b>Tipo:</b> {escape(cat)}<br>
            <b>Municipio:</b> {escape(municipio)}<br>
            <b>Colonia:</b> {escape(colonia)}<br>
            <b>Sector:</b> {escape(safe_text(row.get("__SECTOR__", "")))}<br>
            <b>Cuadrante:</b> {escape(safe_text(row.get("__CUADRANTE__", "")))}<br>
            <b>Latitud:</b> {row.get("__LAT__", "")}<br>
            <b>Longitud:</b> {row.get("__LON__", "")}<br>
            <b>Estatus:</b> <span style="color:{color};font-weight:800">{escape(status)}</span>
        """
        if self.include_popup_details_var.get():
            html += "<hr style='border-color:#334155;'>"
            shown = 0
            for col, val in row.items():
                if str(col).startswith("__"):
                    continue
                txt = safe_text(val)
                if txt:
                    html += f"<b>{escape(str(col))}:</b> {escape(txt)}<br>"
                    shown += 1
                    if shown >= 8:
                        break
        return html + "</div>"

    def get_auto_mode(self, total: int) -> str:
        if total <= SMALL_DATASET_LIMIT:
            return "Todo"
        if total <= MEDIUM_DATASET_LIMIT:
            return "Marcadores + Cluster"
        if total <= LARGE_DATASET_LIMIT:
            return "Cluster"
        return "Calor"

    def make_sample(self, df: pd.DataFrame, size: int) -> pd.DataFrame:
        if len(df) <= size:
            return df
        return df.sample(n=size, random_state=42)

    def add_kml_layers_to_map(self, mapa: folium.Map):
        total = len(self.kml_paths)
        for idx, path in enumerate(self.kml_paths, start=1):
            self.update_progress_async(78 + (idx / max(total, 1)) * 8, f"Cargando KML/KMZ {idx} de {total}: {path.name}")
            try:
                features = load_kml_features(path)
                color = random_color_from_text(path.stem)
                fg = FeatureGroup(name=f"🗂️ {path.stem}", show=True)
                for item in features:
                    name = item.get("name", path.stem)
                    desc = item.get("description", "")
                    popup = folium.Popup(f"<b>{escape(name)}</b><br>{escape(desc)}", max_width=280)

                    if item["type"] == "Point":
                        lat, lon = item["coordinates"][0]
                        folium.CircleMarker([lat, lon], radius=4, color=color, fill=True, fill_color=color, fill_opacity=0.9, popup=popup, tooltip=name).add_to(fg)
                    elif item["type"] == "LineString":
                        folium.PolyLine(item["coordinates"], color=color, weight=3, opacity=0.9, popup=popup, tooltip=name).add_to(fg)
                    elif item["type"] == "Polygon":
                        folium.Polygon(item["coordinates"], color=color, weight=2, fill=True, fill_color=color, fill_opacity=0.16, popup=popup, tooltip=name).add_to(fg)
                fg.add_to(mapa)
            except Exception:
                pass

    def tactical_panels_html(self, stats, counts_cat, counts_mun, table_rows, source_name, mode) -> str:
        total = max(stats.get("total", 0), 1)
        con = stats.get("con_coordenadas", 0)
        sin = stats.get("sin_coordenadas", 0)
        fuera = stats.get("fuera_rango", 0)

        def pct(x):
            return f"{(x / total) * 100:.2f}%"

        cat_items = ""
        for name, qty in counts_cat.head(10).items():
            color = random_color_from_text(str(name))
            cat_items += f'<div class="legend-row"><span class="dot" style="background:{color}"></span><span>{escape(str(name))}</span><b>{qty:,}</b></div>'

        mun_items = ""
        max_mun = max(int(counts_mun.max()) if not counts_mun.empty else 1, 1)
        for name, qty in counts_mun.head(7).items():
            width = max(4, int((qty / max_mun) * 100))
            mun_items += f'<div class="bar-row"><span>{escape(str(name))}</span><div><i style="width:{width}%"></i></div><b>{qty:,}</b></div>'

        table_html = ""
        for r in table_rows:
            status = safe_text(r.get("status", ""))
            cls = "ok" if status == "Con coordenadas" else ("warn" if status.startswith("Sin coordenadas") else "bad")
            table_html += f"""
            <tr>
                <td>{escape(safe_text(r.get("fila", "")))}</td>
                <td>{escape(safe_text(r.get("fecha", "")))}</td>
                <td>{escape(safe_text(r.get("tipo", "")))}</td>
                <td>{escape(safe_text(r.get("municipio", "")))}</td>
                <td>{escape(safe_text(r.get("colonia", "")))}</td>
                <td>{escape(safe_text(r.get("sector", "")))}</td>
                <td>{escape(safe_text(r.get("lat", "")))}</td>
                <td>{escape(safe_text(r.get("lon", "")))}</td>
                <td class="{cls}">{escape(status)}</td>
            </tr>
            """

        now = datetime.now().strftime("%d/%m/%Y<br>%H:%M:%S")
        return f"""
        <style>
            html, body {{ background:#020617 !important; }}
            .leaflet-container {{ background:#020617 !important; font-family:Segoe UI,Arial,sans-serif; }}
            .c5-top {{ position:fixed; top:8px; left:8px; right:300px; height:94px; z-index:99990; display:grid; grid-template-columns:220px repeat(5,1fr); gap:8px; pointer-events:none; }}
            .c5-card {{ background:rgba(2,6,23,.92); border:1px solid rgba(148,163,184,.28); border-radius:10px; color:#e5e7eb; box-shadow:0 10px 28px rgba(0,0,0,.35); padding:12px; pointer-events:auto; }}
            .brand {{ display:flex; align-items:center; gap:12px; }}
            .shield {{ width:55px; height:65px; border:2px solid #dbeafe; border-radius:10px 10px 22px 22px; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:800; }}
            .brand h1 {{ margin:0; font-size:28px; line-height:1; }}
            .brand p {{ margin:4px 0 0; font-size:13px; letter-spacing:.7px; }}
            .metric small {{ display:block; color:#cbd5e1; font-size:11px; text-transform:uppercase; text-align:center; }}
            .metric b {{ display:block; font-size:26px; text-align:center; margin-top:4px; }}
            .metric span {{ display:block; font-size:12px; text-align:center; font-weight:800; margin-top:3px; }}
            .blue {{ color:#3b82f6; }} .green {{ color:#22c55e; }} .yellow {{ color:#f59e0b; }} .red {{ color:#ef4444; }}
            .left-panel {{ position:fixed; top:112px; left:10px; width:300px; bottom:16px; z-index:99998; display:flex; flex-direction:column; gap:8px; pointer-events:none; }}
            .panel {{ background:rgba(2,6,23,.90); border:1px solid rgba(148,163,184,.25); border-radius:10px; color:#e5e7eb; box-shadow:0 10px 28px rgba(0,0,0,.35); padding:12px; pointer-events:auto; }}
            .panel h3 {{ margin:0 0 10px; color:#60a5fa; font-size:14px; text-transform:uppercase; }}
            .filter-box label {{ display:block; color:#cbd5e1; font-size:11px; margin:8px 0 4px; }}
            .filter-box select, .filter-box input {{ width:100%; box-sizing:border-box; background:#0f172a; color:#e5e7eb; border:1px solid #334155; border-radius:6px; padding:7px; }}
            .btn-filter {{ width:100%; background:#2563eb; color:white; border:0; border-radius:6px; padding:10px; margin-top:10px; font-weight:800; }}
            .legend-row {{ display:grid; grid-template-columns:18px 1fr auto; align-items:center; gap:6px; font-size:12px; margin:5px 0; }}
            .dot {{ width:11px; height:11px; border-radius:50%; display:inline-block; }}
            .bar-row {{ display:grid; grid-template-columns:88px 1fr 46px; align-items:center; gap:8px; font-size:11px; margin:8px 0; }}
            .bar-row div {{ height:9px; background:#0f172a; border-radius:10px; overflow:hidden; }}
            .bar-row i {{ height:100%; display:block; background:#2563eb; }}
            .alert-box {{ background:rgba(127,29,29,.35); border-color:rgba(248,113,113,.35); }}
            .alert-box div {{ color:#fca5a5; margin:8px 0; font-size:13px; font-weight:700; }}
            .right-panel {{ display:none !important; position:fixed; top:180px; right:10px; width:275px; z-index:99980; display:flex; flex-direction:column; gap:10px; pointer-events:none; }}
            .right-panel .panel {{ pointer-events:auto; }}
            .layers label {{ display:block; margin:9px 0; font-size:13px; }}
            .ref-row {{ margin:8px 0; font-size:13px; }}
            .bottom-table {{ position:fixed; left:320px; right:10px; bottom:8px; height:190px; z-index:99997; background:rgba(2,6,23,.88); border:1px solid rgba(148,163,184,.25); border-radius:10px; color:#e5e7eb; box-shadow:0 10px 28px rgba(0,0,0,.35); overflow:auto; font-family:Segoe UI,Arial,sans-serif; }}
            .bottom-table h3 {{ margin:10px 14px; color:#60a5fa; font-size:14px; }}
            table {{ width:100%; border-collapse:collapse; font-size:12px; }}
            th, td {{ border-top:1px solid rgba(148,163,184,.17); border-right:1px solid rgba(148,163,184,.12); padding:6px 10px; white-space:nowrap; }}
            th {{ color:#cbd5e1; font-weight:600; }}
            .ok {{ color:#22c55e; font-weight:800; }} .warn {{ color:#f59e0b; font-weight:800; }} .bad {{ color:#ef4444; font-weight:800; }}
        
            .leaflet-top.leaflet-right {{
                top: 112px !important;
                right: 10px !important;
                z-index: 100000 !important;
            }}
            .leaflet-control-layers {{
                background: rgba(2,6,23,.94) !important;
                color: #e5e7eb !important;
                border: 1px solid rgba(148,163,184,.35) !important;
                border-radius: 10px !important;
                box-shadow: 0 10px 28px rgba(0,0,0,.40) !important;
                font-family: Segoe UI, Arial, sans-serif !important;
            }}
            .leaflet-control-layers-expanded {{
                padding: 10px 12px !important;
            }}
            .leaflet-control-layers label {{
                color: #e5e7eb !important;
                font-size: 12px !important;
            }}

            .leaflet-control-layers {{
                min-width: 220px !important;
                background: rgba(2,6,23,.96) !important;
                border: 1px solid rgba(59,130,246,.35) !important;
                border-radius: 12px !important;
                box-shadow: 0 10px 30px rgba(0,0,0,.45) !important;
                padding: 8px !important;
            }}

            .leaflet-control-layers label {{
                padding: 4px 2px !important;
            }}

            .leaflet-control-layers-separator {{
                border-top: 1px solid rgba(148,163,184,.25) !important;
                margin: 8px 0 !important;
            }}


        
            .chart-dashboard {{
                display:grid;
                grid-template-columns: 1fr .75fr 1fr;
                gap:10px;
                padding:10px;
                box-sizing:border-box;
            }}
            .chart-card {{
                background:rgba(2,6,23,.64);
                border:1px solid rgba(148,163,184,.18);
                border-radius:10px;
                padding:10px;
                overflow:auto;
            }}
            .chart-card h3 {{
                margin:0 0 10px;
                color:#60a5fa;
                font-size:14px;
                text-transform:uppercase;
            }}
            .bar-chart-box {{
                display:flex;
                flex-direction:column;
                gap:7px;
                font-size:12px;
            }}
            .bar-chart-row {{
                display:grid;
                grid-template-columns: 210px 1fr 70px;
                gap:8px;
                align-items:center;
            }}
            .bar-chart-name {{
                overflow:hidden;
                text-overflow:ellipsis;
                white-space:nowrap;
                color:#e5e7eb;
                font-weight:700;
            }}
            .bar-chart-track {{
                height:16px;
                background:#0f172a;
                border-radius:999px;
                overflow:hidden;
                border:1px solid rgba(148,163,184,.15);
            }}
            .bar-chart-fill {{
                height:100%;
                border-radius:999px;
                min-width:4px;
            }}
            .bar-chart-value {{
                text-align:right;
                color:#e5e7eb;
                font-weight:900;
            }}
            .donut-wrap {{
                display:grid;
                grid-template-columns: 170px 1fr;
                gap:12px;
                align-items:center;
                min-height:135px;
            }}
            .donut-chart {{
                width:155px;
                height:155px;
                border-radius:50%;
                display:flex;
                align-items:center;
                justify-content:center;
                margin:auto;
                box-shadow:0 0 20px rgba(59,130,246,.18);
            }}
            .donut-chart::after {{
                content: attr(data-total);
                width:82px;
                height:82px;
                border-radius:50%;
                background:#020617;
                color:#e5e7eb;
                display:flex;
                align-items:center;
                justify-content:center;
                font-weight:900;
                font-size:18px;
                border:1px solid rgba(148,163,184,.2);
                white-space:pre;
                text-align:center;
            }}
            .donut-legend {{
                display:flex;
                flex-direction:column;
                gap:6px;
                font-size:12px;
                max-height:145px;
                overflow:auto;
            }}
            .donut-legend-row {{
                display:grid;
                grid-template-columns:14px 1fr auto;
                gap:7px;
                align-items:center;
            }}
            .donut-dot {{
                width:10px;
                height:10px;
                border-radius:50%;
                display:inline-block;
            }}
            .donut-name {{
                overflow:hidden;
                text-overflow:ellipsis;
                white-space:nowrap;
                color:#e5e7eb;
            }}
            .donut-value {{
                color:#cbd5e1;
                font-weight:900;
            }}
            .vertical-chart-box {{
                height:145px;
                display:flex;
                align-items:flex-end;
                gap:8px;
                overflow-x:auto;
                overflow-y:hidden;
                padding:8px 4px 2px;
            }}
            .vertical-bar-item {{
                min-width:48px;
                height:100%;
                display:flex;
                flex-direction:column;
                justify-content:flex-end;
                align-items:center;
                gap:5px;
            }}
            .vertical-bar-value {{
                color:#e5e7eb;
                font-size:11px;
                font-weight:900;
            }}
            .vertical-bar {{
                width:28px;
                min-height:4px;
                border-radius:7px 7px 2px 2px;
                box-shadow:0 0 14px rgba(59,130,246,.25);
            }}
            .vertical-bar-name {{
                width:62px;
                color:#cbd5e1;
                font-size:10px;
                font-weight:800;
                text-align:center;
                overflow:hidden;
                text-overflow:ellipsis;
                white-space:nowrap;
            }}

        </style>

        <div class="c5-top">
            <div class="c5-card brand"><div class="shield">C5 PRO</div><div><h1>C5 PRO</h1><p>MAPA TÁCTICO<br>QUINTANA ROO</p></div></div>
            <div class="c5-card metric"><small>Total registros</small><b>{stats.get("total", 0):,}</b><span class="blue">100%</span></div>
            <div class="c5-card metric"><small>Con coordenadas</small><b>{con:,}</b><span class="green">{pct(con)}</span></div>
            <div class="c5-card metric"><small>Sin coordenadas</small><b>{sin:,}</b><span class="yellow">{pct(sin)}</span></div>
            <div class="c5-card metric"><small>Fuera de rango</small><b class="red">{fuera:,}</b><span class="red">{pct(fuera)}</span></div>
            <div class="c5-card metric"><small>Registros filtrados</small><b id="metric_filtrados" style="font-size:26px">{con:,}</b><span id="metric_ubicaciones" class="blue">Ubicaciones: 0</span></div>
        </div>

        <div class="left-panel">
            <div class="panel filter-box">
                <h3>Filtros</h3>
                <label>Rango de fechas</label><div style="display:grid;grid-template-columns:1fr 1fr;gap:7px;"><input type="date" id="f_fecha_inicio"><input type="date" id="f_fecha_fin"></div>
                <label>Municipio</label>
                <select id="f_municipio"><option value="">Todos</option></select>
                <label>Sector</label>
                <select id="f_sector"><option value="">Todos</option></select>
                <label>Tipo de incidente</label>
                <select id="f_tipo"><option value="">Todos</option></select>
                <button class="btn-filter" id="btn_aplicar_filtros" type="button">Aplicar filtros</button>
                <button class="btn-filter" id="btn_limpiar_filtros" type="button" style="background:#475569;margin-top:6px;">Limpiar filtros</button>
                <div id="filter_count" style="margin-top:8px;color:#cbd5e1;font-size:12px;font-weight:700;">Registros visibles: cargando...</div>
                <label style="display:flex;align-items:center;gap:6px;margin-top:8px;color:#e5e7eb;font-size:12px;font-weight:700;">
                    <input type="checkbox" id="toggle_points_layer" checked>
                    Mostrar puntos / ubicaciones
                </label>
            </div>
            <div class="panel"><h3 id="top10_title">Tipo de incidente (Top 10)</h3><div id="top10_dynamic">{cat_items or "<small>Sin categoría detectada.</small>"}</div></div>
            <div class="panel"><h3>Registros por municipio</h3>{mun_items or "<small>Sin municipio detectado.</small>"}</div>
            <div class="panel alert-box"><h3 style="color:#f87171;">Alertas</h3><div>⚠ {fuera:,} registros fuera de rango geográfico</div><div>⚠ {sin:,} registros sin coordenadas</div></div>
        </div>
        <div class="right-panel">
            <div class="panel layers">
                <h3>Mapa generado</h3><div style="font-size:11px;color:#cbd5e1;line-height:1.45;margin-bottom:10px;">Fuente: {escape(source_name)}<br>Modo: {escape(mode)}<br>Geocodificados: {stats.get("geocodificados", 0):,}</div><h3>Mapas base</h3>
                <div style="font-size:12px;color:#cbd5e1;line-height:1.45;">
                    Cambia el fondo desde el control de capas del mapa.<br>
                    Opciones disponibles:<br>
                    🌑 táctico oscuro<br>
                    🛰️ satelital Esri<br>
                    🗺️ OpenStreetMap<br>
                    ⚪ claro institucional<br>
                    🌍 Voyager<br>
                    ⛰️ topográfico
                </div>
                <h3 style="margin-top:14px;">Capas activas</h3>
                <label>☑ Heatmap</label>
                <label>☑ Puntos / ubicaciones filtradas</label>
                <label>☑ KML/KMZ si se cargaron</label>
            </div>
            <div class="panel">
                <h3>Referencia</h3>
                <div class="ref-row"><span class="dot" style="background:#22c55e"></span> Registro normal</div>
                <div class="ref-row"><span class="dot" style="background:#f59e0b"></span> Geocodificado</div>
                <div class="ref-row"><span class="dot" style="background:#ef4444"></span> Fuera de rango</div>
            </div>
        </div>

        <div class="bottom-table chart-dashboard">
            <div class="chart-card">
                <h3>GRÁFICA DE BARRAS | TOP INCIDENTES</h3>
                <div id="bar_chart_dynamic" class="bar-chart-box"><small>Cargando...</small></div>
            </div>
            <div class="chart-card">
                <h3>GRÁFICA DE ANILLO | DISTRIBUCIÓN</h3>
                <div class="donut-wrap">
                    <div id="donut_chart_dynamic" class="donut-chart"></div>
                    <div id="donut_legend_dynamic" class="donut-legend"></div>
                </div>
            </div>
            <div class="chart-card">
                <h3>BARRAS VERTICALES | MUNICIPIOS</h3>
                <div id="mun_vertical_chart_dynamic" class="vertical-chart-box"><small>Cargando...</small></div>
            </div>
        </div>
        """


    def filter_script_html(self, map_name: str, records: List[Dict]) -> str:
        records_json = json.dumps(records, ensure_ascii=False).replace("</", "<\\/")

        return f"""
        <script type="application/json" id="c5_records_json">{records_json}</script>
        <script>
        (function(){{
            let mapRef = null;
            let records = [];
            let pointLayer = null;

            function boot(){{
                try{{
                    mapRef = {map_name};
                    const node = document.getElementById('c5_records_json');
                    records = JSON.parse(node.textContent || '[]');
                    pointLayer = L.layerGroup();

                    // Capa de puntos controlada desde el checkbox del panel de filtros.
                    mapRef.addLayer(pointLayer);

                    const togglePoints = document.getElementById('toggle_points_layer');
                    if (togglePoints) {{
                        togglePoints.checked = true;
                        togglePoints.addEventListener('change', function() {{
                            if (this.checked) {{
                                if (!mapRef.hasLayer(pointLayer)) {{
                                    mapRef.addLayer(pointLayer);
                                }}
                            }} else {{
                                if (mapRef.hasLayer(pointLayer)) {{
                                    mapRef.removeLayer(pointLayer);
                                }}
                            }}
                        }});
                    }}

                    fillSelect('f_municipio', unique('municipio'));
                    fillSelect('f_sector', unique('sector'));
                    fillSelect('f_tipo', unique('tipo'));
                    setupDateRange();

                    const btnApply = document.getElementById('btn_aplicar_filtros');
                    const btnClear = document.getElementById('btn_limpiar_filtros');
                    if(btnApply) btnApply.addEventListener('click', applyFilters);
                    if(btnClear) btnClear.addEventListener('click', clearFilters);

                    ['f_fecha_inicio','f_fecha_fin','f_municipio','f_sector','f_tipo'].forEach(id=>{{
                        const el = document.getElementById(id);
                        if(el) el.addEventListener('change', applyFilters);
                    }});

                    render(records, false);
                }}catch(err){{
                    console.error('Error C5 PRO:', err);
                    const c = document.getElementById('filter_count');
                    if(c) c.textContent = 'Filtros no disponibles: revisar consola';
                }}
            }}

            function setupDateRange(){{
                const fechas = records.map(r=>String(r.fecha_iso||'')).filter(v=>/^\\d{{4}}-\\d{{2}}-\\d{{2}}$/.test(v)).sort();
                if(!fechas.length) return;
                const ini = document.getElementById('f_fecha_inicio');
                const fin = document.getElementById('f_fecha_fin');
                if(ini){{ ini.min=fechas[0]; ini.max=fechas[fechas.length-1]; ini.value=fechas[0]; }}
                if(fin){{ fin.min=fechas[0]; fin.max=fechas[fechas.length-1]; fin.value=fechas[fechas.length-1]; }}
            }}

            function unique(field){{
                return [...new Set(records.map(r=>String(r[field]||'').trim()).filter(Boolean))]
                    .sort((a,b)=>a.localeCompare(b,'es'));
            }}

            function fillSelect(id, values){{
                const el = document.getElementById(id);
                if(!el) return;
                while(el.options.length > 1) el.remove(1);
                values.forEach(v=>{{
                    const opt=document.createElement('option');
                    opt.value=v; opt.textContent=v; el.appendChild(opt);
                }});
            }}

            function esc(t){{
                return String(t||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;');
            }}

            function colorText(text){{
                const palette=['#ef4444','#f97316','#f59e0b','#22c55e','#06b6d4','#3b82f6','#8b5cf6','#d946ef','#ec4899','#84cc16'];
                let h=0, s=String(text||'SIN DATO');
                for(let i=0;i<s.length;i++) h=s.charCodeAt(i)+((h<<5)-h);
                return palette[Math.abs(h)%palette.length];
            }}

            function colorStatus(r){{
                const s=String(r.status||'');
                if(s==='Fuera de rango') return '#ef4444';
                if(s.startsWith('Sin coordenadas')) return '#f59e0b';
                return r.color || colorText(r.tipo);
            }}

            function countBy(list, field){{
                const c={{}};
                list.forEach(r=>{{
                    const k=String(r[field]||'SIN DATO').trim()||'SIN DATO';
                    c[k]=(c[k]||0)+1;
                }});
                return Object.entries(c).sort((a,b)=>b[1]-a[1]);
            }}

            function updateTop10(list){{
                const box=document.getElementById('top10_dynamic');
                const title=document.getElementById('top10_title');
                if(!box) return;
                const fm=document.getElementById('f_municipio')?.value||'';
                const fs=document.getElementById('f_sector')?.value||'';
                if(title) title.textContent = fm && fs ? `Top 10: ${{fm}} / Sector ${{fs}}` : fm ? `Top 10: ${{fm}}` : fs ? `Top 10: Sector ${{fs}}` : 'Tipo de incidente (Top 10)';
                const entries=countBy(list,'tipo').slice(0,10);
                box.innerHTML = entries.length ? entries.map(([n,q])=>`<div class="legend-row"><span class="dot" style="background:${{list.find(r => String(r.tipo) === String(n))?.color || colorText(n)}}"></span><span>${{esc(n)}}</span><b>${{q.toLocaleString('es-MX')}}</b></div>`).join('') : '<small>Sin registros.</small>';
            }}

            function updateCharts(list){{
                const bars=document.getElementById('bar_chart_dynamic');
                const donut=document.getElementById('donut_chart_dynamic');
                const legend=document.getElementById('donut_legend_dynamic');
                const munBars=document.getElementById('mun_vertical_chart_dynamic');
                const entries=countBy(list,'tipo').slice(0,10);
                const max=entries.length ? Math.max(...entries.map(e=>e[1])) : 1;
                if(bars){{
                    bars.innerHTML = entries.length ? entries.map(([n,q])=>{{
                        const w=Math.max(3,Math.round(q/max*100));
                        const col=colorText(n);
                        return `<div class="bar-chart-row"><div class="bar-chart-name" title="${{esc(n)}}">${{esc(n)}}</div><div class="bar-chart-track"><div class="bar-chart-fill" style="width:${{w}}%;background:${{col}};"></div></div><div class="bar-chart-value">${{q.toLocaleString('es-MX')}}</div></div>`;
                    }}).join('') : '<small>Sin registros para graficar.</small>';
                }}

                const donutEntries=countBy(list,'tipo').slice(0,10);
                const total=donutEntries.reduce((a,e)=>a+e[1],0);
                if(donut && legend){{
                    if(!total){{
                        donut.style.background='#0f172a';
                        donut.setAttribute('data-total','0\\nTotal');
                        legend.innerHTML='<small>Sin registros para graficar.</small>';
                    }}else{{
                        let cur=0, seg=[];
                        donutEntries.forEach(([n,q])=>{{
                            const start=cur, end=cur+(q/total*100), col=colorText(n);
                            seg.push(`${{col}} ${{start}}% ${{end}}%`);
                            cur=end;
                        }});
                        donut.style.background=`conic-gradient(${{seg.join(', ')}})`;
                        donut.setAttribute('data-total',`${{total.toLocaleString('es-MX')}}\\nTotal`);
                        legend.innerHTML=donutEntries.map(([n,q])=>`<div class="donut-legend-row"><span class="donut-dot" style="background:${{colorText(n)}}"></span><span class="donut-name" title="${{esc(n)}}">${{esc(n)}}</span><span class="donut-value">${{q.toLocaleString('es-MX')}} | ${{(q/total*100).toFixed(1)}}%</span></div>`).join('');
                    }}
                }}

                const munEntries=countBy(list,'municipio').slice(0,10);
                const maxMun=munEntries.length ? Math.max(...munEntries.map(e=>e[1])) : 1;
                if(munBars){{
                    munBars.innerHTML = munEntries.length ? munEntries.map(([n,q])=>{{
                        const h=Math.max(6,Math.round(q/maxMun*100));
                        const col=colorText(n);
                        return `<div class="vertical-bar-item">
                            <div class="vertical-bar-value">${{q.toLocaleString('es-MX')}}</div>
                            <div class="vertical-bar" title="${{esc(n)}} | ${{q.toLocaleString('es-MX')}}" style="height:${{h}}%;background:${{col}};"></div>
                            <div class="vertical-bar-name" title="${{esc(n)}}">${{esc(n)}}</div>
                        </div>`;
                    }}).join('') : '<small>Sin municipios para graficar.</small>';
                }}
            }}

            function groupCoords(list){{
                const m=new Map();
                list.forEach(r=>{{
                    const lat=Number(r.lat), lon=Number(r.lon);
                    if(!Number.isFinite(lat)||!Number.isFinite(lon)) return;
                    const k=`${{lat.toFixed(6)}},${{lon.toFixed(6)}}`;
                    if(!m.has(k)) m.set(k,{{lat,lon,registros:[],color:colorStatus(r)}});
                    m.get(k).registros.push(r);
                }});
                return [...m.values()];
            }}

            function popupGroup(g){{
                if(g.registros.length===1) return g.registros[0].popup||'';
                let html=`<div style="font-family:Segoe UI,Arial;min-width:310px;background:#020617;color:#e5e7eb;padding:8px;"><div style="font-size:15px;font-weight:800;color:#38bdf8;margin-bottom:8px;">📍 UBICACIÓN CON ${{g.registros.length}} REGISTROS</div><b>Latitud:</b> ${{g.lat}}<br><b>Longitud:</b> ${{g.lon}}<hr style="border-color:#334155;">`;
                g.registros.slice(0,25).forEach((r,i)=>{{
                    html+=`<div style="margin-bottom:9px;padding-bottom:8px;border-bottom:1px solid #334155;"><div style="font-weight:800;color:${{colorText(r.tipo)}};">${{i+1}}. ${{esc(r.tipo)}}</div><div><b>Municipio:</b> ${{esc(r.municipio)}}</div><div><b>Sector:</b> ${{esc(r.sector)}}</div><div><b>Fila Excel:</b> ${{esc(r.fila)}}</div></div>`;
                }});
                if(g.registros.length>25) html+=`<div style="color:#fbbf24;font-weight:800;">Mostrando 25 de ${{g.registros.length}} registros.</div>`;
                return html+'</div>';
            }}

            function markerGroup(g){{
                const n=g.registros.length;
                if(n===1){{
                    const r=g.registros[0], c=colorStatus(r);
                    const mk=L.circleMarker([g.lat,g.lon],{{radius:6,color:c,fillColor:c,fillOpacity:.92,weight:2}});
                    mk.bindPopup(r.popup||''); mk.bindTooltip(r.tooltip||'Registro'); return mk;
                }}
                const size=n>=10?38:32;
                const icon=L.divIcon({{className:'c5-group-marker',html:`<div style="width:${{size}}px;height:${{size}}px;border-radius:50%;background:#0ea5e9;color:white;border:3px solid #e0f2fe;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:14px;box-shadow:0 0 18px rgba(14,165,233,.85);">${{n}}</div>`,iconSize:[size,size],iconAnchor:[size/2,size/2],popupAnchor:[0,-size/2]}});
                const mk=L.marker([g.lat,g.lon],{{icon}});
                mk.bindPopup(popupGroup(g),{{maxWidth:420}});
                mk.bindTooltip(`${{n}} registros en esta ubicación`);
                return mk;
            }}

            function render(list, moveMap){{
                if(!pointLayer) return;
                pointLayer.clearLayers();
                const valid=list.filter(r=>Number.isFinite(Number(r.lat))&&Number.isFinite(Number(r.lon)));
                const groups=groupCoords(valid);
                const metricFiltrados = document.getElementById('metric_filtrados');
                const metricUbicaciones = document.getElementById('metric_ubicaciones');

                if (metricFiltrados) {{
                    metricFiltrados.textContent = valid.length.toLocaleString('es-MX');
                }}

                if (metricUbicaciones) {{
                    metricUbicaciones.textContent = 'Ubicaciones: ' + groups.length.toLocaleString('es-MX');
                }}
                groups.forEach(g=>pointLayer.addLayer(markerGroup(g)));
                const c=document.getElementById('filter_count');
                if(c) c.textContent=`Registros visibles: ${{valid.length.toLocaleString('es-MX')}} | Ubicaciones: ${{groups.length.toLocaleString('es-MX')}}`;
                updateTop10(valid);
                updateCharts(valid);
                if(moveMap && groups.length){{
                    try{{ mapRef.fitBounds(L.latLngBounds(groups.map(g=>[g.lat,g.lon])).pad(.15)); }}catch(e){{console.warn(e);}}
                }}
            }}

            function applyFilters(){{
                const fInicio=document.getElementById('f_fecha_inicio')?.value||'';
                const fFin=document.getElementById('f_fecha_fin')?.value||'';
                const fm=document.getElementById('f_municipio')?.value||'';
                const fs=document.getElementById('f_sector')?.value||'';
                const ft=document.getElementById('f_tipo')?.value||'';
                const filtered=records.filter(r=>
                    (!fInicio || String(r.fecha_iso||'')>=fInicio) &&
                    (!fFin || String(r.fecha_iso||'')<=fFin) &&
                    (!fm || String(r.municipio||'')===fm) &&
                    (!fs || String(r.sector||'')===fs) &&
                    (!ft || String(r.tipo||'')===ft)
                );
                render(filtered,true);
            }}

            function clearFilters(){{
                ['f_fecha_inicio','f_fecha_fin','f_municipio','f_sector','f_tipo'].forEach(id=>{{const el=document.getElementById(id); if(el) el.value='';}});
                render(records,true);
            }}

            if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',()=>setTimeout(boot,800));
            else setTimeout(boot,800);
        }})();
        </script>
        """
    def start_generate_map(self):
        if self.worker_running:
            return
        if self.df_work is None or self.df_work.empty:
            messagebox.showwarning("Aviso", "Primero carga un archivo.")
            return

        has_latlon = bool(self.lat_combo.get().strip() and self.lon_combo.get().strip())
        has_mun_col = bool(self.municipio_combo.get().strip() and self.colonia_combo.get().strip())
        if not has_latlon and not has_mun_col:
            messagebox.showwarning("Aviso", "Selecciona latitud/longitud o municipio/colonia para cruzar con catálogo.")
            return

        self.set_busy(True)
        self.update_progress(0, "Iniciando generación del mapa táctico...")
        threading.Thread(target=self._generate_map_worker, daemon=True).start()

    def _generate_map_worker(self):
        try:
            category_col = self.category_combo.get().strip()
            popup_col = self.popup_combo.get().strip()
            municipio_col = self.municipio_combo.get().strip()
            colonia_col = self.colonia_combo.get().strip()

            zoom = int(self.zoom_var.get().strip() or "11")
            heat_radius = int(self.heat_radius_var.get().strip() or "22")
            table_limit = int(self.table_limit_var.get().strip() or "150")

            self.update_progress_async(8, "Preparando datos y cruzamiento...")
            work_all, stats = self.prepare_dataframe(self.df_work)
            valid_df = work_all.dropna(subset=["__LAT__", "__LON__"]).copy()
            if valid_df.empty:
                raise ValueError("No hay registros con coordenadas válidas. Carga lat/lon o usa un catálogo de colonias.")

            total_valid = len(valid_df)
            selected_mode = self.map_mode_var.get().strip()
            mode = self.get_auto_mode(total_valid) if selected_mode == "Automático" else selected_mode

            sampled_df = valid_df
            if self.enable_sampling_var.get():
                if total_valid > LARGE_DATASET_LIMIT:
                    sampled_df = self.make_sample(valid_df, 180000 if mode == "Calor" else 120000)
                elif total_valid > MEDIUM_DATASET_LIMIT and mode in ("Marcadores", "Marcadores + Cluster", "Todo"):
                    sampled_df = self.make_sample(valid_df, 50000)

            self.update_progress_async(15, f"Calculando centro del mapa... Modo: {mode}")
            center_lat = float(sampled_df["__LAT__"].mean())
            center_lon = float(sampled_df["__LON__"].mean())

            mapa = folium.Map(
                location=[center_lat, center_lon],
                zoom_start=zoom,
                control_scale=True,
                tiles=None
            )

            # Mapas base seleccionables en el HTML
            folium.TileLayer(
                tiles="CartoDB dark_matter",
                name="🌑 Modo táctico oscuro",
                control=True,
                show=True
            ).add_to(mapa)

            folium.TileLayer(
                tiles="OpenStreetMap",
                name="🗺️ OpenStreetMap",
                control=True,
                show=False
            ).add_to(mapa)

            folium.TileLayer(
                tiles="CartoDB positron",
                name="⚪ Claro institucional",
                control=True,
                show=False
            ).add_to(mapa)

            folium.TileLayer(
                tiles="CartoDB Voyager",
                name="🌍 Voyager",
                control=True,
                show=False
            ).add_to(mapa)

            folium.TileLayer(
                tiles="Esri.WorldImagery",
                attr="Tiles © Esri",
                name="🛰️ Satelital Esri",
                control=True,
                show=False
            ).add_to(mapa)

            folium.TileLayer(
                tiles="https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png",
                attr="Map data © OpenStreetMap contributors, SRTM | Map style © OpenTopoMap",
                name="⛰️ Topográfico",
                control=True,
                show=False
            ).add_to(mapa)

            use_markers = mode in ("Marcadores", "Marcadores + Cluster", "Todo")
            use_cluster = mode in ("Cluster", "Marcadores + Cluster", "Todo")
            use_heat = mode in ("Calor", "Todo")

            fg_valid = FeatureGroup(name="✅ Registros válidos auxiliares", show=False)
            fg_warning = FeatureGroup(name="🟡 Geocodificados auxiliares", show=False)
            fg_bad = FeatureGroup(name="⚠️ Fuera de rango auxiliares", show=False)
            cluster_valid = MarkerCluster(name="Cluster táctico").add_to(fg_valid) if use_cluster else None
            heat_points = []
            filter_records = []
            use_html_dynamic_filters = total_valid <= 100000

            records = sampled_df.to_dict(orient="records")
            step = max(1, len(records) // 80)

            for i, row in enumerate(records, start=1):
                lat = row.get("__LAT__")
                lon = row.get("__LON__")
                category = safe_text(row.get(category_col, "")) if category_col else "SIN CATEGORÍA"
                status = safe_text(row.get("__GEO_STATUS__", ""))
                popup = self.build_popup_html(row, popup_col, category_col, municipio_col, colonia_col)
                tooltip = f"Fila {safe_text(row.get('__FILA_EXCEL__', ''))} | {category}"

                if status == "Fuera de rango":
                    color = "#ef4444"
                    layer = fg_bad
                elif status.startswith("Sin coordenadas"):
                    color = "#f59e0b"
                    layer = fg_warning
                else:
                    color = random_color_from_text(category) if self.color_by_category_var.get() else "#22c55e"
                    layer = fg_valid

                if use_markers and not use_html_dynamic_filters:
                    folium.CircleMarker(
                        location=[lat, lon],
                        radius=5,
                        color=color,
                        fill=True,
                        fill_color=color,
                        fill_opacity=0.88,
                        weight=2,
                        popup=folium.Popup(popup, max_width=360),
                        tooltip=tooltip
                    ).add_to(layer)

                if use_cluster and cluster_valid is not None and status != "Fuera de rango" and not use_html_dynamic_filters:
                    folium.Marker(location=[lat, lon], popup=folium.Popup(popup, max_width=360), tooltip=tooltip).add_to(cluster_valid)

                if status == "Fuera de rango" and not use_html_dynamic_filters:
                    folium.Marker(
                        location=[lat, lon],
                        popup=folium.Popup(popup, max_width=360),
                        tooltip=f"⚠️ Revisar fila {safe_text(row.get('__FILA_EXCEL__', ''))}",
                        icon=folium.Icon(color="red", icon="warning-sign")
                    ).add_to(fg_bad)

                filter_records.append({
                    "lat": float(lat),
                    "lon": float(lon),
                    "tipo": category,
                    "municipio": safe_text(row.get(municipio_col, "")) if municipio_col else "",
                    "colonia": safe_text(row.get(colonia_col, "")) if colonia_col else "",
                    "sector": safe_text(row.get("__SECTOR__", "")),
                    "status": status,
                    "popup": popup,
                    "tooltip": tooltip,
                    "color": color,
                    "fila": safe_text(row.get("__FILA_EXCEL__", "")),
                    "fecha_iso": safe_text(row.get("__FECHA_ISO__", "")),
                })

                if use_heat and status != "Fuera de rango":
                    heat_points.append([lat, lon, 1])

                if i % step == 0 or i == len(records):
                    self.update_progress_async(22 + (i / max(len(records), 1)) * 48, f"Construyendo capas... {i:,} / {len(records):,}")

            fg_valid.add_to(mapa)
            fg_warning.add_to(mapa)
            fg_bad.add_to(mapa)

            if use_heat and heat_points:
                HeatMap(heat_points, name="🔥 Heatmap táctico", radius=heat_radius, blur=18, min_opacity=0.25).add_to(mapa)

            self.add_kml_layers_to_map(mapa)

            self.update_progress_async(88, "Creando paneles tácticos...")
            counts_cat = work_all[category_col].fillna("SIN DATO").astype(str).value_counts() if category_col else pd.Series(dtype=int)
            counts_mun = work_all[municipio_col].fillna("SIN DATO").astype(str).value_counts() if municipio_col else pd.Series(dtype=int)

            table_rows = []
            for _, r in work_all.head(table_limit).iterrows():
                table_rows.append({
                    "fila": r.get("__FILA_EXCEL__", ""),
                    "fecha": r.get("__FECHA_FMT__", ""),
                    "tipo": r.get(category_col, "") if category_col else "",
                    "municipio": r.get(municipio_col, "") if municipio_col else "",
                    "colonia": r.get(colonia_col, "") if colonia_col else "",
                    "sector": r.get("__SECTOR__", ""),
                    "lat": "" if pd.isna(r.get("__LAT__", "")) else r.get("__LAT__", ""),
                    "lon": "" if pd.isna(r.get("__LON__", "")) else r.get("__LON__", ""),
                    "status": r.get("__GEO_STATUS__", ""),
                })

            mapa.get_root().html.add_child(folium.Element(self.tactical_panels_html(
                stats, counts_cat, counts_mun, table_rows,
                self.data_path.name if self.data_path else "", mode
            )))

            if use_html_dynamic_filters:
                mapa.get_root().html.add_child(folium.Element(self.filter_script_html(mapa.get_name(), filter_records)))
            else:
                mapa.get_root().html.add_child(folium.Element("""
                <script>
                    setTimeout(function(){
                        const c = document.getElementById('filter_count');
                        if(c){ c.textContent = 'Filtros desactivados únicamente si la base supera 100,000 registros'; }
                    }, 800);
                </script>
                """))

            folium.LayerControl(collapsed=False, position='topright').add_to(mapa)

            self.update_progress_async(94, "Guardando archivo HTML y reporte...")
            output_dir = Path(r"C:\Users\Propietario\Desktop\LOCAL_GPT_EXCEL\Mapas_Tacticos")
            output_dir.mkdir(parents=True, exist_ok=True)
            base = self.data_path.stem if self.data_path else "mapa"
            sheet = re.sub(r"[^a-zA-Z0-9_-]+", "_", self.sheet_name or "Fuente")

            html_path = output_dir / f"{base}_{sheet}_MAPA_TACTICO.html"
            mapa.save(str(html_path))

            report_path = output_dir / f"{base}_{sheet}_REPORTE_COORDENADAS.xlsx"
            try:
                with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
                    work_all[work_all["__GEO_STATUS__"].eq("Sin coordenadas")].to_excel(writer, index=False, sheet_name="Sin coordenadas")
                    work_all[work_all["__GEO_STATUS__"].eq("Fuera de rango")].to_excel(writer, index=False, sheet_name="Fuera de rango")
                    work_all[work_all["__GEO_STATUS__"].eq("Sin coordenadas (geocodificado)")].to_excel(writer, index=False, sheet_name="Geocodificados")
            except Exception:
                report_path = None

            self.update_progress_async(100, "Mapa táctico generado correctamente.")
            self.root.after(0, lambda: self.status_var.set(
                f"Mapa generado | Total: {stats['total']:,} | Con coord: {stats['con_coordenadas']:,} | Sin coord: {stats['sin_coordenadas']:,} | Fuera rango: {stats['fuera_rango']:,}"
            ))

            if self.open_browser_var.get():
                webbrowser.open(html_path.resolve().as_uri())

            msg = (
                f"Mapa táctico generado correctamente.\n\nHTML:\n{html_path}\n\n"
                f"Total: {stats['total']:,}\n"
                f"Con coordenadas: {stats['con_coordenadas']:,}\n"
                f"Sin coordenadas: {stats['sin_coordenadas']:,}\n"
                f"Fuera de rango: {stats['fuera_rango']:,}\n"
                f"Geocodificados: {stats['geocodificados']:,}"
            )
            if report_path:
                msg += f"\n\nReporte:\n{report_path}"
            self.root.after(0, lambda: messagebox.showinfo("Proceso terminado", msg))

        except Exception:
            error = traceback.format_exc()
            self.root.after(0, lambda: messagebox.showerror("Error al generar mapa", error))
        finally:
            self.root.after(0, lambda: self.set_busy(False))


def main():
    root = tk.Tk()
    TacticalMapApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
