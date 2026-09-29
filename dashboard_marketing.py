#!/usr/bin/env python3
"""Dashboard de Marketing Digital, en un solo archivo y sin paquetes externos.

Ejecuta: python3 dashboard_marketing.py

Al iniciar muestra datos ficticios para probarlo. Puedes cargar .csv o .xlsx
con columnas equivalentes a: date/fecha, visits/visitas,
conversions/conversiones, spend/gasto e ingresos/revenue.
El botón Exportar Excel crea un .xlsx con resumen y datos diarios.
"""

import csv
import posixpath
import random
import tkinter as tk
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, ttk

BG = "#f4f7fb"
INK = "#182230"
MUTED = "#64748b"
BLUE = "#3977f6"
TEAL = "#12a889"
ORANGE = "#f39a38"
GRID = "#e5eaf2"

# Nombres de espacio de Excel Open XML (formato .xlsx).
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)

ALIASES = {
    "date": {"date", "fecha", "día", "dia"},
    "visits": {"visits", "visit", "visitas", "tráfico web", "trafico web", "trafico"},
    "conversions": {"conversions", "conversion", "conversiones", "conversión", "conversion"},
    "spend": {"spend", "ad spend", "gasto", "gasto publicitario", "inversión", "inversion"},
    "revenue": {"revenue", "ingresos", "ingreso", "ventas", "ingresos atribuidos"},
}


def make_demo_data(days=120):
    """Genera cifras ficticias y reproducibles para probar el dashboard."""
    rng = random.Random(18)
    rows = []
    today = date.today()
    for offset in range(days - 1, -1, -1):
        day = today - timedelta(days=offset)
        visits = max(120, int(rng.gauss(780, 115)))
        conversions = max(3, int(visits * rng.uniform(0.025, 0.052)))
        spend = round(rng.uniform(85, 155), 2)
        revenue = round(conversions * rng.uniform(38, 66), 2)
        rows.append({"date": day, "visits": visits, "conversions": conversions,
                     "spend": spend, "revenue": revenue})
    return rows


def _header_map(headers):
    normalized = [str(value).strip().lower() for value in headers]
    positions = {}
    for key, names in ALIASES.items():
        found = next((i for i, name in enumerate(normalized) if name in names), None)
        if found is None:
            raise ValueError("Falta una columna necesaria: " + key)
        positions[key] = found
    return positions


def _as_date(value):
    raw = str(value).strip()
    try:
        return date.fromisoformat(raw[:10])
    except ValueError:
        # Excel suele guardar las fechas como número de días desde 1899-12-30.
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(raw))).date()
        except ValueError as exc:
            raise ValueError(f"Fecha no reconocida: {raw}") from exc


def _number(value, decimal_comma=False):
    raw = str(value).strip()
    if decimal_comma:
        raw = raw.replace(".", "").replace(",", ".")
    return float(raw)


def _rows_from_grid(grid, decimal_comma=False):
    grid = [row for row in grid if any(str(cell).strip() for cell in row)]
    if len(grid) < 2:
        raise ValueError("El archivo no contiene filas de datos.")
    headers = [str(value).strip().lower() for value in grid[0]]
    # Admite archivos Excel exportados con las etiquetas de los indicadores.
    positions = _header_map(headers)
    rows = []
    for line_no, values in enumerate(grid[1:], start=2):
        try:
            def at(key):
                index = positions[key]
                return values[index] if index < len(values) else ""
            rows.append({
                "date": _as_date(at("date")),
                "visits": int(_number(at("visits"), decimal_comma)),
                "conversions": int(_number(at("conversions"), decimal_comma)),
                "spend": _number(at("spend"), decimal_comma),
                "revenue": _number(at("revenue"), decimal_comma),
            })
        except (ValueError, TypeError, IndexError) as exc:
            raise ValueError(f"Revisa la fila {line_no}: {exc}") from exc
    return sorted(rows, key=lambda row: row["date"])


def load_csv(path):
    """Carga un CSV separado por coma o punto y coma."""
    with open(path, "r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        try:
            delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t").delimiter
        except csv.Error:
            delimiter = ","
        grid = list(csv.reader(file, delimiter=delimiter))
    return _rows_from_grid(grid, decimal_comma=(delimiter == ";"))


def load_xlsx(path):
    """Lee la primera hoja de un .xlsx usando solo la biblioteca estándar."""
    def tag(namespace, name):
        return f"{{{namespace}}}{name}"

    with zipfile.ZipFile(path, "r") as book:
        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        sheets = workbook.findall(f"{tag(MAIN_NS, 'sheets')}/{tag(MAIN_NS, 'sheet')}")
        sheet = next((item for item in sheets if item.attrib.get("name") == "Datos diarios"),
                     sheets[0] if sheets else None)
        if sheet is None:
            raise ValueError("El Excel no contiene hojas.")
        rel_id = sheet.attrib.get(tag(DOC_REL_NS, "id"))
        rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels.findall(tag(PKG_REL_NS, "Relationship")):
            if rel.attrib.get("Id") == rel_id:
                target = rel.attrib.get("Target")
                break
        if not target:
            raise ValueError("No se encontró la primera hoja del Excel.")
        worksheet_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        shared = []
        if "xl/sharedStrings.xml" in book.namelist():
            shared_root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            for item in shared_root.findall(tag(MAIN_NS, "si")):
                shared.append("".join(node.text or "" for node in item.iter(tag(MAIN_NS, "t"))))
        root = ET.fromstring(book.read(worksheet_path))
        grid = []
        for row in root.iter(tag(MAIN_NS, "row")):
            values = []
            for cell in row.findall(tag(MAIN_NS, "c")):
                cell_ref = cell.attrib.get("r", "A1")
                letters = "".join(char for char in cell_ref if char.isalpha())
                column = 0
                for char in letters.upper():
                    column = column * 26 + ord(char) - 64
                column -= 1
                while len(values) <= column:
                    values.append("")
                kind = cell.attrib.get("t")
                value_node = cell.find(tag(MAIN_NS, "v"))
                if kind == "inlineStr":
                    value = "".join(node.text or "" for node in cell.iter(tag(MAIN_NS, "t")))
                elif value_node is None:
                    value = ""
                elif kind == "s":
                    value = shared[int(value_node.text)]
                elif kind == "b":
                    value = value_node.text == "1"
                else:
                    value = value_node.text or ""
                values[column] = value
            grid.append(values)
    return _rows_from_grid(grid)


def _excel_col(index):
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _worksheet_xml(rows):
    root = ET.Element(f"{{{MAIN_NS}}}worksheet")
    sheet_data = ET.SubElement(root, f"{{{MAIN_NS}}}sheetData")
    for row_number, values in enumerate(rows, start=1):
        row_node = ET.SubElement(sheet_data, f"{{{MAIN_NS}}}row", {"r": str(row_number)})
        for col_index, value in enumerate(values):
            cell = ET.SubElement(row_node, f"{{{MAIN_NS}}}c",
                                 {"r": f"{_excel_col(col_index)}{row_number}"})
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                value_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
                value_node.text = str(value)
            else:
                cell.set("t", "inlineStr")
                inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
                text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
                text.text = str(value)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def export_xlsx(path, data):
    """Crea un .xlsx con una hoja de resumen y otra con datos diarios."""
    visits = sum(row["visits"] for row in data)
    conversions = sum(row["conversions"] for row in data)
    spend = sum(row["spend"] for row in data)
    revenue = sum(row["revenue"] for row in data)
    cpa = spend / conversions if conversions else 0
    roas = revenue / spend if spend else 0
    conversion_rate = conversions / visits if visits else 0
    summary = [
        ["Métrica", "Resultado"],
        ["Días con datos", len(data)],
        ["Tráfico web (visitas)", visits],
        ["Conversiones", conversions],
        ["Gasto publicitario", round(spend, 2)],
        ["CPA", round(cpa, 2)],
        ["Ingresos atribuidos", round(revenue, 2)],
        ["ROAS", round(roas, 4)],
        ["Tasa de conversión", round(conversion_rate, 4)],
    ]
    daily = [["date", "visits", "conversions", "spend", "revenue"]]
    daily.extend([[row["date"].isoformat(), row["visits"], row["conversions"],
                   round(row["spend"], 2), round(row["revenue"], 2)] for row in data])

    types_root = ET.Element(f"{{{CONTENT_NS}}}Types")
    ET.SubElement(types_root, f"{{{CONTENT_NS}}}Default", {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"})
    ET.SubElement(types_root, f"{{{CONTENT_NS}}}Default", {"Extension": "xml", "ContentType": "application/xml"})
    ET.SubElement(types_root, f"{{{CONTENT_NS}}}Override", {"PartName": "/xl/workbook.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"})
    for index in (1, 2):
        ET.SubElement(types_root, f"{{{CONTENT_NS}}}Override", {"PartName": f"/xl/worksheets/sheet{index}.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})

    package_rels = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    ET.SubElement(package_rels, f"{{{PKG_REL_NS}}}Relationship", {
        "Id": "rId1", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "Target": "xl/workbook.xml"})

    workbook = ET.Element(f"{{{MAIN_NS}}}workbook")
    sheets = ET.SubElement(workbook, f"{{{MAIN_NS}}}sheets")
    ET.SubElement(sheets, f"{{{MAIN_NS}}}sheet", {"name": "Resumen", "sheetId": "1", f"{{{DOC_REL_NS}}}id": "rId1"})
    ET.SubElement(sheets, f"{{{MAIN_NS}}}sheet", {"name": "Datos diarios", "sheetId": "2", f"{{{DOC_REL_NS}}}id": "rId2"})
    workbook_rels = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    for index in (1, 2):
        ET.SubElement(workbook_rels, f"{{{PKG_REL_NS}}}Relationship", {
            "Id": f"rId{index}", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "Target": f"worksheets/sheet{index}.xml"})

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as book:
        book.writestr("[Content_Types].xml", ET.tostring(types_root, encoding="utf-8", xml_declaration=True))
        book.writestr("_rels/.rels", ET.tostring(package_rels, encoding="utf-8", xml_declaration=True))
        book.writestr("xl/workbook.xml", ET.tostring(workbook, encoding="utf-8", xml_declaration=True))
        book.writestr("xl/_rels/workbook.xml.rels", ET.tostring(workbook_rels, encoding="utf-8", xml_declaration=True))
        book.writestr("xl/worksheets/sheet1.xml", _worksheet_xml(summary))
        book.writestr("xl/worksheets/sheet2.xml", _worksheet_xml(daily))


class MarketingDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Marketing Digital | Dashboard")
        self.geometry("1080x730")
        self.minsize(850, 620)
        self.configure(bg=BG)
        self.rows = make_demo_data()
        self.period = tk.StringVar(value="30")
        self.source_label = tk.StringVar(value="Datos de ejemplo · ficticios")
        self._setup_style()
        self._build_ui()
        self.refresh()

    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TCombobox", padding=6, font=("Arial", 10))

    def _build_ui(self):
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=30, pady=(24, 14))
        title_block = tk.Frame(header, bg=BG)
        title_block.pack(side="left", fill="x", expand=True)
        tk.Label(title_block, text="Marketing Digital", font=("Arial", 23, "bold"), fg=INK, bg=BG).pack(anchor="w")
        tk.Label(title_block, textvariable=self.source_label, font=("Arial", 10), fg=MUTED, bg=BG).pack(anchor="w", pady=(5, 0))
        controls = tk.Frame(header, bg=BG)
        controls.pack(side="right", anchor="e")
        tk.Label(controls, text="Periodo", font=("Arial", 10), fg=MUTED, bg=BG).pack(side="left", padx=(0, 8))
        picker = ttk.Combobox(controls, textvariable=self.period, width=8, values=("30", "90", "365"), state="readonly")
        picker.pack(side="left", padx=(0, 8))
        picker.bind("<<ComboboxSelected>>", lambda _event: self.refresh())
        self._button(controls, "Cargar CSV/Excel", self.open_file, BLUE).pack(side="left", padx=(0, 8))
        self._button(controls, "Exportar Excel", self.save_excel, TEAL).pack(side="left")

        self.cards_frame = tk.Frame(self, bg=BG)
        self.cards_frame.pack(fill="x", padx=30, pady=(3, 17))
        for col in range(4):
            self.cards_frame.grid_columnconfigure(col, weight=1, uniform="card")
        self.cards = {}
        specs = [("visits", "TRÁFICO WEB", BLUE), ("cpa", "COSTO POR ADQUISICIÓN", ORANGE),
                 ("conversions", "CONVERSIONES", TEAL), ("roas", "ROAS PUBLICITARIO", "#7656d6")]
        for col, (key, label, accent) in enumerate(specs):
            card = tk.Frame(self.cards_frame, bg="white", highlightbackground="#e7edf5", highlightthickness=1)
            card.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 8, 0))
            tk.Frame(card, bg=accent, height=4).pack(fill="x")
            inner = tk.Frame(card, bg="white")
            inner.pack(fill="both", padx=15, pady=14)
            tk.Label(inner, text=label, font=("Arial", 9, "bold"), fg=MUTED, bg="white", wraplength=175, justify="left").pack(anchor="w")
            value = tk.Label(inner, text="—", font=("Arial", 22, "bold"), fg=INK, bg="white")
            value.pack(anchor="w", pady=(10, 3))
            note = tk.Label(inner, text="En el periodo seleccionado", font=("Arial", 9), fg=MUTED, bg="white")
            note.pack(anchor="w")
            self.cards[key] = (value, note)

        chart_panel = tk.Frame(self, bg="white", highlightbackground="#e7edf5", highlightthickness=1)
        chart_panel.pack(fill="both", expand=True, padx=30, pady=(0, 14))
        chart_header = tk.Frame(chart_panel, bg="white")
        chart_header.pack(fill="x", padx=20, pady=(17, 0))
        tk.Label(chart_header, text="Rendimiento diario", font=("Arial", 14, "bold"), fg=INK, bg="white").pack(side="left")
        legend = tk.Frame(chart_header, bg="white")
        legend.pack(side="right")
        self._legend(legend, BLUE, "Visitas")
        self._legend(legend, ORANGE, "Conversiones")
        self.chart = tk.Canvas(chart_panel, bg="white", height=310, highlightthickness=0)
        self.chart.pack(fill="both", expand=True, padx=10, pady=(3, 8))
        self.chart.bind("<Configure>", lambda _event: self.draw_chart())
        footer = tk.Frame(self, bg=BG)
        footer.pack(fill="x", padx=30, pady=(0, 18))
        tk.Label(footer, text="ROAS = ingresos atribuidos ÷ gasto publicitario  ·  CPA = gasto ÷ conversiones", font=("Arial", 9), fg=MUTED, bg=BG).pack(side="left")
        tk.Label(footer, text="Los datos de ejemplo son ficticios.", font=("Arial", 9), fg=MUTED, bg=BG).pack(side="right")

    @staticmethod
    def _button(parent, text, command, color):
        return tk.Button(parent, text=text, command=command, bg=color, fg="white", activebackground=color,
                         activeforeground="white", relief="flat", bd=0, padx=12, pady=8,
                         font=("Arial", 9, "bold"), cursor="hand2")

    @staticmethod
    def _legend(parent, color, text):
        entry = tk.Frame(parent, bg="white")
        entry.pack(side="left", padx=(12, 0))
        dot = tk.Canvas(entry, width=10, height=10, bg="white", highlightthickness=0)
        dot.pack(side="left")
        dot.create_oval(1, 1, 9, 9, fill=color, outline=color)
        tk.Label(entry, text=text, font=("Arial", 9), fg=MUTED, bg="white").pack(side="left", padx=(4, 0))

    def open_file(self):
        path = filedialog.askopenfilename(title="Selecciona datos de marketing",
                                          filetypes=(("CSV y Excel", "*.csv *.xlsx"),
                                                     ("CSV", "*.csv"), ("Excel", "*.xlsx"),
                                                     ("Todos los archivos", "*.*")))
        if not path:
            return
        try:
            self.rows = load_xlsx(path) if path.lower().endswith(".xlsx") else load_csv(path)
            self.source_label.set("Archivo: " + path.replace("\\", "/").rsplit("/", 1)[-1])
            self.refresh()
        except (OSError, ValueError, KeyError, zipfile.BadZipFile) as exc:
            messagebox.showerror("No se pudo cargar el archivo", str(exc), parent=self)

    def save_excel(self):
        path = filedialog.asksaveasfilename(title="Guardar reporte de marketing en Excel",
                                            defaultextension=".xlsx",
                                            filetypes=(("Libro de Excel", "*.xlsx"),))
        if not path:
            return
        try:
            export_xlsx(path, self.visible_rows())
            messagebox.showinfo("Excel listo", "Se exportaron el resumen y los datos diarios.", parent=self)
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            messagebox.showerror("No se pudo exportar", str(exc), parent=self)

    def visible_rows(self):
        days = int(self.period.get())
        end = max(row["date"] for row in self.rows)
        start = end - timedelta(days=days - 1)
        selected = [row for row in self.rows if start <= row["date"] <= end]
        return selected or self.rows

    def refresh(self):
        data = self.visible_rows()
        visits = sum(row["visits"] for row in data)
        conversions = sum(row["conversions"] for row in data)
        spend = sum(row["spend"] for row in data)
        revenue = sum(row["revenue"] for row in data)
        cpa = spend / conversions if conversions else 0
        roas = revenue / spend if spend else 0
        self.cards["visits"][0].config(text=f"{visits:,}")
        self.cards["cpa"][0].config(text=f"${cpa:,.2f}" if conversions else "—")
        self.cards["conversions"][0].config(text=f"{conversions:,}")
        self.cards["roas"][0].config(text=f"{roas:.2f}×" if spend else "—")
        self.cards["visits"][1].config(text=f"{len(data)} días con datos")
        self.cards["cpa"][1].config(text=f"Gasto total: ${spend:,.2f}")
        self.cards["conversions"][1].config(text=f"Tasa: {conversions / visits:.2%}" if visits else "Tasa: —")
        self.cards["roas"][1].config(text=f"Ingresos: ${revenue:,.2f}")
        self.draw_chart()

    def draw_chart(self):
        canvas = self.chart
        canvas.delete("all")
        data = self.visible_rows()
        if not data:
            return
        width = max(canvas.winfo_width(), 400)
        height = max(canvas.winfo_height(), 220)
        left, right, top, bottom = 56, 52, 22, 38
        plot_w, plot_h = width - left - right, height - top - bottom
        if plot_w <= 0 or plot_h <= 0:
            return
        max_visits = max(row["visits"] for row in data) or 1
        max_conv = max(row["conversions"] for row in data) or 1
        for tick in range(5):
            y = top + plot_h * tick / 4
            canvas.create_line(left, y, width - right, y, fill=GRID)
            canvas.create_text(left - 9, y, text=f"{max_visits * (1 - tick / 4):.0f}", anchor="e", fill=MUTED, font=("Arial", 8))
            canvas.create_text(width - right + 9, y, text=f"{max_conv * (1 - tick / 4):.0f}", anchor="w", fill=MUTED, font=("Arial", 8))
        step = plot_w / len(data)
        bar_w = max(2, min(18, step * 0.62))
        points = []
        for index, row in enumerate(data):
            x = left + step * (index + 0.5)
            bar_h = row["visits"] / max_visits * plot_h
            canvas.create_rectangle(x - bar_w / 2, top + plot_h - bar_h, x + bar_w / 2, top + plot_h, fill=BLUE, outline="")
            y_conv = top + plot_h - row["conversions"] / max_conv * plot_h
            points.extend((x, y_conv))
        if len(points) >= 4:
            canvas.create_line(*points, fill=ORANGE, width=2, smooth=True)
        for index in sorted(set((0, len(data) // 2, len(data) - 1))):
            x = left + step * (index + 0.5)
            canvas.create_text(x, height - 16, text=data[index]["date"].strftime("%d %b"), fill=MUTED, font=("Arial", 8))
        canvas.create_text(14, top + plot_h / 2, text="Visitas", angle=90, fill=BLUE, font=("Arial", 8, "bold"))
        canvas.create_text(width - 7, top + plot_h / 2, text="Conversiones", angle=90, fill=ORANGE, font=("Arial", 8, "bold"))


if __name__ == "__main__":
    MarketingDashboard().mainloop()
