from __future__ import annotations

import io
import sqlite3
import unicodedata
from calendar import monthrange
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "inventario.db"
CURRENCY = "USD"

st.set_page_config(page_title="Inventario y ventas", page_icon="📦", layout="wide")


def get_conn():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def db_connection():
    conn = get_conn()
    try:
        with conn:
            yield conn
    finally:
        conn.close()


def init_db():
    with db_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sku TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL,
                category TEXT NOT NULL DEFAULT '',
                cost REAL NOT NULL DEFAULT 0 CHECK(cost >= 0),
                price REAL NOT NULL DEFAULT 0 CHECK(price >= 0),
                stock INTEGER NOT NULL DEFAULT 0 CHECK(stock >= 0),
                min_stock INTEGER NOT NULL DEFAULT 0 CHECK(min_stock >= 0),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sold_at TEXT NOT NULL,
                customer TEXT NOT NULL DEFAULT '',
                payment_method TEXT NOT NULL,
                total REAL NOT NULL CHECK(total >= 0)
            );
            CREATE TABLE IF NOT EXISTS sale_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_id INTEGER NOT NULL REFERENCES sales(id) ON DELETE CASCADE,
                product_id INTEGER NOT NULL REFERENCES products(id),
                sku TEXT NOT NULL,
                product_name TEXT NOT NULL,
                quantity INTEGER NOT NULL CHECK(quantity > 0),
                unit_price REAL NOT NULL CHECK(unit_price >= 0),
                subtotal REAL NOT NULL CHECK(subtotal >= 0)
            );
            CREATE INDEX IF NOT EXISTS idx_sales_sold_at ON sales(sold_at);
            CREATE INDEX IF NOT EXISTS idx_sale_items_sale_id ON sale_items(sale_id);
            """
        )


def read_products():
    with db_connection() as conn:
        return pd.read_sql_query(
            "SELECT sku AS SKU, name AS Nombre, category AS Categoría, cost AS [Precio compra], "
            "price AS [Precio venta], stock AS Stock, min_stock AS [Stock mínimo], "
            "updated_at AS [Actualizado] FROM products ORDER BY name",
            conn,
        )


def read_sales(start=None, end=None):
    query = "SELECT id, sold_at, customer, payment_method, total FROM sales"
    params = []
    clauses = []
    if start:
        clauses.append("date(sold_at) >= date(?)")
        params.append(str(start))
    if end:
        clauses.append("date(sold_at) <= date(?)")
        params.append(str(end))
    if clauses:
        query += " WHERE " + " AND ".join(clauses)
    query += " ORDER BY sold_at DESC"
    with db_connection() as conn:
        return pd.read_sql_query(query, conn, params=params)


def read_sale_details(start=None, end=None):
    query = """SELECT s.id AS Venta, s.sold_at AS Fecha, s.customer AS Cliente,
               s.payment_method AS [Forma de pago], i.sku AS SKU,
               i.product_name AS Producto, i.quantity AS Cantidad,
               i.unit_price AS [Precio unitario], i.subtotal AS Subtotal
               FROM sale_items i JOIN sales s ON s.id = i.sale_id"""
    params = []
    clauses = []
    if start:
        clauses.append("date(s.sold_at) >= date(?)")
        params.append(str(start))
    if end:
        clauses.append("date(s.sold_at) <= date(?)")
        params.append(str(end))
    if clauses:
        query += " WHERE " + " AND ".join(clauses)
    query += " ORDER BY s.sold_at DESC, s.id DESC"
    with db_connection() as conn:
        return pd.read_sql_query(query, conn, params=params)


def save_product(sku, name, category, cost, price, stock, min_stock):
    now = datetime.now().isoformat(timespec="seconds")
    with db_connection() as conn:
        conn.execute(
            """INSERT INTO products(sku,name,category,cost,price,stock,min_stock,created_at,updated_at)
               VALUES(?,?,?,?,?,?,?,?,?)
               ON CONFLICT(sku) DO UPDATE SET name=excluded.name, category=excluded.category,
               cost=excluded.cost, price=excluded.price, stock=excluded.stock,
               min_stock=excluded.min_stock, updated_at=excluded.updated_at""",
            (sku.strip(), name.strip(), category.strip(), float(cost), float(price), int(stock), int(min_stock), now, now),
        )


def seed_demo_august():
    """Load fictional August 2026 electrical-accessories data into an empty database."""
    products = [
        ("ELE-001", "Cinta aislante 20 m", "Accesorios eléctricos", 1.20, 2.00, 15, 10, 35),
        ("ELE-002", "Tomacorriente doble", "Accesorios eléctricos", 2.50, 4.00, 18, 12, 22),
        ("ELE-003", "Interruptor sencillo", "Accesorios eléctricos", 1.50, 2.80, 20, 10, 25),
        ("ELE-004", "Foco LED 12 W", "Iluminación", 2.00, 3.50, 18, 15, 42),
        ("ELE-005", "Enchufe macho", "Accesorios eléctricos", 1.20, 2.50, 17, 10, 18),
        ("ELE-006", "Extensión 5 m", "Accesorios eléctricos", 5.50, 9.00, 7, 8, 8),
        ("ELE-007", "Breaker 20 A", "Protección eléctrica", 4.50, 7.50, 10, 12, 10),
        ("ELE-008", "Canaleta PVC 2 m", "Instalación", 3.20, 5.50, 12, 10, 13),
        ("ELE-009", "Cable THHN 10 m", "Cableado", 6.00, 9.50, 12, 8, 6),
        ("ELE-010", "Multicontacto 4 salidas", "Accesorios eléctricos", 4.00, 7.00, 8, 10, 12),
    ]
    sale_dates = ["2026-08-03", "2026-08-10", "2026-08-17", "2026-08-24", "2026-08-31"]
    payment_methods = ["Efectivo", "Transferencia", "Tarjeta", "Efectivo", "Otro"]
    with db_connection() as conn:
        if conn.execute("SELECT COUNT(*) FROM products").fetchone()[0] or conn.execute("SELECT COUNT(*) FROM sales").fetchone()[0]:
            raise ValueError("La demo solo se carga cuando el inventario está vacío.")
        try:
            conn.execute("BEGIN IMMEDIATE")
            now = datetime.now().isoformat(timespec="seconds")
            product_ids = {}
            for sku, name, category, cost, price, stock, min_stock, _sold in products:
                cur = conn.execute(
                    "INSERT INTO products(sku,name,category,cost,price,stock,min_stock,created_at,updated_at) "
                    "VALUES(?,?,?,?,?,?,?,?,?)",
                    (sku, name, category, cost, price, stock, min_stock, now, now),
                )
                product_ids[sku] = cur.lastrowid
            for week_index, day in enumerate(sale_dates):
                line_items = []
                for sku, name, _category, _cost, price, _stock, _min_stock, sold in products:
                    base, remainder = divmod(sold, len(sale_dates))
                    quantity = base + (1 if week_index < remainder else 0)
                    if quantity:
                        line_items.append((sku, name, quantity, price))
                total = round(sum(quantity * price for _sku, _name, quantity, price in line_items), 2)
                cur = conn.execute(
                    "INSERT INTO sales(sold_at,customer,payment_method,total) VALUES(?,?,?,?)",
                    (f"{day}T12:00:00", "Cliente de ejemplo", payment_methods[week_index], total),
                )
                sale_id = cur.lastrowid
                for sku, name, quantity, price in line_items:
                    conn.execute(
                        "INSERT INTO sale_items(sale_id,product_id,sku,product_name,quantity,unit_price,subtotal) "
                        "VALUES(?,?,?,?,?,?,?)",
                        (sale_id, product_ids[sku], sku, name, quantity, price, round(quantity * price, 2)),
                    )
            conn.commit()
            return len(products), len(sale_dates), round(sum(row[4] * row[7] for row in products), 2)
        except Exception:
            conn.rollback()
            raise


def record_sale(cart, customer, payment_method):
    if not cart:
        raise ValueError("Agrega al menos un producto a la venta.")
    with db_connection() as conn:
        try:
            conn.execute("BEGIN IMMEDIATE")
            total = round(sum(float(item["unit_price"]) * int(item["quantity"]) for item in cart), 2)
            cur = conn.execute(
                "INSERT INTO sales(sold_at,customer,payment_method,total) VALUES(?,?,?,?)",
                (datetime.now().isoformat(timespec="seconds"), customer.strip(), payment_method, total),
            )
            sale_id = cur.lastrowid
            for item in cart:
                product = conn.execute("SELECT id, name, stock FROM products WHERE sku=?", (item["sku"],)).fetchone()
                if product is None:
                    raise ValueError(f"El producto {item['sku']} ya no existe.")
                qty = int(item["quantity"])
                if qty < 1 or product["stock"] < qty:
                    raise ValueError(f"Stock insuficiente para {product['name']}. Disponible: {product['stock']}.")
                unit_price = float(item["unit_price"])
                subtotal = round(qty * unit_price, 2)
                conn.execute(
                    "INSERT INTO sale_items(sale_id,product_id,sku,product_name,quantity,unit_price,subtotal) "
                    "VALUES(?,?,?,?,?,?,?)",
                    (sale_id, product["id"], item["sku"], product["name"], qty, unit_price, subtotal),
                )
                conn.execute("UPDATE products SET stock=stock-?, updated_at=? WHERE id=?",
                             (qty, datetime.now().isoformat(timespec="seconds"), product["id"]))
            conn.commit()
            return sale_id, total
        except Exception:
            conn.rollback()
            raise


def normalized(value):
    value = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(value.strip().lower().replace("_", " ").split())


def import_products(file):
    df = pd.read_excel(file, sheet_name=0)
    aliases = {
        "sku": "sku", "codigo": "sku", "codigo producto": "sku",
        "nombre": "nombre", "producto": "nombre", "name": "nombre",
        "categoria": "categoria", "category": "categoria",
        "precio compra": "precio compra", "costo": "precio compra", "cost": "precio compra",
        "precio venta": "precio venta", "precio": "precio venta", "price": "precio venta",
        "stock": "stock", "existencia": "stock",
        "stock minimo": "stock minimo", "stock min": "stock minimo", "min stock": "stock minimo",
    }
    rename = {col: aliases.get(normalized(col), normalized(col)) for col in df.columns}
    df = df.rename(columns=rename)
    required = ["sku", "nombre", "categoria", "precio compra", "precio venta", "stock", "stock minimo"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError("Faltan columnas: " + ", ".join(missing))
    df = df[required].dropna(how="all")
    if df.empty:
        raise ValueError("El archivo no contiene filas de productos.")

    parsed = []
    for row_num, row in enumerate(df.to_dict(orient="records"), start=2):
        sku = "" if pd.isna(row["sku"]) else str(row["sku"]).strip()
        name = "" if pd.isna(row["nombre"]) else str(row["nombre"]).strip()
        category = "" if pd.isna(row["categoria"]) else str(row["categoria"]).strip()
        if not sku or not name:
            raise ValueError(f"Fila {row_num}: SKU y Nombre son obligatorios.")
        try:
            cost = float(row["precio compra"])
            price = float(row["precio venta"])
            stock_val = float(row["stock"])
            min_val = float(row["stock minimo"])
        except (TypeError, ValueError):
            raise ValueError(f"Fila {row_num}: revisa precios y cantidades; deben ser números.")
        if cost < 0 or price < 0 or stock_val < 0 or min_val < 0:
            raise ValueError(f"Fila {row_num}: no se aceptan valores negativos.")
        if not stock_val.is_integer() or not min_val.is_integer():
            raise ValueError(f"Fila {row_num}: Stock y Stock mínimo deben ser enteros.")
        parsed.append((sku, name, category, cost, price, int(stock_val), int(min_val)))
    for item in parsed:
        save_product(*item)
    return len(parsed)


def make_export(start=None, end=None):
    products = read_products()
    sales = read_sales(start, end)
    details = read_sale_details(start, end)
    low_stock = products[products["Stock"] <= products["Stock mínimo"]].copy() if not products.empty else products
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        products.to_excel(writer, sheet_name="Productos", index=False)
        sales.to_excel(writer, sheet_name="Ventas", index=False)
        details.to_excel(writer, sheet_name="Detalle de ventas", index=False)
        low_stock.to_excel(writer, sheet_name="Stock bajo", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column_cells in ws.columns:
                letter = column_cells[0].column_letter
                max_len = max((len(str(cell.value or "")) for cell in column_cells), default=10)
                ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 32)
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, float):
                        cell.number_format = '"$"#,##0.00'
    out.seek(0)
    return out


def make_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Productos"
    ws.append(["SKU", "Nombre", "Categoría", "Precio compra", "Precio venta", "Stock", "Stock mínimo"])
    ws.append(["EJ-001", "Producto de ejemplo", "General", 1.50, 2.50, 10, 2])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:G2"
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    for col, width in zip("ABCDEFG", [18, 28, 20, 18, 18, 12, 16]):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=5):
        for cell in row:
            cell.number_format = '\"$\"#,##0.00'
    instructions = wb.create_sheet("Instrucciones")
    instructions.append(["Carga de productos"])
    instructions.append(["Completa la hoja Productos y conserva exactamente los encabezados."])
    instructions.append(["SKU y Nombre son obligatorios; el SKU identifica cada producto."])
    instructions.append(["Precios no negativos; Stock y Stock mínimo deben ser enteros no negativos."])
    instructions.append(["Elimina la fila EJ-001 de ejemplo antes de importar o reemplázala."])
    instructions.column_dimensions["A"].width = 100
    instructions["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    instructions["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    for row in range(2, instructions.max_row + 1):
        instructions.cell(row, 1).alignment = Alignment(wrap_text=True)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


init_db()
st.title("📦 Inventario y ventas")
st.caption("Control de productos, ventas y existencias · Moneda: USD")
notice = st.session_state.pop("demo_loaded_notice", None)
if notice:
    st.success(notice)

tab_dashboard, tab_products, tab_sales, tab_excel = st.tabs(["Dashboard", "Productos", "Ventas", "Excel"])

with tab_dashboard:
    products = read_products()
    sales_all = read_sales()
    today = date.today()
    month_names = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    available_months = sorted(sales_all["sold_at"].dropna().str[:7].unique().tolist(), reverse=True) if not sales_all.empty else []
    current_month = today.strftime("%Y-%m")
    if not available_months:
        available_months = [current_month]
    month_choice = st.selectbox(
        "Mes del dashboard",
        available_months,
        index=0,
        format_func=lambda value: f"{month_names[int(value[5:7]) - 1].capitalize()} {value[:4]}",
    )
    year, month = (int(part) for part in month_choice.split("-"))
    month_start = date(year, month, 1)
    month_end = date(year, month, monthrange(year, month)[1])
    month_label = f"{month_names[month - 1].capitalize()} {year}"
    month_sales = sales_all[sales_all["sold_at"].str[:7] == month_choice] if not sales_all.empty else sales_all
    details_month = read_sale_details(month_start, month_end)
    revenue_month = float(month_sales["total"].sum()) if not month_sales.empty else 0.0
    stock_units = int(products["Stock"].sum()) if not products.empty else 0
    stock_value = float((products["Stock"] * products["Precio compra"]).sum()) if not products.empty else 0.0
    low = products[products["Stock"] <= products["Stock mínimo"]] if not products.empty else products

    if products.empty and sales_all.empty:
        st.info("La app está vacía. Puedes cargar una demostración ficticia de accesorios eléctricos.")
        if st.button("Cargar demo de accesorios eléctricos · agosto 2026", type="primary"):
            try:
                products_added, sales_added, demo_total = seed_demo_august()
                st.session_state["demo_loaded_notice"] = f"Demo lista: {products_added} productos y {sales_added} ventas por ${demo_total:,.2f}."
                st.rerun()
            except Exception as exc:
                st.error(str(exc))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Productos registrados", f"{len(products):,}")
    c2.metric("Unidades en inventario", f"{stock_units:,}")
    c3.metric(f"Ventas · {month_label}", f"${revenue_month:,.2f}")
    c4.metric("Productos con stock bajo", f"{len(low):,}")
    st.caption(f"Valor del inventario al costo: ${stock_value:,.2f} · Datos de ventas del mes seleccionado")
    left, right = st.columns(2)
    with left:
        st.subheader(f"Ventas por semana · {month_label}")
        if not month_sales.empty:
            chart = month_sales.copy()
            chart["Fecha"] = pd.to_datetime(chart["sold_at"]).dt.date
            trend = chart.groupby("Fecha", as_index=False)["total"].sum().set_index("Fecha")
            st.line_chart(trend)
        else:
            st.info("No hay ventas registradas en este mes.")
    with right:
        st.subheader(f"Productos más vendidos · {month_label}")
        if not details_month.empty:
            top = details_month.groupby("Producto", as_index=False)["Cantidad"].sum().sort_values("Cantidad", ascending=False).head(8)
            st.bar_chart(top.set_index("Producto"))
        else:
            st.info("El gráfico aparecerá cuando registres ventas en este mes.")
    st.subheader("Alertas de stock")
    if low.empty:
        st.success("No hay productos bajo el stock mínimo.")
    else:
        st.dataframe(low[["SKU", "Nombre", "Categoría", "Stock", "Stock mínimo"]], use_container_width=True, hide_index=True)
    st.subheader(f"Ventas de {month_label}")
    if month_sales.empty:
        st.info("Todavía no hay ventas para el mes seleccionado.")
    else:
        recent = month_sales.head(8).rename(columns={"id": "Venta", "sold_at": "Fecha", "customer": "Cliente", "payment_method": "Forma de pago", "total": "Total USD"})
        st.dataframe(recent, use_container_width=True, hide_index=True)

with tab_products:
    st.subheader("Registrar o actualizar producto")
    with st.form("product_form", clear_on_submit=True):
        a, b, c = st.columns(3)
        sku = a.text_input("SKU / código *")
        name = b.text_input("Nombre del producto *")
        category = c.text_input("Categoría")
        d, e, f, g = st.columns(4)
        cost = d.number_input("Precio de compra (USD)", min_value=0.0, step=0.01, format="%.2f")
        price = e.number_input("Precio de venta (USD)", min_value=0.0, step=0.01, format="%.2f")
        stock = f.number_input("Stock inicial", min_value=0, step=1)
        min_stock = g.number_input("Stock mínimo", min_value=0, step=1)
        save = st.form_submit_button("Guardar producto")
        if save:
            if not sku.strip() or not name.strip():
                st.error("SKU y nombre son obligatorios.")
            else:
                save_product(sku, name, category, cost, price, stock, min_stock)
                st.success("Producto guardado. Si el SKU ya existía, se actualizó.")
                st.rerun()
    st.subheader("Productos registrados")
    products = read_products()
    if products.empty:
        st.info("Aún no hay productos. Puedes agregarlos aquí o cargarlos desde Excel.")
    else:
        query = st.text_input("Buscar por SKU, nombre o categoría", key="product_search")
        if query.strip():
            mask = products.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False)).any(axis=1)
            products = products[mask]
        st.dataframe(products, use_container_width=True, hide_index=True)

with tab_sales:
    st.subheader("Registrar venta")
    products_df = read_products()
    if products_df.empty:
        st.warning("Primero registra o importa productos para poder vender.")
    else:
        if "cart" not in st.session_state:
            st.session_state.cart = []
        product_map = {f"{r['SKU']} — {r['Nombre']} (stock: {r['Stock']})": r for _, r in products_df.iterrows() if int(r["Stock"]) > 0}
        if not product_map:
            st.warning("No hay productos con stock disponible.")
        else:
            with st.form("add_to_cart"):
                chosen = st.selectbox("Producto", list(product_map.keys()))
                selected = product_map[chosen]
                qty_col, price_col = st.columns(2)
                available = int(selected["Stock"]) - sum(item["quantity"] for item in st.session_state.cart if item["sku"] == selected["SKU"])
                qty = qty_col.number_input("Cantidad", min_value=1, max_value=max(1, available), step=1)
                unit_price = price_col.number_input("Precio unitario (USD)", min_value=0.0, value=float(selected["Precio venta"]), step=0.01, format="%.2f")
                add_item = st.form_submit_button("Agregar a la venta")
                if add_item:
                    already = sum(item["quantity"] for item in st.session_state.cart if item["sku"] == selected["SKU"])
                    if int(qty) + already > int(selected["Stock"]):
                        st.error("La cantidad supera el stock disponible.")
                    else:
                        st.session_state.cart.append({"sku": selected["SKU"], "name": selected["Nombre"], "quantity": int(qty), "unit_price": float(unit_price)})
                        st.rerun()
        if st.session_state.get("cart"):
            st.markdown("**Productos de esta venta**")
            cart_rows = [{"SKU": x["sku"], "Producto": x["name"], "Cantidad": x["quantity"], "Precio unitario": x["unit_price"], "Subtotal": round(x["quantity"] * x["unit_price"], 2)} for x in st.session_state.cart]
            st.dataframe(pd.DataFrame(cart_rows), use_container_width=True, hide_index=True)
            total = sum(row["Subtotal"] for row in cart_rows)
            st.markdown(f"**Total: ${total:,.2f}**")
            with st.form("checkout"):
                customer = st.text_input("Cliente (opcional)")
                payment = st.selectbox("Forma de pago", ["Efectivo", "Transferencia", "Tarjeta", "Otro"])
                c1, c2 = st.columns(2)
                finish = c1.form_submit_button("Confirmar venta y descontar stock", type="primary")
                clear = c2.form_submit_button("Vaciar venta")
                if finish:
                    try:
                        sale_id, final_total = record_sale(st.session_state.cart, customer, payment)
                        st.session_state.cart = []
                        st.success(f"Venta #{sale_id} registrada por ${final_total:,.2f}. El stock ya fue actualizado.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
                if clear:
                    st.session_state.cart = []
                    st.rerun()
    st.divider()
    st.subheader("Historial de ventas")
    start_date, end_date = st.date_input("Filtrar fechas", value=(month_start, month_end), key=f"sales_dates_{month_choice}")
    sales = read_sales(start_date, end_date)
    if sales.empty:
        st.info("No hay ventas en ese rango de fechas.")
    else:
        sales = sales.rename(columns={"id": "Venta", "sold_at": "Fecha", "customer": "Cliente", "payment_method": "Forma de pago", "total": "Total USD"})
        st.dataframe(sales, use_container_width=True, hide_index=True)

with tab_excel:
    st.subheader("Importar productos desde Excel")
    st.download_button("Descargar plantilla Excel", data=make_template(), file_name="plantilla_productos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Columnas requeridas: SKU, Nombre, Categoría, Precio compra, Precio venta, Stock y Stock mínimo. Si el SKU ya existe, se actualizan sus datos y stock.")
    uploaded = st.file_uploader("Selecciona un archivo .xlsx", type=["xlsx"])
    if uploaded is not None and st.button("Importar productos", type="primary"):
        try:
            count = import_products(uploaded)
            st.success(f"Se importaron o actualizaron {count} productos.")
            st.rerun()
        except Exception as exc:
            st.error(f"No se pudo importar el archivo: {exc}")
    st.divider()
    st.subheader("Exportar reportes")
    export_start, export_end = st.date_input("Rango para ventas exportadas", value=(month_start, month_end), key=f"export_dates_{month_choice}")
    workbook = make_export(export_start, export_end)
    st.download_button("Descargar reporte Excel", data=workbook, file_name=f"reporte_inventario_{date.today().isoformat()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("El Excel incluye hojas de Productos, Ventas, Detalle de ventas y Stock bajo.")
