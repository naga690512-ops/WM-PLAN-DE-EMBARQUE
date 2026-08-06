import io
import re
from collections import defaultdict

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Plan de Embarque Walmart", page_icon="📦", layout="centered")

# ============================================================================
# Estándares de formato (validados en producción, ver skill embarques-logistica)
# ============================================================================
HEADER_FILL_COLOR = "1F4E78"
HEADER_FONT_COLOR = "FFFFFF"
HEADER_FONT_NAME = "Calibri"
HEADER_FONT_SIZE = 11


def apply_header_style(ws, row: int = 1):
    header_fill = PatternFill(start_color=HEADER_FILL_COLOR, end_color=HEADER_FILL_COLOR, fill_type="solid")
    header_font = Font(name=HEADER_FONT_NAME, size=HEADER_FONT_SIZE, bold=True, color=HEADER_FONT_COLOR)
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def freeze_header_row(ws, row: int = 2):
    ws.freeze_panes = f"A{row}"


def autofit_columns(ws, min_width: int = 8, max_width: int = 40):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = max(min_width, min(length + 2, max_width))


def format_int_column(ws, col_idx: int, start_row: int = 2, end_row=None):
    """Formato numérico '0' (sin decimales, sin notación científica) -- así viene en el archivo real."""
    end_row = end_row or ws.max_row
    for row_idx in range(start_row, end_row + 1):
        ws.cell(row=row_idx, column=col_idx).number_format = "0"


def style_sheet(ws, ean_col_idxs=None, wrap_col_idxs=None):
    apply_header_style(ws, row=1)
    freeze_header_row(ws, row=2)
    for idx in (ean_col_idxs or []):
        format_int_column(ws, idx)
    for idx in (wrap_col_idxs or []):
        for row_idx in range(2, ws.max_row + 1):
            ws.cell(row=row_idx, column=idx).alignment = Alignment(wrap_text=True, vertical="center")
    autofit_columns(ws)


def to_int(v):
    if v is None or v == "":
        return None
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return v


def safe_sheet_name(name: str) -> str:
    for ch in "[]:*?/\\":
        name = name.replace(ch, "")
    return name[:31]


# ============================================================================
# Lectura y normalización del CSV del portal de Walmart
# ============================================================================
COLUMN_ALIASES = {
    "orden_compra": ["orden compra", "núm. pedido", "num pedido", "no. pedido", "pedido"],
    "cadena": ["cadena"],
    "proveedor": ["proveedor", "# proveedor", "num proveedor", "número de proveedor"],
    "ean": ["ean/upc", "ean-13", "ean", "upc"],
    "estilo": ["estilo"],
    "talla": ["talla"],
    "color": ["color"],
    "tienda": ["tienda"],
    "cantidad": ["cantidad", "qty", "piezas"],
}
REQUIRED = ["orden_compra", "cadena", "ean", "estilo", "talla", "tienda", "cantidad"]


def read_wm_csv(uploaded_file):
    """Devuelve (df_normalizado, df_crudo) -- df_crudo conserva las columnas y el
    orden tal cual vienen del portal, para la hoja de respaldo del CSV original."""
    raw = uploaded_file.read()
    for enc in ("latin1", "utf-8", "cp1252"):
        try:
            df_raw = pd.read_csv(io.BytesIO(raw), encoding=enc)
            break
        except Exception:
            continue
    else:
        raise ValueError("No se pudo leer el archivo con ninguna codificación conocida.")

    lower_map = {c.lower().strip(): c for c in df_raw.columns}
    rename = {}
    missing = []
    for std_name, aliases in COLUMN_ALIASES.items():
        found = None
        for alias in aliases:
            if alias in lower_map:
                found = lower_map[alias]
                break
        if found:
            rename[found] = std_name
        elif std_name in REQUIRED:
            missing.append(std_name)

    if missing:
        raise ValueError(
            "No encontré estas columnas esperadas en el CSV: " + ", ".join(missing) +
            ". Columnas del archivo: " + ", ".join(df_raw.columns.tolist())
        )

    df = df_raw.rename(columns=rename)
    keep = [c for c in COLUMN_ALIASES if c in df.columns]
    df = df[keep].copy()

    df["orden_compra"] = df["orden_compra"].astype(str).str.strip()
    df["cadena"] = df["cadena"].astype(str).str.strip()
    df["tienda"] = df["tienda"].astype(str).str.strip()
    df["talla"] = df["talla"].astype(str).str.strip()
    if "proveedor" in df.columns:
        df["proveedor"] = df["proveedor"].astype(str).str.strip()
    if "color" in df.columns:
        df["color"] = df["color"].astype(str).str.strip()
    else:
        df["color"] = ""

    df["ean"] = df["ean"].apply(to_int)
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce").fillna(0).astype(int)
    df["orden_compra_int"] = df["orden_compra"].apply(to_int)
    df["tienda_int"] = df["tienda"].apply(to_int)

    df["cedis"] = df["cadena"].apply(extract_cedis)

    # limpiar EAN también en el crudo, para que no salga en notación científica
    ean_col_raw = next((k for k, v in rename.items() if v == "ean"), None)
    if ean_col_raw and ean_col_raw in df_raw.columns:
        df_raw[ean_col_raw] = df_raw[ean_col_raw].apply(to_int)

    return df, df_raw


def extract_cedis(cadena: str) -> str:
    """El CEDIS viene como los dígitos al inicio del campo Cadena, ej. '7464 NVA WAL-MART DE...'"""
    m = re.match(r"^\s*(\d+)", str(cadena))
    return m.group(1)[:4] if m else ""


def talla_sort_key(t):
    """Ordena tallas numéricamente aunque traigan prefijo/sufijo (T4, T16, 3EG, etc.)."""
    s = str(t)
    m = re.search(r"\d+", s)
    if m:
        num = int(m.group())
        return (0, num, s[m.end():], s[: m.start()])
    return (1, s)


def count_box_position(rows, key_fields):
    """Calcula 'X de Y' agrupando por key_fields, equivalente a COUNTIFS."""
    groups = defaultdict(list)
    for idx, row in enumerate(rows):
        key = tuple(row[f] for f in key_fields)
        groups[key].append(idx)
    result = {}
    for key, indices in groups.items():
        total = len(indices)
        for pos, idx in enumerate(indices, start=1):
            result[idx] = f"{pos} de {total}"
    return result


# ============================================================================
# Construcción del workbook -- estructura calcada del archivo real del usuario
# (envios_WM.xlsx): Detalle Envíos, Resumen por Tienda, Lineal x Tienda (una
# fila = una CAJA física: agrupada por tienda+estilo+color) y una hoja por PO.
# ============================================================================
DETALLE_HEADERS = ["Núm. Pedido", "Cadena", "# Proveedor", "EAN-13", "Estilo", "Talla", "Color", "Tienda", "Cantidad"]


def build_boxes(df: pd.DataFrame) -> list[dict]:
    """Agrupa el maestro en 'cajas': una caja = una combinación única de
    (Núm. Pedido, Tienda, Estilo, Color), con todas sus tallas/EANs dentro."""
    boxes = []
    grp = df.groupby(["orden_compra", "cadena", "cedis", "tienda", "estilo", "color"], sort=False)
    for (oc, cadena, cedis, tienda, estilo, color), g in grp:
        items = sorted(g.to_dict("records"), key=lambda r: talla_sort_key(r["talla"]))
        boxes.append({
            "orden_compra": oc, "cadena": cadena, "cedis": cedis, "tienda": tienda,
            "estilo": estilo, "color": color,
            "items": [(it["ean"], it["talla"], it["color"], it["cantidad"]) for it in items],
            "total_pzas": sum(it["cantidad"] for it in items),
        })
    return boxes


def write_box_sheet(wb, sheet_name, boxes, box_count_scope, extra_ean_count_col=False):
    """Escribe una hoja de cajas (usada tanto para 'Lineal x Tienda' como para
    cada hoja de PO). box_count_scope: lista de cajas sobre la que se calcula
    'Número de Cajas x tienda' (todo el maestro para Lineal, solo el PO para
    la hoja de PO)."""
    max_pos = max((len(b["items"]) for b in boxes), default=0)

    headers = ["Núm. Pedido", "Cadena", "Cedis", "Tienda", "Estilo", "Total Pzas x Estilo"]
    for p in range(1, max_pos + 1):
        headers += [f"EAN-13 P{p}", f"Talla P{p}", f"Color P{p}", f"Cant P{p}"]
    headers += ["EAN / Cantidad", "Número de Cajas x tienda"]
    if extra_ean_count_col:
        headers.append("Núm. EANs x Caja")

    ws = wb.create_sheet(sheet_name)
    ws.append(headers)

    scope_rows = [{"tienda": b["tienda"]} for b in box_count_scope]
    box_pos = count_box_position(scope_rows, key_fields=("tienda",))
    pos_by_id = {id(b): box_pos[idx] for idx, b in enumerate(box_count_scope)}

    for b in boxes:
        row = [to_int(b["orden_compra"]), b["cadena"], to_int(b["cedis"]), to_int(b["tienda"]), b["estilo"], b["total_pzas"]]
        qr_lines = []
        for p in range(max_pos):
            if p < len(b["items"]):
                ean, talla, color, cant = b["items"][p]
                row += [ean, talla, color, cant]
                qr_lines += [str(ean), str(cant)]
            else:
                row += [None, None, None, None]
        row.append("\n".join(qr_lines))
        row.append(pos_by_id.get(id(b), ""))
        if extra_ean_count_col:
            row.append(len(b["items"]))
        ws.append(row)

    ean_cols = [7 + p * 4 for p in range(max_pos)] if max_pos else []
    qr_col = 6 + max_pos * 4 + 1
    style_sheet(ws, ean_col_idxs=ean_cols, wrap_col_idxs=[qr_col])
    return ws


def build_workbook(df: pd.DataFrame, df_raw: pd.DataFrame, raw_sheet_name: str) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    # ---- 0. CSV crudo tal cual llega del portal (respaldo) ----
    ws0 = wb.create_sheet(safe_sheet_name(raw_sheet_name))
    ws0.append(list(df_raw.columns))
    for _, r in df_raw.iterrows():
        ws0.append(list(r.values))

    # ---- 1. Detalle Envíos (maestro) ----
    ws = wb.create_sheet("Detalle Envíos")
    ws.append(DETALLE_HEADERS)
    for _, r in df.iterrows():
        ws.append([
            to_int(r["orden_compra"]), r["cadena"], r.get("proveedor", ""),
            r["ean"], r["estilo"], r["talla"], r["color"], to_int(r["tienda"]), r["cantidad"],
        ])
    style_sheet(ws, ean_col_idxs=[4])

    # ---- 2. Resumen por Tienda ----
    ws = wb.create_sheet("Resumen por Tienda")
    ws.append(["Núm. Pedido", "Cadena", "Tienda", "Total Piezas", "Núm. Estilos", "Núm. EANs"])
    grp = df.groupby(["orden_compra", "cadena", "tienda"], sort=False)
    for (oc, cadena, tienda), g in grp:
        ws.append([to_int(oc), cadena, to_int(tienda), int(g["cantidad"].sum()), g["estilo"].nunique(), g["ean"].nunique()])
    style_sheet(ws)

    # ---- 3. Lineal x Tienda (una fila = una caja) ----
    all_boxes = build_boxes(df)
    write_box_sheet(wb, "Lineal x Tienda", all_boxes, box_count_scope=all_boxes)

    # ---- 4. Una hoja por PO (mismo formato + Núm. EANs x Caja) ----
    for oc in df["orden_compra"].unique():
        po_boxes = [b for b in all_boxes if b["orden_compra"] == oc]
        write_box_sheet(wb, safe_sheet_name(str(oc)), po_boxes, box_count_scope=po_boxes, extra_ean_count_col=True)

    return wb


# ============================================================================
# QR para CEDIS -- mismo contenido que la columna "EAN / Cantidad" del Excel:
# líneas alternadas EAN/Cantidad separadas por salto de línea simple.
# ============================================================================
def build_qr_content(items):
    lines = []
    for ean, qty in items:
        lines.append(str(ean).strip())
        lines.append(str(qty))
    return "\n".join(lines)


def build_labels_pdf(labels, label_size_cm=(10, 3)):
    import qrcode
    from qrcode.constants import ERROR_CORRECT_M
    from reportlab.lib.units import cm
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader

    width_cm, height_cm = label_size_cm
    page_size = (width_cm * cm, height_cm * cm)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=page_size)

    for label in labels:
        content = label["content"]
        title = label.get("title", "")

        qr = qrcode.QRCode(error_correction=ERROR_CORRECT_M, box_size=10, border=1)
        qr.add_data(content)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")

        img_buf = io.BytesIO()
        img.save(img_buf, format="PNG")
        img_buf.seek(0)

        if title:
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(page_size[0] / 2, page_size[1] - 0.4 * cm, title)

        qr_size = min(page_size[0], page_size[1]) * 0.7
        x = (page_size[0] - qr_size) / 2
        y = (page_size[1] - qr_size) / 2 - 0.2 * cm
        c.drawImage(ImageReader(img_buf), x, y, width=qr_size, height=qr_size)
        c.showPage()

    c.save()
    buf.seek(0)
    return buf


# ============================================================================
# Interfaz
# ============================================================================
st.title("📦 Plan de Embarque Walmart")
st.caption("Sube el CSV del portal de proveedores y descarga el Excel completo (CSV crudo, maestro, resumen, lineal y hojas por PO)")

for key in ("df", "boxes", "wb_bytes"):
    if key not in st.session_state:
        st.session_state[key] = None

uploaded = st.file_uploader("CSV de orden de compra (portal Walmart)", type=["csv"])

if uploaded is not None:
    try:
        df, df_raw = read_wm_csv(uploaded)
    except ValueError as e:
        st.error(str(e))
        df = None

    if df is not None:
        st.session_state.df = df
        st.session_state.boxes = build_boxes(df)
        n_pos = df["orden_compra"].nunique()
        n_tiendas = df["tienda"].nunique()
        n_piezas = int(df["cantidad"].sum())
        n_estilos = df["estilo"].nunique()

        st.success("Archivo leído correctamente")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Pedidos (PO)", n_pos)
        c2.metric("Tiendas", n_tiendas)
        c3.metric("Piezas", f"{n_piezas:,}")
        c4.metric("Estilos", n_estilos)

        with st.expander("Ver pedidos detectados"):
            st.write(sorted(df["orden_compra"].unique().tolist()))

        if st.button("Generar Excel", type="primary"):
            with st.spinner("Generando workbook..."):
                raw_name = uploaded.name.rsplit(".", 1)[0]
                wb = build_workbook(df, df_raw, raw_name)
                out = io.BytesIO()
                wb.save(out)
                out.seek(0)
                st.session_state.wb_bytes = out.getvalue()
            st.success("Listo")

if st.session_state.wb_bytes:
    st.download_button(
        "⬇️ Descargar envios_WM.xlsx",
        data=st.session_state.wb_bytes,
        file_name="envios_WM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()
st.subheader("🔲 Generar QR para CEDIS")
st.caption("Un QR por caja (tienda + estilo + color) -- mismo contenido que la columna \"EAN / Cantidad\" del Excel")

if st.session_state.boxes:
    boxes = st.session_state.boxes
    pos = sorted({b["orden_compra"] for b in boxes})
    sel_po = st.selectbox("Pedido (PO)", pos)

    if st.button("Generar QR de este pedido"):
        with st.spinner("Generando QR..."):
            po_boxes = [b for b in boxes if b["orden_compra"] == sel_po]
            scope_rows = [{"tienda": b["tienda"]} for b in po_boxes]
            box_pos = count_box_position(scope_rows, key_fields=("tienda",))

            labels = []
            for idx, b in enumerate(po_boxes):
                items = [(it[0], it[3]) for it in b["items"]]  # (ean, cantidad)
                content = build_qr_content(items)
                titulo = f"Tienda {b['tienda']} · {b['color']} · Caja {box_pos[idx]}"
                labels.append({"content": content, "title": titulo})
            pdf_buf = build_labels_pdf(labels)
        st.success(f"{len(labels)} etiquetas generadas")
        st.download_button(
            "⬇️ Descargar QR_CEDIS.pdf",
            data=pdf_buf,
            file_name=f"QR_CEDIS_{sel_po}.pdf",
            mime="application/pdf",
        )
else:
    st.info("Primero sube y procesa un CSV arriba.")

st.divider()
st.caption("Sin base de datos · los archivos se procesan en memoria y no se guardan.")
