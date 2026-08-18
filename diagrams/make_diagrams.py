"""Generate the DA1 diagrams for the Warehouse & Supply Chain Management System.

Ten PNGs:
  er_model      - basic ER model (Chen notation), 7 entities
  eer_model     - the SAME model extended with EER features
  sample_data   - sample rows used to work out the functional dependencies
  norm_unf      - the single unnormalized table we start from
  norm_1nf      - relations after 1NF
  norm_2nf      - relations after 2NF
  norm_3nf      - relations after 3NF
  norm_bcnf     - relations after BCNF
  decomp_tree   - decomposition tree, UNF to BCNF, with every column
  final_schema  - final set of tables with keys

Every figure holds the diagram and nothing else: no titles, captions,
legends or commentary. All of the reasoning lives in the report text, so the
two can never drift out of agreement.

Run: python make_diagrams.py
"""
import os
import textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse, Rectangle, Circle, Polygon

matplotlib.rcParams["font.family"] = "sans-serif"
matplotlib.rcParams["font.sans-serif"] = ["Arial", "Helvetica", "DejaVu Sans"]

OUT = os.path.dirname(os.path.abspath(__file__))

INK = "#000000"
GREY = "#808080"
HFILL = "#d9d9d9"      # column-header fill
ZEBRA = "#f4f4f4"      # alternate row fill
SHADE = "#e6e6e6"      # marks the parts EER adds to the ER model

UNITS_PER_INCH = 0.75  # one figure unit is this many inches, kept equal everywhere

# text objects queued for an exact-width underline, drawn once the figure is laid out
_UNDER = []


def _queue_underline(ax, txt, dashed=False, dy=0.075):
    _UNDER.append((ax, txt, dashed, dy))


def flush_underlines(fig):
    """Underline every queued text using its real rendered width."""
    fig.canvas.draw()
    r = fig.canvas.get_renderer()
    for ax, txt, dashed, dy in _UNDER:
        bb = txt.get_window_extent(renderer=r)
        inv = ax.transData.inverted()
        (x0, y0), (x1, _) = inv.transform([[bb.x0, bb.y0], [bb.x1, bb.y1]])
        ax.plot([x0, x1], [y0 - dy, y0 - dy], color=INK, lw=0.9, zorder=7,
                ls=(0, (2.2, 1.6)) if dashed else "solid")
    _UNDER.clear()


def canvas(w, h, y0=0.0):
    """Axes w units wide and h units tall. fit() trims it to the content later."""
    fig, ax = plt.subplots(figsize=(w * UNITS_PER_INCH, h * UNITS_PER_INCH))
    ax.set_xlim(0, w)
    ax.set_ylim(y0, y0 + h)
    ax.axis("off")
    return fig, ax


def fit(ax, pad=0.35):
    """Shrink the axes limits onto what was actually drawn, keeping the scale.

    Units per inch stays fixed, so every figure prints its text at the same
    size however wide its content is.
    """
    fig = ax.figure
    fig.canvas.draw()
    r = fig.canvas.get_renderer()
    inv = ax.transData.inverted()
    xs, ys = [], []
    for a in list(ax.patches) + list(ax.texts) + list(ax.lines):
        bb = a.get_window_extent(renderer=r)
        (x0, y0), (x1, y1) = inv.transform([[bb.x0, bb.y0], [bb.x1, bb.y1]])
        xs += [x0, x1]
        ys += [y0, y1]
    x0, x1 = min(xs) - pad, max(xs) + pad
    y0, y1 = min(ys) - pad, max(ys) + pad
    ax.set_xlim(x0, x1)
    ax.set_ylim(y0, y1)
    fig.set_size_inches((x1 - x0) * UNITS_PER_INCH, (y1 - y0) * UNITS_PER_INCH)


def save(fig, name):
    fit(fig.axes[0])
    flush_underlines(fig)
    fig.savefig(os.path.join(OUT, name + ".png"), dpi=170, bbox_inches="tight",
                facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------- ER primitives
def entity(ax, x, y, text, w=2.6, h=0.95, weak=False, shaded=False):
    """Rectangle entity; weak=True adds the double border of a weak entity."""
    ax.add_patch(Rectangle((x - w / 2, y - h / 2), w, h,
                           fc=SHADE if shaded else "white", ec=INK, lw=1.5, zorder=3))
    if weak:
        ax.add_patch(Rectangle((x - w / 2 + 0.11, y - h / 2 + 0.11), w - 0.22,
                               h - 0.22, fc="none", ec=INK, lw=1.0, zorder=4))
    ax.text(x, y, text, ha="center", va="center", color=INK, fontsize=9.5,
            weight="bold", zorder=5)


def rel(ax, x, y, text, w=2.5, h=1.15, identifying=False, fs=8.2, shaded=False):
    """Diamond relationship; identifying=True adds the inner double diamond."""
    pts = [(x, y + h / 2), (x + w / 2, y), (x, y - h / 2), (x - w / 2, y)]
    ax.add_patch(Polygon(pts, closed=True, fc=SHADE if shaded else "white",
                         ec=INK, lw=1.3, zorder=3))
    if identifying:
        f = 0.76
        ax.add_patch(Polygon([(x, y + h / 2 * f), (x + w / 2 * f, y),
                              (x, y - h / 2 * f), (x - w / 2 * f, y)],
                             closed=True, fc="none", ec=INK, lw=1.0, zorder=4))
    ax.text(x, y, text, ha="center", va="center", color=INK, fontsize=fs,
            weight="bold", zorder=5)


def attr(ax, x, y, text, key=False, partial=False, multi=False, derived=False,
         shaded=False, h=0.62, fs=7.5):
    """Ellipse attribute. Width follows the label so nothing spills out."""
    w = max(1.45, 0.108 * len(text) + 0.52)
    ax.add_patch(Ellipse((x, y), w, h, fc=SHADE if shaded else "white", ec=INK,
                         lw=1.1, ls=(0, (4, 2)) if derived else "solid", zorder=3))
    if multi:
        ax.add_patch(Ellipse((x, y), w - 0.24, h - 0.18, fc="none", ec=INK,
                             lw=0.9, zorder=4))
    t = ax.text(x, y, text, ha="center", va="center", color=INK, fontsize=fs, zorder=5)
    if key or partial:
        _queue_underline(ax, t, dashed=partial)
    return w


def line(ax, p1, p2, label=None, total=False, lw=1.1, color=INK, fs=8.5):
    ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color=color, lw=lw, zorder=1)
    if total:  # a second parallel line marks total participation
        dx, dy = p2[0] - p1[0], p2[1] - p1[1]
        n = (dx ** 2 + dy ** 2) ** 0.5 or 1
        ox, oy = -dy / n * 0.075, dx / n * 0.075
        ax.plot([p1[0] + ox, p2[0] + ox], [p1[1] + oy, p2[1] + oy],
                color=color, lw=lw, zorder=1)
    if label:
        ax.text((p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2 + 0.17, label, ha="center",
                va="center", fontsize=fs, weight="bold", color=INK, zorder=6,
                bbox=dict(fc="white", ec="none", pad=0.8))


# ---------------------------------------------------------------- relation table
def table(ax, x, y, name, cols, rows, colw=None, fs=7.2, hfs=7.6, rh=0.38,
          keys=(), fks=()):
    """Relation drawn as a titled grid. (x, y) is the top-left corner.

    keys  -> column indexes underlined       fks -> column indexes italicised
    """
    colw = colw or [max(0.68, 0.095 * len(c) + 0.42) for c in cols]
    w = sum(colw)
    ax.add_patch(Rectangle((x, y - 0.42), w, 0.42, fc="white", ec=INK, lw=1.2,
                           zorder=3))
    ax.text(x + 0.12, y - 0.21, name, ha="left", va="center", color=INK,
            fontsize=hfs + 0.9, weight="bold", zorder=4)
    cx = x
    for i, c in enumerate(cols):
        ax.add_patch(Rectangle((cx, y - 0.42 - rh), colw[i], rh, fc=HFILL,
                               ec=INK, lw=0.7, zorder=3))
        t = ax.text(cx + colw[i] / 2, y - 0.42 - rh / 2, c, ha="center", va="center",
                    fontsize=hfs, weight="bold", color=INK, zorder=4,
                    style="italic" if i in fks else "normal")
        if i in keys:
            _queue_underline(ax, t, dy=0.055)
        cx += colw[i]
    for r, row in enumerate(rows):
        cx = x
        yy = y - 0.42 - rh * (r + 2)
        for i, val in enumerate(row):
            ax.add_patch(Rectangle((cx, yy), colw[i], rh,
                                   fc="white" if r % 2 == 0 else ZEBRA,
                                   ec=GREY, lw=0.5, zorder=3))
            ax.text(cx + colw[i] / 2, yy + rh / 2, str(val), ha="center", va="center",
                    fontsize=fs, color=INK, zorder=4)
            cx += colw[i]
    return x, x + w, y, y - 0.42 - rh * (len(rows) + 1)


def row_of(ax, y, specs, x0=0.4, gap=0.85):
    """Place several relations side by side. Returns the lowest bottom edge."""
    x, bottom = x0, y
    for s in specs:
        _, r, _, b = table(ax, x, y, s["name"], s["cols"], s["rows"],
                           colw=s.get("colw"), keys=s.get("keys", ()),
                           fks=s.get("fks", ()), fs=s.get("fs", 7.2),
                           hfs=s.get("hfs", 7.6))
        x = r + gap
        bottom = min(bottom, b)
    return bottom


# ============================================================== sample instance
# One small instance used all the way through the normalization. Every relation
# carries at least five rows, which is what makes the dependency check possible.

FLAT_COLS = ["OrderID", "OrderDate", "Status", "CustID", "CName", "Street",
             "City", "State", "ProductID", "PName", "Category", "Qty", "Price"]
FLAT_ROWS = [
    ["O-101", "02-Apr", "SHIPPED", "C-01", "Sharma Retail", "12 MG Rd", "Chennai",
     "Tamil Nadu", "P-100", "Basmati Rice", "Staples", "20", "480"],
    ["O-101", "02-Apr", "SHIPPED", "C-01", "Sharma Retail", "12 MG Rd", "Chennai",
     "Tamil Nadu", "P-104", "Sunflower Oil", "Oils", "50", "132"],
    ["O-102", "05-Apr", "NEW", "C-03", "Metro Mart", "5 Link Rd", "Mumbai",
     "Maharashtra", "P-210", "Frozen Peas", "Frozen", "30", "96"],
    ["O-103", "09-Apr", "SHIPPED", "C-01", "Sharma Retail", "12 MG Rd", "Chennai",
     "Tamil Nadu", "P-100", "Basmati Rice", "Staples", "15", "480"],
    ["O-104", "11-Apr", "NEW", "C-05", "Anand Stores", "7 Anna Salai", "Chennai",
     "Tamil Nadu", "P-108", "Sona Masoori", "Staples", "12", "445"],
    ["O-105", "14-Apr", "SHIPPED", "C-03", "Metro Mart", "5 Link Rd", "Mumbai",
     "Maharashtra", "P-104", "Sunflower Oil", "Oils", "25", "132"],
    ["O-105", "14-Apr", "SHIPPED", "C-03", "Metro Mart", "5 Link Rd", "Mumbai",
     "Maharashtra", "P-210", "Frozen Peas", "Frozen", "40", "96"],
]
FLAT_W = [1.05, 1.15, 1.15, 0.95, 1.8, 1.4, 1.15, 1.5, 1.2, 1.75, 1.2, 0.7, 0.8]

ORDER_MASTER_ROWS = [
    ["O-101", "02-Apr", "SHIPPED", "C-01", "Sharma Retail", "12 MG Rd", "Chennai",
     "Tamil Nadu"],
    ["O-102", "05-Apr", "NEW", "C-03", "Metro Mart", "5 Link Rd", "Mumbai",
     "Maharashtra"],
    ["O-103", "09-Apr", "SHIPPED", "C-01", "Sharma Retail", "12 MG Rd", "Chennai",
     "Tamil Nadu"],
    ["O-104", "11-Apr", "NEW", "C-05", "Anand Stores", "7 Anna Salai", "Chennai",
     "Tamil Nadu"],
    ["O-105", "14-Apr", "SHIPPED", "C-03", "Metro Mart", "5 Link Rd", "Mumbai",
     "Maharashtra"],
    ["O-106", "16-Apr", "NEW", "C-05", "Anand Stores", "7 Anna Salai", "Chennai",
     "Tamil Nadu"],
]
ORDER_MASTER = dict(
    name="ORDER_MASTER",
    cols=["OrderID", "OrderDate", "Status", "CustID", "CName", "Street", "City",
          "State"],
    rows=ORDER_MASTER_ROWS,
    colw=[1.05, 1.15, 1.15, 0.95, 1.8, 1.4, 1.15, 1.5], keys=(0,))

ITEM_FULL = dict(
    name="ORDER_ITEM",
    cols=["OrderID", "ProductID", "PName", "Category", "Qty", "Price"],
    rows=[["O-101", "P-100", "Basmati Rice", "Staples", "20", "480"],
          ["O-101", "P-104", "Sunflower Oil", "Oils", "50", "132"],
          ["O-102", "P-210", "Frozen Peas", "Frozen", "30", "96"],
          ["O-103", "P-100", "Basmati Rice", "Staples", "15", "480"],
          ["O-104", "P-108", "Sona Masoori", "Staples", "12", "445"],
          ["O-105", "P-104", "Sunflower Oil", "Oils", "25", "132"],
          ["O-105", "P-210", "Frozen Peas", "Frozen", "40", "96"],
          ["O-106", "P-100", "Basmati Rice", "Staples", "18", "480"]],
    colw=[1.1, 1.2, 1.8, 1.3, 0.7, 0.85], keys=(0, 1))

ITEM_2NF = dict(
    name="ORDER_ITEM",
    cols=["OrderID", "ProductID", "Qty"],
    rows=[[r[0], r[1], r[4]] for r in ITEM_FULL["rows"]],
    colw=[1.15, 1.25, 0.8], keys=(0, 1), fks=(1,))

PRODUCT = dict(
    name="PRODUCT",
    cols=["ProductID", "PName", "Category", "UnitPrice"],
    rows=[["P-100", "Basmati Rice", "Staples", "480"],
          ["P-104", "Sunflower Oil", "Oils", "132"],
          ["P-108", "Sona Masoori", "Staples", "445"],
          ["P-210", "Frozen Peas", "Frozen", "96"]],
    colw=[1.3, 1.9, 1.35, 1.3], keys=(0,))

SUPPLIER_PHONE = dict(
    name="SUPPLIER_PHONE",
    cols=["SupplierID", "Phone"],
    rows=[["S-01", "9845012345"], ["S-01", "08022334455"], ["S-03", "9820011122"],
          ["S-05", "9840199887"], ["S-05", "04428451212"]],
    colw=[1.55, 1.85], keys=(0, 1))

DELIVERY_DUTY = dict(
    name="DELIVERY_DUTY",
    cols=["OrderID", "DriverID", "City"],
    rows=[["O-101", "D-05", "Chennai"], ["O-101", "D-09", "Mumbai"],
          ["O-102", "D-09", "Mumbai"], ["O-103", "D-05", "Chennai"],
          ["O-104", "D-05", "Chennai"], ["O-105", "D-09", "Mumbai"],
          ["O-106", "D-11", "Chennai"]],
    colw=[1.2, 1.3, 1.45], keys=(0, 2))

ORDERS = dict(
    name="ORDERS",
    cols=["OrderID", "OrderDate", "Status", "CustID"],
    rows=[[r[0], r[1], r[2], r[3]] for r in ORDER_MASTER_ROWS],
    colw=[1.1, 1.2, 1.25, 1.0], keys=(0,), fks=(3,))

CUSTOMER = dict(
    name="CUSTOMER",
    cols=["CustID", "CName", "Street", "City"],
    rows=[["C-01", "Sharma Retail", "12 MG Rd", "Chennai"],
          ["C-03", "Metro Mart", "5 Link Rd", "Mumbai"],
          ["C-05", "Anand Stores", "7 Anna Salai", "Chennai"]],
    colw=[1.05, 1.9, 1.5, 1.25], keys=(0,), fks=(3,))

CITY = dict(
    name="CITY", cols=["City", "State"],
    rows=[["Chennai", "Tamil Nadu"], ["Mumbai", "Maharashtra"]],
    colw=[1.35, 1.7], keys=(0,))

DRIVER_CITY = dict(
    name="DRIVER_CITY", cols=["DriverID", "City"],
    rows=[["D-05", "Chennai"], ["D-09", "Mumbai"], ["D-11", "Chennai"]],
    colw=[1.4, 1.55], keys=(0,), fks=(1,))

ORDER_DRIVER = dict(
    name="ORDER_DRIVER", cols=["OrderID", "DriverID"],
    rows=[[r[0], r[1]] for r in DELIVERY_DUTY["rows"]],
    colw=[1.35, 1.45], keys=(0, 1))


# ================================================================ 1. ER MODEL
def er_model():
    fig, ax = canvas(23.2, 13.6, y0=0.6)

    TOP, MID, BOT = 9.4, 6.5, 3.6

    for n, x, y in [("SUPPLIER", 2.3, TOP), ("PRODUCT", 8.2, TOP),
                    ("WAREHOUSE", 14.3, TOP), ("EMPLOYEE", 20.9, TOP),
                    ("CUSTOMER", 2.3, BOT), ("ORDERS", 8.2, BOT),
                    ("SHIPMENT", 14.3, BOT)]:
        entity(ax, x, y, n)

    # horizontal relationships
    rel(ax, 5.25, TOP, "SUPPLIES")
    line(ax, (3.6, TOP), (4.0, TOP), "M")
    line(ax, (6.5, TOP), (6.9, TOP), "N")
    rel(ax, 11.25, TOP, "STORED_IN")
    line(ax, (9.5, TOP), (10.0, TOP), "M")
    line(ax, (12.5, TOP), (13.0, TOP), "N")
    rel(ax, 17.6, TOP, "WORKS_AT")
    line(ax, (15.6, TOP), (16.35, TOP), "1")
    line(ax, (18.85, TOP), (19.6, TOP), "N", total=True)
    rel(ax, 5.25, BOT, "PLACES")
    line(ax, (3.6, BOT), (4.0, BOT), "1")
    line(ax, (6.5, BOT), (6.9, BOT), "N", total=True)
    rel(ax, 11.25, BOT, "SHIPPED_BY")
    line(ax, (9.5, BOT), (10.0, BOT), "1")
    line(ax, (12.5, BOT), (13.0, BOT), "N", total=True)

    # vertical relationships
    rel(ax, 8.2, MID, "CONTAINS")
    line(ax, (8.2, TOP - 0.48), (8.2, MID + 0.58), "N")
    line(ax, (8.2, MID - 0.58), (8.2, BOT + 0.48), "M", total=True)
    rel(ax, 14.3, MID, "DISPATCHED\nFROM", fs=7.6)
    line(ax, (14.3, TOP - 0.48), (14.3, MID + 0.58), "1")
    line(ax, (14.3, MID - 0.58), (14.3, BOT + 0.48), "N", total=True)

    # attributes on relationships
    attr(ax, 5.25, TOP + 1.55, "LeadTimeDays")
    line(ax, (5.25, TOP + 0.58), (5.25, TOP + 1.24), lw=0.8, color=GREY)
    attr(ax, 11.25, TOP + 1.55, "Quantity")
    line(ax, (11.25, TOP + 0.58), (11.25, TOP + 1.24), lw=0.8, color=GREY)
    attr(ax, 10.3, MID, "Qty")
    line(ax, (9.58, MID), (9.45, MID), lw=0.8, color=GREY)

    # entity attributes: two staggered rows above / below
    def fan(cx, ey, items, up=True):
        ya, yb = (12.15, 11.15) if up else (2.35, 1.35)
        anchor = ey + 0.48 if up else ey - 0.48
        for i, (txt, kw) in enumerate(items):
            xx = cx + (-1.15 if i % 2 == 0 else 1.15)
            yy = ya if i < 2 else yb
            attr(ax, xx, yy, txt, **kw)
            line(ax, (xx, yy - 0.31 if up else yy + 0.31), (cx, anchor),
                 lw=0.8, color=GREY)

    fan(2.3, TOP, [("SupplierID", {"key": True}), ("SName", {}),
                   ("City", {}), ("Phone", {})])
    fan(8.2, TOP, [("ProductID", {"key": True}), ("PName", {}),
                   ("Category", {}), ("UnitPrice", {})])
    fan(14.3, TOP, [("WarehouseID", {"key": True}), ("WName", {}),
                    ("City", {}), ("Capacity", {})])
    fan(20.9, TOP, [("EmpID", {"key": True}), ("EName", {}),
                    ("Role", {}), ("Salary", {})])
    fan(2.3, BOT, [("CustomerID", {"key": True}), ("CName", {}),
                   ("Address", {}), ("Phone", {})], up=False)
    fan(8.2, BOT, [("OrderID", {"key": True}), ("OrderDate", {}),
                   ("Status", {}), ("TotalAmt", {})], up=False)
    fan(14.3, BOT, [("ShipmentID", {"key": True}), ("DispatchDate", {}),
                    ("DeliveryDate", {}), ("Status", {})], up=False)

    save(fig, "er_model")


# =============================================================== 2. EER MODEL
def eer_model():
    fig, ax = canvas(27.0, 15.2, y0=-0.6)

    TOP, MID, BOT = 9.8, 6.5, 3.4

    for n, x, y in [("SUPPLIER", 2.3, TOP), ("PRODUCT", 8.2, TOP),
                    ("WAREHOUSE", 14.3, TOP), ("EMPLOYEE", 20.9, TOP),
                    ("CUSTOMER", 2.3, BOT), ("SHIPMENT", 14.3, BOT),
                    ("ORDERS", 8.2, BOT)]:
        entity(ax, x, y, n)

    rel(ax, 5.25, TOP, "SUPPLIES")
    line(ax, (3.6, TOP), (4.0, TOP), "M")
    line(ax, (6.5, TOP), (6.9, TOP), "N")
    rel(ax, 11.25, TOP, "STORED_IN")
    line(ax, (9.5, TOP), (10.0, TOP), "M")
    line(ax, (12.5, TOP), (13.0, TOP), "N")
    rel(ax, 17.6, TOP, "WORKS_AT")
    line(ax, (15.6, TOP), (16.35, TOP), "1")
    line(ax, (18.85, TOP), (19.6, TOP), "N", total=True)
    rel(ax, 5.25, BOT, "PLACES")
    line(ax, (3.6, BOT), (4.0, BOT), "1")
    line(ax, (6.5, BOT), (6.9, BOT), "N", total=True)
    rel(ax, 11.25, BOT, "SHIPPED_BY")
    line(ax, (9.5, BOT), (10.0, BOT), "1")
    line(ax, (12.5, BOT), (13.0, BOT), "N", total=True)
    rel(ax, 14.3, MID, "DISPATCHED\nFROM", fs=7.6)
    line(ax, (14.3, TOP - 0.48), (14.3, MID + 0.58), "1")
    line(ax, (14.3, MID - 0.58), (14.3, BOT + 0.48), "N", total=True)

    # ---- added 1: the M:N CONTAINS becomes a weak entity ORDER_ITEM
    rel(ax, 8.2, 8.2, "FOR", w=2.0, h=1.0, shaded=True)
    entity(ax, 8.2, 6.6, "ORDER_ITEM", w=3.0, h=0.9, weak=True, shaded=True)
    rel(ax, 8.2, 5.0, "HAS", w=2.0, h=1.0, identifying=True, shaded=True)
    line(ax, (8.2, TOP - 0.48), (8.2, 8.71), "1")
    line(ax, (8.2, 7.69), (8.2, 7.06), "N", total=True)
    line(ax, (8.2, 6.14), (8.2, 5.51), "N", total=True)
    line(ax, (8.2, 4.49), (8.2, BOT + 0.48), "1")
    attr(ax, 5.70, 6.60, "ItemNo", partial=True, shaded=True)
    line(ax, (6.43, 6.60), (6.70, 6.60), lw=0.8, color=GREY)
    attr(ax, 11.05, 7.62, "Qty")
    line(ax, (10.6, 7.34), (9.6, 7.02), lw=0.8, color=GREY)

    # ---- added 2: EMPLOYEE specialization, total and disjoint
    ax.add_patch(Circle((20.9, 7.95), 0.31, fc=SHADE, ec=INK, lw=1.2, zorder=4))
    ax.text(20.9, 7.95, "d", ha="center", va="center", fontsize=10.5,
            weight="bold", color=INK, zorder=5)
    line(ax, (20.9, TOP - 0.48), (20.9, 8.26), total=True)
    entity(ax, 18.7, 6.25, "MANAGER", w=2.5, h=0.88, shaded=True)
    entity(ax, 23.1, 6.25, "DRIVER", w=2.5, h=0.88, shaded=True)
    line(ax, (20.66, 7.70), (18.7, 6.69))
    line(ax, (21.14, 7.70), (23.1, 6.69))
    for xx, t in [(17.55, "Level"), (19.75, "Bonus")]:
        attr(ax, xx, 4.85, t, shaded=True)
        line(ax, (xx, 5.16), (18.7, 5.81), lw=0.8, color=GREY)
    for xx, t in [(22.05, "LicenseNo"), (24.35, "Expiry")]:
        attr(ax, xx, 4.85, t, shaded=True)
        line(ax, (xx, 5.16), (23.1, 5.81), lw=0.8, color=GREY)

    # ---- added 3: PRODUCT specialization, partial
    ax.add_patch(Circle((5.00, 7.80), 0.31, fc=SHADE, ec=INK, lw=1.2, zorder=4))
    ax.text(5.00, 7.80, "d", ha="center", va="center", fontsize=10.5,
            weight="bold", color=INK, zorder=5)
    line(ax, (7.10, 9.33), (5.24, 7.99))
    entity(ax, 2.60, 6.40, "PERISHABLE\nPRODUCT", w=2.8, h=0.98, shaded=True)
    line(ax, (4.78, 7.62), (3.50, 6.89))
    attr(ax, 1.60, 7.75, "ShelfLifeDays", shaded=True)
    line(ax, (1.80, 7.44), (2.20, 6.89), lw=0.8, color=GREY)
    attr(ax, 1.50, 5.15, "StorageTempC", shaded=True)
    line(ax, (1.80, 5.46), (2.20, 5.91), lw=0.8, color=GREY)

    # ---- added 4: recursive SUPERVISES on EMPLOYEE
    rel(ax, 25.10, TOP, "SUPERVISES", w=2.7, h=1.2, fs=7.8, shaded=True)
    line(ax, (22.2, TOP + 0.26), (23.90, TOP + 0.10))
    line(ax, (22.2, TOP - 0.26), (23.90, TOP - 0.10))
    ax.text(23.05, TOP + 0.66, "1  supervisor", fontsize=7.3, color=INK,
            weight="bold", ha="center", zorder=6)
    ax.text(23.05, TOP - 0.70, "N  subordinate", fontsize=7.3, color=INK,
            weight="bold", ha="center", zorder=6)

    # ---- attributes, top row
    def fan_up(cx, items, ya=12.5, yb=11.5):
        for i, (txt, kw) in enumerate(items):
            xx = cx + (-1.15 if i % 2 == 0 else 1.15)
            yy = ya if i < 2 else yb
            attr(ax, xx, yy, txt, **kw)
            line(ax, (xx, yy - 0.31), (cx, TOP + 0.48), lw=0.8, color=GREY)

    fan_up(2.3, [("SupplierID", {"key": True}), ("SName", {}),
                 ("City", {}), ("Phone", {"multi": True, "shaded": True})])
    fan_up(8.2, [("ProductID", {"key": True}), ("PName", {}),
                 ("Category", {}), ("UnitPrice", {})])
    fan_up(14.3, [("WarehouseID", {"key": True}), ("WName", {}),
                  ("City", {}), ("Capacity", {})])
    fan_up(20.9, [("EmpID", {"key": True}), ("EName", {}),
                  ("Role", {}), ("Salary", {})])

    # ---- attributes, bottom row (CUSTOMER shows a composite attribute)
    attr(ax, 1.15, 2.25, "CustomerID", key=True)
    line(ax, (1.15, 2.56), (2.3, BOT - 0.48), lw=0.8, color=GREY)
    attr(ax, 3.45, 2.25, "CName")
    line(ax, (3.45, 2.56), (2.3, BOT - 0.48), lw=0.8, color=GREY)
    attr(ax, 2.3, 1.20, "Address", shaded=True)
    line(ax, (2.3, 1.51), (2.3, 1.94), lw=0.8, color=GREY)
    for xx, t in [(0.85, "Street"), (2.3, "City"), (3.8, "Pincode")]:
        attr(ax, xx, 0.15, t, h=0.56, fs=7.2, shaded=True)
        line(ax, (xx, 0.43), (2.3, 0.89), lw=0.8, color=GREY)

    for xx, yy, t, kw in [(7.05, 2.25, "OrderID", {"key": True}),
                          (9.35, 2.25, "OrderDate", {}),
                          (7.05, 1.20, "Status", {}),
                          (9.35, 1.20, "TotalAmt", {"derived": True,
                                                    "shaded": True})]:
        attr(ax, xx, yy, t, **kw)
        line(ax, (xx, yy + 0.31), (8.2, BOT - 0.48), lw=0.8, color=GREY)

    for xx, yy, t, kw in [(13.15, 2.25, "ShipmentID", {"key": True}),
                          (15.45, 2.25, "Status", {}),
                          (13.15, 1.20, "DispatchDate", {}),
                          (15.45, 1.20, "DeliveryDate", {})]:
        attr(ax, xx, yy, t, **kw)
        line(ax, (xx, yy + 0.31), (14.3, BOT - 0.48), lw=0.8, color=GREY)

    save(fig, "eer_model")


# ====================================================== 3. SAMPLE DATA FOR FDs
def sample_data():
    fig, ax = canvas(20.0, 9.6, y0=1.0)

    b = row_of(ax, 10.2, [dict(name="ORDER_LINE (all order and product facts in "
                                    "one relation)",
                               cols=FLAT_COLS, rows=FLAT_ROWS, colw=FLAT_W,
                               fs=6.8, hfs=7.0)])
    row_of(ax, b - 0.9, [DELIVERY_DUTY, SUPPLIER_PHONE])

    save(fig, "sample_data")


# ============================================================ 4. UNF
def norm_unf():
    fig, ax = canvas(21.0, 6.0, y0=1.6)

    table(ax, 0.4, 6.9, "ORDER_REGISTER   (one row per order)",
          ["OrderID", "OrderDate", "Status", "CustID", "CName", "Street", "City",
           "State", "SupplierPhones",
           "Items (ProductID, PName, Category, Qty, Price)",
           "Delivery (DriverID, City)"],
          [["O-101", "02-Apr", "SHIPPED", "C-01", "Sharma Retail", "12 MG Rd",
            "Chennai", "Tamil Nadu", "9845012345,\n08022334455",
            "(P-100, Basmati Rice, Staples, 20, 480)\n"
            "(P-104, Sunflower Oil, Oils, 50, 132)",
            "(D-05, Chennai)\n(D-09, Mumbai)"],
           ["O-102", "05-Apr", "NEW", "C-03", "Metro Mart", "5 Link Rd", "Mumbai",
            "Maharashtra", "9820011122",
            "(P-210, Frozen Peas, Frozen, 30, 96)", "(D-09, Mumbai)"],
           ["O-104", "11-Apr", "NEW", "C-05", "Anand Stores", "7 Anna Salai",
            "Chennai", "Tamil Nadu", "9840199887,\n04428451212",
            "(P-108, Sona Masoori, Staples, 12, 445)", "(D-05, Chennai)"]],
          colw=[1.0, 1.1, 1.1, 0.9, 1.75, 1.4, 1.15, 1.45, 2.3, 5.3, 2.0],
          fs=6.6, hfs=6.8, rh=1.05)

    save(fig, "norm_unf")


# ============================================================ 5. 1NF
def norm_1nf():
    fig, ax = canvas(21.0, 11.0, y0=0.4)
    b = row_of(ax, 10.6, [ORDER_MASTER, DELIVERY_DUTY, SUPPLIER_PHONE])
    row_of(ax, b - 0.9, [ITEM_FULL])
    save(fig, "norm_1nf")


# ============================================================ 6. 2NF
def norm_2nf():
    fig, ax = canvas(21.0, 11.0, y0=0.4)
    b = row_of(ax, 10.6, [ORDER_MASTER, DELIVERY_DUTY, SUPPLIER_PHONE])
    row_of(ax, b - 0.9, [ITEM_2NF, PRODUCT])
    save(fig, "norm_2nf")


# ============================================================ 7. 3NF
def norm_3nf():
    fig, ax = canvas(21.0, 11.6, y0=0.4)
    b = row_of(ax, 11.2, [ORDERS, CUSTOMER, CITY, DELIVERY_DUTY])
    row_of(ax, b - 0.9, [ITEM_2NF, PRODUCT, SUPPLIER_PHONE])
    save(fig, "norm_3nf")


# ============================================================ 8. BCNF
def norm_bcnf():
    fig, ax = canvas(21.0, 11.6, y0=0.4)
    b = row_of(ax, 11.2, [ORDERS, CUSTOMER, CITY, DRIVER_CITY])
    row_of(ax, b - 0.9, [ITEM_2NF, PRODUCT, SUPPLIER_PHONE, ORDER_DRIVER])
    save(fig, "norm_bcnf")


# ================================================= 9. DECOMPOSITION TREE
# (name, columns, parent name in the previous stage, key)
TREE = [
    ("UNF", [
        ("ORDER_REGISTER",
         ["OrderID", "OrderDate", "Status", "CustID", "CName", "Street", "City",
          "State", "SupplierPhones *", "Items *", "Delivery *"], None,
         "no atomic key yet   (* = repeating group)"),
    ]),
    ("1NF", [
        ("ORDER_MASTER", ["OrderID", "OrderDate", "Status", "CustID", "CName",
                          "Street", "City", "State"], "ORDER_REGISTER", "OrderID"),
        ("ORDER_ITEM", ["OrderID", "ProductID", "PName", "Category", "Qty",
                        "Price"], "ORDER_REGISTER", "OrderID + ProductID"),
        ("SUPPLIER_PHONE", ["SupplierID", "Phone"], "ORDER_REGISTER",
         "SupplierID + Phone"),
        ("DELIVERY_DUTY", ["OrderID", "DriverID", "City"], "ORDER_REGISTER",
         "OrderID + City"),
    ]),
    ("2NF", [
        ("ORDER_MASTER", ["OrderID", "OrderDate", "Status", "CustID", "CName",
                          "Street", "City", "State"], "ORDER_MASTER", "OrderID"),
        ("ORDER_ITEM", ["OrderID", "ProductID", "Qty"], "ORDER_ITEM",
         "OrderID + ProductID"),
        ("PRODUCT", ["ProductID", "PName", "Category", "UnitPrice"], "ORDER_ITEM",
         "ProductID"),
        ("SUPPLIER_PHONE", ["SupplierID", "Phone"], "SUPPLIER_PHONE",
         "SupplierID + Phone"),
        ("DELIVERY_DUTY", ["OrderID", "DriverID", "City"], "DELIVERY_DUTY",
         "OrderID + City"),
    ]),
    ("3NF", [
        ("ORDERS", ["OrderID", "OrderDate", "Status", "CustID"], "ORDER_MASTER",
         "OrderID"),
        ("CUSTOMER", ["CustID", "CName", "Street", "City"], "ORDER_MASTER",
         "CustID"),
        ("CITY", ["City", "State"], "ORDER_MASTER", "City"),
        ("ORDER_ITEM", ["OrderID", "ProductID", "Qty"], "ORDER_ITEM",
         "OrderID + ProductID"),
        ("PRODUCT", ["ProductID", "PName", "Category", "UnitPrice"], "PRODUCT",
         "ProductID"),
        ("SUPPLIER_PHONE", ["SupplierID", "Phone"], "SUPPLIER_PHONE",
         "SupplierID + Phone"),
        ("DELIVERY_DUTY", ["OrderID", "DriverID", "City"], "DELIVERY_DUTY",
         "OrderID + City"),
    ]),
    ("BCNF", [
        ("ORDERS", ["OrderID", "OrderDate", "Status", "CustID"], "ORDERS",
         "OrderID"),
        ("CUSTOMER", ["CustID", "CName", "Street", "City"], "CUSTOMER", "CustID"),
        ("CITY", ["City", "State"], "CITY", "City"),
        ("ORDER_ITEM", ["OrderID", "ProductID", "Qty"], "ORDER_ITEM",
         "OrderID + ProductID"),
        ("PRODUCT", ["ProductID", "PName", "Category", "UnitPrice"], "PRODUCT",
         "ProductID"),
        ("SUPPLIER_PHONE", ["SupplierID", "Phone"], "SUPPLIER_PHONE",
         "SupplierID + Phone"),
        ("DRIVER_CITY", ["DriverID", "City"], "DELIVERY_DUTY", "DriverID"),
        ("ORDER_DRIVER", ["OrderID", "DriverID"], "DELIVERY_DUTY",
         "OrderID + DriverID"),
    ]),
]


def decomp_tree():
    W, H = 21.6, 13.2
    fig, ax = canvas(W, H)

    NW = 3.55                                   # node width
    colx = [0.35 + i * 4.25 for i in range(len(TREE))]
    leaves = TREE[-1][1]
    slot = (H - 1.9) / len(leaves)

    y = {}
    for i, node in enumerate(leaves):
        y[(len(TREE) - 1, node[0])] = H - 1.45 - slot * (i + 0.5)
    for li in range(len(TREE) - 2, -1, -1):
        for node in TREE[li][1]:
            kids = [k for k in TREE[li + 1][1] if k[2] == node[0]]
            ys = [y[(li + 1, k[0])] for k in kids] or [H / 2]
            y[(li, node[0])] = sum(ys) / len(ys)

    prev_cols = {}
    for li, (stage, nodes) in enumerate(TREE):
        ax.text(colx[li] + NW / 2, H - 0.55, stage, ha="center", va="center",
                fontsize=13, weight="bold", color=INK)
        ax.plot([colx[li], colx[li] + NW], [H - 0.95, H - 0.95], color=INK, lw=1.0)
        for name, cols, parent, keytext in nodes:
            cy = y[(li, name)]
            changed = li == 0 or cols != prev_cols.get(parent)
            body = textwrap.wrap(", ".join(cols), 44)
            kl = textwrap.wrap("key: " + keytext, 44)
            h = 0.52 + 0.30 * len(body) + 0.30 * len(kl) + 0.16
            x = colx[li]
            ax.add_patch(Rectangle((x, cy - h / 2), NW, h, fc="white", ec=INK,
                                   lw=1.5 if changed else 0.7, zorder=3))
            yy = cy + h / 2 - 0.32
            ax.text(x + 0.14, yy, name, fontsize=8.4, weight="bold", color=INK,
                    va="center", zorder=4)
            yy -= 0.34
            for ln in body:
                ax.text(x + 0.14, yy, ln, fontsize=6.9, color=INK, va="center",
                        zorder=4)
                yy -= 0.30
            yy -= 0.02
            for ln in kl:
                ax.text(x + 0.14, yy, ln, fontsize=6.9, color="#444444",
                        style="italic", va="center", zorder=4)
                yy -= 0.30
            if parent is not None:                       # elbow back to the parent
                py = y[(li - 1, parent)]
                x0, x1 = colx[li - 1] + NW, x
                xm = (x0 + x1) / 2
                lw = 1.2 if changed else 0.6
                ax.plot([x0, xm], [py, py], color=INK, lw=lw, zorder=1)
                ax.plot([xm, xm], [py, cy], color=INK, lw=lw, zorder=1)
                ax.plot([xm, x1], [cy, cy], color=INK, lw=lw, zorder=1)
        prev_cols = {n[0]: n[1] for n in nodes}

    save(fig, "decomp_tree")


# ============================================================ 10. FINAL SCHEMA
def final_schema():
    fig, ax = canvas(20.4, 11.4, y0=1.2)

    groups = [
        ("Master data", [
            ("CITY", ["City", "State"], (0,), ()),
            ("SUPPLIER", ["SupplierID", "SName", "City"], (0,), (2,)),
            ("SUPPLIER_PHONE", ["SupplierID", "Phone"], (0, 1), (0,)),
            ("PRODUCT", ["ProductID", "PName", "Category", "UnitPrice"], (0,), ()),
            ("PERISHABLE_PRODUCT", ["ProductID", "ShelfLifeDays", "StorageTempC"],
             (0,), (0,)),
        ]),
        ("Warehouse and staff", [
            ("WAREHOUSE", ["WarehouseID", "WName", "City", "Capacity", "ManagerID"],
             (0,), (2, 4)),
            ("EMPLOYEE", ["EmpID", "EName", "Role", "Salary", "WarehouseID",
                          "SupervisorID"], (0,), (4, 5)),
            ("MANAGER", ["EmpID", "Level", "Bonus"], (0,), (0,)),
            ("DRIVER", ["EmpID", "LicenseNo", "Expiry"], (0,), (0,)),
            ("DRIVER_CITY", ["DriverID", "City"], (0,), (0, 1)),
        ]),
        ("Stock and orders", [
            ("SUPPLIES", ["SupplierID", "ProductID", "LeadTimeDays", "SupplyPrice"],
             (0, 1), (0, 1)),
            ("STOCK", ["WarehouseID", "ProductID", "Quantity"], (0, 1), (0, 1)),
            ("CUSTOMER", ["CustID", "CName", "Street", "City", "Phone"], (0,), (3,)),
            ("ORDERS", ["OrderID", "OrderDate", "Status", "CustID"], (0,), (3,)),
            ("ORDER_ITEM", ["OrderID", "ItemNo", "ProductID", "Qty"], (0, 1),
             (0, 2)),
            ("ORDER_DRIVER", ["OrderID", "DriverID"], (0, 1), (0, 1)),
            ("SHIPMENT", ["ShipmentID", "OrderID", "WarehouseID", "DispatchDate",
                          "DeliveryDate"], (0,), (1, 2)),
        ]),
    ]

    GW = 6.2       # group width
    x = 0.4
    for gname, tables in groups:
        ax.text(x, 10.62, gname, fontsize=11, weight="bold", color=INK, zorder=4)
        ax.plot([x, x + GW], [10.42, 10.42], color=INK, lw=1.2, zorder=4)
        y = 9.95
        for tname, cols, keys, fks in tables:
            rh = 0.38
            cw = GW / len(cols)
            ax.add_patch(Rectangle((x, y - 0.40), GW, 0.40, fc="white", ec=INK,
                                   lw=1.1, zorder=3))
            ax.text(x + 0.10, y - 0.20, tname, ha="left", va="center", color=INK,
                    fontsize=8.5, weight="bold", zorder=4)
            for i, c in enumerate(cols):
                cx = x + cw * i
                ax.add_patch(Rectangle((cx, y - 0.40 - rh), cw, rh, fc=HFILL,
                                       ec=GREY, lw=0.5, zorder=3))
                t = ax.text(cx + cw / 2, y - 0.40 - rh / 2, c, ha="center",
                            va="center", fontsize=6.5 if len(c) > 11 else 7.1,
                            color=INK, zorder=4,
                            style="italic" if i in fks else "normal",
                            weight="bold" if i in keys else "normal")
                if i in keys:
                    _queue_underline(ax, t, dy=0.055)
            y -= (0.40 + rh + 0.34)
        x += GW + 0.60

    save(fig, "final_schema")


FIGURES = [er_model, eer_model, sample_data, norm_unf, norm_1nf, norm_2nf,
           norm_3nf, norm_bcnf, decomp_tree, final_schema]
NAMES = [f.__name__ for f in FIGURES]

if __name__ == "__main__":
    for fn in FIGURES:
        fn()
    for n in NAMES:
        p = os.path.join(OUT, n + ".png")
        assert os.path.getsize(p) > 5000, f"{n}.png missing or too small"
    print(f"{len(NAMES)} diagrams written to {OUT}")
    print("self-check OK")
