"""Generate DA1 diagrams for the Warehouse & Supply Chain Management System.

Eight PNGs:
  er_model      - basic ER model (Chen notation), 7 entities
  eer_model     - the SAME model extended with EER features
  norm_unf      - the single unnormalized table we start from
  norm_1nf      - after 1NF
  norm_2nf      - after 2NF
  norm_3nf      - after 3NF
  norm_bcnf     - the BCNF case and its fix
  final_schema  - final set of tables with keys

Run: python make_diagrams.py
"""
import os
import textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import (FancyBboxPatch, Ellipse, Rectangle, FancyArrowPatch,
                                Circle, Polygon)

OUT = os.path.dirname(os.path.abspath(__file__))

NAVY = "#1f3a5f"
TEAL = "#2a9d8f"
GREEN = "#43aa8b"
AMBER = "#e9c46a"
RED = "#c1121f"
GREY = "#adb5bd"
INK = "#0d1b2a"
LIGHT = "#eef2f7"

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
        ax.plot([x0, x1], [y0 - dy, y0 - dy], color=INK, lw=1.0, zorder=7,
                ls=(0, (2.2, 1.6)) if dashed else "solid")
    _UNDER.clear()


# ---------------------------------------------------------------- primitives
def entity(ax, x, y, text, w=2.6, h=0.95, weak=False, sub=False):
    """Rectangle entity; weak=True adds the double border of a weak entity."""
    ax.add_patch(Rectangle((x - w / 2, y - h / 2), w, h, fc=TEAL if sub else NAVY,
                           ec=INK, lw=1.6, zorder=3))
    if weak:
        ax.add_patch(Rectangle((x - w / 2 + 0.1, y - h / 2 + 0.1), w - 0.2, h - 0.2,
                               fc="none", ec="white", lw=1.4, zorder=4))
    ax.text(x, y, text, ha="center", va="center", color="white", fontsize=9.5,
            weight="bold", zorder=5)


def rel(ax, x, y, text, w=2.5, h=1.15, identifying=False, fs=8.2):
    """Diamond relationship; identifying=True adds the inner double diamond."""
    pts = [(x, y + h / 2), (x + w / 2, y), (x, y - h / 2), (x - w / 2, y)]
    ax.add_patch(Polygon(pts, closed=True, fc=LIGHT, ec=INK, lw=1.5, zorder=3))
    if identifying:
        f = 0.76
        ax.add_patch(Polygon([(x, y + h / 2 * f), (x + w / 2 * f, y),
                              (x, y - h / 2 * f), (x - w / 2 * f, y)],
                             closed=True, fc="none", ec=INK, lw=1.2, zorder=4))
    ax.text(x, y, text, ha="center", va="center", color=NAVY, fontsize=fs,
            weight="bold", zorder=5)


def attr(ax, x, y, text, key=False, partial=False, multi=False, derived=False,
         h=0.62, fs=7.5):
    """Ellipse attribute. Width follows the label so nothing spills out."""
    w = max(1.45, 0.108 * len(text) + 0.52)
    ax.add_patch(Ellipse((x, y), w, h, fc="white", ec=NAVY, lw=1.3,
                         ls=(0, (4, 2)) if derived else "solid", zorder=3))
    if multi:
        ax.add_patch(Ellipse((x, y), w - 0.24, h - 0.18, fc="none", ec=NAVY,
                             lw=1.1, zorder=4))
    t = ax.text(x, y, text, ha="center", va="center", color=INK, fontsize=fs, zorder=5)
    if key or partial:
        _queue_underline(ax, t, dashed=partial)
    return w


def line(ax, p1, p2, label=None, total=False, lw=1.3, color=INK, fs=8.5):
    ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color=color, lw=lw, zorder=1)
    if total:  # a second parallel line marks total participation
        dx, dy = p2[0] - p1[0], p2[1] - p1[1]
        n = (dx ** 2 + dy ** 2) ** 0.5 or 1
        ox, oy = -dy / n * 0.075, dx / n * 0.075
        ax.plot([p1[0] + ox, p2[0] + ox], [p1[1] + oy, p2[1] + oy],
                color=color, lw=lw, zorder=1)
    if label:
        ax.text((p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2 + 0.17, label, ha="center",
                va="center", fontsize=fs, weight="bold", color=RED, zorder=6,
                bbox=dict(fc="white", ec="none", pad=0.8))


def note(ax, x, y, text, color=TEAL, fs=7.8, ha="center"):
    ax.text(x, y, text, ha=ha, va="center", fontsize=fs, color=color,
            style="italic", zorder=6)


def banner(ax, x, y, text, w, color=AMBER, fs=9.5, tc="#3a2f00", h=0.56):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                                boxstyle="round,pad=0.02,rounding_size=0.1",
                                fc=color, ec="none", zorder=3))
    ax.text(x, y, text, ha="center", va="center", fontsize=fs, weight="bold",
            color=tc, zorder=4)


def infobox(ax, x, y, w, title, lines, fc="#eaf6f4", ec=TEAL, tc=None,
            fs=9.0, tfs=10.5, wrapat=None, footer=None, fcolor="#8a6d1a"):
    """Auto-height box. y is the TOP edge. lines is a list of bullet strings."""
    tc = tc or ec
    wrapat = wrapat or int(w * 6.4)
    rows = []
    for ln in lines:
        parts = textwrap.wrap(ln, wrapat) or [""]
        rows.append(parts)
    nlines = sum(len(p) for p in rows)
    fl = textwrap.wrap(footer, wrapat) if footer else []
    h = 0.62 + 0.40 * nlines + 0.16 * len(rows) + (0.30 + 0.34 * len(fl) if fl else 0)
    ax.add_patch(FancyBboxPatch((x, y - h), w, h,
                                boxstyle="round,pad=0.06,rounding_size=0.12",
                                fc=fc, ec=ec, lw=1.3, zorder=3))
    ax.text(x + 0.28, y - 0.36, title, fontsize=tfs, weight="bold", color=tc, zorder=4)
    yy = y - 0.92
    for parts in rows:
        for i, p in enumerate(parts):
            ax.text(x + 0.34 if i == 0 else x + 0.62, yy, ("•  " if i == 0 else "") + p,
                    fontsize=fs, color=INK, va="center", zorder=4)
            yy -= 0.40
        yy -= 0.16
    for p in fl:
        ax.text(x + 0.34, yy - 0.10, p, fontsize=fs - 0.4, color=fcolor,
                style="italic", va="center", zorder=4)
        yy -= 0.34
    return y - h


# The report gives each figure a heading and a caption, and the deck gives each
# slide a title bar, so a title drawn inside the PNG would only repeat it. The
# subtitle stays: it carries the rule the figure is demonstrating.
DRAW_TITLE = False


def canvas(xlim, ylim, figsize, title, subtitle=None):
    fig, ax = plt.subplots(figsize=figsize)
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.axis("off")
    cx = (xlim[0] + xlim[1]) / 2
    if DRAW_TITLE:
        ax.text(cx, ylim[1] - 0.40, title, ha="center", va="center", fontsize=15.5,
                weight="bold", color=NAVY)
    if subtitle:
        ax.text(cx, ylim[1] - 1.00, subtitle, ha="center", va="center",
                fontsize=11 if not DRAW_TITLE else 10,
                color="#555", style="italic")
    return fig, ax


def save(fig, name):
    flush_underlines(fig)
    fig.savefig(os.path.join(OUT, name + ".png"), dpi=170, bbox_inches="tight",
                facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------- relation table
def table(ax, x, y, name, cols, rows, colw=None, fs=7.4, hfs=7.8, rh=0.38,
          headcolor=NAVY, keys=(), fks=()):
    """Relation drawn as a titled grid. (x, y) is the top-left corner.

    keys  -> column indexes underlined       fks -> column indexes italicised
    """
    colw = colw or [max(0.66, 0.09 * len(c) + 0.36) for c in cols]
    w = sum(colw)
    ax.add_patch(Rectangle((x, y - 0.44), w, 0.44, fc=headcolor, ec=INK, lw=1.1, zorder=3))
    ax.text(x + 0.12, y - 0.22, name, ha="left", va="center", color="white",
            fontsize=hfs + 0.7, weight="bold", zorder=4)
    cx = x
    for i, c in enumerate(cols):
        ax.add_patch(Rectangle((cx, y - 0.44 - rh), colw[i], rh, fc="#dfe7f1",
                               ec=INK, lw=0.8, zorder=3))
        t = ax.text(cx + colw[i] / 2, y - 0.44 - rh / 2, c, ha="center", va="center",
                    fontsize=hfs, weight="bold", color=INK, zorder=4,
                    style="italic" if i in fks else "normal")
        if i in keys:
            _queue_underline(ax, t, dy=0.055)
        cx += colw[i]
    for r, row in enumerate(rows):
        cx = x
        yy = y - 0.44 - rh * (r + 2)
        for i, val in enumerate(row):
            ax.add_patch(Rectangle((cx, yy), colw[i], rh,
                                   fc="white" if r % 2 == 0 else "#f7f9fc",
                                   ec=GREY, lw=0.6, zorder=3))
            ax.text(cx + colw[i] / 2, yy + rh / 2, str(val), ha="center", va="center",
                    fontsize=fs, color=INK, zorder=4)
            cx += colw[i]
    return x, x + w, y, y - 0.44 - rh * (len(rows) + 1)


def fdbox(ax, x, y, lines, w=6.4, title="Functional dependencies"):
    h = 0.36 * len(lines) + 0.72
    ax.add_patch(FancyBboxPatch((x, y - h), w, h,
                                boxstyle="round,pad=0.05,rounding_size=0.12",
                                fc="#fff5e0", ec=AMBER, lw=1.3, zorder=3))
    ax.text(x + 0.22, y - 0.34, title, ha="left", va="center", fontsize=8.8,
            weight="bold", color="#7a5c00", zorder=4)
    for i, (txt, bad) in enumerate(lines):
        ax.text(x + 0.32, y - 0.76 - 0.36 * i, txt, ha="left", va="center",
                fontsize=8.3, color=RED if bad else INK,
                weight="bold" if bad else "normal", zorder=4)
    return y - h


# ================================================================ 1. ER MODEL
def er_model():
    fig, ax = canvas((0, 24.5), (0.6, 14.2), (18, 10.4),
                     "ER Model — Warehouse & Supply Chain Management System",
                     "7 entities   ·   7 relationships   ·   Chen notation")

    TOP, MID, BOT = 9.4, 6.5, 3.6
    EX = {"SUPPLIER": 2.3, "PRODUCT": 8.2, "WAREHOUSE": 14.3, "EMPLOYEE": 20.9,
          "CUSTOMER": 2.3, "ORDERS": 8.2, "SHIPMENT": 14.3}

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
    line(ax, (5.25, TOP + 0.58), (5.25, TOP + 1.24), lw=0.9, color=GREY)
    attr(ax, 11.25, TOP + 1.55, "Quantity")
    line(ax, (11.25, TOP + 0.58), (11.25, TOP + 1.24), lw=0.9, color=GREY)
    attr(ax, 10.3, MID, "Qty")
    line(ax, (9.58, MID), (9.45, MID), lw=0.9, color=GREY)

    # entity attributes: two staggered rows above / below
    def fan(cx, ey, items, up=True):
        ya, yb = (12.15, 11.15) if up else (2.35, 1.35)
        anchor = ey + 0.48 if up else ey - 0.48
        for i, (txt, kw) in enumerate(items):
            xx = cx + (-1.15 if i % 2 == 0 else 1.15)
            yy = ya if i < 2 else yb
            attr(ax, xx, yy, txt, **kw)
            line(ax, (xx, yy - 0.31 if up else yy + 0.31), (cx, anchor),
                 lw=0.9, color=GREY)

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

    # legend, left middle
    ax.add_patch(FancyBboxPatch((0.7, 4.55), 5.6, 3.75,
                                boxstyle="round,pad=0.06,rounding_size=0.12",
                                fc="#f7f9fc", ec=GREY, lw=1.2, zorder=3))
    ax.text(0.98, 7.95, "How to read this diagram", fontsize=10.2, weight="bold",
            color=NAVY, zorder=4)
    for i, t in enumerate(["Rectangle  =  entity",
                           "Diamond  =  relationship",
                           "Ellipse  =  attribute",
                           "Underlined  =  primary key",
                           "1 / M / N  =  how many can take part",
                           "Double line  =  must take part"]):
        ax.text(1.05, 7.45 - 0.46 * i, t, fontsize=8.8, color=INK, zorder=4)

    # relationship summary, right middle
    ax.add_patch(FancyBboxPatch((16.7, 4.55), 7.4, 3.75,
                                boxstyle="round,pad=0.06,rounding_size=0.12",
                                fc="#f7f9fc", ec=GREY, lw=1.2, zorder=3))
    ax.text(16.98, 7.95, "The 7 relationships", fontsize=10.2, weight="bold",
            color=NAVY, zorder=4)
    for i, (r, p, c) in enumerate([
            ("SUPPLIES", "SUPPLIER – PRODUCT", "M:N"),
            ("STORED_IN", "PRODUCT – WAREHOUSE", "M:N"),
            ("CONTAINS", "PRODUCT – ORDERS", "M:N"),
            ("WORKS_AT", "WAREHOUSE – EMPLOYEE", "1:N"),
            ("PLACES", "CUSTOMER – ORDERS", "1:N"),
            ("SHIPPED_BY", "ORDERS – SHIPMENT", "1:N"),
            ("DISPATCHED_FROM", "WAREHOUSE – SHIPMENT", "1:N")]):
        yy = 7.48 - 0.45 * i
        ax.text(17.02, yy, r, fontsize=7.9, weight="bold", color=NAVY, zorder=4)
        ax.text(19.55, yy, p, fontsize=7.9, color=INK, zorder=4)
        ax.text(23.55, yy, c, fontsize=7.9, weight="bold", color=RED,
                ha="right", zorder=4)

    save(fig, "er_model")


# =============================================================== 2. EER MODEL
def eer_model():
    fig, ax = canvas((0, 27.0), (-0.55, 14.6), (18, 10.4),
                     "EER Model — the same design, extended with EER features",
                     "added: specialization  ·  weak entity  ·  multivalued, composite "
                     "and derived attributes  ·  recursive relationship")

    TOP, MID, BOT = 9.8, 6.5, 3.4

    for n, x, y in [("SUPPLIER", 2.3, TOP), ("PRODUCT", 8.2, TOP),
                    ("WAREHOUSE", 14.3, TOP), ("EMPLOYEE", 20.9, TOP),
                    ("CUSTOMER", 2.3, BOT), ("SHIPMENT", 14.3, BOT)]:
        entity(ax, x, y, n)
    entity(ax, 8.2, BOT, "ORDERS")

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

    # ---- NEW 1: the M:N CONTAINS becomes a weak entity ORDER_ITEM
    rel(ax, 8.2, 8.2, "FOR", w=2.0, h=1.0)
    entity(ax, 8.2, 6.6, "ORDER_ITEM", w=3.0, h=0.9, weak=True)
    rel(ax, 8.2, 5.0, "HAS", w=2.0, h=1.0, identifying=True)
    line(ax, (8.2, TOP - 0.48), (8.2, 8.71), "1")
    line(ax, (8.2, 7.69), (8.2, 7.06), "N", total=True)
    line(ax, (8.2, 6.14), (8.2, 5.51), "N", total=True)
    line(ax, (8.2, 4.49), (8.2, BOT + 0.48), "1")
    attr(ax, 5.70, 6.60, "ItemNo", partial=True)
    line(ax, (6.43, 6.60), (6.70, 6.60), lw=0.9, color=GREY)
    attr(ax, 11.05, 7.62, "Qty")
    line(ax, (10.6, 7.34), (9.6, 7.02), lw=0.9, color=GREY)
    note(ax, 9.95, 5.0, "identifying relationship", ha="left")
    note(ax, 9.90, 6.60, "weak entity", ha="left")

    # ---- NEW 2: EMPLOYEE specialization — total, disjoint
    ax.add_patch(Circle((20.9, 7.95), 0.31, fc="white", ec=INK, lw=1.4, zorder=4))
    ax.text(20.9, 7.95, "d", ha="center", va="center", fontsize=10.5,
            weight="bold", color=INK, zorder=5)
    line(ax, (20.9, TOP - 0.48), (20.9, 8.26), total=True)
    entity(ax, 18.7, 6.25, "MANAGER", w=2.5, h=0.88, sub=True)
    entity(ax, 23.1, 6.25, "DRIVER", w=2.5, h=0.88, sub=True)
    line(ax, (20.66, 7.70), (18.7, 6.69))
    line(ax, (21.14, 7.70), (23.1, 6.69))
    for xx, t in [(17.55, "Level"), (19.75, "Bonus")]:
        attr(ax, xx, 4.85, t)
        line(ax, (xx, 5.16), (18.7, 5.81), lw=0.9, color=GREY)
    for xx, t in [(22.05, "LicenseNo"), (24.35, "Expiry")]:
        attr(ax, xx, 4.85, t)
        line(ax, (xx, 5.16), (23.1, 5.81), lw=0.9, color=GREY)
    note(ax, 20.9, 3.75, "total + disjoint  —  every employee is exactly one of these")

    # ---- NEW 3: PRODUCT specialization — partial
    ax.add_patch(Circle((5.00, 7.80), 0.31, fc="white", ec=INK, lw=1.4, zorder=4))
    ax.text(5.00, 7.80, "d", ha="center", va="center", fontsize=10.5,
            weight="bold", color=INK, zorder=5)
    line(ax, (7.10, 9.33), (5.24, 7.99))
    entity(ax, 2.60, 6.40, "PERISHABLE\nPRODUCT", w=2.8, h=0.98, sub=True)
    line(ax, (4.78, 7.62), (3.50, 6.89))
    attr(ax, 1.60, 7.75, "ShelfLifeDays")
    line(ax, (1.80, 7.44), (2.20, 6.89), lw=0.9, color=GREY)
    attr(ax, 1.50, 5.15, "StorageTempC")
    line(ax, (1.80, 5.46), (2.20, 5.91), lw=0.9, color=GREY)
    note(ax, 2.60, 4.35, "partial  —  most products\nare not perishable")

    # ---- NEW 4: recursive SUPERVISES on EMPLOYEE
    rel(ax, 25.10, TOP, "SUPERVISES", w=2.7, h=1.2, fs=7.8)
    line(ax, (22.2, TOP + 0.26), (23.90, TOP + 0.10))
    line(ax, (22.2, TOP - 0.26), (23.90, TOP - 0.10))
    ax.text(23.05, TOP + 0.66, "1  supervisor", fontsize=7.3, color=RED,
            weight="bold", ha="center", zorder=6)
    ax.text(23.05, TOP - 0.70, "N  subordinate", fontsize=7.3, color=RED,
            weight="bold", ha="center", zorder=6)

    # ---- attributes, top row
    def fan_up(cx, items, ya=12.5, yb=11.5):
        for i, (txt, kw) in enumerate(items):
            xx = cx + (-1.15 if i % 2 == 0 else 1.15)
            yy = ya if i < 2 else yb
            attr(ax, xx, yy, txt, **kw)
            line(ax, (xx, yy - 0.31), (cx, TOP + 0.48), lw=0.9, color=GREY)

    fan_up(2.3, [("SupplierID", {"key": True}), ("SName", {}),
                 ("City", {}), ("Phone", {"multi": True})])
    note(ax, 4.35, 11.5, "multivalued", ha="left")
    fan_up(8.2, [("ProductID", {"key": True}), ("PName", {}),
                 ("Category", {}), ("UnitPrice", {})])
    fan_up(14.3, [("WarehouseID", {"key": True}), ("WName", {}),
                  ("City", {}), ("Capacity", {})])
    fan_up(20.9, [("EmpID", {"key": True}), ("EName", {}),
                  ("Role", {}), ("Salary", {})])

    # ---- attributes, bottom row  (CUSTOMER shows a composite attribute)
    attr(ax, 1.15, 2.25, "CustomerID", key=True)
    line(ax, (1.15, 2.56), (2.3, BOT - 0.48), lw=0.9, color=GREY)
    attr(ax, 3.45, 2.25, "CName")
    line(ax, (3.45, 2.56), (2.3, BOT - 0.48), lw=0.9, color=GREY)
    attr(ax, 2.3, 1.20, "Address")
    line(ax, (2.3, 1.51), (2.3, 1.94), lw=0.9, color=GREY)
    for xx, t in [(0.85, "Street"), (2.3, "City"), (3.8, "Pincode")]:
        attr(ax, xx, 0.15, t, h=0.56, fs=7.2)
        line(ax, (xx, 0.43), (2.3, 0.89), lw=0.9, color=GREY)
    note(ax, 4.55, 1.20, "composite", ha="left")

    for xx, yy, t, kw in [(7.05, 2.25, "OrderID", {"key": True}),
                          (9.35, 2.25, "OrderDate", {}),
                          (7.05, 1.20, "Status", {}),
                          (9.35, 1.20, "TotalAmt", {"derived": True})]:
        attr(ax, xx, yy, t, **kw)
        line(ax, (xx, yy + 0.31), (8.2, BOT - 0.48), lw=0.9, color=GREY)
    note(ax, 10.45, 1.20, "derived", ha="left")

    for xx, yy, t, kw in [(13.15, 2.25, "ShipmentID", {"key": True}),
                          (15.45, 2.25, "Status", {}),
                          (13.15, 1.20, "DispatchDate", {}),
                          (15.45, 1.20, "DeliveryDate", {})]:
        attr(ax, xx, yy, t, **kw)
        line(ax, (xx, yy + 0.31), (14.3, BOT - 0.48), lw=0.9, color=GREY)

    # ---- what changed
    infobox(ax, 17.3, 3.30, 9.3, "What EER added to the ER model", [
        "ORDER_ITEM — a weak entity replacing the M:N CONTAINS",
        "EMPLOYEE split into MANAGER and DRIVER (specialization)",
        "PERISHABLE_PRODUCT — a partial specialization of PRODUCT",
        "Phone multivalued, Address composite, TotalAmt derived",
        "SUPERVISES — EMPLOYEE related back to itself",
    ], fs=8.6, tfs=10.0, wrapat=70)

    save(fig, "eer_model")


# ============================================================ 3. UNF
def norm_unf():
    fig, ax = canvas((0, 20), (0.85, 9.4), (17, 7.4),
                     "Step 0 — Unnormalized Form (UNF)",
                     "one table holding everything, exactly as a clerk keeps it "
                     "in a spreadsheet")

    l, r, t, b = table(
        ax, 0.5, 7.9, "ORDER_REGISTER   (one row per order)",
        ["OrderID", "OrderDate", "CustID", "CustName", "CustCity", "CustState",
         "SupplierPhones", "Items (ProductID, PName, Category, Qty, Price)"],
        [["O-101", "02-Apr", "C-01", "Sharma Retail", "Chennai", "Tamil Nadu",
          "9845012345, 08022334455",
          "(P-100, Basmati Rice, Staples, 20, 480)\n(P-104, Sunflower Oil, Oils, 50, 132)"],
         ["O-102", "05-Apr", "C-03", "Metro Mart", "Mumbai", "Maharashtra",
          "9820011122", "(P-210, Frozen Peas, Frozen, 30, 96)"]],
        colw=[1.15, 1.25, 1.0, 1.85, 1.4, 1.6, 3.1, 5.6], fs=7.0, hfs=7.2, rh=0.52)

    ax.annotate("", xy=(9.9, b + 0.08), xytext=(9.9, b - 0.42),
                arrowprops=dict(arrowstyle="-|>", color=RED, lw=1.4))
    ax.text(9.6, b - 0.62, "one cell holds many phone numbers", fontsize=8.8,
            color=RED, style="italic", ha="right")
    ax.annotate("", xy=(15.5, b + 0.08), xytext=(15.5, b - 0.42),
                arrowprops=dict(arrowstyle="-|>", color=RED, lw=1.4))
    ax.text(15.8, b - 0.62, "one cell holds a whole group of products", fontsize=8.8,
            color=RED, style="italic", ha="left")

    infobox(ax, 0.5, b - 1.25, 9.0, "What goes wrong", [
        "Cannot add a new customer until they place an order.",
        "Deleting order O-102 wipes out Metro Mart completely.",
        "Renaming a customer means editing every one of their orders.",
        "City and State are repeated on every order row.",
    ], fc="#fdeaea", ec=RED,
        footer="These are called insert, delete and update anomalies.", fcolor=RED)

    infobox(ax, 10.5, b - 1.25, 9.0, "The two rules 1NF will force on us", [
        "Every cell must hold exactly one value.",
        "A repeating group must move into its own table.",
    ], footer="Normalization removes the anomalies one step at a time, and each "
              "step below shows the tables before and after.")

    save(fig, "norm_unf")


# ============================================================ 4. 1NF
def norm_1nf():
    fig, ax = canvas((0, 20), (1.90, 10.0), (17, 7.0),
                     "Step 1 — First Normal Form (1NF)",
                     "make every cell atomic: the repeating group and the phone list "
                     "each move to their own table")

    table(ax, 0.5, 8.5, "ORDER_MASTER",
          ["OrderID", "OrderDate", "CustID", "CustName", "CustCity", "CustState"],
          [["O-101", "02-Apr", "C-01", "Sharma Retail", "Chennai", "Tamil Nadu"],
           ["O-102", "05-Apr", "C-03", "Metro Mart", "Mumbai", "Maharashtra"]],
          colw=[1.1, 1.2, 1.0, 1.8, 1.35, 1.55], keys=(0,))

    table(ax, 11.2, 8.5, "SUPPLIER_PHONE",
          ["SupplierID", "Phone"],
          [["S-01", "9845012345"], ["S-01", "08022334455"], ["S-03", "9820011122"]],
          colw=[1.5, 1.8], keys=(0, 1))

    table(ax, 0.5, 4.9, "ORDER_ITEM",
          ["OrderID", "ProductID", "PName", "Category", "Qty", "Price"],
          [["O-101", "P-100", "Basmati Rice", "Staples", "20", "480"],
           ["O-101", "P-104", "Sunflower Oil", "Oils", "50", "132"],
           ["O-102", "P-210", "Frozen Peas", "Frozen", "30", "96"]],
          colw=[1.1, 1.2, 1.8, 1.3, 0.7, 0.85], keys=(0, 1))
    note(ax, 0.5, 2.55, "key = OrderID + ProductID  (both columns together)",
         color=NAVY, fs=8.6, ha="left")

    infobox(ax, 8.6, 4.9, 5.3, "Now in 1NF", [
        "Every cell holds one value.",
        "Product lines became ORDER_ITEM.",
        "Phone numbers became SUPPLIER_PHONE.",
    ], fs=8.6, wrapat=36)

    infobox(ax, 14.5, 4.9, 5.0, "But still repeating", [
        "'Basmati Rice' and 'Staples' repeat on every order containing P-100.",
    ], fc="#fdeaea", ec=RED, fs=8.6, wrapat=34,
        footer="2NF fixes exactly this.", fcolor=RED)

    save(fig, "norm_1nf")


# ============================================================ 5. 2NF
def norm_2nf():
    fig, ax = canvas((0, 20), (1.20, 10.8), (17, 7.9),
                     "Step 2 — Second Normal Form (2NF)",
                     "no partial dependency: a column must depend on the whole key, "
                     "not just part of it")

    fdbox(ax, 0.5, 9.5, [
        ("{OrderID, ProductID}  ->  Qty, Price        depends on the WHOLE key  —  fine",
         False),
        ("ProductID  ->  PName, Category              depends on PART of the key  —  PARTIAL",
         True),
    ], w=11.4, title="ORDER_ITEM, whose key is {OrderID, ProductID}")

    ax.text(12.3, 9.1, "PName and Category describe the\nproduct, not the order. They move\n"
                       "to a PRODUCT table keyed by\nProductID alone.",
            fontsize=9.2, color=INK, va="top", zorder=4)

    ax.annotate("", xy=(5.8, 7.15), xytext=(5.8, 7.85),
                arrowprops=dict(arrowstyle="-|>", color=TEAL, lw=2.4))
    ax.text(6.15, 7.5, "split", fontsize=9.6, weight="bold", color=TEAL)

    table(ax, 0.5, 7.0, "ORDER_ITEM   (after)",
          ["OrderID", "ProductID", "Qty", "Price"],
          [["O-101", "P-100", "20", "480"],
           ["O-101", "P-104", "50", "132"],
           ["O-102", "P-210", "30", "96"]],
          colw=[1.15, 1.25, 0.8, 0.9], keys=(0, 1))

    table(ax, 5.5, 7.0, "PRODUCT   (new)",
          ["ProductID", "PName", "Category", "UnitPrice"],
          [["P-100", "Basmati Rice", "Staples", "480"],
           ["P-104", "Sunflower Oil", "Oils", "132"],
           ["P-210", "Frozen Peas", "Frozen", "96"]],
          colw=[1.3, 1.9, 1.35, 1.25], keys=(0,))

    table(ax, 12.6, 7.0, "ORDER_MASTER   (unchanged)",
          ["OrderID", "OrderDate", "CustID", "CustName", "CustCity"],
          [["O-101", "02-Apr", "C-01", "Sharma Retail", "Chennai"],
           ["O-102", "05-Apr", "C-03", "Metro Mart", "Mumbai"]],
          colw=[1.05, 1.15, 0.95, 1.75, 1.3], keys=(0,))

    infobox(ax, 0.5, 4.5, 9.0, "What we gained", [
        "A product's name is now stored exactly once.",
        "Renaming a product is a one-row edit.",
        "A product can exist before anyone orders it.",
    ])

    infobox(ax, 10.2, 4.5, 9.3, "Still wrong in ORDER_MASTER", [
        "OrderID  ->  CustID  ->  CustName, CustCity",
        "The customer's name is reached only through CustID — an "
        "indirect, or transitive, route.",
    ], fc="#fdeaea", ec=RED, footer="3NF removes it.", fcolor=RED)

    save(fig, "norm_2nf")


# ============================================================ 6. 3NF
def norm_3nf():
    fig, ax = canvas((0, 20), (1.20, 10.8), (17, 7.9),
                     "Step 3 — Third Normal Form (3NF)",
                     "no transitive dependency: a non-key column must not be "
                     "determined by another non-key column")

    fdbox(ax, 0.5, 9.5, [
        ("OrderID   ->  OrderDate, CustID          direct from the key  —  fine", False),
        ("CustID    ->  CustName, CustCity         reached via CustID  —  TRANSITIVE", True),
        ("CustCity  ->  CustState                  reached via CustCity  —  TRANSITIVE", True),
    ], w=11.4, title="ORDER_MASTER, whose key is {OrderID}")

    ax.text(12.3, 9.1, "Two indirect routes exist. Each\nnon-key determinant gets its own\n"
                       "table, and the old table keeps\nonly a link to it.",
            fontsize=9.2, color=INK, va="top", zorder=4)

    ax.annotate("", xy=(5.8, 7.15), xytext=(5.8, 7.85),
                arrowprops=dict(arrowstyle="-|>", color=TEAL, lw=2.4))
    ax.text(6.15, 7.5, "split twice", fontsize=9.6, weight="bold", color=TEAL)

    table(ax, 0.5, 7.0, "ORDERS",
          ["OrderID", "OrderDate", "Status", "CustID", "WarehouseID"],
          [["O-101", "02-Apr", "SHIPPED", "C-01", "W-01"],
           ["O-102", "05-Apr", "NEW", "C-03", "W-01"]],
          colw=[1.1, 1.2, 1.25, 1.0, 1.45], keys=(0,), fks=(3, 4))

    table(ax, 7.0, 7.0, "CUSTOMER   (new)",
          ["CustID", "CName", "Street", "City", "Phone"],
          [["C-01", "Sharma Retail", "12 MG Rd", "Chennai", "9840011223"],
           ["C-03", "Metro Mart", "5 Link Rd", "Mumbai", "9820044556"]],
          colw=[1.0, 1.85, 1.4, 1.25, 1.6], keys=(0,), fks=(3,))

    table(ax, 15.4, 7.0, "CITY   (new)",
          ["City", "State"], [["Chennai", "Tamil Nadu"], ["Mumbai", "Maharashtra"]],
          colw=[1.35, 1.7], keys=(0,))

    infobox(ax, 0.5, 4.5, 9.0, "What we gained", [
        "A customer's details are stored exactly once.",
        "A city's state is stored exactly once.",
        "Correcting an address is a one-row edit.",
        "Two orders can no longer disagree about a customer.",
    ])

    infobox(ax, 10.2, 4.5, 9.3, "Almost done", [
        "All 16 tables of the design are now in 3NF.",
        "One table still hides a problem 3NF cannot see, because 3NF "
        "allows a non-key determinant so long as what it determines is "
        "part of some key.",
    ], fc="#fff5e0", ec=AMBER, tc="#8a6d1a",
        footer="BCNF closes exactly that loophole.")

    save(fig, "norm_3nf")


# ============================================================ 7. BCNF
def norm_bcnf():
    fig, ax = canvas((0, 20), (0.55, 11.6), (17, 8.7),
                     "Step 4 — Boyce-Codd Normal Form (BCNF)",
                     "the strict rule: the left-hand side of every dependency must "
                     "be a key")

    ax.text(0.5, 9.85, "The rule drivers are assigned by:", fontsize=9.8,
            weight="bold", color=NAVY)
    ax.text(0.5, 9.42, "each driver covers exactly one city   ·   within a city, an "
                       "order always goes to the same driver", fontsize=9.2, color=INK)

    l, r, t, b = table(ax, 0.5, 8.85, "DELIVERY_DUTY",
                       ["OrderID", "DriverID", "City"],
                       [["O-101", "D-05", "Chennai"], ["O-101", "D-09", "Mumbai"],
                        ["O-102", "D-09", "Mumbai"], ["O-103", "D-05", "Chennai"]],
                       colw=[1.2, 1.3, 1.45])

    note(ax, (l + r) / 2, b - 0.42, "'D-09 covers Mumbai' is stored twice", color=RED,
         fs=8.8)
    note(ax, (l + r) / 2, b - 1.15, "change one row and not the other,\n"
                                    "and the table contradicts itself", color=RED, fs=8.4)

    fdbox(ax, 5.6, 8.85, [
        ("{OrderID, City}  ->  DriverID        left side IS a key  —  fine", False),
        ("DriverID  ->  City                   left side is NOT a key  —  BREAKS BCNF",
         True),
    ], w=10.6)

    ax.text(5.9, 6.55, "Candidate keys:   {OrderID, City}   and   {OrderID, DriverID}",
            fontsize=9.4, weight="bold", color=INK)
    ax.text(5.9, 5.95, "Every column belongs to some key, so 3NF is satisfied.\n"
                       "BCNF is not, because DriverID is not a key on its own.",
            fontsize=9.2, color=INK, va="top")

    ax.annotate("", xy=(8.4, 4.55), xytext=(8.4, 5.30),
                arrowprops=dict(arrowstyle="-|>", color=TEAL, lw=2.4))
    ax.text(8.8, 4.9, "split on the bad dependency", fontsize=9.5, weight="bold",
            color=TEAL)

    table(ax, 2.6, 4.30, "DRIVER_CITY",
          ["DriverID", "City"], [["D-05", "Chennai"], ["D-09", "Mumbai"]],
          colw=[1.4, 1.55], keys=(0,))

    table(ax, 7.4, 4.30, "ORDER_DRIVER",
          ["OrderID", "DriverID"],
          [["O-101", "D-05"], ["O-101", "D-09"], ["O-102", "D-09"], ["O-103", "D-05"]],
          colw=[1.35, 1.45], keys=(0, 1))

    infobox(ax, 11.8, 4.30, 8.0, "Check the split", [
        "Both tables are now in BCNF.",
        "Lossless — the join on DriverID returns the four rows.",
        "A driver's city is now stored only once.",
    ], fs=8.8, wrapat=58,
        footer="Trade-off: {OrderID, City} -> DriverID now needs a trigger in DA2.")

    banner(ax, 6.0, 1.15, "Result:  all 16 tables are in BCNF", 9.4, color=GREEN,
           tc="white", fs=12.5, h=0.68)

    save(fig, "norm_bcnf")


# ============================================================ 8. FINAL SCHEMA
def final_schema():
    fig, ax = canvas((0, 20.4), (1.60, 12.3), (17, 8.6),
                     "Final Relational Schema — 16 tables, all in BCNF",
                     "underlined = primary key        italic = foreign key")

    groups = [
        ("Master data", TEAL, [
            ("CITY", ["City", "State"], (0,), ()),
            ("SUPPLIER", ["SupplierID", "SName", "City"], (0,), (2,)),
            ("SUPPLIER_PHONE", ["SupplierID", "Phone"], (0, 1), (0,)),
            ("PRODUCT", ["ProductID", "PName", "Category", "UnitPrice"], (0,), ()),
            ("PERISHABLE_PRODUCT", ["ProductID", "ShelfLifeDays", "StorageTempC"],
             (0,), (0,)),
        ]),
        ("Warehouse & staff", NAVY, [
            ("WAREHOUSE", ["WarehouseID", "WName", "City", "Capacity", "ManagerID"],
             (0,), (2, 4)),
            ("EMPLOYEE", ["EmpID", "EName", "Role", "Salary", "WarehouseID",
                          "SupervisorID"], (0,), (4, 5)),
            ("MANAGER", ["EmpID", "Level", "Bonus"], (0,), (0,)),
            ("DRIVER", ["EmpID", "LicenseNo", "Expiry"], (0,), (0,)),
            ("DRIVER_CITY", ["DriverID", "City"], (0,), (0, 1)),
        ]),
        ("Stock & orders", GREEN, [
            ("SUPPLIES", ["SupplierID", "ProductID", "LeadTimeDays", "SupplyPrice"],
             (0, 1), (0, 1)),
            ("STOCK", ["WarehouseID", "ProductID", "Quantity"], (0, 1), (0, 1)),
            ("CUSTOMER", ["CustID", "CName", "Street", "City", "Phone"], (0,), (3,)),
            ("ORDERS", ["OrderID", "OrderDate", "Status", "CustID", "WarehouseID"],
             (0,), (3, 4)),
            ("ORDER_ITEM", ["OrderID", "ItemNo", "ProductID", "Qty", "Price"],
             (0, 1), (0, 2)),
            ("SHIPMENT", ["ShipmentID", "OrderID", "DispatchDate", "DeliveryDate",
                          "DriverID"], (0,), (1, 4)),
        ]),
    ]

    GW = 6.2       # group width
    x = 0.4
    for gname, gcolor, tables in groups:
        banner(ax, x + GW / 2, 10.5, gname, GW, color=gcolor, tc="white", fs=11)
        y = 9.85
        for tname, cols, keys, fks in tables:
            rh = 0.38
            cw = GW / len(cols)
            ax.add_patch(Rectangle((x, y - 0.40), GW, 0.40, fc=gcolor, ec=INK,
                                   lw=1.0, zorder=3))
            ax.text(x + 0.10, y - 0.20, tname, ha="left", va="center", color="white",
                    fontsize=8.5, weight="bold", zorder=4)
            for i, c in enumerate(cols):
                cx = x + cw * i
                ax.add_patch(Rectangle((cx, y - 0.40 - rh), cw, rh, fc="#f2f6fb",
                                       ec=GREY, lw=0.6, zorder=3))
                t = ax.text(cx + cw / 2, y - 0.40 - rh / 2, c, ha="center", va="center",
                            fontsize=6.5 if len(c) > 11 else 7.1, color=INK, zorder=4,
                            style="italic" if i in fks else "normal",
                            weight="bold" if i in keys else "normal")
                if i in keys:
                    _queue_underline(ax, t, dy=0.055)
            y -= (0.40 + rh + 0.34)
        x += GW + 0.60

    ax.text(10.2, 2.55, "Every table has one subject, every fact is stored in exactly "
                        "one place, and no table can contradict another.",
            ha="center", fontsize=10.2, color=NAVY, weight="bold")
    ax.text(10.2, 2.00, "1NF  atomic cells      ·      2NF  no partial dependency"
                        "      ·      3NF  no transitive dependency"
                        "      ·      BCNF  every determinant is a key",
            ha="center", fontsize=9, color="#555", style="italic")

    save(fig, "final_schema")


if __name__ == "__main__":
    names = ["er_model", "eer_model", "norm_unf", "norm_1nf", "norm_2nf",
             "norm_3nf", "norm_bcnf", "final_schema"]
    for fn in (er_model, eer_model, norm_unf, norm_1nf, norm_2nf, norm_3nf,
               norm_bcnf, final_schema):
        fn()
    for n in names:
        p = os.path.join(OUT, n + ".png")
        assert os.path.getsize(p) > 5000, f"{n}.png missing or too small"
    print(f"{len(names)} diagrams written to {OUT}")
    print("self-check OK")
