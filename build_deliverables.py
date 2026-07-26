"""Build DA1_Report.docx and DA1_Presentation.pptx from the diagram PNGs.

Run: python build_deliverables.py     (run make_diagrams.py first)
"""
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

ROOT = os.path.dirname(os.path.abspath(__file__))
DIAG = os.path.join(ROOT, "diagrams")

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
TEAL = RGBColor(0x2A, 0x9D, 0x8F)
PNAVY = PRGB(0x1F, 0x3A, 0x5F)
PTEAL = PRGB(0x2A, 0x9D, 0x8F)
PWHITE = PRGB(0xFF, 0xFF, 0xFF)
PDARK = PRGB(0x22, 0x22, 0x22)
PMIST = PRGB(0xCF, 0xE3, 0xDF)

TITLE = "Warehouse and Supply Chain Management System"
TEAM = [("Safwaan Mohamed S", "25BCE1188"),
        ("Avi Dhandhania", "25BCE1207"),
        ("Priyal Bhalla", "25BCE1261")]

ENTITIES = [
    ("WAREHOUSE", "WarehouseID", "WName, City, Capacity"),
    ("PRODUCT", "ProductID", "PName, Category, UnitPrice"),
    ("SUPPLIER", "SupplierID", "SName, City, Phone (many per supplier)"),
    ("CUSTOMER", "CustomerID", "CName, Address (Street + City + Pincode), Phone"),
    ("ORDERS", "OrderID", "OrderDate, Status, TotalAmt (calculated)"),
    ("SHIPMENT", "ShipmentID", "DispatchDate, DeliveryDate, Status"),
    ("EMPLOYEE", "EmpID", "EName, Role, Salary"),
]

RELATIONSHIPS = [
    ("SUPPLIES", "SUPPLIER – PRODUCT", "M:N", "A supplier supplies many products; "
     "a product comes from many suppliers. Carries LeadTimeDays."),
    ("STORED_IN", "PRODUCT – WAREHOUSE", "M:N", "A product is stocked in many "
     "warehouses. Carries Quantity."),
    ("CONTAINS", "PRODUCT – ORDERS", "M:N", "An order contains many products. "
     "Carries Qty."),
    ("WORKS_AT", "WAREHOUSE – EMPLOYEE", "1:N", "Every employee is posted to "
     "exactly one warehouse."),
    ("PLACES", "CUSTOMER – ORDERS", "1:N", "Every order belongs to exactly one "
     "customer."),
    ("SHIPPED_BY", "ORDERS – SHIPMENT", "1:N", "One order can go out in more than "
     "one shipment."),
    ("DISPATCHED_FROM", "WAREHOUSE – SHIPMENT", "1:N", "Every shipment leaves from "
     "one warehouse."),
]

EER_ADDED = [
    ("Weak entity", "ORDER_ITEM", "An order line means nothing without its order, "
     "and ItemNo (1, 2, 3 …) is only unique inside one order. So its key is "
     "OrderID + ItemNo. This replaces the M:N CONTAINS relationship."),
    ("Specialization\n(total, disjoint)", "EMPLOYEE →\nMANAGER, DRIVER",
     "Every employee is exactly one of the two, never both and never neither. "
     "A manager has Level and Bonus; a driver has LicenseNo and Expiry."),
    ("Specialization\n(partial)", "PRODUCT →\nPERISHABLE_PRODUCT",
     "Only some products are perishable. Only those need ShelfLifeDays and "
     "StorageTempC, so the rest are not forced to store empty columns."),
    ("Multivalued attribute", "SUPPLIER.Phone",
     "One supplier can have several phone numbers."),
    ("Composite attribute", "CUSTOMER.Address",
     "Address is made of Street, City and Pincode, which we sometimes need "
     "separately."),
    ("Derived attribute", "ORDERS.TotalAmt",
     "It can always be worked out by adding up the order lines, so storing it "
     "would only create a second copy that can go stale."),
    ("Recursive relationship", "EMPLOYEE\nSUPERVISES EMPLOYEE",
     "An employee reports to another employee, so the relationship joins one "
     "entity to itself."),
]

MAPPING_RULES = [
    "Each entity becomes one table, and its key attribute becomes the primary key.",
    "Each 1:N relationship becomes a foreign key stored on the 'many' side.",
    "Each M:N relationship becomes a new table whose primary key is the two keys "
    "put together.",
    "Each multivalued attribute becomes its own small table.",
    "Each composite attribute becomes one column per part.",
    "A derived attribute is not stored at all; it is calculated when needed.",
    "A weak entity becomes a table keyed by its owner's key plus its own partial key.",
    "A specialization becomes one table per subclass, each sharing the "
    "superclass's key.",
]

SCHEMA = [
    ("CITY", "City", "—"),
    ("SUPPLIER", "SupplierID", "City"),
    ("SUPPLIER_PHONE", "SupplierID + Phone", "SupplierID"),
    ("PRODUCT", "ProductID", "—"),
    ("PERISHABLE_PRODUCT", "ProductID", "ProductID"),
    ("WAREHOUSE", "WarehouseID", "City, ManagerID"),
    ("EMPLOYEE", "EmpID", "WarehouseID, SupervisorID"),
    ("MANAGER", "EmpID", "EmpID"),
    ("DRIVER", "EmpID", "EmpID"),
    ("DRIVER_CITY", "DriverID", "DriverID, City"),
    ("SUPPLIES", "SupplierID + ProductID", "SupplierID, ProductID"),
    ("STOCK", "WarehouseID + ProductID", "WarehouseID, ProductID"),
    ("CUSTOMER", "CustID", "City"),
    ("ORDERS", "OrderID", "CustID, WarehouseID"),
    ("ORDER_ITEM", "OrderID + ItemNo", "OrderID, ProductID"),
    ("SHIPMENT", "ShipmentID", "OrderID, DriverID"),
]

NORM_STEPS = [
    ("norm_unf.png", "7.1  Where we start — Unnormalized Form (UNF)",
     "Everything sits in one wide table, the way a clerk would keep it in a "
     "spreadsheet. One cell holds several phone numbers, and another holds a whole "
     "group of product lines. Because customer details are repeated on every order "
     "row, the same fact is stored many times — and that is what causes the insert, "
     "delete and update problems listed in the figure."),
    ("norm_1nf.png", "7.2  First Normal Form (1NF)",
     "1NF asks for two things: every cell must hold exactly one value, and a "
     "repeating group must move into its own table. So the product lines become "
     "ORDER_ITEM and the phone numbers become SUPPLIER_PHONE. ORDER_ITEM needs "
     "OrderID and ProductID together as its key, because neither one alone "
     "identifies a row."),
    ("norm_2nf.png", "7.3  Second Normal Form (2NF)",
     "2NF removes partial dependency, which means a column must depend on the "
     "whole key and not just a part of it. In ORDER_ITEM the key is "
     "{OrderID, ProductID}, but PName and Category depend on ProductID alone. "
     "They describe the product, not the order, so they move into a new PRODUCT "
     "table. A product name is now written down exactly once."),
    ("norm_3nf.png", "7.4  Third Normal Form (3NF)",
     "3NF removes transitive dependency, which means one ordinary column must not "
     "decide another. In ORDER_MASTER the customer's name is reached only through "
     "CustID, and the state is reached only through the city. Both indirect routes "
     "are pulled out into CUSTOMER and CITY. After this step every one of our 16 "
     "tables is in 3NF."),
    ("norm_bcnf.png", "7.5  Boyce-Codd Normal Form (BCNF)",
     "BCNF applies one strict rule: the left-hand side of every dependency must be "
     "a key. DELIVERY_DUTY passes 3NF only because every column happens to belong "
     "to some key, yet DriverID still decides City while not being a key itself — "
     "so a driver's city is stored twice and the two copies can disagree. "
     "Splitting on that dependency gives DRIVER_CITY and ORDER_DRIVER, and the "
     "join on DriverID rebuilds the original rows exactly, so nothing is lost."),
]


def img(name):
    return os.path.join(DIAG, name)


# =====================================================================  DOCX
def build_docx():
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    for s in d.sections:
        s.top_margin = s.bottom_margin = Inches(0.8)
        s.left_margin = s.right_margin = Inches(0.7)

    def h(text, level):
        p = d.add_heading(text, level)
        for r in p.runs:
            r.font.color.rgb = NAVY
        return p

    def para(text="", bold=False, italic=False, size=11, align=None):
        p = d.add_paragraph()
        if align is not None:
            p.alignment = align
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        return p

    def bullets(items):
        for it in items:
            p = d.add_paragraph(style="List Bullet")
            p.add_run(it).font.size = Pt(10.5)

    def table(headers, rows, widths=None, fs=9.5):
        t = d.add_table(rows=1, cols=len(headers))
        t.style = "Light Grid Accent 1"
        for i, hd in enumerate(headers):
            c = t.rows[0].cells[i]
            c.text = ""
            r = c.paragraphs[0].add_run(hd)
            r.bold = True
            r.font.size = Pt(fs)
        for row in rows:
            cells = t.add_row().cells
            for i, val in enumerate(row):
                cells[i].text = ""
                cells[i].paragraphs[0].add_run(str(val)).font.size = Pt(fs)
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Inches(w)
        return t

    def figure(name, caption, width=7.0):
        d.add_picture(img(name), width=Inches(width))
        d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = para(caption, italic=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        return p

    # ---------------- title page
    para()
    para(TITLE, bold=True, size=26, align=WD_ALIGN_PARAGRAPH.CENTER)
    d.paragraphs[-1].runs[0].font.color.rgb = NAVY
    para("Digital Assignment 1 (DA1)", size=15, italic=True,
         align=WD_ALIGN_PARAGRAPH.CENTER)
    para("ER / EER Model and Normalization up to BCNF", size=13,
         align=WD_ALIGN_PARAGRAPH.CENTER)
    d.paragraphs[-1].runs[0].font.color.rgb = TEAL
    d.paragraphs[-1].runs[0].bold = True
    para()
    para("Database Management Systems", size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    para("Vellore Institute of Technology, Chennai", size=12,
         align=WD_ALIGN_PARAGRAPH.CENTER)
    para()
    para("Submitted by", size=12, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for name, reg in TEAM:
        para(f"{name} — {reg}", size=13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        d.paragraphs[-1].runs[0].font.color.rgb = NAVY
    para()
    para("Submission deadline: 31 July 2026", size=11, italic=True,
         align=WD_ALIGN_PARAGRAPH.CENTER)
    d.add_page_break()

    # ---------------- 1
    h("1. Problem Statement, Objectives and Scope", 1)
    h("1.1 Problem Statement", 2)
    para("A warehouse business has to know three things at all times: what stock it "
         "holds and where, who supplies each product, and where every customer order "
         "has reached. In most small and mid-sized operations these three answers "
         "live in separate places — a purchase register, a stock sheet on the "
         "warehouse floor, and a delivery log kept by the transport team. Nothing "
         "keeps the three in agreement.")
    para("The result is that the same fact ends up written down in several places. "
         "A customer's address appears on every order they have ever placed, and a "
         "product's name appears on every order line that includes it. When one copy "
         "is corrected and another is missed, the data starts contradicting itself, "
         "and the following problems follow directly:")
    bullets([
        "A new supplier or customer cannot be recorded until they appear on an order.",
        "Deleting the last order for a customer erases that customer completely.",
        "Correcting a product name or an address means editing many rows, and any "
        "row that is missed leaves the database in two minds.",
        "Nobody can say reliably how much of a product is on hand, because the "
        "figure is kept in more than one place.",
    ])
    para("These are known as insert, delete and update anomalies. They are not "
         "caused by missing data — they are caused by the same data being stored "
         "more than once. Normalization is the process that removes that "
         "duplication, and this assignment applies it step by step.")

    h("1.2 Objectives", 2)
    bullets([
        "O1. Identify the entities, attributes and relationships the system needs.",
        "O2. Draw an ER model for the design using Chen notation.",
        "O3. Extend that same model into an EER model, adding specialization, a "
        "weak entity, and multivalued, composite and derived attributes.",
        "O4. Convert the EER model into relational tables using the standard "
        "mapping rules.",
        "O5. Normalize every table step by step from UNF through 1NF, 2NF and 3NF "
        "up to BCNF, showing the tables at each stage.",
        "O6. Confirm that no table in the final design can hold a contradiction.",
    ])

    h("1.3 Scope", 2)
    para("In scope: ", bold=True)
    d.paragraphs[-1].add_run(
        "seven entities covering warehouses, products, suppliers, customers, "
        "orders, shipments and employees; the relationships between them; the "
        "conversion of the design into tables; and normalization up to BCNF.")
    para("Out of scope for DA1: ", bold=True)
    d.paragraphs[-1].add_run(
        "writing the SQL and PL/SQL (that is DA2), and rebuilding the project on a "
        "modern database technology (that is DA3). Billing, tax, user logins and "
        "demand forecasting are left out of the design altogether so that the "
        "model stays small enough to explain in full.")

    # ---------------- 2
    h("2. Entities and Attributes", 1)
    para("The design uses seven entities. The key attribute of each one is the "
         "column that identifies a single row, and it is shown underlined in every "
         "diagram in this report.")
    table(["Entity", "Key attribute", "Other attributes"],
          ENTITIES, widths=[1.4, 1.3, 3.9])

    # ---------------- 3
    h("3. Relationships", 1)
    para("Seven relationships connect the entities. '1:N' means one row on the left "
         "can be linked to many rows on the right. 'M:N' means many on both sides.")
    table(["Relationship", "Between", "Type", "Meaning"],
          RELATIONSHIPS, widths=[1.3, 1.6, 0.5, 3.2], fs=9)

    # ---------------- 4
    h("4. ER Model", 1)
    para("The figure below shows the whole design as one ER diagram in Chen "
         "notation: rectangles are entities, diamonds are relationships and "
         "ellipses are attributes. The 1, M and N marks on each line say how many "
         "rows can take part, and a double line means taking part is compulsory — "
         "for example, every order must belong to a customer.")
    figure("er_model.png", "Figure 1 — ER model: 7 entities and 7 relationships")

    # ---------------- 5
    h("5. EER Model", 1)
    para("An EER model is the same ER model with a few extra ideas added, so it can "
         "describe situations plain ER cannot. Figure 2 is the identical design from "
         "Figure 1 with those additions marked in green.")
    figure("eer_model.png", "Figure 2 — the same model, extended with EER features")
    h("5.1 What EER Added and Why", 2)
    table(["EER feature", "Where", "Why it is needed"],
          EER_ADDED, widths=[1.4, 1.5, 3.7], fs=9)
    para("The weak entity is worth singling out. ORDER_ITEM is drawn with a double "
         "border because it cannot exist on its own, and the diamond joining it to "
         "ORDERS is drawn as a double diamond to show that this is the relationship "
         "that gives it its identity. Its partial key, ItemNo, is underlined with a "
         "dashed line because it is only unique within one order.", italic=True)

    # ---------------- 6
    h("6. Converting the EER Model into Tables", 1)
    para("The EER model is converted into relational tables using the standard "
         "mapping rules:")
    bullets(MAPPING_RULES)
    para("Applying these rules produces 16 tables. Those tables are the starting "
         "point for the normalization in the next section, and the finished set is "
         "shown in Figure 8.")

    # ---------------- 7
    h("7. Normalization", 1)
    para("Normalization is done by starting from one badly designed table and "
         "splitting it, one rule at a time, until no table can store the same fact "
         "twice. Each step below shows the tables before and after the split, with "
         "the exact dependency that forced it.")
    para("A functional dependency, written X → Y, means that if you know X then Y "
         "is fixed. For example CustID → CName: once you know the customer's ID, "
         "the name is decided. Every normal form is a rule about which "
         "dependencies are allowed to exist in a table.", italic=True)
    for name, heading, text in NORM_STEPS:
        h(heading, 2)
        para(text)
        figure(name, "Figure " + str(3 + NORM_STEPS.index((name, heading, text)))
               + " — " + heading.split("  ", 1)[1])

    # ---------------- 8
    h("8. Final Schema", 1)
    para("The finished design has 16 tables and every one of them is in BCNF, which "
         "also means every one is in 1NF, 2NF and 3NF.")
    figure("final_schema.png", "Figure 8 — final relational schema, all 16 tables "
                               "in BCNF")
    h("8.1 Keys of Every Table", 2)
    table(["Table", "Primary key", "Foreign keys"], SCHEMA,
          widths=[2.1, 2.2, 2.6], fs=9.5)
    para("Every figure in this report is also supplied as a full-resolution image "
         "in the diagrams folder, and each one fills a whole slide in the "
         "accompanying presentation.", italic=True, size=9.5)
    h("8.2 Why the Design Is Now Safe", 2)
    bullets([
        "Every fact is stored in exactly one place, so two rows can never disagree.",
        "Correcting a product name, a customer address or a city's state is a "
        "single-row edit.",
        "A supplier, product or customer can be added before it appears on any "
        "order.",
        "Deleting an order cannot wipe out the customer who placed it.",
        "TotalAmt is calculated from the order lines rather than stored, so it can "
        "never fall out of step with them.",
    ])
    para("One trade-off is worth recording. Splitting DELIVERY_DUTY into "
         "DRIVER_CITY and ORDER_DRIVER means the rule '{OrderID, City} → DriverID' "
         "can no longer be checked inside a single table. It will be enforced by a "
         "trigger in DA2. This is the normal price of BCNF, and it is worth paying "
         "here because the alternative allows a driver to be recorded against two "
         "different cities at once.")

    # ---------------- 9
    h("9. Conclusion and Next Steps", 1)
    para("DA1 delivered an ER model of seven entities and seven relationships, the "
         "same model extended into an EER model with specialization, a weak entity "
         "and multivalued, composite and derived attributes, and a full "
         "normalization from UNF to BCNF in which every step is shown with the "
         "tables before and after. The result is 16 tables, all in BCNF.")
    para("DA2 (due 18 September 2026): ", bold=True)
    d.paragraphs[-1].add_run(
        "create the 16 tables in SQL with their keys and constraints, load sample "
        "data, add the trigger that enforces the dependency lost at BCNF, and write "
        "PL/SQL procedures for receiving stock and dispatching an order.")
    para("DA3 (due 23 October 2026): ", bold=True)
    d.paragraphs[-1].add_run(
        "rebuild the same project on one modern database technology. The current "
        "plan is a graph database (Neo4j), because tracing a delivery back through "
        "shipment, order, warehouse and supplier is a chain of hops that a graph "
        "query expresses in one line, whereas SQL needs a fresh join for every "
        "extra step.")

    out = os.path.join(ROOT, "docs", "DA1_Report.docx")
    d.save(out)
    return out


# =====================================================================  PPTX
def build_pptx():
    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    BLANK = prs.slide_layouts[6]
    W = prs.slide_width

    def bg(slide, color):
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = color

    def textbox(slide, l, t, w, h, text, size, color=PDARK, bold=False,
                align=PP_ALIGN.LEFT, italic=False):
        tb = slide.shapes.add_textbox(PInches(l), PInches(t), PInches(w), PInches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        f = r.font
        f.size = PPt(size)
        f.bold = bold
        f.italic = italic
        f.color.rgb = color
        f.name = "Calibri"
        return tb

    def header(slide, title, subtitle=None, band=1.15):
        shp = slide.shapes.add_shape(1, 0, 0, W, PInches(band))
        shp.fill.solid()
        shp.fill.fore_color.rgb = PNAVY
        shp.line.fill.background()
        shp.shadow.inherit = False
        textbox(slide, 0.5, 0.20, 12.4, 0.8, title, 27, PWHITE, bold=True)
        if subtitle:
            textbox(slide, 0.5, band + 0.06, 12.4, 0.45, subtitle, 14, PTEAL,
                    italic=True)

    def bullets_slide(title, items, subtitle=None, size=17):
        s = prs.slides.add_slide(BLANK)
        bg(s, PWHITE)
        header(s, title, subtitle)
        top = 1.95 if subtitle else 1.45
        tb = s.shapes.add_textbox(PInches(0.7), PInches(top), PInches(12.0),
                                  PInches(7.2 - top))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            sub = it.startswith("  ")
            r = p.add_run()
            r.text = ("–  " if sub else "•  ") + it.strip()
            r.font.size = PPt(size - 2 if sub else size)
            r.font.color.rgb = PDARK if not sub else PRGB(0x55, 0x55, 0x55)
            r.font.name = "Calibri"
            p.space_after = PPt(9)
            if sub:
                p.level = 1
        return s

    def image_slide(title, image, subtitle=None):
        s = prs.slides.add_slide(BLANK)
        bg(s, PWHITE)
        band = 1.0
        header(s, title, subtitle, band=band)
        top = band + (0.52 if subtitle else 0.14)
        iw, ih = Image.open(image).size
        maxw, maxh = 12.6, 7.34 - top
        scale = min(maxw / iw, maxh / ih)
        dw, dh = iw * scale, ih * scale
        s.shapes.add_picture(image, PInches((13.333 - dw) / 2),
                             PInches(top + (maxh - dh) / 2), PInches(dw), PInches(dh))
        return s

    def table_slide(title, headers, rows, widths, subtitle=None, fs=12):
        s = prs.slides.add_slide(BLANK)
        bg(s, PWHITE)
        header(s, title, subtitle)
        top = 1.95 if subtitle else 1.5
        rowh = 0.46
        est = rowh * (len(rows) + 1)
        top += max(0.0, (7.30 - top - est) / 2)
        total = sum(widths)
        shape = s.shapes.add_table(len(rows) + 1, len(headers),
                                   PInches((13.333 - total) / 2), PInches(top),
                                   PInches(total), PInches(est))
        tbl = shape.table
        for i, w in enumerate(widths):
            tbl.columns[i].width = PInches(w)
        for i, hd in enumerate(headers):
            c = tbl.cell(0, i)
            c.text = hd
            pr = c.text_frame.paragraphs[0]
            pr.runs[0].font.size = PPt(fs)
            pr.runs[0].font.bold = True
            pr.runs[0].font.color.rgb = PWHITE
            c.fill.solid()
            c.fill.fore_color.rgb = PNAVY
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
        for ri, row in enumerate(rows, start=1):
            for ci, val in enumerate(row):
                c = tbl.cell(ri, ci)
                c.text = str(val)
                pr = c.text_frame.paragraphs[0]
                pr.runs[0].font.size = PPt(fs - 1)
                pr.runs[0].font.color.rgb = PDARK
                c.fill.solid()
                c.fill.fore_color.rgb = PRGB(0xF2, 0xF6, 0xFB) if ri % 2 else PWHITE
                c.vertical_anchor = MSO_ANCHOR.MIDDLE
        return s

    # ---------------- title
    s = prs.slides.add_slide(BLANK)
    bg(s, PNAVY)
    textbox(s, 0.8, 1.9, 11.7, 1.7, TITLE, 40, PWHITE, bold=True,
            align=PP_ALIGN.CENTER)
    textbox(s, 0.8, 3.65, 11.7, 0.6, "Digital Assignment 1  •  Database Management "
            "Systems", 20, PMIST, align=PP_ALIGN.CENTER)
    textbox(s, 0.8, 4.35, 11.7, 0.6, "ER / EER Model and Normalization up to BCNF",
            19, PTEAL, bold=True, align=PP_ALIGN.CENTER)
    textbox(s, 0.8, 5.75, 11.7, 0.9,
            "   •   ".join(f"{n} ({r})" for n, r in TEAM), 15, PWHITE, bold=True,
            align=PP_ALIGN.CENTER)

    # ---------------- agenda
    bullets_slide("Agenda", [
        "The problem, objectives and scope",
        "Entities, attributes and relationships",
        "The ER model",
        "The EER model — the same design, extended",
        "Turning the model into tables",
        "Normalization: UNF → 1NF → 2NF → 3NF → BCNF",
        "The final schema, and what comes next",
    ])

    # ---------------- problem
    bullets_slide("1. The Problem", [
        "A warehouse must always know: what stock is where, who supplies it, and "
        "where each order has reached.",
        "In practice those three answers live in three separate places, and "
        "nothing keeps them in agreement.",
        "So the same fact gets stored twice — a customer's address on every order, "
        "a product's name on every order line.",
        "Correct one copy, miss another, and the data contradicts itself.",
    ], subtitle="The trouble is not missing data — it is the same data stored more "
                "than once")

    bullets_slide("1. What That Duplication Costs", [
        "Insert problem — a new customer cannot be recorded until they place an "
        "order.",
        "Delete problem — deleting a customer's last order erases the customer.",
        "Update problem — renaming a product means editing many rows; miss one and "
        "the two disagree.",
        "Unreliable stock — the quantity on hand is kept in more than one place.",
        "These are the insert, delete and update anomalies. Normalization removes "
        "them.",
    ])

    bullets_slide("1. Objectives and Scope", [
        "Identify the entities, attributes and relationships.",
        "Draw the ER model, then extend the same model into an EER model.",
        "Convert the model into relational tables.",
        "Normalize every table from UNF up to BCNF, showing each step.",
        "In scope: 7 entities — warehouse, product, supplier, customer, orders, "
        "shipment, employee.",
        "Out of scope: SQL and PL/SQL (DA2), and a modern database (DA3).",
    ])

    # ---------------- entities & relationships
    table_slide("2. The Seven Entities",
                ["Entity", "Key attribute", "Other attributes"],
                ENTITIES, [2.4, 2.5, 6.6],
                subtitle="the key attribute identifies one row, and is underlined "
                         "in every diagram", fs=13)

    table_slide("3. The Seven Relationships",
                ["Relationship", "Between", "Type"],
                [(r[0], r[1], r[2]) for r in RELATIONSHIPS], [3.4, 4.4, 1.6],
                subtitle="1:N means one row links to many; M:N means many on both "
                         "sides", fs=14)

    # ---------------- ER / EER
    image_slide("4. The ER Model", img("er_model.png"))
    image_slide("5. The EER Model — the Same Design, Extended",
                img("eer_model.png"))

    table_slide("5. What EER Added, and Why",
                ["EER feature", "Where"],
                [(a.replace("\n", " "), b.replace("\n", " "))
                 for a, b, _ in EER_ADDED], [4.6, 5.6],
                subtitle="each addition describes something plain ER cannot",
                fs=13)

    bullets_slide("5. The Weak Entity, in One Slide", [
        "ORDER_ITEM cannot exist on its own — an order line means nothing without "
        "its order.",
        "ItemNo (1, 2, 3 …) is only unique inside one order, so it is a partial key "
        "— shown with a dashed underline.",
        "Its real key is therefore OrderID + ItemNo.",
        "Drawn with a double border; the diamond joining it to ORDERS is a double "
        "diamond, meaning that relationship gives it its identity.",
        "It replaces the M:N CONTAINS relationship from the ER model.",
    ], subtitle="why ORDER_ITEM is drawn differently from every other box")

    # ---------------- mapping
    bullets_slide("6. Turning the Model into Tables", MAPPING_RULES,
                  subtitle="the standard mapping rules — they produce 16 tables",
                  size=16)

    # ---------------- normalization
    bullets_slide("7. How Normalization Works", [
        "A functional dependency X → Y means: if you know X, then Y is fixed.",
        "  Example — CustID → CName: once you know the customer ID, the name is "
        "decided.",
        "Each normal form is a rule about which dependencies a table may contain.",
        "1NF — every cell holds one value.",
        "2NF — no column depends on only part of the key.",
        "3NF — no ordinary column decides another ordinary column.",
        "BCNF — the left side of every dependency must be a key.",
    ], subtitle="start with one bad table, then split it one rule at a time")

    for name, heading, _ in NORM_STEPS:
        num, rest = heading.split("  ", 1)
        image_slide("7. " + rest, img(name))

    # ---------------- final
    image_slide("8. Final Schema — 16 Tables, All in BCNF",
                img("final_schema.png"))

    bullets_slide("8. Why the Design Is Now Safe", [
        "Every fact is stored in exactly one place, so two rows can never disagree.",
        "Correcting a product name, an address or a city's state is a single-row "
        "edit.",
        "A supplier, product or customer can exist before appearing on any order.",
        "Deleting an order cannot wipe out the customer who placed it.",
        "TotalAmt is calculated, not stored, so it can never fall out of step.",
        "One trade-off: the rule {OrderID, City} → DriverID now needs a trigger, "
        "which DA2 will add.",
    ])

    bullets_slide("9. What Comes Next", [
        "DA2 — 18 September 2026",
        "  Create the 16 tables in SQL with keys and constraints, load sample data, "
        "add the BCNF trigger, and write PL/SQL for receiving stock and dispatching "
        "orders.",
        "DA3 — 23 October 2026",
        "  Rebuild on a graph database (Neo4j). Tracing a delivery back through "
        "shipment, order, warehouse and supplier is a chain of hops — one line in a "
        "graph query, but a fresh join for every step in SQL.",
    ], size=17)

    # ---------------- closing
    s = prs.slides.add_slide(BLANK)
    bg(s, PNAVY)
    textbox(s, 0.8, 2.7, 11.7, 1.2, "Thank You", 44, PWHITE, bold=True,
            align=PP_ALIGN.CENTER)
    textbox(s, 0.8, 4.0, 11.7, 0.7, "7 entities  ·  7 relationships  ·  16 tables  "
            "·  all in BCNF", 19, PTEAL, italic=True, align=PP_ALIGN.CENTER)
    textbox(s, 0.8, 5.6, 11.7, 0.7,
            "   •   ".join(f"{n} ({r})" for n, r in TEAM), 14, PMIST,
            align=PP_ALIGN.CENTER)

    out = os.path.join(ROOT, "presentation", "DA1_Presentation.pptx")
    prs.save(out)
    return out, len(prs.slides.__iter__.__self__._sldIdLst)


if __name__ == "__main__":
    for f in ["er_model", "eer_model", "norm_unf", "norm_1nf", "norm_2nf",
              "norm_3nf", "norm_bcnf", "final_schema"]:
        assert os.path.exists(img(f + ".png")), f"missing {f}.png — run make_diagrams.py"
    print("DOCX:", build_docx())
    path, n = build_pptx()
    print("PPTX:", path, f"({n} slides)")
