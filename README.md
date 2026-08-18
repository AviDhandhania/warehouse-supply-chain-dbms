# Warehouse and Supply Chain Management System — DBMS DA1

ER / EER model, functional dependency analysis, and normalization up to BCNF.

**Team:** Safwaan Mohamed S (25BCE1188) · Avi Dhandhania (25BCE1207) · Priyal Bhalla (25BCE1261)

## Deliverables

| File | What it is |
|---|---|
| `docs/DA1_Report.docx` | The report — 11 sections, 11 figures |
| `presentation/DA1_Presentation.pptx` | The deck — 36 slides, 16:9 |
| `diagrams/*.png` | The 11 figures at full resolution |

## Rebuilding

```
cd diagrams && python make_diagrams.py     # writes the 11 PNGs
cd ..       && python build_deliverables.py # writes the docx and pptx
```

Needs `matplotlib`, `python-docx`, `python-pptx`, `pillow`. Close the docx and pptx
in Office first, or the build hits a permission error.

## The design

9 entities — WAREHOUSE, PRODUCT, SUPPLIER, CUSTOMER, ORDERS, SHIPMENT,
GOODS_RETURN, INSPECTION, EMPLOYEE — and 12 relationships between them, one of
which (MANAGES) is 1:1. Mapping the EER model to tables and then normalizing
gives 21 tables, all of which are in BCNF: 8 out of the normalization and 13
straight from the mapping.

The EER model carries every construct: specialization (EMPLOYEE to
MANAGER/DRIVER, total and disjoint; PRODUCT to PERISHABLE_PRODUCT and
HAZARDOUS_PRODUCT, partial and overlapping), generalization (SHIPMENT +
GOODS_RETURN up into GOODS_MOVEMENT), aggregation (INSPECTION against the boxed
SUPPLIER-SUPPLIES-PRODUCT arrangement), the weak entity ORDER_ITEM, and the
multivalued, composite and derived attributes. Specialization and
generalization draw the same and map the same; the difference is that one
reasons downwards and the other upwards.

Each of the three hierarchies is marked three ways, and the report explains all
three: the letter in the circle (`d` disjoint for EMPLOYEE and GOODS_MOVEMENT,
`o` overlapping for PRODUCT, since a vaccine is both perishable and hazardous),
the line into the circle (double for total, single for partial), and the subset
symbol on each subclass line, pointing at the superclass.

Section 7 of the report fills each relation with sample rows (7 in ORDER_LINE, 7 in
DELIVERY_DUTY, 5 in SUPPLIER_PHONE) and reads the functional dependencies off them:
8 that hold (F1–F8) and 8 that the data rules out (N1–N8). Every split in section 8
is then justified by one of those refs.

The normalization chain does not touch any of the new material, which is why
`sample_data`, the five `norm_*` figures and `decomp_tree` are unchanged: the
lineage runs through ORDERS, ORDER_ITEM, PRODUCT, CUSTOMER, CITY,
SUPPLIER_PHONE, DRIVER_CITY and ORDER_DRIVER, and the 1:1, the generalization
and the aggregation all land in the other thirteen tables.

## The figures

1. `er_model` — the whole design as one ER diagram (Chen notation)
2. `eer_model` — the same design with the EER features added, drawn shaded
3. `eer_aggregation` — the aggregation alone, so the box round SUPPLIES stays legible
4. `sample_data` — the rows the dependency analysis is read from
5. `norm_unf` — the one unnormalized table everything starts from
6. `norm_1nf` — the relations after 1NF
7. `norm_2nf` — after 2NF
8. `norm_3nf` — after 3NF
9. `norm_bcnf` — after BCNF
10. `decomp_tree` — UNF → BCNF as a tree, every column at every stage
11. `final_schema` — all 21 tables with their keys

Every PNG holds the diagram and nothing else: no titles, captions, legends or
commentary. All of the reasoning, including the notation legends, lives in the
report text, so the two can never drift out of agreement.

## Still to fill in

The report cover page has the course name but no course code, faculty name or slot.
