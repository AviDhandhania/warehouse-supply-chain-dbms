# Warehouse and Supply Chain Management System — DBMS DA1

ER / EER model, functional dependency analysis, and normalization up to BCNF.

**Team:** Safwaan Mohamed S (25BCE1188) · Avi Dhandhania (25BCE1207) · Priyal Bhalla (25BCE1261)

## Deliverables

| File | What it is |
|---|---|
| `docs/DA1_Report.docx` | The report — 11 sections, 10 figures |
| `presentation/DA1_Presentation.pptx` | The deck — 33 slides, 16:9 |
| `diagrams/*.png` | The 10 figures at full resolution |

## Rebuilding

```
cd diagrams && python make_diagrams.py     # writes the 10 PNGs
cd ..       && python build_deliverables.py # writes the docx and pptx
```

Needs `matplotlib`, `python-docx`, `python-pptx`, `pillow`. Close the docx and pptx
in Office first, or the build hits a permission error.

## The design

7 entities — WAREHOUSE, PRODUCT, SUPPLIER, CUSTOMER, ORDERS, SHIPMENT, EMPLOYEE —
and 7 relationships between them. Mapping the EER model to tables and then
normalizing gives 17 tables, all of which are in BCNF.

Section 7 of the report fills each relation with sample rows (7 in ORDER_LINE, 7 in
DELIVERY_DUTY, 5 in SUPPLIER_PHONE) and reads the functional dependencies off them:
8 that hold (F1–F8) and 8 that the data rules out (N1–N8). Every split in section 8
is then justified by one of those refs.

## The figures

1. `er_model` — the whole design as one ER diagram (Chen notation)
2. `eer_model` — the same design with the EER features added, drawn shaded
3. `sample_data` — the rows the dependency analysis is read from
4. `norm_unf` — the one unnormalized table everything starts from
5. `norm_1nf` — the relations after 1NF
6. `norm_2nf` — after 2NF
7. `norm_3nf` — after 3NF
8. `norm_bcnf` — after BCNF
9. `decomp_tree` — UNF → BCNF as a tree, every column at every stage
10. `final_schema` — all 17 tables with their keys

Tables only inside the normalization figures; the reasoning lives in the report
text, so the two never drift out of agreement.

## Still to fill in

The report cover page has the course name but no course code, faculty name or slot.
