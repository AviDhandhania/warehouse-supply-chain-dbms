# Warehouse and Supply Chain Management System — DBMS DA1

ER / EER model and normalization up to BCNF.

**Team:** Safwaan Mohamed S (25BCE1188) · Avi Dhandhania (25BCE1207) · Priyal Bhalla (25BCE1261)

## Deliverables

| File | What it is |
|---|---|
| `docs/DA1_Report.docx` | The report — 10 pages, 8 figures |
| `presentation/DA1_Presentation.pptx` | The deck — 22 slides, 16:9 |
| `diagrams/*.png` | The 8 figures at full resolution |

## Rebuilding

```
cd diagrams && python make_diagrams.py     # writes the 8 PNGs
cd ..       && python build_deliverables.py # writes the docx and pptx
```

Needs `matplotlib`, `python-docx`, `python-pptx`, `pillow`.

## The design

7 entities — WAREHOUSE, PRODUCT, SUPPLIER, CUSTOMER, ORDERS, SHIPMENT, EMPLOYEE —
and 7 relationships between them. Mapping the EER model to tables gives 16 tables,
all of which are in BCNF.

## The figures

1. `er_model` — the whole design as one ER diagram (Chen notation)
2. `eer_model` — the same design with the EER features added, marked in green
3. `norm_unf` — the one unnormalized table everything starts from
4. `norm_1nf` — after 1NF
5. `norm_2nf` — after 2NF
6. `norm_3nf` — after 3NF
7. `norm_bcnf` — the BCNF case and its fix
8. `final_schema` — all 16 tables with their keys

## Still to fill in

The report cover page has the course name but no course code, faculty name or slot.
