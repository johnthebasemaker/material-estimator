# 🏗 Smart Material Estimator

A Streamlit desktop web application for priority-based material allocation and
lining project feasibility planning.

---

## Prerequisites

- Python 3.10+
- The three source Excel files (File A, B, C)

---

## Setup

```bash
# 1. Clone / place all four files in the same folder:
#    app.py · validate_data.py · allocation_engine.py · requirements.txt

# 2. Create a virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Launch the app
streamlit run app.py
```

The browser will open automatically at http://localhost:8501

---

## File Structure

```
smart-material-estimator/
├── app.py                  ← Streamlit UI (this file)
├── validate_data.py        ← Data loading, cleaning & validation
├── allocation_engine.py    ← Allocation & suggestion engine
├── requirements.txt
└── README.md
```

---

## How to Use

### Step 1 — Upload your data
In the left sidebar, upload:
- **File A** — `Materials_DetailsAvailable_Qty.xlsx`  (Inventory)
- **File B** — `For_1_SQM.xlsx`  (Lining system recipes)
- **File C** — `Equipment.xlsx`  (Equipment list with surface areas)

Click **Load / Refresh Data** — the app validates all joins and runs the full pipeline.

### Step 2 — Set build priority (Tab 1)
Drag and drop equipment tags to define your build order.
Top = highest priority → gets materials first.
Click **Apply & Recalculate**.

### Step 3 — Review feasibility (Tab 2)
The colour-coded table shows:
- ✅ **100% Fully Ready to Build** — all materials fully covered
- 🟡 **Partially Ready (X%)** — some materials short
- 🔴 **Blocked by Shortages** — at least one material at zero

Use the equipment inspector to see per-material fulfillment rates.
Read the Predictive Suggestion Engine panel for the best pause scenario.

### Step 4 — Procure shortages (Tab 3)
Download the **Procurement Shopping List** as Excel.
The stacked bar chart shows Available vs Shortage per material.

---

## Key Logic Notes

| Concept | Implementation |
|---|---|
| Demand per equipment | `For_1_SQM × Surface_Area_SQM`, aggregated across all lining systems |
| Sequential allocation | Materials distributed top-priority-first; no material is split between equipment unless there is surplus |
| Duplicate PO rows (File A) | Summed by `Material_Code` |
| Comma-separated material codes (File B) | Exploded to individual rows |
| Missing inventory records | Treated as `Available_Qty = 0`, appear on shopping list |

---

## Dependencies

| Package | Purpose |
|---|---|
| `streamlit` | UI framework |
| `pandas` | Data engine |
| `openpyxl` | Excel read/write |
| `plotly` | Interactive charts |
| `streamlit-sortables` | Drag-and-drop priority list (graceful fallback if absent) |