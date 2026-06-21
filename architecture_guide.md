# Smart Material Estimator & Planner — Architecture & Rules

## 🏗️ Project Overview
An enterprise-grade Streamlit application for tracking, estimating, and allocating lining materials for construction equipment. It acts as an ERP system managing Orders, Receipts, Consumption drafts, and global allocations via a SQLite database.

## 💾 Database Architecture (`sme_database.db`)
* **`inventory`**: Master list of materials, Available Qty, and Ordered Qty.
* **`recipe`**: How much of each material is needed for 1 SQM of a specific `lining_system_code`.
* **`equipment`**: Master list of equipment tags, locations, and their `surface_area_sqm`.
* **`sqm_progress`**: Tracks completed SQM. `original_sqm` minus `done_sqm` = `remaining_sqm`.
* **`consumption_log`**: History of approved daily work.
* **`draft_consumption`**: Temporary shopping-cart table tied to a user session UUID to survive browser reloads.
* **`receipt_log`**: History of new material inventory received, linked to `order_id`.
* **`orders_log`**: Tracks generated Purchase Orders, material pending quantities, and fulfillment status.

## ⚙️ Core Mechanics
* **Data Loader:** `app.py` reads from SQLite via a cached function.
* **Allocation Engine:** `cascade_allocate()` loops through equipment and dynamically deducts material requirements from a global inventory pool. 
* **State Updates:** Any time an `INSERT`, `UPDATE`, or `DELETE` is executed against the database, the code MUST execute this exact sequence:
  1. `conn.commit()`
  2. `st.cache_data.clear()`
  3. `st.rerun()` 
  This forces the Streamlit UI to instantly reflect live database changes.

## 🎨 UI & Styling Rules (STRICT)
* **Adaptive Theme:** The app perfectly supports both Light and Dark mode. 
* **NO HARDCODED COLORS:** Never use hex codes (e.g., `#FFFFFF` or `#000000`) for text. Always use Streamlit's native CSS variables injected in the style block: `var(--t0)` to `var(--t5)` for text, and `var(--bg0)` to `var(--bg4)` for backgrounds.
* **NO BLIND OVERALL QUANTITIES:** Never display aggregate metric cards that sum raw material quantities (e.g., mixing KG, LTR, and NOS) across different materials. Summing mixed physical units is misleading. Only aggregate areas (SQM), completion percentages, or order counts. Never sum raw physical units globally.
* **Area vs. Material Units:** Be extremely careful to distinguish between **Area Metrics** (which must explicitly be labeled with "SQM") and **Material Quantities** (which are raw units like KG/NOS and should NOT be labeled as SQM).
* **Tables:** Always use Streamlit's native `st.dataframe` or `st.data_editor` configured with `use_container_width=True`. Hide index columns unless strictly necessary.