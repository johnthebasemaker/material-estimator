# Smart Material Estimator — Architecture & Rules

## 🏗️ Project Overview
An enterprise-grade Streamlit application for tracking, estimating, and allocating lining materials for construction equipment. It uses a cascading priority allocation engine and a SQLite database.

## 💾 Database Architecture (`sme_database.db`)
The application was migrated from static Excel files to SQLite. 
* **`inventory`**: Master list of materials, Available Qty, and Ordered Qty.
* **`recipe`**: How much of each material is needed for 1 SQM of a specific `lining_system_code`.
* **`equipment`**: Master list of equipment tags, locations, and their `surface_area_sqm`.
* **`sqm_progress`**: Tracks completed SQM. `original_sqm` minus `done_sqm` = `remaining_sqm`.
* **`consumption_log`**: History of daily work entered by users.
* **`receipt_log`**: History of new material inventory received.

## ⚙️ Core Mechanics
* **Data Loader:** `app.py` reads from SQLite via a cached function: `@st.cache_data(show_spinner="Loading project data…") def load_all():`.
* **Allocation Engine:** `cascade_allocate()` loops through selected equipment and dynamically deducts material requirements from a global inventory pool. It is cached to prevent UI lag.
* **State Updates:** Any time an `INSERT` or `UPDATE` is executed against the database, the code MUST call `conn.commit()` followed immediately by `st.cache_data.clear()` and `st.rerun()` so the Streamlit UI reflects the live database changes.

## 🎨 UI & Styling Rules (STRICT)
* **Adaptive Theme:** The app perfectly supports both Light and Dark mode. 
* **NO HARDCODED COLORS:** Never use hex codes (e.g., `#FFFFFF` or `#000000`) for text. Always use Streamlit's native CSS variables injected in the style block: `var(--t0)` to `var(--t5)` for text, and `var(--bg0)` to `var(--bg4)` for backgrounds.
* **Tables:** Always use Streamlit's native `st.dataframe` with Pandas `.style` for background colors. Do not use Plotly `go.Table` as it breaks in Dark Mode.