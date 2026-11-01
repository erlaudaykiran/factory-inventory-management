# Factory Inventory Management Portal

A centralized, multi-user web-based inventory management portal built for environmental chamber manufacturing, designed to replace and synchronize individual Excel workbooks.

## Key Features

1. **Role-Based Access for 4 Personnel:**
   - **1 Administrator (`admin`):** Full privileges to record Stock IN (Receipts / GRN) and Stock OUT (Issues to Job/Chamber), manage items, view logs, and export reports.
   - **3 Viewers (`viewer1`, `viewer2`, `viewer3`):** Read-only live access to search, filter, monitor closing stock levels, inspect reorder alerts, and export data.
2. **Centralized Real-Time Synchronized Database:**
   - Any transaction recorded by the Administrator immediately updates the central database (SQLite).
   - All 4 users see the exact same real-time stock levels, closing stock calculations, and reorder warnings.
3. **Full Excel Catalog Import:**
   - Automatically imports all **284 items** across all 7 manufacturing categories from `Inventory management.xlsx`:
     1. `01 General Fabrication`
     2. `02 Refrigeration`
     3. `03 Heating`
     4. `04 De-humidity`
     5. `05 Altitude`
     6. `06 Electrical & Control`
     7. `07 Miscellaneous`
4. **Automated Stock Calculations & Validations:**
   - `Closing Stock = Opening Stock + Total Received (IN) - Total Issued (OUT)`
   - Immediate **REORDER** flag when `Closing Stock <= Minimum Level`.
   - Prevents issuing more quantity than available in stock.
5. **Auditing & Exports:**
   - Complete stock movement audit log with Date, Doc/GRN No., Item Code, Quantity, Job/Chamber No., Issued To, and User remarks.
   - Export to Excel (`.xlsx`) and CSV with one click.

---

## Default Access Credentials

| User ID | Role | Default Password | Permissions |
| :--- | :--- | :--- | :--- |
| **`admin`** | Administrator | `admin123` | Full Access (Record IN/OUT, manage stock, export) |
| **`viewer1`** | Viewer | `view123` | View-Only (Live dashboard, stock log, search, export) |
| **`viewer2`** | Viewer | `view123` | View-Only (Live dashboard, stock log, search, export) |
| **`viewer3`** | Viewer | `view123` | View-Only (Live dashboard, stock log, search, export) |

> **Note:** On Render or any production host, you can customize these usernames and passwords anytime using environment variables:
> - `ADMIN_ID`, `ADMIN_PASSWORD`
> - `VIEWER2_ID`, `VIEWER2_PASSWORD`
> - `VIEWER3_ID`, `VIEWER3_PASSWORD`
> - `VIEWER4_ID`, `VIEWER4_PASSWORD`

---

## Running Locally

1. Install requirements:
   ```bash
   pip install -r requirements.txt
   ```
2. Start the application:
   ```bash
   python app.py
   ```
3. Open your browser and navigate to:
   ```
   http://localhost:5000
   ```

---

## Deploying to Render

1. Push this repository to GitHub: `https://github.com/erlaudaykiran/factory-inventory-management`
2. Go to your Render Dashboard (`https://dashboard.render.com`).
3. Click **New +** > **Web Service**.
4. Select `factory-inventory-management` from your connected GitHub account.
5. Render will automatically detect `render.yaml` or you can configure:
   - **Environment:** Python
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Disk (Optional for persistence):** Mount `/var/data` (1 GB) with `DATABASE_PATH=/var/data/inventory.db`
6. Click **Deploy Web Service**.
