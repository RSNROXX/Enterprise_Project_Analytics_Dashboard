# Enterprise Performance Management & Analytics Dashboard

A centralized data intelligence platform designed to automate the tracking, calculation, and visualization of Key Performance Indicators (KPIs) across Sales, Design, and Operations departments.

This system replaces fragmented spreadsheet workflows with a robust Django-based architecture, serving as the organization's single source of truth for employee and project performance.

## 📈 Business Impact & ROI

This tool was developed to address critical data fragmentation issues. Since its deployment, it has delivered measurable operational improvements:

* **~75% Process Automation:** Successfully automated approximately ~75% of the manual data aggregation work previously required by the MIS team, reducing weekly reporting time from hours to seconds.
* **High-Velocity Value Delivery:** The system was designed with an agile feedback loop; the core dashboard met business requirements and achieved stakeholder sign-off by the **2nd commit**.
* **High Leadership Adoption:** Currently utilized **twice weekly** by Department Heads and Managers to monitor project health and resource allocation.
* **Standardized Scoring:** Replaced subjective manual grading with a deterministic, weighted scoring engine (`pandas` based) that ensures fair performance evaluation across thousands of projects.

## ✨ Key Features

### 1. 📊 Advanced Data Engineering
* **Multi-Sheet Excel Ingestion:** A robust ETL pipeline using `pandas` and `openpyxl` to parse complex, multi-sheet Excel files (Sales, Design, Operations).
* **Intelligent Data Cleaning:** Automated normalization of inconsistent column headers, date formats, and empty values (`NaN`) to ensure database integrity.
* **Weighted Scoring Engine:** A custom logic layer that calculates a **0-100 Project Health Score** based on dynamic, stage-specific metrics (Pre-Sales vs. Execution).

### 2. 🎨 User Experience & Interface
* **Adaptive "Soft Dark" UI:** A custom-engineered theme using CSS variables (`--card-bg`, `--text-main`) designed specifically to reduce eye strain for managers reviewing data on large monitors.
* **Role-Based Filtering:** Context-aware dashboards that dynamically adjust views for specific roles (e.g., Sales Leads vs. Project Managers).
* **Gamified Leaderboards:** A "Hall of Fame" module that aggregates individual credits to rank top performers by department.

### 3. 📝 Reporting & Exports
* **Deep-Dive Analytics:** Detailed report views with **sticky headers** and cross-linking to individual Project Scorecards for audit trails.
* **Automated Exporting:** One-click generation of formatted Excel reports for offline analysis.
* **Live Web Tables:** Renders dataframes directly to HTML tables for quick reviews without downloading.

## 🛠️ Tech Stack

* **Backend:** Python 3.10+, Django 5.0
* **Data Processing:** Pandas, NumPy, OpenPyXL
* **Frontend:** HTML5, Bootstrap 5, Vanilla JavaScript, CSS Variables
* **Database:** SQLite (Dev) / PostgreSQL (Production)

## ⚙️ Local Installation

Follow these steps to run the project locally on your machine.

**1. Clone the repository**
```bash
git clone [https://github.com/YOUR-USERNAME/Enterprise-Project-Dashboard.git](https://github.com/YOUR-USERNAME/Enterprise-Project-Dashboard.git)
cd Enterprise-Project-Dashboard
```

**2. Create a Virtual Environment**
```bash
# Windows
python -m venv venv
source venv/Scripts/activate

# Mac/Linux
python3 -m venv venv
source venv/bin/activate
```

**3. Install Dependencies**
```bash
pip install -r requirements.txt
```

**4. Apply Database Migrations**
```bash
python manage.py makemigrations
python manage.py migrate
```

**5. Create a Superuser (Admin)**
```bash
python manage.py createsuperuser
```

**6. Run the Server**
```bash
python manage.py runserver
```
Access the dashboard at `http://127.0.0.1:8000`.

## 🔒 Privacy & Security Note

This repository contains the **source code logic only**.
* **No Data:** All database files (`db.sqlite3`) and raw Excel datasets are excluded via `.gitignore` to ensure client data privacy.
* **No Secrets:** Environment variables (Secret Keys, Debug Mode) should be configured locally in a `.env` file for production use.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

Author: [Rishikesh Naware](https://www.linkedin.com/in/rishikesh-naware/)

