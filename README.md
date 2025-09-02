# Excel Comparison Tool

This project provides a web-based and API-driven tool for comparing two Excel files. It highlights differences between sheets and generates a downloadable comparison report.

## Features

- Compare two Excel files (.xlsx) via a web interface or HTTP API
- Highlights differences and similarities in rows and cells
- Supports custom sheet configurations
- Returns a downloadable Excel report with summary and details

## Requirements

- Python 3.8+
- See [requirements.txt](requirements.txt) for dependencies:
  - fastapi
  - uvicorn
  - pandas
  - openpyxl
  - python-multipart

## Usage

### 1. Install dependencies

```sh
pip install -r requirements.txt
```

### 2. Run the API server

```sh
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

### 3. Open the Web Interface

Open [index.html](index.html) in your browser.  
Upload two Excel files and click "Compare Excels" to download the result.

### 4. API Endpoints

- **GET /health**  
  Health check endpoint.

- **POST /compare**  
  Upload two Excel files and (optionally) a JSON `sheets_config`.  
  Returns a comparison Excel file.

#### Example cURL

```sh
curl -X POST "http://localhost:8000/compare" \
-F "original_file=@Tax_Report_hemal_patel_28052025 - 4.0.xlsx" \
-F "website_file=@Tax_Report_hemal_patel_21082025.xlsx" \
-F 'sheets_config={"Gain Summary":{"header_row":2,"data_start_row":3},"8938":{"header_row":6,"data_start_row":7},"FBAR":{"header_row":2,"data_start_row":3}}'
```

## Project Structure

- [main.py](main.py): FastAPI backend and Excel comparison logic
- [index.html](index.html): Web frontend
- [requirements.txt](requirements.txt): Python dependencies
- `uploads/`: Directory for uploaded files

## Notes

- If `sheets_config` is omitted, sensible defaults are used.
- The output Excel file contains a summary and detailed comparison sheets.

# Excel Comparison API

**Excel Comparison API (FastAPI)**  
Exposes your existing Excel comparison logic as an HTTP API.

---

## **Features**

- Compare two Excel files (`.xlsx`) in-memory.
- Highlights differences in numeric and text values.
- Supports multiple sheets with customizable header and data start rows.
- Generates a summary sheet with common rows, differences, and totals.
- Returns a downloadable Excel file with differences highlighted.

---

## **Endpoints**

| Method | URL        | Description |
|--------|-----------|-------------|
| GET    | `/health` | Simple health check |
| POST   | `/compare` | Upload two Excel files (and optional JSON `sheets_config`) to get compared `.xlsx` |

---

## **Run Locally**

1. Clone the repository or copy the project folder:

```bash
git clone <repo-url>
cd excel_task_copy
