# Product CSV Comparator

A local Flask web application that compares two product CSV files by SKU and generates an Excel report plus a TXT report displayed in the browser.

## Run on Windows

1. Install Python 3.10+.
2. Open Command Prompt/PowerShell in this folder.
3. Create a virtual environment (recommended):

```bash
python -m venv venv
```

4. Activate it:

```powershell
venv\Scripts\activate
```

5. Install dependencies:

```bash
pip install -r requirements.txt
```

6. Start the application:

```bash
python app.py
```

7. Open:

http://127.0.0.1:5000

## Supported SKU columns

The app detects any of these SKU column names in both files:

- `sku`
- `SKU`
- `#SKU`

## Output

The generated Excel report contains:

- Summary
- New SKUs
- Removed SKUs
- Column Changes
- Detailed Changes
- Duplicate SKUs
- Common SKU Summary

The TXT report is shown on the webpage and can also be downloaded.
