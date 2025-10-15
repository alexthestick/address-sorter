## Address Sorter

Automates sorting of addresses into categories with ROE deduplication and anomaly detection.

### Quick Start (macOS)

Option A: Web UI (recommended)

- Install Python 3.10+ from `https://www.python.org/downloads/` if needed
- Double-click `Address Sorter Web.command` to launch the browser UI
- Upload your CSV/XLSX, click "Process Addresses", preview sheets, then click "Download Sorted Excel"

Option B: Classic (file picker)

- Double-click `Address Sorter.command`
- Pick your input file (CSV or XLSX) when prompted
- Choose where to save the output Excel file

Alternatively, via Terminal:
```bash
cd "/Users/alexcoluna/Desktop/Project Folder/Address Sorter"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# Web UI
streamlit run app.py
# Classic CLI
python3 address_sorter.py path/to/input.xlsx  # or .csv
```

### Input Requirements
- Required columns: `ID`, `Street Address`, `Unit Number`, `Building Type`, `Subname`
- Optional columns (used if present): `City`, `Zip`, `Plus 4 Code`, `Zone`, `Street Name`

### Output
An Excel workbook with sheets: `All`, `Public`, `Commercial`, `ROE`, `Competitive`, `Other`, `Remove`, `Flagged for Review`, `Unit Count`.

### Notes
- If launched without CLI arguments, a GUI file picker will open (Tkinter).
- Output is saved as `<input_basename>_sorted.xlsx` if no output filename is chosen.

### Deploy to Streamlit Community Cloud
1. Push these files to a public GitHub repo: `app.py`, `address_sorter.py`, `requirements.txt`, `runtime.txt`, optional `.streamlit/config.toml`.
2. Go to the Streamlit Community Cloud dashboard and click "New app".
3. Select your repo/branch and set the entry point to `app.py`.
4. Deploy. You’ll get a shareable URL.

Tips:
- The included `runtime.txt` pins Python 3.11. Adjust if needed.
- `.streamlit/config.toml` raises `maxUploadSize` (MB). If uploads still fail, reduce input size or optimize locally.
