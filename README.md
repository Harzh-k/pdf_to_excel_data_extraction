# IRDAI PDF → Excel Extractor

## Project Structure

```
your_project/
├── app.py                          ← Streamlit UI
├── extract_tables_smart_merged.py  ← CLI version
├── requirements.txt
├── README.md
└── src/
    ├── __init__.py                 ← empty file (must exist)
    └── extractor.py                ← extraction engine
```

---

## Step-by-Step Setup & Run Guide

### Step 1 — Create project folder

```bash
mkdir irdai_extractor
cd irdai_extractor
```

### Step 2 — Create src folder and empty __init__.py

```bash
mkdir src
touch src/__init__.py
```

### Step 3 — Copy your files into the folder

Place these files exactly as shown:
```
irdai_extractor/
├── app.py
├── extract_tables_smart_merged.py
├── requirements.txt
└── src/
    ├── __init__.py
    └── extractor.py
```

### Step 4 — Create a virtual environment (recommended)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Mac / Linux
python3 -m venv venv
source venv/bin/activate
```

### Step 5 — Install dependencies

```bash
pip install -r requirements.txt
```

### Step 6 — Run the Streamlit UI

```bash
streamlit run app.py
```

Your browser will open automatically at:
```
http://localhost:8501
```

---

## Using the App

1. **Upload PDF** — drag & drop or click the uploader on the left
2. **Set output filename** — optional, defaults to `IRDAI_Extract.xlsx`
3. **Click "Extract to Excel"** — watch the progress bar and live log
4. **Download** — click the blue download button when complete

---

## Using the CLI (no UI)

```bash
python extract_tables_smart_merged.py path/to/disclosure.pdf output.xlsx
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError: pdfplumber` | Run `pip install -r requirements.txt` |
| `ModuleNotFoundError: src.extractor` | Make sure `src/__init__.py` exists |
| Browser doesn't open | Go to `http://localhost:8501` manually |
| Port 8501 in use | Run `streamlit run app.py --server.port 8502` |
| Large PDF is slow | Normal — allow 1–2 min per 100 pages |

---

## Supported Forms

All IRDAI FORM L-* disclosure types including:
- `L-1-A-RA` Revenue Account (20-column multi-company layout)
- `L-2-A-PL` Profit & Loss Account
- `L-3-A-BS` Balance Sheet
- `L-4` through `L-45` all other schedules