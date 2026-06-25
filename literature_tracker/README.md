# Literature Tracker — CGM + Multi-Omics Datasets

Auto-updating literature search for Dr. Wu's lab.
Searches PubMed weekly for new publicly available datasets relevant to:
1. CGM data with paired dietary/meal records
2. Multi-omics acute time series in humans (exercise, meals, vaccination)

---

## What this does

Every Monday at 9am, GitHub automatically:
1. Searches PubMed using targeted queries
2. Filters results for papers that describe actual open datasets
3. Adds new candidates as highlighted rows in the Excel spreadsheet
4. Commits the updated spreadsheet back to this repo

New rows are highlighted in **green** so they're easy to spot.
They are marked `[NEW]` and some fields will say `?` — those need
manual review and filling in (N, population, omics types, etc.).

---

## File structure

```
literature_tracker/
├── search_and_update.py          # Main script — runs the search
├── requirements.txt              # Python packages needed
├── README.md                     # This file
├── .github/
│   └── workflows/
│       └── weekly_update.yml     # GitHub Actions schedule config
└── data/
    ├── literature_tables.xlsx    # THE TABLE (auto-updated weekly)
    ├── search_log.json           # Tracks which PMIDs have been seen
    └── last_run_summary.json     # Summary of most recent run
```

---

## Setup instructions (one-time)

### 1. Create a GitHub account
Go to github.com and sign up (free).

### 2. Create a new repository
- Click the "+" button → "New repository"
- Name it something like `wu-lab-literature-tracker`
- Set it to **Public** or **Private** (either works)
- Click "Create repository"

### 3. Upload these files
You can drag and drop files directly on GitHub, or use Git if you know it.
Upload everything in this folder, keeping the folder structure the same.
The `.github/workflows/` folder is important — don't skip it.

### 4. Upload the spreadsheet
Put your `literature_tables.xlsx` file in a folder called `data/`
(create the folder on GitHub by uploading a file named `data/literature_tables.xlsx`).

### 5. Give GitHub Actions permission to push commits
- Go to your repo → Settings → Actions → General
- Scroll to "Workflow permissions"
- Select **"Read and write permissions"**
- Click Save

### 6. Test it manually
- Go to the Actions tab in your repo
- Click "Weekly Literature Search" on the left
- Click "Run workflow" → "Run workflow"
- Watch it run (takes about 1-2 minutes)
- Check the `data/` folder — the spreadsheet should be updated

After that, it runs automatically every Monday. You don't need to do anything.

---

## Customizing the search queries

In `search_and_update.py`, near the top you'll see:

```python
CGM_QUERIES = [
    "CGM dietary dataset open access meal times",
    ...
]

OMICS_QUERIES = [
    "multi-omics acute time series human open access perturbation",
    ...
]
```

You can add, remove, or change these queries anytime.
The script won't add duplicate papers — it remembers what it's seen before.

---

## Understanding the output

When the script runs, new rows are added to the spreadsheet with:
- **Green highlight** — so you can spot them immediately
- **`[NEW]` prefix** in the dataset name
- **`?`** in fields that need manual filling (N, population, etc.)
- The **PMID and DOI** so you can look up the paper

Your job after each weekly run is to:
1. Open the spreadsheet
2. Review the green rows
3. Look up each paper and fill in the `?` fields
4. Delete any rows that aren't actually relevant after reading the abstract

---

## Running manually (without GitHub)

If you have Python installed locally:

```bash
# Install dependencies
pip install -r requirements.txt

# Make sure your spreadsheet is in data/literature_tables.xlsx
# Then run:
python search_and_update.py
```

---

*Wu Lab — Shreya Bharathi Srikanth, 2026*
