# Literature Tracker — CGM + Multi-Omics Datasets

Auto-updating literature search for Dr. Wu's lab.
Will Search PubMed weekly for new publicly available datasets relevant to:
1. CGM data with paired dietary/meal records
2. Multi-omics acute time series in humans (exercise, meals, vaccination)

---

New rows will be highlighted in green, ? means something to look up individually

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
