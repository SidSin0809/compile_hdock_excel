# compile_hdock_excel
Scrapes HDOCK job result pages, extracts the “Summary of the Top 10 Models” tables for any number of complexes, and builds a single-sheet Excel workbook with embedded download links to each job’s full results archive.

# compile_hdock_excel.py
Scrape any number of **HDOCK** job result pages, capture each job’s  
**“Summary of the Top 10 Models”** table, and combine everything into a single
Excel workbook.  
Every complex is stacked in one worksheet:
<ComplexID> Rank / Docking Score / Confidence Score / Ligand RMSD / Interface residues All results package ← clickable hyperlink to all_results.tar.gz (blank line)

Features
1. Single-sheet output – Easy filtering, sorting, or pivot-table analysis.
2. Embedded hyperlinks – One-click download of all_results.tar.gz for
each complex.
3. Flexible scraping
3.1 Tries the plaintext ranked_poses.txt first (fast).
3.2 Falls back to the HTML result page if needed.
3.3 Handles both row-oriented (native) and column-oriented table variants.
4. Polite scraping (1s delay) and descriptive logging.

# Installation
git clone https://github.com/SidSin0809/hdock_batch.git

cd hdock_batch

pip install -r requirements.txt

pandas>=2.0
openpyxl>=3.1
requests>=2.30
beautifulsoup4>=4.12

All dependencies are pure-Python and installable with pip install -r requirements.txt.

# Usage
1. Prepare an input list
Create hdock_urls.txt (or any filename you like) with one job per line:

6PB0-CRH http://hdock.phys.hust.edu.cn/data/xxxxxxxxxxxxx/

6WZG-Secretin http://hdock.phys.hust.edu.cn/data/xxxxxxxxxxxxx/

Lines starting with # are ignored
First column = sheet header / complex ID
Second column = base URL of the job directory (trailing “/” optional).

2. Run the script

python compile_hdock_excel.py \
      -i hdock_urls.txt \
      -o compiled_hdock_results.xlsx

You’ll see progress messages in the console.
If a job directory is unreachable or malformed, it’s reported and skipped.

3. Open the workbook

compiled_hdock_results.xlsx → worksheet Summary

Each complex block begins with a bold ID row.

The five-row matrix (Rank → Interface residues) follows.

The “All results package” row contains an Excel HYPERLINK formula pointing to all_results.tar.gz.

# Troubleshooting
“Top-10 table not found”

Check the URL and confirm the job is finished (HDOCK sometimes cleans
old jobs).

Empty workbook

All jobs failed to parse—run with one URL first for diagnostics.

