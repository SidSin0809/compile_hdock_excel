#!/usr/bin/env python3
"""
compile_hdock_excel.py  – single-sheet + download link
=========================================================
Layout for each complex block
-----------------------------
```
<ComplexID>
Rank / Docking ... / Interface rows (5 rows)
All results package   https://…/all_results.tar.gz
<blank line>
```
"""
from __future__ import annotations

import argparse
import sys
import time
from io import StringIO
from pathlib import Path
from typing import Optional
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

# -------- Config --------
HEADERS = {"User-Agent": "Mozilla/5.0 Chrome/126"}
TIMEOUT = 20
PAUSE = 1.0
SHEET = "Summary"

# -------- Network helpers --------

def _get_text(url: str) -> str:
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
    except requests.RequestException as exc:
        raise RuntimeError(f"{url} – {exc}") from exc
    return r.text

# -------- TXT parser --------

def _parse_ranked_txt(txt: str) -> Optional[pd.DataFrame]:
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    if not lines:
        return None
    if not lines[0].split()[0].isdigit():
        lines.insert(0, "rank dock conf rmsd")
    else:
        lines.insert(0, "rank dock conf rmsd")
    df = pd.read_csv(StringIO("\n".join(lines)), delim_whitespace=True)
    cmap = {}
    for c in df.columns:
        cl = c.lower()
        if cl.startswith("rank") or cl == "#":
            cmap[c] = "Rank"
        elif "dock" in cl:
            cmap[c] = "Docking Score"
        elif "conf" in cl:
            cmap[c] = "Confidence Score"
        elif "rmsd" in cl:
            cmap[c] = "Ligand RMSD (Å)"
    if {"Rank", "Docking Score", "Confidence Score", "Ligand RMSD (Å)"} <= set(cmap.values()):
        df.rename(columns=cmap, inplace=True)
        return df[["Rank", "Docking Score", "Confidence Score", "Ligand RMSD (Å)"]].head(10)
    return None

# -------- HTML parser --------

def _find_top10_table(soup: BeautifulSoup):
    label = soup.find(lambda t: t.name in {"strong", "h1", "h2", "h3", "h4"} and "top" in t.get_text().lower() and "10" in t.get_text())
    if label:
        tbl = label.find_next("table")
        if tbl:
            return tbl
    for tbl in soup.find_all("table"):
        th = tbl.find("th")
        if th and "rank" in th.get_text(strip=True).lower():
            text = tbl.get_text(" ", strip=True).lower()
            if "docking" in text or "confidence" in text:
                return tbl
    raise RuntimeError("Top-10 table not found")


def _parse_html(html: str) -> Optional[pd.DataFrame]:
    soup = BeautifulSoup(html, "html.parser")
    try:
        tbl = _find_top10_table(soup)
    except RuntimeError:
        return None
    df = pd.read_html(StringIO(str(tbl)), flavor="bs4")[0]
    if str(df.columns[0]).strip().lower() == "rank":
        df.rename(columns={df.columns[0]: "Metric"}, inplace=True)
        return df
    if isinstance(df.iat[0, 0], str) and df.iat[0, 0].strip().lower() == "rank":
        df.columns = [c.strip() for c in df.columns]
        ren = {c: ("Rank" if c.lower()=="rank" else "Docking Score" if "dock" in c.lower() else "Confidence Score" if "conf" in c.lower() else "Ligand RMSD (Å)" if "rmsd" in c.lower() else c) for c in df.columns}
        df.rename(columns=ren, inplace=True)
        return df[["Rank", "Docking Score", "Confidence Score", "Ligand RMSD (Å)"]].head(10)
    return None

# -------- wide conversion --------

def _to_wide(tidy: pd.DataFrame) -> pd.DataFrame:
    if "Metric" in tidy.columns:
        return tidy
    ranks = [str(int(r)) for r in tidy["Rank"].tolist()]
    wide = pd.DataFrame(index=["Rank", "Docking Score", "Confidence Score", "Ligand rmsd (Å)", "Interface residues"], columns=ranks)
    wide.loc["Rank"] = ranks
    wide.loc["Docking Score"] = tidy["Docking Score"].tolist()
    wide.loc["Confidence Score"] = tidy["Confidence Score"].tolist()
    wide.loc["Ligand rmsd (Å)"] = tidy["Ligand RMSD (Å)"].tolist()
    wide.loc["Interface residues"] = [f"model_{i}" for i in range(1, len(ranks)+1)]
    wide.reset_index(inplace=True)
    wide.rename(columns={"index": "Metric"}, inplace=True)
    return wide

# -------- scrape orchestrator --------

def scrape(url: str) -> pd.DataFrame:
    for n in ("ranked_poses.txt", "ranked.txt"):
        try:
            txt = _get_text(urljoin(url, n))
            tidy = _parse_ranked_txt(txt)
            if tidy is not None:
                return _to_wide(tidy)
        except RuntimeError:
            pass
    for n in ("", "result.html", "index.html", "results.html"):
        try:
            html = _get_text(urljoin(url, n))
        except RuntimeError:
            continue
        tidy_or_wide = _parse_html(html)
        if tidy_or_wide is not None:
            return tidy_or_wide if "Metric" in tidy_or_wide.columns else _to_wide(tidy_or_wide)
    raise RuntimeError("No parsable Top-10 data")

# -------- compile to single sheet --------

def compile_excel(in_file: Path, out_xlsx: Path):
    blocks = []
    for ln, raw in enumerate(in_file.read_text().splitlines(), 1):
        raw = raw.strip()
        if not raw or raw.startswith("#"):
            continue
        try:
            cid, url = raw.split(maxsplit=1)
        except ValueError:
            print(f"[WARN] line {ln}: bad format", file=sys.stderr)
            continue
        if not urlparse(url).path.endswith("/"):
            url += "/"
        print(f"[INFO] {cid}: scraping …", file=sys.stderr)
        try:
            wide = scrape(url)
            blocks.append((cid, url, wide))
            print(f"[INFO] {cid}: ✔", file=sys.stderr)
        except Exception as exc:
            print(f"[ERROR] {cid}: {exc}", file=sys.stderr)
        time.sleep(PAUSE)
    if not blocks:
        raise SystemExit("No data parsed – workbook not created.")

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xl:
        rptr = 0
        for cid, base_url, wide in blocks:
            # Header row (complex ID)
            hdr = pd.DataFrame([[cid] + [""]*(wide.shape[1]-1)], columns=wide.columns)
            hdr.to_excel(xl, sheet_name=SHEET, startrow=rptr, index=False, header=False)
            rptr += 1
            wide.to_excel(xl, sheet_name=SHEET, startrow=rptr, index=False, header=False)
            rptr += len(wide)
            # Hyperlink row
            link = urljoin(base_url, "all_results.tar.gz")
            formula = f'=HYPERLINK("{link}","all_results.tar.gz")'
            link_row = pd.DataFrame([["All results package", formula] + [""]*(wide.shape[1]-2)], columns=wide.columns)
            link_row.to_excel(xl, sheet_name=SHEET, startrow=rptr, index=False, header=False)
            rptr += 2  # blank line
    print(f"[DONE] Workbook → {out_xlsx}", file=sys.stderr)

# -------- CLI --------

def main():
    p = argparse.ArgumentParser(description="Compile HDOCK results into single-sheet Excel with download links")
    p.add_argument("-i", "--input", default="hdock_urls.txt", type=Path)
    p.add_argument("-o", "--output", default="compiled_hdock_results.xlsx", type=Path)
    args = p.parse_args()
    if not args.input.exists():
        p.error("input file not found")
    compile_excel(args.input, args.output)

if __name__ == "__main__":
    main()
