#!/usr/bin/env python3
"""
SPC Tenders Daily Synchronization Entrypoint
Discovers tenders from the Sri Lanka State Pharmaceuticals Corporation portal,
identifies unprocessed tenders against the manifest, extracts line items into
stylized Excel sheets, creates a release zip archive, and updates master catalogs.
"""

import os
import sys
import re
import json
import zipfile
import argparse
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Any, Set

from src.scraper import discover_all_online_tenders, download_pdf
from src.extractor import SPCTenderTableExtractor
from src.excel_builder import generate_tender_workbook, build_consolidated_workbook, generate_csv


def normalize_tender_no(tno: str) -> str:
    """Normalizes tender number for robust string comparison across 2-digit and 4-digit years."""
    if not tno:
        return ""
    # Strip whitespace, convert underscores to slashes, and uppercase
    norm = "".join(tno.replace("_", "/").split()).upper()
    # Normalize 4-digit years to 2-digit (e.g. /2027 -> /27)
    norm = re.sub(r'/20(\d{2})$', r'/\1', norm)
    # Strip trailing .PDF if filename passed
    if norm.endswith(".PDF"):
        norm = norm[:-4]
    return norm


def load_manifest(manifest_path: str) -> List[Dict[str, Any]]:
    if not os.path.exists(manifest_path):
        return []
    try:
        with open(manifest_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[WARN] Error reading manifest at {manifest_path}: {e}")
        return []


def save_manifest(manifest: List[Dict[str, Any]], manifest_path: str):
    os.makedirs(os.path.dirname(os.path.abspath(manifest_path)), exist_ok=True)
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)


def write_github_output(key_values: Dict[str, str]):
    gh_output_path = os.environ.get("GITHUB_OUTPUT")
    if not gh_output_path:
        return
    with open(gh_output_path, "a", encoding="utf-8") as f:
        for k, v in key_values.items():
            f.write(f"{k}={v}\n")


def main():
    parser = argparse.ArgumentParser(description="Synchronize SPC Tenders and extract tables.")
    parser.add_argument("--data-dir", default="data", help="Directory containing manifest and extracted data.")
    parser.add_argument("--output-dir", default="output", help="Directory for release artifacts.")
    parser.add_argument("--downloads-dir", default="downloads", help="Directory for downloaded PDFs.")
    parser.add_argument("--dry-run", action="store_true", help="Only check for new tenders without processing.")
    parser.add_argument("--force-rebuild", action="store_true", help="Rebuild consolidated files from existing data.")
    args = parser.parse_args()

    data_dir = os.path.abspath(args.data_dir)
    output_dir = os.path.abspath(args.output_dir)
    downloads_dir = os.path.abspath(args.downloads_dir)

    os.makedirs(data_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(downloads_dir, exist_ok=True)

    manifest_path = os.path.join(data_dir, "manifest.json")
    jsonl_path = os.path.join(data_dir, "extracted_tenders.jsonl")
    csv_path = os.path.join(data_dir, "extracted_items.csv")

    manifest = load_manifest(manifest_path)
    known_tenders: Set[str] = set()
    for t in manifest:
        if t.get("tender_no"):
            known_tenders.add(normalize_tender_no(t["tender_no"]))
        if t.get("filename"):
            known_tenders.add(normalize_tender_no(t["filename"]))

    if os.path.exists(jsonl_path):
        try:
            with open(jsonl_path, "r", encoding="utf-8") as f:
                for line in f:
                    if line.strip():
                        rec = json.loads(line)
                        meta_tno = rec.get("metadata", {}).get("procurement_no")
                        if meta_tno:
                            known_tenders.add(normalize_tender_no(meta_tno))
                        fn = rec.get("file_name")
                        if fn:
                            known_tenders.add(normalize_tender_no(fn))
        except Exception as e:
            print(f"[WARN] Error reading jsonl for known tenders: {e}")
    print(f"[INFO] Loaded manifest containing {len(manifest)} existing tenders ({len(known_tenders)} unique IDs).")

    if args.force_rebuild:
        print("[INFO] Rebuilding consolidated files from extracted_tenders.jsonl...")
        if os.path.exists(jsonl_path):
            all_records = []
            with open(jsonl_path, "r", encoding="utf-8") as f:
                for line in f:
                    if line.strip():
                        all_records.append(json.loads(line))
            cons_xlsx = os.path.join(output_dir, "all_tenders_consolidated.xlsx")
            cons_csv = os.path.join(output_dir, "extracted_items.csv")
            build_consolidated_workbook(all_records, cons_xlsx)
            generate_csv(all_records, cons_csv)
            # Also sync master CSV to data_dir
            generate_csv(all_records, csv_path)
            print(f"[SUCCESS] Rebuilt master consolidated workbook and CSV with {len(all_records)} tenders.")
        else:
            print(f"[ERROR] Cannot rebuild: {jsonl_path} does not exist.")
        if not args.dry_run and not sys.argv[1:]:
            return 0

    # 1. Discover all online tenders
    print("[INFO] Querying SPC portal for active DHS/SS and DHS/SA tenders...")
    online_tenders = discover_all_online_tenders()
    print(f"[INFO] Discovered {len(online_tenders)} matching tenders online.")

    # 2. Identify new tenders
    new_tenders = []
    for t in online_tenders:
        norm_no = normalize_tender_no(t.get("tender_no", ""))
        if norm_no and norm_no not in known_tenders:
            new_tenders.append(t)

    print(f"[INFO] New, unprocessed tenders detected: {len(new_tenders)}")

    if not new_tenders:
        print("[SUCCESS] All online tenders have already been processed. Nothing to do.")
        write_github_output({
            "has_new": "false",
            "new_count": "0"
        })
        return 0

    print(f"[INFO] Listing {len(new_tenders)} newly detected tenders:")
    for nt in new_tenders:
        print(f"  • {nt.get('tender_no')} | Closing: {nt.get('closing_date')} | Category: {nt.get('category')}")

    if args.dry_run:
        print("[INFO] Dry run requested. Halting before download and extraction.")
        write_github_output({
            "has_new": "true",
            "new_count": str(len(new_tenders)),
            "dry_run": "true"
        })
        return 0

    # 3. Process each new tender
    now_utc = datetime.now(timezone.utc)
    date_str = now_utc.strftime("%Y-%m-%d")
    tag_name = f"v{now_utc.strftime('%Y.%m.%d')}"
    release_title = f"SPC Tenders Update — {date_str} ({len(new_tenders)} new tenders)"

    new_tenders_dir = os.path.join(output_dir, f"new_tenders_{date_str}")
    os.makedirs(new_tenders_dir, exist_ok=True)

    extractor = SPCTenderTableExtractor()
    new_records = []

    for idx, tender_info in enumerate(new_tenders, start=1):
        tno = tender_info.get("tender_no", "")
        doc_url = tender_info.get("document_url", "")
        fn = tender_info.get("filename", f"{tno.replace('/', '_')}.pdf")
        pdf_path = os.path.join(downloads_dir, fn)

        print(f"[{idx}/{len(new_tenders)}] Downloading PDF for {tno}...")
        dl_ok = download_pdf(doc_url, pdf_path)
        if not dl_ok:
            print(f"[ERROR] Could not download PDF for {tno} from {doc_url}")
            tender_record = {
                "file_path": pdf_path,
                "file_name": fn,
                "status": "DOWNLOAD_FAILED",
                "error": "Failed to download PDF from SPC server",
                "metadata": {"procurement_no": tno, "closing_date": tender_info.get("closing_date", "")},
                "items": []
            }
        else:
            print(f"[{idx}/{len(new_tenders)}] Extracting tables from {fn}...")
            tender_record = extractor.process_pdf(pdf_path)

        # Generate individual workbook
        xlsx_name = fn.replace(".pdf", ".xlsx")
        xlsx_path = os.path.join(new_tenders_dir, xlsx_name)
        generate_tender_workbook(tender_record, xlsx_path)

        new_records.append(tender_record)
        manifest.append(tender_info)
        known_tenders.add(normalize_tender_no(tno))

    # 4. Create ZIP archive of all newly extracted tables
    zip_name = f"new_tenders_{date_str}.zip"
    zip_path = os.path.join(output_dir, zip_name)
    print(f"[INFO] Creating zip archive of new tenders: {zip_path}")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(new_tenders_dir):
            for file in sorted(files):
                if file.endswith(".xlsx"):
                    file_path = os.path.join(root, file)
                    zf.write(file_path, arcname=file)

    # 5. Append new records to JSONL dataset
    print(f"[INFO] Appending {len(new_records)} records to {jsonl_path}...")
    with open(jsonl_path, "a", encoding="utf-8") as f:
        for rec in new_records:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    # 6. Save updated manifest
    print(f"[INFO] Saving updated manifest with {len(manifest)} total tenders...")
    save_manifest(manifest, manifest_path)

    # 7. Rebuild master consolidated workbook & CSV
    print("[INFO] Rebuilding consolidated master workbook & CSV across all tenders...")
    all_records = []
    with open(jsonl_path, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                all_records.append(json.loads(line))

    consolidated_xlsx = os.path.join(output_dir, "all_tenders_consolidated.xlsx")
    consolidated_csv = os.path.join(output_dir, "extracted_items.csv")
    build_consolidated_workbook(all_records, consolidated_xlsx)
    generate_csv(all_records, consolidated_csv)
    # Also update repository copy of extracted_items.csv
    generate_csv(all_records, csv_path)

    # 8. Write release notes and outputs for GitHub Actions
    release_notes_file = os.path.join(output_dir, "release_notes.md")
    notes_content = [
        f"## 📋 SPC Tenders Daily Extraction Update ({date_str})",
        "",
        f"Automatically detected and processed **{len(new_tenders)} new tender(s)** from Sri Lanka State Pharmaceuticals Corporation.",
        "",
        "### 📦 Newly Processed Tenders",
        "| Tender No | Category | Closing Date | Items Extracted | Status |",
        "| :--- | :---: | :---: | :---: | :---: |"
    ]
    for nr in new_records:
        meta = nr.get("metadata", {})
        tno = meta.get("procurement_no") or nr.get("file_name", "").replace(".pdf", "")
        cat = "DHS/SS" if "DHS/SS" in tno.upper() or "DHS_SS" in nr.get("file_name", "") else "DHS/SA"
        closing = meta.get("closing_date") or "N/A"
        icnt = len(nr.get("items", []))
        st = "✅ Success" if nr.get("status") == "SUCCESS" else "⚠️ Failed"
        notes_content.append(f"| `{tno}` | {cat} | {closing} | **{icnt}** | {st} |")

    notes_content.extend([
        "",
        "### 📁 Attached Release Assets",
        f"- **`{zip_name}`**: Zip archive containing individual stylized Excel workbooks for the **{len(new_tenders)} newly processed tenders**.",
        f"- **`all_tenders_consolidated.xlsx`**: Updated master multi-tab spreadsheet containing all **{len(all_records)}** catalogued tenders and their items.",
        f"- **`extracted_items.csv`**: Full consolidated CSV export with every extracted line item across all tenders.",
        "",
        "---",
        "*Automated synchronization powered by [SPC Tenders Pipeline](https://github.com/budd9442/spc-tenders).*"
    ])

    with open(release_notes_file, "w", encoding="utf-8") as f:
        f.write("\n".join(notes_content))

    write_github_output({
        "has_new": "true",
        "new_count": str(len(new_tenders)),
        "tag_name": tag_name,
        "release_title": release_title,
        "zip_path": zip_path,
        "consolidated_xlsx": consolidated_xlsx,
        "extracted_csv": consolidated_csv,
        "release_notes_file": release_notes_file
    })

    print(f"[SUCCESS] Completed daily sync. Extracted {len(new_tenders)} new tenders. Assets ready at {output_dir}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
