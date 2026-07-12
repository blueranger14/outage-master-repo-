"""
merge_zip_to_master.py
Orchestrator untuk zip source (WhatsApp export): baca semua .zip dalam
inbox/zip_history/, extract .txt, parse guna parse_zip_source.py, merge
ke master_outage.xlsx, then archive zip ke processed/zip_history/.

Beza dengan merge_to_master.py (email source):
- Setiap zip diproses BERASINGAN (tak perlu tunggu pasangan
  Borneo+Peninsular macam versi lama github_process.py).
- DATE_FILTER / SEVERITY_FILTER dikawal dalam parse_zip_source.py
  (edit terus kat situ untuk toggle).
- Zip source TAK bagi Outage End/Status/District - row output cuma ada
  field yang zip source memang tahu, supaya "fill blank only" logic
  dalam excel_utils.py tak accidentally overwrite value dari email source.

Run terus: python merge_zip_to_master.py
"""

import sys
import shutil
import tempfile
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent / "sources"))

from common.excel_utils import (  # noqa: E402
    get_or_create_workbook,
    build_index,
    write_row,
    append_row,
    save_workbook,
)
from sources.parse_zip_source import process_zip  # noqa: E402
from sort_master import sort_master  # noqa: E402


def process_inbox(inbox_dir: Path, processed_dir: Path, master_path: Path):
    """
    Proses semua *.zip dalam inbox_dir, merge ke master_path,
    archive zip yang berjaya diproses ke processed_dir.

    Return: dict summary {files_processed, files_skipped, rows_inserted, rows_updated}
    """
    inbox_dir = Path(inbox_dir)
    processed_dir = Path(processed_dir)
    master_path = Path(master_path)

    processed_dir.mkdir(parents=True, exist_ok=True)

    zip_files = sorted(inbox_dir.glob("*.zip"))

    summary = {
        "files_processed": 0,
        "files_skipped": 0,
        "rows_inserted": 0,
        "rows_updated": 0,
    }

    if not zip_files:
        print(f"Takde file .zip dalam {inbox_dir}, tiada apa nak diproses.")
        return summary

    wb, ws = get_or_create_workbook(str(master_path))
    index = build_index(ws)  # O(n) SEKALI je (bukan per row)
    next_row_idx = ws.max_row + 1  # panggil SEKALI je, track sendiri lepas ni

    files_to_archive = []

    with tempfile.TemporaryDirectory() as tmp_root:
        for zip_path in zip_files:
            print(f"\n[ZIP] {zip_path.name}")
            extract_subdir = Path(tmp_root) / zip_path.stem
            extract_subdir.mkdir(parents=True, exist_ok=True)

            try:
                rows = process_zip(str(zip_path), str(extract_subdir))
            except Exception as e:
                print(f"  [SKIP] {zip_path.name} — ERROR semasa proses: {e}")
                summary["files_skipped"] += 1
                continue

            if not rows:
                print(f"  [SKIP] {zip_path.name} — tiada row dihasilkan (tiada mesej match filter, atau tiada Site ID).")
                summary["files_skipped"] += 1
                continue

            for row in rows:
                inc_no = row.get("INC No")
                site_id = row.get("Site ID")
                if not inc_no or not site_id:
                    continue  # skip row tanpa key wajib

                key = (inc_no, site_id)
                existing_row_idx = index.get(key)
                if existing_row_idx:
                    write_row(ws, existing_row_idx, row, fill_blank_only=True)
                    summary["rows_updated"] += 1
                else:
                    append_row(ws, row, row_idx=next_row_idx)
                    index[key] = next_row_idx
                    next_row_idx += 1
                    summary["rows_inserted"] += 1

            print(f"  [OK] {zip_path.name} — {len(rows)} site row(s) diproses")
            summary["files_processed"] += 1
            files_to_archive.append(zip_path)

    save_workbook(wb, str(master_path))

    # Archive lepas save berjaya (elak file hilang dari inbox kalau save gagal)
    for zip_path in files_to_archive:
        dest = processed_dir / zip_path.name
        shutil.move(str(zip_path), str(dest))

    # Auto-sort master (terkini di atas) — HANYA kalau ada row baru/diupdate,
    # elak sort/save berulang tanpa perlu bila takde perubahan langsung.
    if summary["rows_inserted"] > 0 or summary["rows_updated"] > 0:
        print("\n[SORT] Menyusun master (Outage Start terkini di atas)...")
        sort_master(str(master_path))

    return summary


def main():
    repo_root = Path(__file__).resolve().parent.parent

    parser = argparse.ArgumentParser(description="Merge zip outage source ke master Excel")
    parser.add_argument(
        "--inbox", default=str(repo_root / "inbox" / "zip_history"),
        help="Folder zip masuk (default: inbox/zip_history)"
    )
    parser.add_argument(
        "--processed", default=str(repo_root / "processed" / "zip_history"),
        help="Folder archive lepas diproses (default: processed/zip_history)"
    )
    parser.add_argument(
        "--master", default=str(repo_root / "master" / "master_outage.xlsx"),
        help="Path master Excel (default: master/master_outage.xlsx)"
    )
    args = parser.parse_args()

    print(f"Inbox   : {args.inbox}")
    print(f"Master  : {args.master}")
    print(f"Archive : {args.processed}")

    summary = process_inbox(args.inbox, args.processed, args.master)

    print()
    print("=== Ringkasan ===")
    print(f"  Fail diproses : {summary['files_processed']}")
    print(f"  Fail di-skip  : {summary['files_skipped']}")
    print(f"  Row baru      : {summary['rows_inserted']}")
    print(f"  Row diupdate  : {summary['rows_updated']}")


if __name__ == "__main__":
    main()
