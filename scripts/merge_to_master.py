"""
merge_to_master.py
Orchestrator utama untuk email source: baca semua raw HTML dalam
inbox/email_html/, parse guna parse_email_source.py, merge/update ke
master_outage.xlsx, then archive raw file ke processed/email_html/.

Design:
- Load master workbook SEKALI je (bukan per-file), semua row diproses,
  save SEKALI di hujung — elak I/O berulang & git diff besar tak perlu.
- Kalau parsing satu file gagal (INC No tak jumpa), file tu SKIP (tak
  archive, tak proses) supaya boleh disemak manual — tak silently hilang.
- Fail lain still diproses walaupun satu file gagal (tak stop whole run).

Run terus: python merge_to_master.py
(path inbox/processed/master boleh override guna CLI arg, tengok main())
"""

import sys
import shutil
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
from sources.parse_email_source import parse_email_html  # noqa: E402


def process_inbox(inbox_dir: Path, processed_dir: Path, master_path: Path):
    """
    Proses semua *.html dalam inbox_dir, merge ke master_path,
    archive file yang berjaya diproses ke processed_dir.

    Return: dict summary {files_processed, files_skipped, rows_inserted, rows_updated}
    """
    inbox_dir = Path(inbox_dir)
    processed_dir = Path(processed_dir)
    master_path = Path(master_path)

    processed_dir.mkdir(parents=True, exist_ok=True)

    html_files = sorted(inbox_dir.glob("*.html"))

    summary = {
        "files_processed": 0,
        "files_skipped": 0,
        "rows_inserted": 0,
        "rows_updated": 0,
    }

    if not html_files:
        print(f"Takde file .html dalam {inbox_dir}, tiada apa nak diproses.")
        return summary

    wb, ws = get_or_create_workbook(str(master_path))
    index = build_index(ws)  # O(n) SEKALI je (bukan per row)
    next_row_idx = ws.max_row + 1  # panggil SEKALI je, track sendiri lepas ni

    files_to_archive = []

    for html_file in html_files:
        html_content = html_file.read_text(encoding="utf-8")

        # Filename dipakai sebagai fallback "subject" (kalau Power Automate
        # save file guna nama yang mengandungi INC No, contoh INC000102173805.html)
        subject_guess = html_file.stem

        rows = parse_email_html(html_content, subject=subject_guess)

        if not rows:
            print(f"  [SKIP] {html_file.name} — INC No tak dapat di-extract, biar dalam inbox untuk semakan manual.")
            summary["files_skipped"] += 1
            continue

        for row in rows:
            key = (row["INC No"], row["Site ID"])
            existing_row_idx = index.get(key)
            if existing_row_idx:
                write_row(ws, existing_row_idx, row, fill_blank_only=True)
                summary["rows_updated"] += 1
            else:
                append_row(ws, row, row_idx=next_row_idx)
                index[key] = next_row_idx
                next_row_idx += 1
                summary["rows_inserted"] += 1

        print(f"  [OK] {html_file.name} — {len(rows)} site row(s) diproses (INC {rows[0]['INC No']})")
        summary["files_processed"] += 1
        files_to_archive.append(html_file)

    save_workbook(wb, str(master_path))

    # Archive lepas save berjaya (elak file hilang dari inbox kalau save gagal)
    for html_file in files_to_archive:
        dest = processed_dir / html_file.name
        shutil.move(str(html_file), str(dest))

    return summary


def main():
    repo_root = Path(__file__).resolve().parent.parent

    parser = argparse.ArgumentParser(description="Merge email outage HTML ke master Excel")
    parser.add_argument(
        "--inbox", default=str(repo_root / "inbox" / "email_html"),
        help="Folder raw HTML masuk (default: inbox/email_html)"
    )
    parser.add_argument(
        "--processed", default=str(repo_root / "processed" / "email_html"),
        help="Folder archive lepas diproses (default: processed/email_html)"
    )
    parser.add_argument(
        "--master", default=str(repo_root / "master" / "master_outage.xlsx"),
        help="Path master Excel (default: master/master_outage.xlsx)"
    )
    args = parser.parse_args()

    print(f"Inbox   : {args.inbox}")
    print(f"Master  : {args.master}")
    print(f"Archive : {args.processed}")
    print()

    summary = process_inbox(args.inbox, args.processed, args.master)

    print()
    print("=== Ringkasan ===")
    print(f"  Fail diproses : {summary['files_processed']}")
    print(f"  Fail di-skip  : {summary['files_skipped']}")
    print(f"  Row baru      : {summary['rows_inserted']}")
    print(f"  Row diupdate  : {summary['rows_updated']}")


if __name__ == "__main__":
    main()
