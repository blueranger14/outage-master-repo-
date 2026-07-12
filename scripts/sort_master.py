"""
sort_master.py
Susun semula master_outage.xlsx ikut "Outage Start" — TERKINI DI ATAS
(descending). Row tanpa Outage Start (blank/kosong) diletak di HUJUNG
sekali (bukan atas, supaya tak kacau susunan tarikh yang sah).

Guna semula excel_utils.py (write_row) untuk regenerate formula Duration
+ number format betul mengikut row position BARU (formula lama
"=IF(AND(E5<>"",F5<>""),...)" tak valid lepas row tu pindah posisi,
kena regenerate ikut row index baru).

Run terus: python scripts/sort_master.py
(boleh override path guna --master)
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).resolve().parent))

from common.excel_utils import (  # noqa: E402
    get_or_create_workbook,
    write_row,
    save_workbook,
)
from common.schema import COLUMNS, FORMULA_COLUMNS  # noqa: E402

# Row tanpa Outage Start diletak PALING BAWAH (bukan atas) — guna
# datetime.min sebagai fallback sort key supaya dia sentiasa "paling lama"
# dalam susunan descending (terkini di atas).
FALLBACK_DATE = datetime.min


def sort_master(master_path: str):
    master_path = Path(master_path)

    if not master_path.exists():
        print(f"Master file tak wujud: {master_path}")
        return 0

    wb, ws = get_or_create_workbook(str(master_path))

    if ws.max_row < 2:
        print("Master file kosong, tiada apa nak disusun.")
        return 0

    # Baca semua data row (row 2 hingga akhir) jadi list of dict
    read_columns = [c for c in COLUMNS if c not in FORMULA_COLUMNS]
    rows = []
    for row_idx in range(2, ws.max_row + 1):
        row_data = {}
        has_value = False
        for col_name in read_columns:
            col_idx = COLUMNS.index(col_name) + 1
            value = ws.cell(row=row_idx, column=col_idx).value
            row_data[col_name] = value
            if value not in (None, ""):
                has_value = True
        if has_value:  # skip row kosong sepenuhnya (kalau ada)
            rows.append(row_data)

    total = len(rows)
    print(f"Total row dibaca: {total}")

    # Sort by Outage Start, TERKINI DI ATAS (descending). Row tanpa
    # Outage Start (None) jatuh ke fallback (paling bawah).
    def sort_key(row):
        val = row.get("Outage Start")
        if isinstance(val, datetime):
            return val
        return FALLBACK_DATE

    rows.sort(key=sort_key, reverse=True)

    # Padam semua data row sedia ada (row 2 hingga akhir), kekalkan header
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # Tulis semula ikut susunan baru
    for i, row_data in enumerate(rows, start=2):
        write_row(ws, i, row_data, fill_blank_only=False)

    save_workbook(wb, str(master_path))
    print(f"Selesai — {total} row disusun semula (Outage Start terkini di atas).")
    return total


def main():
    repo_root = Path(__file__).resolve().parent.parent
    parser = argparse.ArgumentParser(description="Sort master_outage.xlsx by Outage Start (terkini di atas)")
    parser.add_argument(
        "--master", default=str(repo_root / "master" / "master_outage.xlsx"),
        help="Path master Excel (default: master/master_outage.xlsx)"
    )
    args = parser.parse_args()
    sort_master(args.master)


if __name__ == "__main__":
    main()
