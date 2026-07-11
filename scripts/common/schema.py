"""
schema.py
Definisi standard columns untuk master_outage.xlsx.

Semua source parser (parse_email_source.py, parse_zip_source.py, dsb)
KENA hasilkan output row dengan columns ni (guna nama key yang sama),
supaya merge_to_master.py boleh proses semua source secara sama rata.
"""

# Column order macam yang akan tertulis dalam Excel (kiri ke kanan)
COLUMNS = [
    "INC No",
    "Site ID",
    "Site Name",
    "Svc Imp",
    "Outage Start",
    "Outage End",
    "Duration (Hour)",
    "Severity",
    "Region",
    "Status",
    "District",
]

# Key untuk dedup / match row sedia ada vs row baru
DEDUP_KEYS = ["INC No", "Site ID"]

# Columns yang datetime object (untuk formatting khas dalam excel_utils.py)
DATETIME_COLUMNS = ["Outage Start", "Outage End"]

# Column formula (dikira, bukan raw value dari source)
FORMULA_COLUMNS = ["Duration (Hour)"]

# Number format string untuk setiap column (dipakai oleh excel_utils.py)
NUMBER_FORMATS = {
    "Outage Start": "dd/mm/yyyy hh:mm",
    "Outage End": "dd/mm/yyyy hh:mm",
    "Duration (Hour)": "[h]:mm",
}

# Columns yang WAJIB ada value (row akan di-skip kalau kosong)
REQUIRED_FIELDS = ["INC No", "Site ID"]
