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

# Merge behaviour bila row sedia ada dijumpai (match by DEDUP_KEYS):
#
# ALWAYS_OVERWRITE_FIELDS -> value terbaru dari source SENTIASA menang,
#   walaupun cell sedia ada dah ada value. Sebab field ni memang jangka
#   BERUBAH sepanjang lifecycle incident (contoh: Status OPEN -> CLOSED).
ALWAYS_OVERWRITE_FIELDS = ["Status", "Outage End", "Severity"]

# FILL_BLANK_ONLY_FIELDS -> hanya isi kalau cell sedia ada KOSONG.
#   Kalau dah ada value, KEKALKAN (elak accidental overwrite data yang
#   dah betul dengan value dari source lain yang mungkin kurang tepat).
FILL_BLANK_ONLY_FIELDS = [
    c for c in COLUMNS
    if c not in ALWAYS_OVERWRITE_FIELDS and c not in DEDUP_KEYS and c not in FORMULA_COLUMNS
]
