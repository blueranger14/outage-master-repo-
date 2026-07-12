r"""
parse_zip_source.py
Parse WhatsApp chat export (.zip -> .txt) yang mengandungi mesej
"CD OUTAGE *" (ALERT/UPDATE/INFO/CLOSURE), hasilkan rows yang match
dengan schema.py punya COLUMNS, untuk terus merge ke master_outage.xlsx.

Refactor dari github_process.py (versi lama, tulis ke fail harian
berasingan). Perubahan utama:
- Output terus match schema.py (row per site), untuk merge ke SATU
  master file (bukan CD_OUTAGE_DDMMYYYY.xlsx berasingan lagi).
- Setiap zip diproses BERASINGAN (tak perlu tunggu pasangan
  Borneo+Peninsular macam versi lama).
- Field Outage End / Status / District SENGAJA tak dimasukkan dalam
  row output (zip source memang tak ada maklumat ni) - supaya
  merge_to_master.py punya "fill blank only" logic tak accidentally
  overwrite value yang dah diisi oleh source lain (email).
- Dedup + write-to-excel logic dah pindah ke common/excel_utils.py +
  merge_zip_to_master.py (orchestrator), tak lagi dalam file ni.

CONFIG FILTER (edit terus di bawah untuk toggle manual):
- DATE_FILTER: None = proses SEMUA tarikh. Uncomment baris "YESTERDAY"
  untuk filter ke semalam sahaja.
- SEVERITY_FILTER: None = proses SEMUA severity. Uncomment salah satu
  set() untuk filter ke severity tertentu sahaja.
"""

import re
import os
import zipfile
import zoneinfo
from datetime import datetime, timedelta

TZ = zoneinfo.ZoneInfo("Asia/Kuala_Lumpur")

# ─── CONFIG FILTER (toggle manual — uncomment baris yang nak guna) ─────────

# -- Date filter --
# DATE_FILTER = None  # None = proses SEMUA tarikh (default)
DATE_FILTER = (datetime.now(TZ) - timedelta(days=1)).date()  # <- semalam sahaja

# -- Severity filter --
SEVERITY_FILTER = None  # None = proses SEMUA severity (default)
# SEVERITY_FILTER = {"Critical"}
# SEVERITY_FILTER = {"Major"}
# SEVERITY_FILTER = {"Minor"}
# SEVERITY_FILTER = {"Major", "Critical"}  # contoh gabungan


# ─── REGEX PATTERNS ─────────────────────────────────────────────────────────

MSG_START_RE = re.compile(r'^\[\d{1,2}/\d{1,2}/\d{4}, \d{1,2}:\d{2}:\d{2}\] ')
# Tolerant terhadap teks tersasul antara "@" dan waktu (contoh: NOC
# accidentally taip INC No atau ulang tarikh dalam baris Outage Start):
#   "Outage Start: 26/04/2026 @ INC000000000105 outage 09:39"
#   "Outage Start: 19/06/2026 @ 19/6/2026 @ 13:51"
# Group 1 = tarikh, Group 2 = teks pelik (kalau ada, untuk warning), Group 3 = waktu
OUTAGE_START_RE = re.compile(r'Outage Start:\s*(\d{1,2}/\d{1,2}/\d{4})\s*@\s*(.*?)\s*(\d{1,2}:\d{2})')
# Tolerant terhadap typo biasa bila NOC taip manual, contoh:
#   (INC000099786348z)  - ada huruf lebih kat belakang sebelum ")"
#   (NC000100048024)    - huruf "I" tertinggal kat depan
INC_NO_RE = re.compile(r'\(I?NC(\d+)[A-Za-z]?\)')

SVC_IMP_RE  = re.compile(r'^Svc Imp:[ \t]*(.+)?$', re.MULTILINE)
SEVERITY_RE = re.compile(r'^Severity:[ \t]*(.+)?$', re.MULTILINE)
REGION_RE   = re.compile(r'^Region[ \t]*:[ \t]*(.+)?$', re.MULTILINE)
CP_RE       = re.compile(r'^CP[ \t]*:[ \t]*(.+)?$', re.MULTILINE)

# Pattern untuk kenal pasti token "Site ID". DUA bentuk:
#   1. Digit dulu, huruf lepas (contoh: 9227B, 7039A) - format Borneo biasa
#   2. SATU huruf prefix + digit + optional huruf suffix (contoh: Q02279,
#      S00174, C00665, D00853, J02163, T00072, dsb - Peninsular guna
#      pelbagai prefix ikut region/cluster, bukan cuma Q/S)
#
# PENTING: prefix mesti TEPAT SATU huruf. Token dengan 2+ huruf prefix
# (contoh DQKCH0478, ME1019345177, WL1042439732) BUKAN site code - itu
# Circuit ID/PTN number (muncul dalam block "Circuit ID:" berasingan,
# tiada nama site lepas dia). Regex ni sengaja EXCLUDE pattern macam tu -
# jangan ubah [A-Za-z] tunggal ni jadi terima >1 huruf prefix.
SITE_CODE_RE = re.compile(r'^([0-9]{3,6}[A-Za-z]{1,2}|[A-Za-z][0-9]{3,6}[A-Za-z]?)$')

# Pattern Marketing Cluster code (contoh: C047A, D018A, M077C, Q124E) -
# SECARA SINTAKS ia match SITE_CODE_RE juga (satu huruf + 3 digit + satu
# huruf), tapi ia CLUSTER bukan Site ID sebenar. Dipakai untuk EXCLUDE
# token ni bila muncul sebagai token PERTAMA dalam format underscore
# (contoh "D018A_D00833OD_KAMPUNG..." - "D018A" tu cluster, "D00833" tu
# baru Site ID sebenar).
MARKETING_CLUSTER_RE = re.compile(r'^[A-Za-z]\d{3}[A-Za-z]$')

# Format "<digit code>-<nama>" (contoh: 20135-MILE_86_SDK-ACC1,
# 43191-TMN_DELIMA2) - dijumpai dalam sesetengah Peninsular data untuk
# site fixed-line/microwave. Kod = digit tulen (4-6 digit), tiada huruf.
DASH_NUMERIC_RE = re.compile(r'^(\d{4,6})-(.+)$')

# Kod site + suffix teknikal (contoh: D00833OD, W01515IB, W00350IB1) -
# dipakai untuk extract Site ID sebenar dari segment ke-2 underscore
# token bila segment pertama tu Marketing Cluster (rujuk atas).
CODE_WITH_TECH_SUFFIX_RE = re.compile(r'^([A-Za-z]\d{4,6})[A-Za-z0-9]{0,4}$')

# Panjang digit biasa untuk nombor INC (based on real data pattern - hampir
# semua INC No ada 12 digit). Kalau lain, kemungkinan besar typo - script
# akan bagi WARNING sahaja, tak "baiki" sendiri.
EXPECTED_INC_DIGIT_LEN = 12


# ─── ZIP EXTRACTION ──────────────────────────────────────────────────────────

def extract_txt_from_zip(zip_path, extract_to):
    """Extract semua fail .txt dari dalam zip_path ke extract_to.
    Return list path fail .txt yang berjaya di-extract."""
    txt_paths = []
    with zipfile.ZipFile(zip_path, 'r') as zf:
        for name in zf.namelist():
            if name.lower().endswith('.txt'):
                zf.extract(name, extract_to)
                txt_paths.append(os.path.join(extract_to, name))
    return txt_paths


# ─── PARSING HELPERS ─────────────────────────────────────────────────────────

def split_messages(raw_text):
    """Split a raw WhatsApp export into individual messages (multi-line aware)."""
    raw_text = raw_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = raw_text.split("\n")

    messages = []
    current = []
    for line in lines:
        if MSG_START_RE.match(line):
            if current:
                messages.append("\n".join(current))
            current = [line]
        else:
            current.append(line)
    if current:
        messages.append("\n".join(current))
    return messages


def _field(regex, text, default=""):
    m = regex.search(text)
    if not m or m.group(1) is None:
        return default
    return m.group(1).strip()


def parse_cp_field(cp_raw):
    """CP field format: '<Site ID> <Site Name>', kadang beberapa CP
    dipisah koma: 'S00174 NAME1, S00220 NAME2'. Return list of (code, name)."""
    cp_raw = cp_raw.strip()
    if not cp_raw or cp_raw == "-":
        return []
    entries = []
    for part in cp_raw.split(","):
        part = part.strip()
        if not part or part == "-":
            continue
        tokens = part.split(None, 1)
        code = tokens[0]
        name = tokens[1].strip() if len(tokens) > 1 else ""
        entries.append((code, name))
    return entries


def try_extract_underscore_token(token):
    """
    Cuba extract (site_id, site_name) dari SATU token tanpa space (guna
    underscore atau dash sebagai separator). Return None kalau tak
    confident (elak teka salah - lebih baik skip drpd data silap).

    Handle 3 kes:
    1. Format dash-numeric: "20135-MILE_86_SDK-ACC1" -> ('20135', 'MILE 86 SDK-ACC1')
    2. Underscore, segment PERTAMA dah site code sah (bukan cluster):
       "1203K_NIC_W066N_..." -> ('1203K', 'NIC W066N ...')
    3. Underscore, segment PERTAMA cluster, kod sebenar dalam segment KEDUA:
       "D018A_D00833OD_KAMPUNG_SUNGAI_PERIA" -> ('D00833', 'KAMPUNG SUNGAI PERIA')
    """
    # --- Kes 1: dash-numeric ---
    m = DASH_NUMERIC_RE.match(token)
    if m:
        code = m.group(1)
        name = m.group(2).replace('_', ' ').strip()
        return code, name

    if '_' not in token:
        return None

    segments = token.split('_')
    if not segments or not segments[0]:
        return None

    first_seg = segments[0]

    # --- Kes 2: segment pertama terus site code sah ---
    if SITE_CODE_RE.match(first_seg) and not MARKETING_CLUSTER_RE.match(first_seg):
        name = ' '.join(segments[1:]).replace('_', ' ').strip()
        return first_seg, name

    # --- Kes 3: segment pertama BUKAN site code sah - scan SEMUA segment
    # lain (bukan cuma segment ke-2) untuk cari kod dengan tech suffix.
    # Kenapa loop (bukan cuma check index [1]): kadang ada LEBIH satu
    # segment "non-code" sebelum kod sebenar muncul, contoh
    # "B1_D018A_D00570OD_..." - "B1" (marker) lepas tu "D018A" (cluster)
    # BARU "D00570OD" (kod sebenar) di segment KE-3.
    # CODE_WITH_TECH_SUFFIX_RE anchored ketat (huruf tunggal + 4-6 digit)
    # jadi cluster (cuma 3 digit) & marker pendek takkan tersalah match.
    for idx, seg in enumerate(segments[1:], start=1):
        m2 = CODE_WITH_TECH_SUFFIX_RE.match(seg)
        if m2:
            code = m2.group(1)
            name = ' '.join(segments[idx + 1:]).replace('_', ' ').strip()
            return code, name

    return None  # tak confident - biar skip drpd teka salah


def build_site_map(msg):
    """Scan SELURUH mesej untuk sebarang baris berbentuk '<Site ID> <Site
    Name>' - tak kira dia dalam block CP, Affected Sites, atau vendor block
    (Sacofa/CTSB) - semua dianggap 1 record. Return dict {site_id: site_name}.

    Handle juga baris SATU TOKEN (tiada space) yang guna underscore/dash
    sebagai separator (contoh format "1203K_NIC_..." atau
    "20135-MILE_86_SDK-ACC1") - rujuk try_extract_underscore_token().
    """
    sites = {}
    for line in msg.split("\n"):
        line = line.strip()
        if not line or ":" in line:
            continue

        parts = line.split(None, 1)
        code = parts[0]

        if SITE_CODE_RE.match(code):
            name = parts[1].strip() if len(parts) > 1 else ""
            if code not in sites or (not sites[code] and name):
                sites[code] = name
            continue

        # Code (token pertama) tak match terus - cuba underscore/dash
        # extraction KALAU line ni memang satu token je (tiada space lain
        # selepas token pertama, sebab kalau ada, itu mungkin format lain
        # yang kita tak faham, elak teka)
        if len(parts) == 1:
            extracted = try_extract_underscore_token(code)
            if extracted:
                ex_code, ex_name = extracted
                if ex_code not in sites or (not sites[ex_code] and ex_name):
                    sites[ex_code] = ex_name

    # Field 'CP :' ada colon kat depan, jadi terlepas dari scan di atas -
    # parse & tambah secara eksplisit.
    cp_raw = _field(CP_RE, msg)
    for code, name in parse_cp_field(cp_raw):
        if not SITE_CODE_RE.match(code):
            continue
        if code not in sites or (not sites[code] and name):
            sites[code] = name

    return sites


def parse_outage_message(msg):
    """Return dict {'_start_date': date, '_severity': str, 'rows': [...]}
    untuk 1 mesej CD OUTAGE, atau None kalau tiada Outage Start yang boleh
    diparse. Setiap row dict = 1 site, keys match schema.py (tanpa Outage
    End / Status / District - zip source memang tak ada maklumat ni)."""
    start_match = OUTAGE_START_RE.search(msg)
    if not start_match:
        return None
    try:
        start_date = datetime.strptime(start_match.group(1), "%d/%m/%Y").date()
        start_dt = datetime.strptime(
            f"{start_match.group(1)} {start_match.group(3)}", "%d/%m/%Y %H:%M"
        )
    except Exception:
        return None

    if start_match.group(2):
        print(f"  [WARNING] Ada teks pelik dalam baris 'Outage Start' antara '@' dan waktu "
              f"({start_match.group(1)} @ ... {start_match.group(3)}): {start_match.group(2)!r}. "
              f"Waktu masih berjaya diparse, tapi sila semak manual kalau perlu.")

    inc_match = INC_NO_RE.search(msg)
    inc_no = f"INC{inc_match.group(1)}" if inc_match else ""
    if inc_match and len(inc_match.group(1)) != EXPECTED_INC_DIGIT_LEN:
        print(f"  [WARNING] INC No panjang tak biasa ({len(inc_match.group(1))} digit, "
              f"biasanya {EXPECTED_INC_DIGIT_LEN}): {inc_no} - Outage Start "
              f"{start_match.group(1)} @ {start_match.group(3)}. Sila semak manual.")

    severity = _field(SEVERITY_RE, msg)

    common = {
        "INC No": inc_no,
        "Svc Imp": _field(SVC_IMP_RE, msg),
        "Outage Start": start_dt,
        "Severity": severity,
        "Region": _field(REGION_RE, msg),
    }

    site_map = build_site_map(msg)

    rows = []
    for code, name in site_map.items():
        rows.append({**common, "Site ID": code, "Site Name": name})

    return {"_start_date": start_date, "_severity": severity, "rows": rows}


def _passes_filters(parsed):
    """Check DATE_FILTER & SEVERITY_FILTER (module-level config di atas)."""
    if DATE_FILTER is not None and parsed["_start_date"] != DATE_FILTER:
        return False
    if SEVERITY_FILTER is not None and parsed["_severity"] not in SEVERITY_FILTER:
        return False
    return True


def process_file(path):
    """Read one WhatsApp export .txt, return list of row dicts (schema.py
    format) untuk mesej CD OUTAGE yang pass DATE_FILTER/SEVERITY_FILTER."""
    print(f"\n[READ] {path}")
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        raw = f.read()

    all_msgs = split_messages(raw)
    # Scan SEMUA jenis "CD OUTAGE *" (ALERT/UPDATE/INFO/CLOSURE), bukan
    # ALERT sahaja - mesej UPDATE/INFO/CLOSURE pun ada full Site ID +
    # Outage Start dalam format yang sama, dan dedupe (INC No + Site ID)
    # jaga takde duplicate row walaupun INC sama muncul dalam >1 jenis.
    outage_msgs = [m for m in all_msgs if "CD OUTAGE" in m]
    print(f"  -> Total messages: {len(all_msgs)} | CD OUTAGE (ALERT/UPDATE/INFO/CLOSURE): {len(outage_msgs)}")

    rows, skipped_no_date, skipped_filtered, skipped_no_site = [], 0, 0, 0
    for m in outage_msgs:
        parsed = parse_outage_message(m)
        if parsed is None:
            skipped_no_date += 1
            continue
        if not _passes_filters(parsed):
            skipped_filtered += 1
            continue
        if not parsed["rows"]:
            skipped_no_site += 1
            continue
        rows.extend(parsed["rows"])

    print(f"  -> Rows dihasilkan: {len(rows)}")
    if skipped_no_date:
        print(f"  (skipped {skipped_no_date} alert(s) - no parseable Outage Start date)")
    if skipped_filtered:
        print(f"  (skipped {skipped_filtered} alert(s) - tak match DATE_FILTER/SEVERITY_FILTER)")
    if skipped_no_site:
        print(f"  (skipped {skipped_no_site} alert(s) - no parseable Site ID)")
    return rows


def process_zip(zip_path, tmp_extract_dir):
    """Extract + parse satu zip. Return list of row dicts (schema.py format)."""
    txt_paths = extract_txt_from_zip(zip_path, tmp_extract_dir)
    if not txt_paths:
        print(f"  ⚠️  No .txt found inside {os.path.basename(zip_path)}")
        return []

    rows = []
    for txt_path in txt_paths:
        rows.extend(process_file(txt_path))
    return rows


# ─── Quick test bila run terus (bukan import) ───────────────────────────────
if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python parse_zip_source.py <input.zip|input.txt>")
        sys.exit(1)

    input_path = sys.argv[1]
    if input_path.lower().endswith(".zip"):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp_dir:
            rows = process_zip(input_path, tmp_dir)
    else:
        rows = process_file(input_path)

    print(f"\nTotal rows: {len(rows)}")
    for r in rows[:10]:
        print(r)
