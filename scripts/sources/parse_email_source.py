"""
parse_email_source.py
Parse HTML body email outage (CelcomDigi NOC format) jadi normalized rows
yang match dengan schema.py punya COLUMNS.

Struktur HTML yang dijangka (dari Power Automate "Get email body"):
1. Header box (bold, background gelap) mengandungi:
   "Incident No: INC000102173805<br>TT Status: CLOSED" + "Severity: Major"
2. Detail table — setiap <tr> ada 2 <td>: label (bold, warna #4472C4) & value.
   Fields: Outage Description, Region, Severity, Outage Start Time,
   Outage End Time, Customer Impact, Impacted Service, Progress, Sites.
3. Sites — nested <table> di dalam <td> row "Sites". Header row lepas tu
   data rows: Site ID, Site Name, Marketing Cluster, District.

Kalau struktur berubah / field tak jumpa, function akan return None untuk
field tu (bukan raise error) — supaya satu field hilang tak crash whole parse.
"""

import re
import sys
from datetime import datetime
from pathlib import Path
from bs4 import BeautifulSoup

# Supaya "from common.schema import ..." boleh jalan tak kira script ni
# dipanggil dari mana (repo root, scripts/sources/, atau GitHub Actions runner)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.schema import COLUMNS, FORMULA_COLUMNS  # noqa: E402

# Column output dari parser ni = semua COLUMNS KECUALI Duration (Hour),
# sebab tu formula yang dikira semasa tulis ke excel (bukan dari source)
OUTPUT_COLUMNS = [c for c in COLUMNS if c not in FORMULA_COLUMNS]


# ---------------------------------------------------------------------------
# Helper: extract text bersih dari satu <td> (strip whitespace berlebihan)
# ---------------------------------------------------------------------------
def _clean_text(td) -> str:
    if td is None:
        return ""
    text = td.get_text(separator=" ", strip=True)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# ---------------------------------------------------------------------------
# 1. Extract Incident No + TT Status dari header box
# ---------------------------------------------------------------------------
def extract_incident_header(soup):
    """
    Cari <b><span> yang ada text 'Incident No:' — extract INC No & TT Status
    dari situ (dipisah <br> dalam satu text block).
    Return: (inc_no, tt_status)
    """
    inc_no, tt_status = None, None

    # Cari semua <b> tag, check text dia contain "Incident No:"
    for b_tag in soup.find_all("b"):
        text = b_tag.get_text(separator="|", strip=True)  # | ganti <br>
        if "Incident No" in text:
            # Cari pattern INC diikuti digit
            m = re.search(r"Incident No:\s*(INC\d+)", text)
            if m:
                inc_no = m.group(1)

            # TT Status punya VALUE (contoh "CLOSED") selalunya dalam <b> tag
            # SIBLING berasingan (warna berbeza), bukan dalam <b> yang sama
            # macam "TT Status:" label. Kalau text ni ada "TT Status:" tapi
            # value kosong lepas tu, cari next sibling <b> tag punya text.
            m2 = re.search(r"TT Status:\s*([A-Za-z]*)", text)
            if m2 and m2.group(1):
                tt_status = m2.group(1).strip().upper()
            elif "TT Status" in text:
                sibling = b_tag.find_next_sibling("b")
                if sibling:
                    sib_text = sibling.get_text(strip=True)
                    if sib_text:
                        tt_status = sib_text.strip().upper()

            if inc_no:
                break

    return inc_no, tt_status


# ---------------------------------------------------------------------------
# 2. Extract Severity dari header box (kanan atas, berasingan dari status)
# ---------------------------------------------------------------------------
def extract_header_severity(soup):
    for span in soup.find_all("span"):
        text = span.get_text(strip=True)
        if text.startswith("Severity:"):
            m = re.search(r"Severity:\s*(\w*)", text)
            if m and m.group(1):
                return m.group(1)
    # fallback: cari <b><span> lepas "Severity:" punya sibling
    for tag in soup.find_all(string=re.compile(r"Severity:")):
        parent = tag.find_parent("td") or tag.find_parent("p")
        if parent:
            m = re.search(r"Severity:\s*([A-Za-z]+)", parent.get_text(" ", strip=True))
            if m:
                return m.group(1)
    return None


# ---------------------------------------------------------------------------
# 3. Extract detail table (label-value pairs) — EXCLUDE nested Sites table
# ---------------------------------------------------------------------------
def extract_detail_fields(soup):
    """
    Loop semua <tr> yang ada label bold warna #4472C4 dalam <td> pertama.
    Return dict {label: value_text}.
    Untuk row "Sites", value tak diambil di sini (nested table, handled
    berasingan oleh extract_sites_table).
    """
    fields = {}

    for tr in soup.find_all("tr"):
        tds = tr.find_all("td", recursive=False)
        if len(tds) != 2:
            continue

        label_td, value_td = tds[0], tds[1]

        # Label mesti ada <b> tag (bold) untuk qualify sebagai label row
        bold = label_td.find("b")
        if not bold:
            continue
        label = _clean_text(label_td)
        if not label:
            continue

        # Kalau value_td ada nested <table>, ni row "Sites" — skip value text
        # (diambil oleh extract_sites_table)
        if value_td.find("table"):
            fields[label] = None
            continue

        value = _clean_text(value_td)
        fields[label] = value

    return fields


# ---------------------------------------------------------------------------
# 4. Extract Sites table (nested table dalam row "Sites")
# ---------------------------------------------------------------------------
def extract_sites_table(soup):
    """
    Cari <td> yang label dia "Sites", masuk ke <td> sebelah (value),
    cari nested <table> di situ, parse header + data rows.

    Handle DUA variant struktur:
    1. Ada header row ("Site ID", "Site Name", dsb) -> map by header name
    2. TAKDE header row (terus data) -> guna posisi column tetap:
       [Site ID, Site Name, Marketing Cluster, District]
       (variant ni jumpa dalam email dari client "Microsoft Word 15
       (filtered medium)" generator, struktur border pun beza sikit
       tapi tak relevant untuk parsing)

    Return: list of dict [{"Site ID": ..., "Site Name": ..., "District": ...}, ...]
    """
    # Label header yang dijangka (lowercase, untuk case-insensitive match)
    EXPECTED_HEADERS = {"site id", "site name", "marketing cluster", "district"}
    DEFAULT_COLUMN_ORDER = ["Site ID", "Site Name", "Marketing Cluster", "District"]

    sites = []

    for tr in soup.find_all("tr"):
        tds = tr.find_all("td", recursive=False)
        if len(tds) != 2:
            continue
        label_td, value_td = tds[0], tds[1]
        label = _clean_text(label_td)
        if label != "Sites":
            continue

        nested_table = value_td.find("table")
        if not nested_table:
            continue

        rows = nested_table.find_all("tr")
        if not rows:
            continue

        # --- Detect sama ada row pertama tu header atau terus data ---
        first_row_cells = rows[0].find_all("td")
        first_row_texts = {_clean_text(c).lower() for c in first_row_cells}
        is_header_row = len(first_row_texts & EXPECTED_HEADERS) >= 2

        if is_header_row:
            headers = [_clean_text(c) for c in first_row_cells]
            data_rows = rows[1:]
        else:
            # Takde header -> guna posisi column tetap, SEMUA row = data
            headers = DEFAULT_COLUMN_ORDER[: len(first_row_cells)]
            data_rows = rows

        for row in data_rows:
            cells = row.find_all("td")
            if len(cells) != len(headers):
                continue
            row_data = dict(zip(headers, [_clean_text(c) for c in cells]))
            sites.append({
                "Site ID": row_data.get("Site ID", "").strip(),
                "Site Name": row_data.get("Site Name", "").strip(),
                "District": row_data.get("District", "").strip(),
            })

        break  # dah jumpa Sites table, tak perlu cari lagi

    return sites


# ---------------------------------------------------------------------------
# 5. Parse datetime string "11/07/2026 @ 15:33" -> datetime object
# ---------------------------------------------------------------------------
def parse_outage_datetime(text):
    if not text:
        return None
    text = text.strip()
    # Format utama: dd/mm/yyyy @ HH:MM
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s*@\s*(\d{1,2}):(\d{2})", text)
    if m:
        day, month, year, hour, minute = map(int, m.groups())
        try:
            return datetime(year, month, day, hour, minute)
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------
# 6. Fallback: extract INC No dari subject line (kalau body punya tak jumpa)
# ---------------------------------------------------------------------------
def extract_inc_from_subject(subject: str):
    if not subject:
        return None
    m = re.search(r"(INC\d+)", subject)
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# MAIN: parse_email_html(html, subject="") -> list[dict] (row per site)
# ---------------------------------------------------------------------------
def parse_email_html(html: str, subject: str = ""):
    """
    Entry point utama. Return list of dict, satu dict per site,
    dengan keys match schema.py COLUMNS (kecuali Duration (Hour),
    tu formula, dikira semasa write ke excel).
    """
    soup = BeautifulSoup(html, "html.parser")

    inc_no, tt_status = extract_incident_header(soup)
    if not inc_no:
        inc_no = extract_inc_from_subject(subject)

    severity_header = extract_header_severity(soup)
    detail = extract_detail_fields(soup)

    svc_imp = detail.get("Outage Description")
    region = detail.get("Region")
    severity = detail.get("Severity") or severity_header
    outage_start_raw = detail.get("Outage Start Time")
    outage_end_raw = detail.get("Outage End Time")

    outage_start = parse_outage_datetime(outage_start_raw)
    outage_end = parse_outage_datetime(outage_end_raw)

    sites = extract_sites_table(soup)

    if not inc_no:
        # Tak dapat INC No langsung (body & subject dua-dua fail) —
        # tak boleh proses, return list kosong (row akan di-skip oleh caller)
        return []

    if not sites:
        # Takde sites table jumpa — masih return satu row dengan Site ID kosong
        # supaya info incident tak hilang terus (boleh review manual nanti)
        sites = [{"Site ID": "", "Site Name": "", "District": ""}]

    # Nilai mentah setiap field, keyed macam nama column dalam schema.py
    raw_values = {
        "Svc Imp": svc_imp,
        "Outage Start": outage_start,
        "Outage End": outage_end,
        "Severity": severity,
        "Region": region,
        "Status": tt_status,
    }

    rows = []
    for site in sites:
        row_values = {
            "INC No": inc_no,
            "Site ID": site["Site ID"],
            "Site Name": site["Site Name"],
            "District": site["District"],
            **raw_values,
        }
        # Susun ikut OUTPUT_COLUMNS (dari schema.py) supaya konsisten dengan
        # semua source parser lain — sesiapa baca schema.py, itulah source of truth
        ordered_row = {col: row_values.get(col) for col in OUTPUT_COLUMNS}
        rows.append(ordered_row)

    return rows


# ---------------------------------------------------------------------------
# Quick test bila run terus (bukan import)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    with open("test_sample_email.html", "r", encoding="utf-8") as f:
        html_content = f.read()

    subject = (
        "[MAJOR] Outage Closure: INC000102173805 - Service interruption of "
        "13 (2G)(4G) xD sites at certain parts of Sarikei / Kanowit 1, "
        "Sibu 6, Kapit area, Sarawak"
    )

    results = parse_email_html(html_content, subject=subject)

    print(f"Total rows extracted: {len(results)}\n")
    for r in results:
        print(r)
