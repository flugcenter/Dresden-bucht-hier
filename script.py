import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1ofCTU1sES9tMBjS-hj2ruNtxeudRHEP1/export?format=csv"

BLOCKED_STATUS_WORDS = ("storniert", "ausgebucht", "abgesagt")


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_label(text: str) -> str:
    text = clean_text(text).lower()
    text = text.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    text = " ".join(text.split())
    return text


def to_int_or_none(value):
    if pd.isna(value):
        return None

    text = str(value).strip()
    if not text:
        return None

    # deutsche/uneinheitliche Schreibweisen abfangen
    text = text.replace("\xa0", "").replace(" ", "")
    text = text.replace(".", "").replace(",", ".")

    try:
        return int(float(text))
    except Exception:
        return None


def parse_start_date(date_text: str):
    """
    Erwartete Formate z. B.:
    - 29.04.-06.05.26
    - 29.04. - 06.05.26
    - 29.04.–06.05.26

    Es wird nur das Startdatum links verwendet.
    """
    text = clean_text(date_text)
    if not text:
        return None

    text = text.replace("–", "-").replace("—", "-")

    try:
        parts = text.split("-")
        if not parts:
            return None

        start_part = parts[0].strip().rstrip(".")
        start_bits = [x for x in start_part.split(".") if x]

        if len(start_bits) < 2:
            return None

        day = int(start_bits[0])
        month = int(start_bits[1])

        digits = "".join(ch for ch in text if ch.isdigit())
        if len(digits) < 2:
            return None

        year_suffix = digits[-2:]
        year = 2000 + int(year_suffix)

        return datetime(year, month, day)
    except Exception:
        return None


def is_blocked_status(status_text: str) -> bool:
    status = clean_text(status_text).lower()
    if not status:
        return False
    return any(word in status for word in BLOCKED_STATUS_WORDS)


def find_row_index_by_label(raw, possible_labels):
    """
    Sucht in Spalte A (Spalte 0) nach einer Beschriftung.
    Gibt den Zeilenindex zurück oder None.
    """
    wanted = {normalize_label(label) for label in possible_labels}

    for row in range(raw.shape[0]):
        label = normalize_label(raw.iat[row, 0])
        if label in wanted:
            return row

    return None


def find_last_nonempty_row_for_column(raw, col):
    """
    Letzte inhaltlich gefüllte Zeile einer Reisespalte.
    Nützlich für Status in der letzten Zeile.
    """
    for row in range(raw.shape[0] - 1, -1, -1):
        if clean_text(raw.iat[row, col]) != "":
            return row
    return None


def main():
    raw = pd.read_csv(SHEET_CSV_URL, header=None)

    # Feste Kopfzeilen laut deiner Struktur
    ROW_DESTINATION = 0      # 1. Zeile im Sheet
    ROW_DATE = 1             # 2. Zeile im Sheet
    ROW_RESPONSIBLE = 2      # 3. Zeile im Sheet

    FIRST_DATA_COL = 1       # Spalte A = Bezeichnungen, ab Spalte B beginnen die Reisen

    # Dynamisch per Beschriftung suchen
    row_booked = find_row_index_by_label(raw, [
        "gebuchte teilnehmer",
        "gebuchte tn",
        "teilnehmer gebucht",
        "gebucht"
    ])

    row_max = find_row_index_by_label(raw, [
        "max-tn",
        "max tn",
        "max. tn",
        "maximalteilnehmer",
        "maximal-teilnehmer",
        "maximale teilnehmerzahl",
        "teilnehmer max"
    ])

    if row_booked is None:
        raise ValueError("Zeile 'Gebuchte Teilnehmer' wurde in Spalte A nicht gefunden.")

    if row_max is None:
        raise ValueError("Zeile 'Max-TN' / 'Maximalteilnehmer' wurde in Spalte A nicht gefunden.")

    today = datetime.today().date()
    cutoff = today + timedelta(days=7)

    rows = []

    for col in range(FIRST_DATA_COL, raw.shape[1]):
        destination = clean_text(raw.iat[ROW_DESTINATION, col])
        date_text = clean_text(raw.iat[ROW_DATE, col])
        responsible = clean_text(raw.iat[ROW_RESPONSIBLE, col])

        if not destination:
            continue

        start_date = parse_start_date(date_text)
        if start_date is None:
            continue

        # Filter: Startdatum > heute + 7 Tage
        if start_date.date() <= cutoff:
            continue

        booked = to_int_or_none(raw.iat[row_booked, col])
        max_tn = to_int_or_none(raw.iat[row_max, col])

        # Status aus der letzten gefüllten Zeile der jeweiligen Reisespalte
        last_row = find_last_nonempty_row_for_column(raw, col)
        status = ""
        if last_row is not None:
            status = clean_text(raw.iat[last_row, col])

        # Falls in der letzten Zeile ein Status steht und geblockt ist -> raus
        if is_blocked_status(status):
            continue

        # Freie Plätze nur berechnen, wenn Max-TN vorhanden
        if max_tn is None:
            free_places = "auf Anfrage"
        else:
            if booked is None:
                free_places = max_tn
            else:
                free_places = max_tn - booked

        rows.append({
            "Reiseziel": destination,
            "Reisedatum": date_text,
            "Verantwortlich": responsible,
            "Gebuchte TN": "" if booked is None else booked,
            "Max-TN": "" if max_tn is None else max_tn,
            "Freie Plätze": free_places,
        })

    df = pd.DataFrame(rows)

    if not df.empty:
        df["_sort_date"] = df["Reisedatum"].apply(parse_start_date)
        df = df.sort_values("_sort_date", kind="stable").drop(columns=["_sort_date"])

    generated_at = datetime.now().strftime("%d.%m.%Y %H:%M")

    if df.empty:
        table_html = "<p>Zurzeit keine passenden Reisen vorhanden.</p>"
    else:
        table_html = df.to_html(index=False, border=0, classes="reise-table")

    html = f"""<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Dresden bucht hier – Aktuelle Reisen</title>
  <style>
    body {{
      font-family: Arial, Helvetica, sans-serif;
      margin: 24px;
      line-height: 1.4;
      color: #222;
      background: #fff;
    }}
    h1 {{
      margin-bottom: 8px;
    }}
    .meta {{
      color: #666;
      margin-bottom: 18px;
    }}
    table.reise-table {{
      border-collapse: collapse;
      width: 100%;
      max-width: 1200px;
    }}
    .reise-table th, .reise-table td {{
      border: 1px solid #ccc;
      padding: 8px 10px;
      text-align: left;
      vertical-align: top;
    }}
    .reise-table th {{
      background: #f3f3f3;
    }}
    .reise-table tr:nth-child(even) {{
      background: #fafafa;
    }}
    .hint {{
      margin-top: 18px;
      color: #666;
      font-size: 14px;
    }}
  </style>
</head>
<body>
  <h1>Aktuelle Reisen</h1>
  <div class="meta">Automatisch aktualisiert: {generated_at}</div>
  {table_html}
  <div class="hint">
    Filter: Startdatum &gt; heute + 7 Tage, Status in der letzten gefüllten Zeile nicht
    storniert/ausgebucht/abgesagt, freie Plätze nur bei vorhandener Max-TN,
    Verantwortliche aus Zeile 3.
  </div>
</body>
</html>
"""

    Path("index.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
