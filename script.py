import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1ofCTU1sES9tMBjS-hj2ruNtxeudRHEP1/export?format=csv"

BLOCKED_STATUS_WORDS = ("storniert", "ausgebucht", "abgesagt")


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_int_or_none(value):
    if pd.isna(value):
        return None

    text = str(value).strip()
    if not text:
        return None

    text = text.replace(".", "").replace(",", ".")

    try:
        return int(float(text))
    except Exception:
        return None


def parse_start_date(date_text: str):
    """
    Erwartet typischerweise Formate wie:
    - 29.04.-06.05.26
    - 29.04. - 06.05.26
    - 29.04.–06.05.26

    Verwendet nur das Startdatum links und das Jahr aus dem gesamten Text.
    """
    if not date_text:
        return None

    text = str(date_text).strip()
    if not text:
        return None

    text = text.replace("–", "-").replace("—", "-")

    try:
        parts = text.split("-")
        if not parts:
            return None

        start_part = parts[0].strip().rstrip(".")
        start_bits = start_part.split(".")
        start_bits = [x for x in start_bits if x]

        if len(start_bits) < 2:
            return None

        day = int(start_bits[0])
        month = int(start_bits[1])

        year = None

        # Jahr aus dem Text holen, bevorzugt die letzten 2 Ziffern
        digits = "".join(ch for ch in text if ch.isdigit())
        if len(digits) >= 2:
            year_suffix = digits[-2:]
            year = 2000 + int(year_suffix)

        if year is None:
            return None

        return datetime(year, month, day)
    except Exception:
        return None


def is_blocked_status(status_text: str) -> bool:
    status = clean_text(status_text).lower()
    if not status:
        return False
    return any(word in status for word in BLOCKED_STATUS_WORDS)


def main():
    raw = pd.read_csv(SHEET_CSV_URL, header=None)

    # Struktur laut Tabelle:
    # Zeile 0  = Reiseziel
    # Zeile 1  = Reisedatum
    # Zeile 2  = Verantwortliche
    # Zeile 31 = Gebuchte TN
    # Zeile 33 = Max-TN
    # Zeile 34 = Status / letzte inhaltliche Zeile
    #
    # Spalte 0 enthält die Zeilenbezeichnungen und wird ignoriert.
    FIRST_DATA_COL = 1

    today = datetime.today().date()
    cutoff = today + timedelta(days=7)

    rows = []

    for col in range(FIRST_DATA_COL, raw.shape[1]):
        destination = clean_text(raw.iat[0, col])
        date_text = clean_text(raw.iat[1, col])
        responsible = clean_text(raw.iat[2, col])
        booked = to_int_or_none(raw.iat[31, col])
        max_tn = to_int_or_none(raw.iat[33, col])
        status = clean_text(raw.iat[34, col])

        if not destination:
            continue

        start_date = parse_start_date(date_text)
        if start_date is None:
            continue

        # Filter: Startdatum > heute + 7 Tage
        if start_date.date() <= cutoff:
            continue

        # Filter: Status in letzter Zeile darf NICHT storniert/ausgebucht/abgesagt enthalten
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
        # Sortierung nach Startdatum
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
    Filter: Startdatum &gt; heute + 7 Tage, Status nicht storniert/ausgebucht/abgesagt,
    freie Plätze nur bei vorhandener Max-TN, Verantwortliche aus Zeile 3.
  </div>
</body>
</html>
"""

    Path("index.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
