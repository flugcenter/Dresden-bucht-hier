import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1ofCTU1sES9tMBjS-hj2ruNtxeudRHEP1/export?format=csv"

STATUS_BLOCKLIST = {"storniert", "ausgebucht", "abgesagt"}

def parse_date_range(date_text: str):
    """
    Erwartet z. B. '29.04.-06.05.26'
    Nimmt das Startdatum links vom Bindestrich.
    """
    if pd.isna(date_text):
        return None

    text = str(date_text).strip()
    if not text:
        return None

    try:
        start_part = text.split("-")[0].strip()
        # Format: TT.MM.
        day, month = start_part.split(".")[:2]

        # Jahr aus dem rechten Teil holen, z. B. ...05.26
        year_suffix = text[-2:]
        year = 2000 + int(year_suffix)

        return datetime(year, int(month), int(day))
    except Exception:
        return None

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def to_int_or_none(value):
    if pd.isna(value):
        return None
    text = str(value).strip().replace(",", ".")
    if text == "":
        return None
    try:
        return int(float(text))
    except Exception:
        return None

def main():
    raw = pd.read_csv(SHEET_CSV_URL, header=None)

    # Spaltenstruktur:
    # Zeile 0 = Reiseziel
    # Zeile 1 = Reisedatum
    # Zeile 2 = verantwortlich im Verein
    # Zeile 31 = gebuchte Teilnehmer
    # Zeile 33 = Maximalteilnehmer
    # Zeile 34 = Status / letzte inhaltliche Zeile
    #
    # Erste Spalte enthält nur Bezeichnungen und wird ignoriert.
    col_start = 1
    col_end = raw.shape[1]

    today = datetime.today().date()
    cutoff = today + timedelta(days=7)

    results = []

    for col in range(col_start, col_end):
        destination = clean_text(raw.iat[0, col])
        date_text = clean_text(raw.iat[1, col])
        responsible = clean_text(raw.iat[2, col])
        booked = to_int_or_none(raw.iat[31, col])
        max_tn = to_int_or_none(raw.iat[33, col])
        status = clean_text(raw.iat[34, col]).lower()

        if not destination:
            continue

        start_date = parse_date_range(date_text)
        if start_date is None:
            continue

        if start_date.date() <= cutoff:
            continue

        if status in STATUS_BLOCKLIST:
            continue

        if max_tn is None:
            free_places = "auf Anfrage"
        else:
            if booked is None:
                free_places = max_tn
            else:
                free_places = max_tn - booked

        results.append({
            "Reiseziel": destination,
            "Reisedatum": date_text,
            "Verantwortlich": responsible,
            "Gebuchte TN": "" if booked is None else booked,
            "Max-TN": "" if max_tn is None else max_tn,
            "Freie Plätze": free_places,
        })

    df = pd.DataFrame(results)

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
