import json
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1ofCTU1sES9tMBjS-hj2ruNtxeudRHEP1/export?format=csv"

BLOCKED_STATUS_WORDS = ("storniert", "ausgebucht", "abgesagt")


def clean(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_label(text):
    text = clean(text).lower()
    text = text.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    return " ".join(text.split())


def to_int(value):
    text = clean(value)

    if not text:
        return None

    text = text.replace("\xa0", "").replace(" ", "")
    text = text.replace(".", "").replace(",", ".")

    try:
        return int(float(text))
    except Exception:
        return None


def parse_date(text):
    text = clean(text)

    if not text:
        return None

    text = text.replace("–", "-").replace("—", "-")

    try:
        start = text.split("-")[0].strip().rstrip(".")
        parts = [p for p in start.split(".") if p]

        if len(parts) < 2:
            return None

        day = int(parts[0])
        month = int(parts[1])

        digits = "".join(ch for ch in text if ch.isdigit())

        if len(digits) < 2:
            return None

        year = 2000 + int(digits[-2:])

        return datetime(year, month, day)

    except Exception:
        return None


def is_blocked(status):
    status = clean(status).lower()

    if not status:
        return False

    return any(word in status for word in BLOCKED_STATUS_WORDS)


def find_row(raw, labels):
    wanted = {normalize_label(x) for x in labels}

    for i in range(raw.shape[0]):
        if normalize_label(raw.iat[i, 0]) in wanted:
            return i

    return None


def last_filled_row(raw, col):
    for i in range(raw.shape[0] - 1, -1, -1):
        if clean(raw.iat[i, col]):
            return i

    return None


def free_class(free_value):
    if isinstance(free_value, int):

        if free_value <= 0:
            return "free-full"

        if free_value <= 3:
            return "free-low"

    return "free-ok"


def main():

    raw = pd.read_csv(SHEET_CSV_URL, header=None)

    ROW_DEST = 0
    ROW_DATE = 1
    ROW_RESPONSIBLE = 2
    FIRST_COL = 1

    row_booked = find_row(
        raw,
        ["gebuchte teilnehmer", "gebuchte tn", "gebucht"]
    )

    row_max = find_row(
        raw,
        ["max-tn", "max tn", "maximalteilnehmer", "maximal teilnehmer"]
    )

    if row_booked is None:
        row_booked = 32

    if row_max is None:
        row_max = 34

    today = datetime.today().date()
    cutoff = today + timedelta(days=7)

    data = []

    for col in range(FIRST_COL, raw.shape[1]):

        title = clean(raw.iat[ROW_DEST, col])
        date_text = clean(raw.iat[ROW_DATE, col])
        responsible = clean(raw.iat[ROW_RESPONSIBLE, col])

        if not title:
            continue

        start = parse_date(date_text)

        if start is None:
            continue

        if start.date() <= cutoff:
            continue

        booked = to_int(raw.iat[row_booked, col])
        max_tn = to_int(raw.iat[row_max, col])

        last = last_filled_row(raw, col)
        status = clean(raw.iat[last, col]) if last is not None else ""

        if is_blocked(status):
            continue

        if max_tn is None:
            free_value = "auf Anfrage"
        else:
            free_value = max_tn if booked is None else max_tn - booked

        data.append({
            "titel": title,
            "termin": date_text,
            "reisebuero": responsible,
            "frei": free_value,
            "sort_date": start.strftime("%Y-%m-%d")
        })

    data.sort(key=lambda x: x["sort_date"])

    # Öffentliche JSON-Ausgabe
    json_output = [
        {
            "titel": item["titel"],
            "termin": item["termin"],
            "frei": item["frei"]
        }
        for item in data
    ]

    Path("reisen.json").write_text(
        json.dumps(json_output, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    now = datetime.now().strftime("%d.%m.%Y %H:%M")

    html = f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>Dresden bucht hier – Aktuelle Reisen</title>

<style>

body {{
    font-family: Arial, sans-serif;
    margin: 0;
    background: #f3f5f7;
    color: #222;
}}

.header {{
    background: linear-gradient(135deg, #005ea8, #003f70);
    color: white;
    padding: 28px 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.15);
}}

.header-inner {{
    max-width: 1050px;
    margin: auto;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 20px;
    flex-wrap: wrap;
}}

.header h1 {{
    margin: 0;
    font-size: 34px;
}}

.subtitle {{
    margin-top: 8px;
    font-size: 15px;
    opacity: 0.9;
}}

.info-button {{
    background: white;
    color: #005ea8;
    text-decoration: none;
    padding: 14px 20px;
    border-radius: 10px;
    font-weight: bold;
    transition: 0.2s ease;
    white-space: nowrap;
}}

.info-button:hover {{
    background: #eef5fb;
}}

.container {{
    max-width: 1050px;
    margin: 25px auto;
    padding: 0 14px;
}}

.card {{
    background: white;
    border-radius: 14px;
    padding: 18px 20px;
    margin-bottom: 16px;
    box-shadow: 0 3px 12px rgba(0,0,0,0.08);
    transition: 0.15s ease;
    border-left: 5px solid #005ea8;
}}

.card:hover {{
    transform: translateY(-2px);
}}

.title {{
    font-size: 22px;
    font-weight: 700;
    margin-bottom: 10px;
    color: #003b6f;
}}

.meta {{
    color: #555;
    font-size: 15px;
    margin-bottom: 6px;
}}

.free-ok {{
    color: #137333;
    font-weight: 700;
    margin-top: 10px;
}}

.free-low {{
    color: #b26a00;
    font-weight: 700;
    margin-top: 10px;
}}

.free-full {{
    color: #b00020;
    font-weight: 700;
    margin-top: 10px;
}}

.footer {{
    text-align: center;
    margin: 40px 0 25px;
    color: #666;
    font-size: 13px;
    line-height: 1.6;
}}

.footer a {{
    color: #005ea8;
    text-decoration: none;
}}

.footer a:hover {{
    text-decoration: underline;
}}

.status {{
    margin-top: 8px;
    font-size: 13px;
    opacity: 0.85;
}}

@media (max-width: 700px) {{

    .header h1 {{
        font-size: 28px;
    }}

    .header-inner {{
        flex-direction: column;
        align-items: flex-start;
    }}

    .card {{
        padding: 16px;
    }}

    .title {{
        font-size: 20px;
    }}

}}

</style>
</head>

<body>

<div class="header">

    <div class="header-inner">

        <div>
            <h1>Dresden bucht hier</h1>

            <div class="subtitle">
                Interne Mitarbeiterübersicht Studienreisen
            </div>

            <div class="status">
                Stand: {now}
            </div>
        </div>

        <a class="info-button"
           href="https://www.dresden-bucht-hier.de/"
           target="_blank">
           Weitere Informationen
        </a>

    </div>

</div>

<div class="container">
"""

    if not data:
        html += """
<p>Zurzeit keine passenden Reisen vorhanden.</p>
"""
    else:
        for item in data:

            cls = free_class(item["frei"])

            html += f"""
<div class="card">

  <div class="title">
    {item['titel']}
  </div>

  <div class="meta">
    {item['termin']}
  </div>

  <div class="meta">
    Reisebüro: {item['reisebuero']}
  </div>

  <div class="{cls}">
    Noch frei: {item['frei']}
  </div>

</div>
"""

    html += """
</div>

<div class="footer">

    <strong>Dresdner Reisebüros e.V.</strong><br>

    <a href="https://www.dresden-bucht-hier.de/" target="_blank">
        www.dresden-bucht-hier.de
    </a>

    <br><br>

    <a href="https://www.dresden-bucht-hier.de/#impressum" target="_blank">
        Impressum
    </a>

</div>

</body>
</html>
"""

    Path("index.html").write_text(
        html,
        encoding="utf-8"
    )


if __name__ == "__main__":
    main()
