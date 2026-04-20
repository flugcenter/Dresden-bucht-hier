import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1ofCTU1sES9tMBjS-hj2ruNtxeudRHEP1/export?format=csv"

BLOCKED_STATUS_WORDS = ("storniert", "ausgebucht", "abgesagt")


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def norm(t):
    t = clean(t).lower()
    t = t.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    return " ".join(t.split())


def to_int(v):
    t = clean(v)
    if not t:
        return None
    t = t.replace(".", "").replace(",", ".")
    try:
        return int(float(t))
    except:
        return None


def parse_date(t):
    t = clean(t)
    if not t:
        return None
    t = t.replace("–", "-")
    try:
        start = t.split("-")[0].strip()
        d, m = start.split(".")[:2]
        digits = "".join(ch for ch in t if ch.isdigit())
        year = 2000 + int(digits[-2:])
        return datetime(year, int(m), int(d))
    except:
        return None


def is_blocked(s):
    s = clean(s).lower()
    return any(w in s for w in BLOCKED_STATUS_WORDS)


def find_row(raw, words):
    wanted = {norm(x) for x in words}
    for i in range(raw.shape[0]):
        if norm(raw.iat[i, 0]) in wanted:
            return i
    return None


def last_row(raw, col):
    for i in range(raw.shape[0]-1, -1, -1):
        if clean(raw.iat[i, col]):
            return i
    return None


def main():
    raw = pd.read_csv(SHEET_CSV_URL, header=None)

    ROW_DEST = 0
    ROW_DATE = 1
    ROW_RESP = 2
    FIRST_COL = 1

    # 🔴 dynamisch suchen
    row_booked = find_row(raw, ["gebuchte teilnehmer", "gebuchte tn", "gebucht"])
    row_max = find_row(raw, ["max-tn", "max tn", "maximalteilnehmer"])

    # 🔴 FALLBACK falls nicht gefunden
    if row_booked is None:
        row_booked = 32   # entspricht ungefähr deiner bisherigen Zeile 33
    if row_max is None:
        row_max = 34

    today = datetime.today().date()
    cutoff = today + timedelta(days=7)

    data = []

    for col in range(FIRST_COL, raw.shape[1]):
        ziel = clean(raw.iat[ROW_DEST, col])
        datum = clean(raw.iat[ROW_DATE, col])
        resp = clean(raw.iat[ROW_RESP, col])

        if not ziel:
            continue

        start = parse_date(datum)
        if not start or start.date() <= cutoff:
            continue

        booked = to_int(raw.iat[row_booked, col])
        max_tn = to_int(raw.iat[row_max, col])

        last = last_row(raw, col)
        status = clean(raw.iat[last, col]) if last else ""

        if is_blocked(status):
            continue

        if max_tn is None:
            frei = "auf Anfrage"
        else:
            frei = max_tn if booked is None else max_tn - booked

        data.append({
            "ziel": ziel,
            "datum": datum,
            "resp": resp,
            "booked": "" if booked is None else booked,
            "max": "" if max_tn is None else max_tn,
            "frei": frei
        })

    df = pd.DataFrame(data)

    if not df.empty:
        df["_d"] = df["datum"].apply(parse_date)
        df = df.sort_values("_d").drop(columns=["_d"])

    now = datetime.now().strftime("%d.%m.%Y %H:%M")

    html = f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dresden bucht hier – Reisen</title>

<style>
body {{
    font-family: Arial;
    margin: 0;
    background: #f5f5f5;
}}

.header {{
    background: white;
    padding: 15px;
    border-bottom: 1px solid #ddd;
}}

.container {{
    max-width: 1000px;
    margin: 20px auto;
    padding: 10px;
}}

.card {{
    background: white;
    padding: 12px;
    margin-bottom: 10px;
    border-radius: 8px;
}}

.title {{
    font-weight: bold;
}}

.free-ok {{ color: green; }}
.free-low {{ color: orange; }}
.free-full {{ color: red; }}

.footer {{
    text-align: center;
    margin: 30px;
    color: #666;
}}
</style>
</head>

<body>

<div class="header">
<h2>Aktuelle Reisen</h2>
<div>Stand: {now}</div>
</div>

<div class="container">
"""

    if df.empty:
        html += "<p>Keine passenden Reisen</p>"
    else:
        for _, r in df.iterrows():
            frei = r["frei"]

            cls = "free-ok"
            if isinstance(frei, int):
                if frei <= 3:
                    cls = "free-low"
                if frei <= 0:
                    cls = "free-full"

            html += f"""
<div class="card">
<div class="title">{r['ziel']}</div>
<div>{r['datum']} | {r['resp']}</div>
<div>Gebucht: {r['booked']} | Max: {r['max']}</div>
<div class="{cls}">Freie Plätze: {frei}</div>
</div>
"""

    html += """
</div>

<div class="footer">
<strong>Dresdner Reisebüros e.V.</strong><br>
<a href="https://www.dresden-bucht-hier.de/#impressum" target="_blank">
Impressum
</a>
</div>

</body>
</html>
"""

    Path("index.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
