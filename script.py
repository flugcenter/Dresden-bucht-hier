import json
import hashlib
import re
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd


SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1ofCTU1sES9tMBjS-hj2ruNtxeudRHEP1/export?format=csv"

BLOCKED_STATUS_WORDS = ("storniert", "ausgebucht", "abgesagt")

CARD_COLORS = [
    ("#0f5fa8", "#edf6ff"),
    ("#8a3d72", "#fff1f8"),
    ("#2f7d4a", "#eefaf2"),
    ("#a45b12", "#fff6e9"),
    ("#6350a3", "#f5f1ff"),
    ("#147a82", "#edfbfc"),
    ("#9a3f3f", "#fff1f1"),
]

COUNTRY_CODE_RULES = [
    (("menorca", "mallorca", "spanien", "andalus", "kanaren", "teneriffa", "gran canaria", "ibiza"), "es"),
    (("verona", "apulien", "italien", "sizilien", "sardinien", "toscana", "rom", "venedig", "kalabrien"), "it"),
    (("jersey",), "je"),
    (("indien",), "in"),
    (("island",), "is"),
    (("norwegen", "fjord"), "no"),
    (("schweden",), "se"),
    (("finnland",), "fi"),
    (("daenemark", "dänemark"), "dk"),
    (("portugal", "madeira", "azoren"), "pt"),
    (("frankreich", "provence", "normandie", "bretagne", "paris", "korsika"), "fr"),
    (("griechenland", "kreta", "rhodos", "korfu"), "gr"),
    (("tuerkei", "türkei", "kappadokien"), "tr"),
    (("kroatien",), "hr"),
    (("slowenien",), "si"),
    (("albanien",), "al"),
    (("montenegro",), "me"),
    (("bosnien",), "ba"),
    (("serbien",), "rs"),
    (("nordmazedonien", "mazedonien"), "mk"),
    (("bulgarien",), "bg"),
    (("rumaenien", "rumänien"), "ro"),
    (("ungarn", "budapest"), "hu"),
    (("tschechien", "prag", "boehmen", "böhmen"), "cz"),
    (("polen", "breslau", "warschau", "krakau"), "pl"),
    (("oesterreich", "österreich", "wien", "tirol"), "at"),
    (("schweiz",), "ch"),
    (("niederlande", "holland", "amsterdam"), "nl"),
    (("belgien", "bruessel", "brüssel"), "be"),
    (("irland",), "ie"),
    (("schottland", "england", "grossbritannien", "großbritannien", "london", "wales"), "gb"),
    (("marokko",), "ma"),
    (("aegypten", "ägypten", "kairo"), "eg"),
    (("tunesien",), "tn"),
    (("suedafrika", "südafrika"), "za"),
    (("namibia",), "na"),
    (("kenia",), "ke"),
    (("tansania", "sansibar"), "tz"),
    (("japan",), "jp"),
    (("china",), "cn"),
    (("vietnam",), "vn"),
    (("thailand",), "th"),
    (("sri lanka",), "lk"),
    (("nepal",), "np"),
    (("indonesien", "bali"), "id"),
    (("malaysia",), "my"),
    (("singapur",), "sg"),
    (("kanada",), "ca"),
    (("usa", "vereinigte staaten", "kalifornien", "new york", "florida"), "us"),
    (("mexiko",), "mx"),
    (("kuba",), "cu"),
    (("brasilien",), "br"),
    (("argentinien",), "ar"),
    (("chile",), "cl"),
    (("peru",), "pe"),
    (("australien",), "au"),
    (("neuseeland",), "nz"),
]


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


def trip_duration_days(text):
    text = clean(text)
    if not text:
        return None

    text = text.replace("–", "-").replace("—", "-")

    try:
        if "-" not in text:
            return None

        start_text, end_text = text.split("-", 1)
        start_numbers = [int(x) for x in re.findall(r"\d+", start_text)]
        end_numbers = [int(x) for x in re.findall(r"\d+", end_text)]

        if len(start_numbers) < 1 or len(end_numbers) < 2:
            return None

        end_day = end_numbers[0]
        end_month = end_numbers[1]

        if len(end_numbers) >= 3:
            end_year = end_numbers[2]
            if end_year < 100:
                end_year += 2000
        else:
            digits = "".join(ch for ch in text if ch.isdigit())
            if len(digits) < 2:
                return None
            end_year = 2000 + int(digits[-2:])

        start_day = start_numbers[0]
        start_month = start_numbers[1] if len(start_numbers) >= 2 else end_month

        if len(start_numbers) >= 3:
            start_year = start_numbers[2]
            if start_year < 100:
                start_year += 2000
        else:
            start_year = end_year

        start_date = datetime(start_year, start_month, start_day)
        end_date = datetime(end_year, end_month, end_day)

        if end_date < start_date and len(start_numbers) < 3:
            start_date = datetime(end_year - 1, start_month, start_day)

        days = (end_date.date() - start_date.date()).days + 1
        return days if days > 0 else None

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


def color_for_title(title):
    digest = hashlib.md5(title.encode("utf-8")).hexdigest()
    index = int(digest[:8], 16) % len(CARD_COLORS)
    return CARD_COLORS[index]


def country_code_for_title(title):
    normalized = normalize_label(title)
    for keywords, code in COUNTRY_CODE_RULES:
        for keyword in keywords:
            if normalize_label(keyword) in normalized:
                return code
    return None


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

    row_min = find_row(
        raw,
        ["mindestteilnehmer", "mindest teilnehmer", "mindest-tn", "mindest tn"]
    )

    row_max = find_row(
        raw,
        ["max-tn", "max tn", "maximalteilnehmer", "maximal teilnehmer"]
    )

    if row_booked is None:
        row_booked = 32
    if row_min is None:
        row_min = 33
    if row_max is None:
        row_max = 34

    today = datetime.now(ZoneInfo("Europe/Berlin")).date()
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
        min_tn = to_int(raw.iat[row_min, col])
        max_tn = to_int(raw.iat[row_max, col])

        if min_tn is None and booked is not None and booked > 0:
            min_tn = 1

        last = last_filled_row(raw, col)
        status = clean(raw.iat[last, col]) if last is not None else ""

        if is_blocked(status):
            continue

        if max_tn is None:
            free_value = "auf Anfrage"
        else:
            free_value = max_tn if booked is None else max_tn - booked

        accent, tint = color_for_title(title)

        data.append({
            "titel": title,
            "termin": date_text,
            "dauer": trip_duration_days(date_text),
            "reisebuero": responsible,
            "gebucht": booked if booked is not None else 0,
            "max_tn": max_tn,
            "frei": free_value,
            "country_code": country_code_for_title(title),
            "accent": accent,
            "tint": tint,
            "sort_date": start.strftime("%Y-%m-%d")
        })

    data.sort(key=lambda x: x["sort_date"])

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

    now = datetime.now(ZoneInfo("Europe/Berlin")).strftime("%d.%m.%Y %H:%M")

    html = f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dresden bucht hier – Aktuelle Reisen</title>
<style>
* {{ box-sizing: border-box; }}
body {{
    font-family: Arial, Helvetica, sans-serif;
    margin: 0;
    background: #f4f6f8;
    color: #1f2937;
}}
.header {{
    background: linear-gradient(135deg, #075895, #0a6da9 55%, #0a4879);
    color: white;
    padding: 18px 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.12);
}}
.header-inner {{
    max-width: 1120px;
    margin: auto;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 16px;
}}
.header h1 {{
    margin: 0;
    font-size: 28px;
    line-height: 1.1;
}}
.subtitle {{
    margin-top: 4px;
    font-size: 14px;
    opacity: 0.92;
}}
.status {{
    margin-top: 5px;
    font-size: 12px;
    opacity: 0.78;
}}
.info-button {{
    background: white;
    color: #075895;
    text-decoration: none;
    padding: 10px 15px;
    border-radius: 9px;
    font-size: 14px;
    font-weight: 700;
    white-space: nowrap;
    box-shadow: 0 1px 4px rgba(0,0,0,0.1);
}}
.info-button:hover {{ background: #eef7ff; }}
.container {{
    max-width: 1120px;
    margin: 16px auto 0;
    padding: 0 12px;
}}
.card {{
    display: grid;
    grid-template-columns: minmax(0, 1fr) auto;
    align-items: center;
    gap: 16px;
    background: white;
    border-radius: 11px;
    padding: 11px 14px;
    margin-bottom: 9px;
    box-shadow: 0 1px 5px rgba(0,0,0,0.07);
    border-left: 5px solid var(--accent);
}}
.card-main {{ min-width: 0; }}
.title-row {{
    display: flex;
    align-items: center;
    gap: 9px;
}}
.country-flag {{
    width: 28px;
    height: 19px;
    object-fit: cover;
    border-radius: 3px;
    box-shadow: 0 0 0 1px rgba(0,0,0,0.12);
    flex: 0 0 auto;
}}
.title {{
    color: var(--accent);
    font-size: 18px;
    line-height: 1.2;
    font-weight: 800;
}}
.details {{
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 5px 14px;
    margin-top: 6px;
    color: #5f6874;
    font-size: 13px;
}}
.detail strong {{ color: #303846; }}
.booking-box {{
    display: flex;
    align-items: stretch;
    background: var(--tint);
    border-radius: 9px;
    overflow: hidden;
    min-width: 250px;
}}
.booking-item {{
    min-width: 118px;
    padding: 8px 12px;
    text-align: center;
}}
.booking-item + .booking-item {{
    border-left: 1px solid rgba(31,41,55,0.12);
}}
.booking-number {{
    color: var(--accent);
    font-size: 20px;
    line-height: 1.1;
    font-weight: 800;
}}
.booking-label {{
    color: #596273;
    font-size: 11px;
    margin-top: 2px;
}}
.free-ok {{ color: #137333; }}
.free-low {{ color: #b26a00; }}
.free-full {{ color: #b00020; }}
.footer {{
    max-width: 1120px;
    margin: 20px auto 22px;
    padding: 0 12px;
    display: flex;
    justify-content: space-between;
    gap: 14px;
    flex-wrap: wrap;
    color: #6b7280;
    font-size: 12px;
}}
.footer a {{
    color: #075895;
    text-decoration: none;
}}
.footer a:hover {{ text-decoration: underline; }}
@media (max-width: 760px) {{
    .header-inner {{
        align-items: flex-start;
        flex-direction: column;
    }}
    .header h1 {{ font-size: 24px; }}
    .card {{
        grid-template-columns: 1fr;
        gap: 9px;
    }}
    .booking-box {{
        min-width: 0;
        width: 100%;
    }}
    .booking-item {{
        flex: 1;
        min-width: 0;
    }}
}}
</style>
</head>
<body>
<div class="header">
    <div class="header-inner">
        <div>
            <h1>Dresden bucht hier</h1>
            <div class="subtitle">Interne Mitarbeiterübersicht</div>
            <div class="status">Stand: {now}</div>
        </div>
        <a class="info-button"
           href="https://www.dresden-bucht-hier.de/"
           target="_blank"
           rel="noopener noreferrer">
            Weitere Informationen
        </a>
    </div>
</div>
<div class="container">
"""

    if not data:
        html += "<p>Zurzeit keine passenden Reisen vorhanden.</p>"
    else:
        for item in data:
            cls = free_class(item["frei"])
            duration_text = f" · {item['dauer']} Tage" if item["dauer"] is not None else ""
            flag_html = ""
            if item["country_code"]:
                code = item["country_code"]
                flag_html = f'<img class="country-flag" src="https://flagcdn.com/w40/{code}.png" alt="Flagge" loading="lazy">'

            html += f"""
<div class="card" style="--accent:{item['accent']}; --tint:{item['tint']};">
    <div class="card-main">
        <div class="title-row">
            {flag_html}
            <div class="title">{item['titel']}</div>
        </div>
        <div class="details">
            <span class="detail">📅 {item['termin']}{duration_text}</span>
            <span class="detail"><strong>Reisebüro:</strong> {item['reisebuero']}</span>
        </div>
    </div>
    <div class="booking-box">
        <div class="booking-item">
            <div class="booking-number">{item['gebucht']}</div>
            <div class="booking-label">Plätze gebucht</div>
        </div>
"""

            if item["max_tn"] is not None:
                html += f"""
        <div class="booking-item">
            <div class="booking-number {cls}">{item['frei']}</div>
            <div class="booking-label">Noch frei</div>
        </div>
"""

            html += """
    </div>
</div>
"""

    html += f"""
</div>
<div class="footer">
    <div>
        <strong>Dresdner Reisebüros e.V.</strong>
        &nbsp;·&nbsp;
        <a href="https://www.dresden-bucht-hier.de/" target="_blank" rel="noopener noreferrer">www.dresden-bucht-hier.de</a>
        &nbsp;·&nbsp;
        <a href="https://www.dresden-bucht-hier.de/#impressum" target="_blank" rel="noopener noreferrer">Impressum</a>
    </div>
    <div>Stand: {now}</div>
</div>
</body>
</html>
"""

    Path("index.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
