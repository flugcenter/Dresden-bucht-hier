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
