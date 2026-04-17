import pandas as pd
import json
import os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ========================
# CONFIG
# ========================
DASHBOARD_URL = "https://SEU-APP.streamlit.app"

import os

DASHBOARD_URL = "https://SEU-APP.streamlit.app"

import os
EMAIL_FROM = "davi.spinconsulting@gmail.com"
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# ========================
# LOAD DATA
# ========================
df = pd.read_csv("initiatives.csv")

df["baseline_annual_spend"] = pd.to_numeric(df["baseline_annual_spend"], errors="coerce").fillna(0)
df["expected_saving_pct"] = pd.to_numeric(df["expected_saving_pct"], errors="coerce").fillna(0)
df["realized_saving_value"] = pd.to_numeric(df["realized_saving_value"], errors="coerce").fillna(0)

df["expected_saving_value"] = df["baseline_annual_spend"] * df["expected_saving_pct"]

# ========================
# KPIs
# ========================
total_spend = df["baseline_annual_spend"].sum()
pipeline = df["expected_saving_value"].sum()
realized = df["realized_saving_value"].sum()

# ========================
# SNAPSHOT
# ========================
os.makedirs("snapshots", exist_ok=True)

snapshot_file = "snapshots/latest_snapshot.csv"

if os.path.exists(snapshot_file):
    old = pd.read_csv(snapshot_file)

    old_pipeline = old["expected_saving_value"].sum()
    old_realized = old["realized_saving_value"].sum()

    delta_pipeline = pipeline - old_pipeline
    delta_realized = realized - old_realized

    # novas iniciativas
    new_ids = set(df["initiative_id"]) - set(old["initiative_id"])
    new_items = df[df["initiative_id"].isin(new_ids)]

    # concluídas
    closed_items = df[df["stage"] == "closed"]

else:
    delta_pipeline = 0
    delta_realized = 0
    new_items = pd.DataFrame()
    closed_items = pd.DataFrame()

# salva snapshot novo
df.to_csv(snapshot_file, index=False)

# ========================
# EMAIL HTML
# ========================
html = f"""
<h2>Compras 2.0 – Weekly Report</h2>

<p><b>Spend:</b> R$ {total_spend:,.0f}</p>
<p><b>Pipeline:</b> R$ {pipeline:,.0f}</p>
<p><b>Realized:</b> R$ {realized:,.0f}</p>

<hr>

<h3>Variação semanal</h3>
<p>Pipeline: {delta_pipeline:,.0f}</p>
<p>Realized: {delta_realized:,.0f}</p>

<hr>

<h3>Novas iniciativas</h3>
{new_items[['title']].to_html(index=False) if not new_items.empty else 'Nenhuma'}

<h3>Iniciativas concluídas</h3>
{closed_items[['title']].to_html(index=False) if not closed_items.empty else 'Nenhuma'}

<hr>

<p><a href="{DASHBOARD_URL}">Abrir Dashboard</a></p>
"""

# ========================
# EMAIL SEND
# ========================
with open("email_recipients.json") as f:
    recipients = json.load(f)["recipients"]

msg = MIMEMultipart()
msg["From"] = EMAIL_FROM
msg["Subject"] = "Compras 2.0 - Weekly Report"

msg.attach(MIMEText(html, "html"))

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(EMAIL_FROM, EMAIL_PASSWORD)

for r in recipients:
    msg["To"] = r
    server.sendmail(EMAIL_FROM, r, msg.as_string())

server.quit()

print("Email enviado")
