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
DASHBOARD_URL = "https://compras20.streamlit.app/"

import os

DASHBOARD_URL = "https://compras20.streamlit.app/"

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
<div style="font-family: Arial; max-width: 800px; margin:auto;">

<h2 style="color:#1f2937;">📊 Compras 2.0 – Weekly Report</h2>

<div style="background:#0f172a; color:white; padding:15px; border-radius:8px;">
<b>Pipeline:</b> <span style="color:#fbbf24;">R$ {pipeline:,.0f}</span> |
<b>Realizado:</b> <span style="color:#22c55e;">R$ {realized:,.0f}</span>
</div>

<br>

<h3 style="color:#374151;">📈 Variação semanal</h3>

<ul>
<li>Pipeline: <b style="color:#fbbf24;">R$ {delta_pipeline:,.0f}</b></li>
<li>Realizado: <b style="color:#22c55e;">R$ {delta_realized:,.0f}</b></li>
</ul>

<hr>

<h3>🆕 Novas iniciativas</h3>
{new_items[['title']].to_html(index=False) if not new_items.empty else "<p>Nenhuma</p>"}

<h3>✅ Concluídas</h3>
{closed_items[['title']].to_html(index=False) if not closed_items.empty else "<p>Nenhuma</p>"}

<hr>

<div style="text-align:center; margin-top:20px;">
<a href="{DASHBOARD_URL}" 
style="background:#2563eb;color:white;padding:12px 20px;text-decoration:none;border-radius:6px;">
👉 Abrir Dashboard Completo
</a>
</div>

</div>
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
