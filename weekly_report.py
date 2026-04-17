import json
import os
from datetime import datetime
import smtplib

import pandas as pd
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from playwright.sync_api import sync_playwright

# ========================
# CONFIG
# ========================
DASHBOARD_URL = "https://compras20.streamlit.app/"

EMAIL_FROM = "davi.spinconsulting@gmail.com"
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

SNAPSHOT_DIR = "snapshots"
LATEST_SNAPSHOT_FILE = os.path.join(SNAPSHOT_DIR, "latest_snapshot.csv")


# ========================
# HELPERS
# ========================
def brl(value: float) -> str:
    return f"R$ {value:,.0f}"


def pct(value: float) -> str:
    return f"{value:.1%}"


def take_screenshot(output_path: str = "dashboard.png") -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1440, "height": 2200})
        page.goto(DASHBOARD_URL, wait_until="networkidle")
        page.wait_for_timeout(5000)
        page.screenshot(path=output_path, full_page=True)
        browser.close()


def load_recipients() -> list[str]:
    with open("email_recipients.json", "r", encoding="utf-8") as f:
        return json.load(f)["recipients"]


def load_current_data() -> pd.DataFrame:
    df = pd.read_csv("initiatives.csv")

    df["initiative_id"] = pd.to_numeric(df["initiative_id"], errors="coerce").fillna(0).astype(int)
    df["baseline_annual_spend"] = pd.to_numeric(df["baseline_annual_spend"], errors="coerce").fillna(0.0)
    df["expected_saving_pct"] = pd.to_numeric(df["expected_saving_pct"], errors="coerce").fillna(0.0)
    df["realized_saving_value"] = pd.to_numeric(df["realized_saving_value"], errors="coerce").fillna(0.0)

    df["expected_saving_value"] = df["baseline_annual_spend"] * df["expected_saving_pct"]

    for col in ["title", "stage", "owner", "category", "type", "negotiation_lever"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str)

    return df


def load_previous_snapshot() -> pd.DataFrame:
    if not os.path.exists(LATEST_SNAPSHOT_FILE):
        return pd.DataFrame()

    old = pd.read_csv(LATEST_SNAPSHOT_FILE)

    if "initiative_id" in old.columns:
        old["initiative_id"] = pd.to_numeric(old["initiative_id"], errors="coerce").fillna(0).astype(int)
    if "baseline_annual_spend" in old.columns:
        old["baseline_annual_spend"] = pd.to_numeric(old["baseline_annual_spend"], errors="coerce").fillna(0.0)
    if "expected_saving_pct" in old.columns:
        old["expected_saving_pct"] = pd.to_numeric(old["expected_saving_pct"], errors="coerce").fillna(0.0)
    if "realized_saving_value" in old.columns:
        old["realized_saving_value"] = pd.to_numeric(old["realized_saving_value"], errors="coerce").fillna(0.0)

    if "expected_saving_value" not in old.columns:
        old["expected_saving_value"] = old["baseline_annual_spend"] * old["expected_saving_pct"]

    for col in ["title", "stage"]:
        if col not in old.columns:
            old[col] = ""
        old[col] = old[col].fillna("").astype(str)

    return old


def save_snapshot(df: pd.DataFrame) -> None:
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)

    dated_file = os.path.join(
        SNAPSHOT_DIR,
        f"{datetime.now().strftime('%Y-%m-%d')}_snapshot.csv"
    )

    df.to_csv(LATEST_SNAPSHOT_FILE, index=False)
    df.to_csv(dated_file, index=False)


def build_weekly_changes(current: pd.DataFrame, previous: pd.DataFrame) -> dict:
    total_spend = current["baseline_annual_spend"].sum()
    pipeline = current["expected_saving_value"].sum()
    realized = current["realized_saving_value"].sum()

    expected_pct_total = pipeline / total_spend if total_spend > 0 else 0
    realized_pct_total = realized / total_spend if total_spend > 0 else 0
    capture_rate = realized / pipeline if pipeline > 0 else 0

    if previous.empty:
        delta_pipeline = 0.0
        delta_realized = 0.0
        new_items = current.head(0).copy()
        closed_this_week = current.head(0).copy()
    else:
        old_pipeline = previous["expected_saving_value"].sum()
        old_realized = previous["realized_saving_value"].sum()

        delta_pipeline = pipeline - old_pipeline
        delta_realized = realized - old_realized

        previous_ids = set(previous["initiative_id"].tolist())
        current_ids = set(current["initiative_id"].tolist())

        new_ids = current_ids - previous_ids
        new_items = current[current["initiative_id"].isin(new_ids)].copy()

        previous_stage = previous[["initiative_id", "stage"]].rename(columns={"stage": "old_stage"})
        merged = current.merge(previous_stage, on="initiative_id", how="left")

        closed_this_week = merged[
            (merged["stage"].str.lower() == "closed") &
            (merged["old_stage"].fillna("").str.lower() != "closed")
        ].copy()

    return {
        "total_spend": total_spend,
        "pipeline": pipeline,
        "realized": realized,
        "expected_pct_total": expected_pct_total,
        "realized_pct_total": realized_pct_total,
        "capture_rate": capture_rate,
        "delta_pipeline": delta_pipeline,
        "delta_realized": delta_realized,
        "new_items": new_items,
        "closed_this_week": closed_this_week,
    }


def dataframe_to_html_table(df: pd.DataFrame, columns: list[str], empty_message: str) -> str:
    if df.empty:
        return f"<p style='margin:0;color:#6b7280;'>{empty_message}</p>"

    temp = df[columns].copy()
    return temp.to_html(index=False, border=0)


def build_email_html(summary: dict) -> str:
    total_spend = summary["total_spend"]
    pipeline = summary["pipeline"]
    realized = summary["realized"]
    expected_pct_total = summary["expected_pct_total"]
    realized_pct_total = summary["realized_pct_total"]
    capture_rate = summary["capture_rate"]
    delta_pipeline = summary["delta_pipeline"]
    delta_realized = summary["delta_realized"]
    new_items = summary["new_items"]
    closed_this_week = summary["closed_this_week"]

    pipeline_color = "#fbbf24"
    realized_color = "#22c55e"
    delta_pipeline_color = "#f59e0b" if delta_pipeline >= 0 else "#ef4444"
    delta_realized_color = "#22c55e" if delta_realized >= 0 else "#ef4444"

    new_items_html = dataframe_to_html_table(
        new_items,
        ["title", "owner", "category", "expected_saving_value"],
        "Nenhuma nova iniciativa."
    )

    closed_items_html = dataframe_to_html_table(
        closed_this_week,
        ["title", "owner", "category", "expected_saving_value", "realized_saving_value"],
        "Nenhuma iniciativa concluída na semana."
    )

    html = f"""
    <html>
    <body style="font-family: Arial, Helvetica, sans-serif; background:#f8fafc; margin:0; padding:24px;">
      <div style="max-width:900px; margin:auto; background:white; border:1px solid #e5e7eb; border-radius:12px; overflow:hidden;">

        <div style="background:#0f172a; color:white; padding:20px 24px;">
          <h2 style="margin:0 0 8px 0;">📊 Compras 2.0 – Weekly Report</h2>
          <div style="font-size:16px; line-height:1.6;">
            O pipeline atual soma <b style="color:{pipeline_color};">{brl(pipeline)}</b>,
            com captura realizada de <b style="color:{realized_color};">{brl(realized)}</b>
            e realização de <b style="color:{pipeline_color};">{pct(capture_rate)}</b> do projetado.
          </div>
        </div>

        <div style="padding:20px 24px;">

          <div style="display:flex; gap:12px; margin-bottom:18px;">
            <div style="flex:1; background:#f8fafc; border:1px solid #e5e7eb; padding:14px; border-radius:10px;">
              <div style="font-size:12px; color:#6b7280;">Spend anual total</div>
              <div style="font-size:22px; font-weight:700;">{brl(total_spend)}</div>
            </div>

            <div style="flex:1; background:#f8fafc; border:1px solid #e5e7eb; padding:14px; border-radius:10px;">
              <div style="font-size:12px; color:#6b7280;">Saving esperado</div>
              <div style="font-size:22px; font-weight:700; color:{pipeline_color};">{brl(pipeline)}</div>
              <div style="font-size:13px; color:#6b7280;">{pct(expected_pct_total)} do spend</div>
            </div>

            <div style="flex:1; background:#f8fafc; border:1px solid #e5e7eb; padding:14px; border-radius:10px;">
              <div style="font-size:12px; color:#6b7280;">Saving realizado</div>
              <div style="font-size:22px; font-weight:700; color:{realized_color};">{brl(realized)}</div>
              <div style="font-size:13px; color:#6b7280;">{pct(realized_pct_total)} do spend</div>
            </div>
          </div>

          <h3 style="margin-bottom:8px; color:#111827;">📈 Variação semanal</h3>
          <div style="display:flex; gap:12px; margin-bottom:18px;">
            <div style="flex:1; background:#fff7ed; border:1px solid #fed7aa; padding:14px; border-radius:10px;">
              <div style="font-size:12px; color:#6b7280;">Variação Pipeline</div>
              <div style="font-size:22px; font-weight:700; color:{delta_pipeline_color};">{brl(delta_pipeline)}</div>
            </div>

            <div style="flex:1; background:#f0fdf4; border:1px solid #bbf7d0; padding:14px; border-radius:10px;">
              <div style="font-size:12px; color:#6b7280;">Variação Realizado</div>
              <div style="font-size:22px; font-weight:700; color:{delta_realized_color};">{brl(delta_realized)}</div>
            </div>
          </div>

          <h3 style="margin-bottom:8px; color:#111827;">🆕 Novas iniciativas</h3>
          <div style="margin-bottom:20px;">
            {new_items_html}
          </div>

          <h3 style="margin-bottom:8px; color:#111827;">✅ Iniciativas concluídas na semana</h3>
          <div style="margin-bottom:20px;">
            {closed_items_html}
          </div>

          <h3 style="margin-bottom:8px; color:#111827;">📸 Dashboard</h3>
          <div style="margin-bottom:20px;">
            <img src="cid:dashboard" style="width:100%; border:1px solid #e5e7eb; border-radius:10px;" />
          </div>

          <div style="text-align:center; margin-top:24px;">
            <a href="{DASHBOARD_URL}"
               style="background:#2563eb; color:white; padding:12px 22px; text-decoration:none; border-radius:8px; font-weight:700;">
              👉 Abrir Dashboard Completo
            </a>
          </div>

        </div>
      </div>
    </body>
    </html>
    """
    return html


def send_email(html: str, recipients: list[str], image_path: str = "dashboard.png") -> None:
    if not EMAIL_PASSWORD:
        raise ValueError("EMAIL_PASSWORD não foi encontrado. Verifique o GitHub Secret.")

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(EMAIL_FROM, EMAIL_PASSWORD)

    for recipient in recipients:
        msg = MIMEMultipart("related")
        msg["From"] = EMAIL_FROM
        msg["To"] = recipient
        msg["Subject"] = "Compras 2.0 - Weekly Report"

        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(html, "html", "utf-8"))
        msg.attach(alt)

        with open(image_path, "rb") as img:
            image = MIMEImage(img.read())
            image.add_header("Content-ID", "<dashboard>")
            image.add_header("Content-Disposition", "inline", filename="dashboard.png")
            msg.attach(image)

        server.sendmail(EMAIL_FROM, recipient, msg.as_string())

    server.quit()


# ========================
# MAIN
# ========================
def main():
    print("Loading current data...")
    current = load_current_data()

    print("Loading previous snapshot...")
    previous = load_previous_snapshot()

    print("Taking dashboard screenshot...")
    take_screenshot("dashboard.png")

    print("Building weekly changes...")
    summary = build_weekly_changes(current, previous)

    print("Building email html...")
    html = build_email_html(summary)

    print("Loading recipients...")
    recipients = load_recipients()

    print("Sending email...")
    send_email(html, recipients, "dashboard.png")

    print("Saving snapshot...")
    save_snapshot(current)

    print("Weekly report sent successfully.")


if __name__ == "__main__":
    main()
