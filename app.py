import re
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Compras 2.0", layout="wide")

CONFIDENCE_FACTORS = {
    "low": 0.50,
    "medium": 0.75,
    "high": 1.00,
}

TYPE_OPTIONS = ["Gross", "Net"]

LEVER_OPTIONS = [
    "Aggregate Cost",
    "Change Supplier",
    "Change Relationship",
    "Cost Analysis",
    "Change Process",
    "Change Specification",
    "Global Procurement",
    "Avoid Cost",
    "E-Business",
    "Low-cost Countries Sourcing",
]

STAGE_OPTIONS = [
    "idea",
    "analysis",
    "sourcing",
    "negotiation",
    "implementation",
    "closed",
]

CONFIDENCE_OPTIONS = ["low", "medium", "high"]

REQUIRED_COLUMNS = [
    "initiative_id",
    "title",
    "category",
    "type",
    "negotiation_lever",
    "owner",
    "stage",
    "start_date",
    "expected_end_date",
    "baseline_annual_spend",
    "expected_saving_pct",
    "confidence_level",
    "realized_saving_value",
]


# ========================
# Data loading / parsing
# ========================
def load_data():
    try:
        initiatives = pd.read_csv("initiatives.csv")
    except FileNotFoundError:
        initiatives = pd.DataFrame(columns=REQUIRED_COLUMNS)

    for col in REQUIRED_COLUMNS:
        if col not in initiatives.columns:
            initiatives[col] = ""

    return initiatives[REQUIRED_COLUMNS]


def read_uploaded_csv(uploaded_file):
    try:
        return pd.read_csv(uploaded_file, encoding="utf-8", sep=None, engine="python")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, encoding="latin1", sep=None, engine="python")


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_brl_number(value):
    if pd.isna(value):
        return 0.0

    s = str(value).strip()
    if s == "":
        return 0.0

    s = s.replace("R$", "").replace("r$", "").strip()
    s = s.replace(" ", "")
    s = s.replace(".", "")
    s = s.replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)

    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_pct(value):
    if pd.isna(value):
        return 0.0

    s = str(value).strip()
    if s == "":
        return 0.0

    s = s.replace("%", "").replace(" ", "")
    s = s.replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)

    try:
        num = float(s)
    except ValueError:
        return 0.0

    if num > 1:
        return num / 100

    return num


def map_confidence(value):
    s = normalize_text(value).lower()

    mapping = {
        "low": "low",
        "baixo": "low",
        "baixa": "low",
        "medium": "medium",
        "moderado": "medium",
        "médio": "medium",
        "medio": "medium",
        "high": "high",
        "alto": "high",
        "alta": "high",
        "complexo": "high",
    }

    return mapping.get(s, "medium")


def map_stage(value):
    s = normalize_text(value).lower()

    mapping = {
        "idea": "idea",
        "ideia": "idea",
        "a ser iniciado": "idea",
        "analysis": "analysis",
        "análise": "analysis",
        "analise": "analysis",
        "sourcing": "sourcing",
        "negotiation": "negotiation",
        "negociação": "negotiation",
        "negociacao": "negotiation",
        "implementation": "implementation",
        "implementação": "implementation",
        "implementacao": "implementation",
        "closed": "closed",
        "fechado": "closed",
        "em progresso": "implementation",
        "em andamento": "implementation",
    }

    return mapping.get(s, "idea")


def map_type(value):
    s = normalize_text(value).strip().lower()

    mapping = {
        "gross": "Gross",
        "net": "Net",
    }

    return mapping.get(s, "Gross")


def map_negotiation_lever(value):
    s = normalize_text(value).strip().lower()

    mapping = {
        "aggregate cost": "Aggregate Cost",
        "change supplier": "Change Supplier",
        "change relationship": "Change Relationship",
        "cost analysis": "Cost Analysis",
        "change process": "Change Process",
        "change specification": "Change Specification",
        "global procurement": "Global Procurement",
        "avoid cost": "Avoid Cost",
        "e-business": "E-Business",
        "ebusiness": "E-Business",
        "low-cost countries sourcing": "Low-cost Countries Sourcing",
        "low cost countries sourcing": "Low-cost Countries Sourcing",
    }

    return mapping.get(s, "Cost Analysis")


def normalize_uploaded_initiatives(df):
    result = df.copy()

    for col in REQUIRED_COLUMNS:
        if col not in result.columns:
            result[col] = ""

    result["initiative_id"] = pd.to_numeric(result["initiative_id"], errors="coerce")
    missing_id_mask = result["initiative_id"].isna()

    next_id = 1
    if (~missing_id_mask).any():
        next_id = int(result["initiative_id"].dropna().max()) + 1

    for idx in result.index[missing_id_mask]:
        result.at[idx, "initiative_id"] = next_id
        next_id += 1

    result["initiative_id"] = result["initiative_id"].astype(int)

    text_cols = ["title", "category", "owner", "start_date", "expected_end_date"]
    for col in text_cols:
        result[col] = result[col].apply(normalize_text)

    result["type"] = result["type"].apply(map_type)
    result["negotiation_lever"] = result["negotiation_lever"].apply(map_negotiation_lever)
    result["stage"] = result["stage"].apply(map_stage)
    result["confidence_level"] = result["confidence_level"].apply(map_confidence)

    result["baseline_annual_spend"] = result["baseline_annual_spend"].apply(parse_brl_number)
    result["expected_saving_pct"] = result["expected_saving_pct"].apply(parse_pct)
    result["realized_saving_value"] = result["realized_saving_value"].apply(parse_brl_number)

    return result[REQUIRED_COLUMNS]


def enrich_initiatives(df):
    if df.empty:
        return df

    result = df.copy()

    for col in REQUIRED_COLUMNS:
        if col not in result.columns:
            result[col] = ""

    result["baseline_annual_spend"] = pd.to_numeric(
        result["baseline_annual_spend"], errors="coerce"
    ).fillna(0.0)

    result["expected_saving_pct"] = pd.to_numeric(
        result["expected_saving_pct"], errors="coerce"
    ).fillna(0.0)

    result["realized_saving_value"] = pd.to_numeric(
        result["realized_saving_value"], errors="coerce"
    ).fillna(0.0)

    result["confidence_level"] = result["confidence_level"].astype(str).str.lower().str.strip()
    result["stage"] = result["stage"].astype(str).str.lower().str.strip()

    result["expected_saving_value"] = (
        result["baseline_annual_spend"] * result["expected_saving_pct"]
    )

    result["confidence_factor"] = (
        result["confidence_level"].map(CONFIDENCE_FACTORS).fillna(0.75)
    )

    result["weighted_saving_value"] = (
        result["expected_saving_value"] * result["confidence_factor"]
    )

    return result


# ========================
# Formatting / dashboard helpers
# ========================
def brl(value):
    return f"R$ {value:,.0f}"


def pct_display(value):
    return f"{value * 100:.1f}%"


def format_ratio(value):
    if pd.isna(value):
        return "-"
    return f"{value:.0%}"


def ratio_color(pct):
    if pd.isna(pct):
        return "#d9d9d9"
    if pct >= 1.0:
        return "#00b050"
    if pct >= 0.9:
        return "#92d050"
    if pct >= 0.7:
        return "#ffd966"
    return "#ff6b6b"


def build_summary(df, group_col, label_col=None, top_n=None):
    temp = (
        df.groupby(group_col, dropna=False, as_index=False)
        .agg(
            projected=("expected_saving_value", "sum"),
            realized=("realized_saving_value", "sum"),
        )
    )

    temp["ratio"] = temp.apply(
        lambda r: (r["realized"] / r["projected"]) if r["projected"] > 0 else 0,
        axis=1,
    )

    if label_col and label_col != group_col:
        temp = temp.rename(columns={group_col: label_col})

    temp = temp.sort_values("projected", ascending=False)

    if top_n is not None:
        temp = temp.head(top_n)

    return temp


def render_exec_table(df, first_col_name):
    rows_html = ""

    total_projected = df["projected"].sum()
    total_realized = df["realized"].sum()
    total_ratio = total_realized / total_projected if total_projected > 0 else 0

    for _, row in df.iterrows():
        label = row[first_col_name]
        projected = brl(row["projected"])
        realized = brl(row["realized"])
        ratio = format_ratio(row["ratio"])
        bg = ratio_color(row["ratio"])

        rows_html += (
            f"<tr>"
            f"<td style='padding:6px 8px;border:1px solid #444;'>{label}</td>"
            f"<td style='padding:6px 8px;border:1px solid #444;text-align:right;'>{projected}</td>"
            f"<td style='padding:6px 8px;border:1px solid #444;text-align:right;'>{realized}</td>"
            f"<td style='padding:6px 8px;border:1px solid #444;text-align:center;background:{bg};font-weight:700;color:#000;'>{ratio}</td>"
            f"</tr>"
        )

    rows_html += (
        f"<tr style='font-weight:700;background:#111;'>"
        f"<td style='padding:6px 8px;border:1px solid #444;'>TOTAL</td>"
        f"<td style='padding:6px 8px;border:1px solid #444;text-align:right;'>{brl(total_projected)}</td>"
        f"<td style='padding:6px 8px;border:1px solid #444;text-align:right;'>{brl(total_realized)}</td>"
        f"<td style='padding:6px 8px;border:1px solid #444;text-align:center;background:{ratio_color(total_ratio)};color:#000;'>{format_ratio(total_ratio)}</td>"
        f"</tr>"
    )

    html = (
        f"<table style='width:100%; border-collapse:collapse; font-size:13px;'>"
        f"<thead>"
        f"<tr style='background:#f1c232; color:#000; font-weight:700;'>"
        f"<th style='padding:8px;border:1px solid #444;text-align:left;'>{first_col_name}</th>"
        f"<th style='padding:8px;border:1px solid #444;text-align:right;'>Projetado</th>"
        f"<th style='padding:8px;border:1px solid #444;text-align:right;'>Realizado</th>"
        f"<th style='padding:8px;border:1px solid #444;text-align:center;'>%</th>"
        f"</tr>"
        f"</thead>"
        f"<tbody>{rows_html}</tbody>"
        f"</table>"
    )

    st.markdown(html, unsafe_allow_html=True)


# ========================
# Dictionary / export
# ========================
def build_dictionary_df():
    return pd.DataFrame([
        {
            "Campo": "initiative_id",
            "Descrição": "Identificador único da iniciativa",
            "Como preencher": "Número inteiro ou deixar vazio no upload",
            "Exemplo": "1, 2, 3",
            "Observação": "Gerado automaticamente se não informado"
        },
        {
            "Campo": "title",
            "Descrição": "Nome da iniciativa",
            "Como preencher": "Texto curto e objetivo",
            "Exemplo": "Renegociação frete Sul",
            "Observação": "Nome que aparecerá nos relatórios"
        },
        {
            "Campo": "category",
            "Descrição": "Categoria de gasto",
            "Como preencher": "Texto",
            "Exemplo": "Logistics, Packaging, Facilities",
            "Observação": "Usado em agrupamentos e análises"
        },
        {
            "Campo": "type",
            "Descrição": "Tipo de saving",
            "Como preencher": "Selecionar Gross ou Net",
            "Exemplo": "Gross",
            "Observação": "Gross = saving bruto; Net = saving líquido"
        },
        {
            "Campo": "negotiation_lever",
            "Descrição": "Alavanca principal de negociação ou sourcing associada à iniciativa",
            "Como preencher": "Selecionar uma das 10 alavancas padrão",
            "Exemplo": "Cost Analysis",
            "Observação": "Usado para classificar a lógica principal da iniciativa"
        },
        {
            "Campo": "owner",
            "Descrição": "Responsável pela iniciativa",
            "Como preencher": "Nome do comprador ou líder",
            "Exemplo": "João Silva",
            "Observação": ""
        },
        {
            "Campo": "stage",
            "Descrição": "Estágio atual da iniciativa",
            "Como preencher": "Selecionar da lista padrão",
            "Exemplo": "idea, analysis, implementation, closed",
            "Observação": "Usado para leitura de pipeline e status"
        },
        {
            "Campo": "start_date",
            "Descrição": "Data de início da iniciativa",
            "Como preencher": "Texto livre no padrão desejado",
            "Exemplo": "jan/25",
            "Observação": "Pode ser refinado para data estruturada depois"
        },
        {
            "Campo": "expected_end_date",
            "Descrição": "Data esperada de conclusão",
            "Como preencher": "Texto livre no padrão desejado",
            "Exemplo": "abr/25",
            "Observação": "Usado para acompanhamento"
        },
        {
            "Campo": "baseline_annual_spend",
            "Descrição": "Gasto anual base da iniciativa",
            "Como preencher": "Valor numérico em R$",
            "Exemplo": "1000000",
            "Observação": "Base para cálculo do saving esperado"
        },
        {
            "Campo": "expected_saving_pct",
            "Descrição": "Percentual de saving esperado",
            "Como preencher": "Percentual humano: 5 = 5%",
            "Exemplo": "5, 8.5, 12",
            "Observação": "Não preencher em decimal e não usar valor em R$"
        },
        {
            "Campo": "confidence_level",
            "Descrição": "Nível de confiança da captura",
            "Como preencher": "low, medium ou high",
            "Exemplo": "medium",
            "Observação": "Usado para cálculo do weighted saving"
        },
        {
            "Campo": "realized_saving_value",
            "Descrição": "Saving já realizado pela iniciativa",
            "Como preencher": "Valor numérico em R$, pode ficar em branco/zero",
            "Exemplo": "25000",
            "Observação": "Pode ser zero enquanto a iniciativa não gerou captura"
        },
        {
            "Campo": "expected_saving_value",
            "Descrição": "Saving esperado em valor absoluto",
            "Como preencher": "Calculado automaticamente",
            "Exemplo": "50000",
            "Observação": "baseline_annual_spend × expected_saving_pct"
        },
        {
            "Campo": "weighted_saving_value",
            "Descrição": "Saving esperado ponderado por confiança",
            "Como preencher": "Calculado automaticamente",
            "Exemplo": "37500",
            "Observação": "expected_saving_value × fator de confiança"
        },
    ])


def dataframe_to_excel_bytes(df_dict):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output.getvalue()


def build_export_base(initiatives_df):
    export_df = initiatives_df.copy()

    for col in ["expected_saving_value", "confidence_factor", "weighted_saving_value"]:
        if col in export_df.columns:
            export_df = export_df.drop(columns=[col])

    if "expected_saving_pct" in export_df.columns:
        export_df["expected_saving_pct"] = pd.to_numeric(
            export_df["expected_saving_pct"], errors="coerce"
        ).fillna(0.0) * 100

    return export_df


# ========================
# App state
# ========================
df = enrich_initiatives(load_data())

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Dashboard",
    "Importar base",
    "Nova iniciativa",
    "Dicionário",
    "Editar",
    "Exportar",
])

# ========================
# Dashboard
# ========================
with tab1:
    st.title("Compras 2.0 – Executive Dashboard")
    st.caption("Resumo executivo de pipeline, captura e principais alavancas")

    if df.empty:
        st.warning("Nenhuma base de iniciativas carregada ainda.")
    else:
        total_spend = df["baseline_annual_spend"].sum()
        total_projected = df["expected_saving_value"].sum()
        total_weighted = df["weighted_saving_value"].sum()
        total_realized = df["realized_saving_value"].sum()

        capture_rate = total_realized / total_projected if total_projected > 0 else 0
        expected_pct_total = total_projected / total_spend if total_spend > 0 else 0
        realized_pct_total = total_realized / total_spend if total_spend > 0 else 0

        closed_projected = df.loc[df["stage"] == "closed", "expected_saving_value"].sum()
        open_projected = df.loc[df["stage"] != "closed", "expected_saving_value"].sum()

        st.markdown(
            f"""
            <div style="
                padding:14px;
                border-radius:10px;
                background:#0f172a;
                border:1px solid #334155;
                font-size:20px;
                font-weight:700;
                margin-bottom:12px;">
            O pipeline atual soma <span style="color:#f1c232;">{brl(total_projected)}</span>,
            com captura realizada de <span style="color:#00b050;">{brl(total_realized)}</span>
            e realização de <span style="color:#f1c232;">{format_ratio(capture_rate)}</span> do projetado.
            </div>
            """,
            unsafe_allow_html=True,
        )

        mini1, mini2, mini3 = st.columns(3)
        mini1.metric("Spend anual total", brl(total_spend))
        mini2.metric("Saving esperado", brl(total_projected), format_ratio(expected_pct_total))
        mini3.metric("Saving realizado", brl(total_realized), format_ratio(realized_pct_total))

        st.divider()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Projetado", brl(total_projected))
        c2.metric("Ponderado", brl(total_weighted))
        c3.metric("Pipeline Fechado", brl(closed_projected))
        c4.metric("Pipeline Aberto", brl(open_projected))

        if total_realized > 0 and closed_projected == 0:
            st.warning("Há saving realizado informado, mas nenhuma iniciativa está marcada como 'closed'.")

        st.markdown("### Resumo executivo")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            type_summary = build_summary(df, "type", label_col="Tipo")
            render_exec_table(type_summary.rename(columns={"Tipo": "Tipo"}), "Tipo")

        with col2:
            category_summary = build_summary(df, "category", label_col="Categoria", top_n=10)
            render_exec_table(category_summary.rename(columns={"Categoria": "Categoria"}), "Categoria")

        with col3:
            lever_summary = build_summary(df, "negotiation_lever", label_col="Alavanca", top_n=10)
            render_exec_table(lever_summary.rename(columns={"Alavanca": "Alavanca"}), "Alavanca")

        with col4:
            owner_summary = build_summary(df, "owner", label_col="Comprador", top_n=10)
            render_exec_table(owner_summary.rename(columns={"Comprador": "Comprador"}), "Comprador")

        st.markdown("### Top 15 iniciativas")

        top15 = (
            df[[
                "title",
                "expected_saving_value",
                "realized_saving_value",
                "type",
                "negotiation_lever",
                "owner",
                "stage",
            ]]
            .copy()
            .sort_values("expected_saving_value", ascending=False)
            .head(15)
        )

        top15["ratio"] = top15.apply(
            lambda r: (r["realized_saving_value"] / r["expected_saving_value"])
            if r["expected_saving_value"] > 0 else 0,
            axis=1,
        )

        top15_view = top15.copy()
        top15_view["expected_saving_value"] = top15_view["expected_saving_value"].apply(brl)
        top15_view["realized_saving_value"] = top15_view["realized_saving_value"].apply(brl)
        top15_view["ratio"] = top15_view["ratio"].apply(format_ratio)

        st.dataframe(
            top15_view.rename(columns={
                "title": "Iniciativa",
                "expected_saving_value": "Projetado",
                "realized_saving_value": "Realizado",
                "ratio": "%",
                "type": "Tipo",
                "negotiation_lever": "Alavanca",
                "owner": "Comprador",
                "stage": "Stage",
            }),
            use_container_width=True,
            hide_index=True,
        )

        with st.expander("Ver base detalhada completa"):
            detail = df.copy()
            detail["baseline_annual_spend"] = detail["baseline_annual_spend"].apply(brl)
            detail["expected_saving_pct"] = detail["expected_saving_pct"].apply(pct_display)
            detail["expected_saving_value"] = detail["expected_saving_value"].apply(brl)
            detail["weighted_saving_value"] = detail["weighted_saving_value"].apply(brl)
            detail["realized_saving_value"] = detail["realized_saving_value"].apply(brl)

            st.dataframe(
                detail.rename(columns={
                    "initiative_id": "ID",
                    "title": "Initiative",
                    "category": "Category",
                    "type": "Saving Type",
                    "negotiation_lever": "Negotiation Lever",
                    "owner": "Owner",
                    "stage": "Stage",
                    "start_date": "Start Date",
                    "expected_end_date": "Expected End Date",
                    "baseline_annual_spend": "Baseline Spend",
                    "expected_saving_pct": "Expected Saving %",
                    "confidence_level": "Confidence",
                    "expected_saving_value": "Expected Saving",
                    "weighted_saving_value": "Weighted Saving",
                    "realized_saving_value": "Realized Saving",
                }),
                use_container_width=True,
                hide_index=True,
            )

# ========================
# Importar base
# ========================
with tab2:
    st.header("Importar iniciativas")

    template = pd.DataFrame(
        {
            "initiative_id": [],
            "title": [],
            "category": [],
            "type": [],
            "negotiation_lever": [],
            "owner": [],
            "stage": [],
            "start_date": [],
            "expected_end_date": [],
            "baseline_annual_spend": [],
            "expected_saving_pct": [],
            "confidence_level": [],
            "realized_saving_value": [],
        }
    )

    st.download_button(
        label="Baixar template",
        data=template.to_csv(index=False),
        file_name="template_initiatives.csv",
        mime="text/csv",
        key="download_template",
    )

    uploaded = st.file_uploader(
        "Upload CSV",
        type=["csv"],
        key="upload_initiatives",
    )

    if uploaded is not None:
        try:
            uploaded_df = read_uploaded_csv(uploaded)

            st.write("Preview original:")
            st.dataframe(uploaded_df.head(), use_container_width=True)

            missing_cols = [col for col in REQUIRED_COLUMNS if col not in uploaded_df.columns]

            if missing_cols:
                st.error(f"Colunas faltando: {missing_cols}")
            else:
                cleaned_df = normalize_uploaded_initiatives(uploaded_df)

                preview_df = cleaned_df.copy()
                preview_df["baseline_annual_spend"] = preview_df["baseline_annual_spend"].apply(brl)
                preview_df["expected_saving_pct"] = preview_df["expected_saving_pct"].apply(pct_display)
                preview_df["realized_saving_value"] = preview_df["realized_saving_value"].apply(brl)

                st.write("Preview normalizado:")
                st.dataframe(preview_df.head(), use_container_width=True)

                if st.button("Importar base", key="btn_import_base"):
                    cleaned_df.to_csv("initiatives.csv", index=False)
                    st.success("Base importada com sucesso.")
                    st.info("Atualize a página para refletir os novos dados no dashboard.")

        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")

# ========================
# Nova iniciativa
# ========================
with tab3:
    st.header("Nova iniciativa")

    current_max_id = 0
    if not df.empty and "initiative_id" in df.columns:
        ids = pd.to_numeric(df["initiative_id"], errors="coerce").fillna(0)
        current_max_id = int(ids.max())

    with st.form("form_new_initiative"):
        title = st.text_input("Título")
        category = st.text_input("Categoria")
        type_ = st.selectbox("Tipo de saving", TYPE_OPTIONS)
        negotiation_lever = st.selectbox("Alavanca de negociação", LEVER_OPTIONS)
        owner = st.text_input("Responsável")
        stage = st.selectbox("Estágio", STAGE_OPTIONS)
        baseline = st.number_input("Baseline anual (R$)", min_value=0.0, value=0.0, step=1000.0)
        saving_pct_input = st.number_input(
            "Saving esperado (%)",
            min_value=0.0,
            max_value=100.0,
            value=0.0,
            step=0.5,
        )
        confidence = st.selectbox("Confiança", CONFIDENCE_OPTIONS)
        realized_saving_input = st.number_input(
            "Saving realizado (R$)",
            min_value=0.0,
            value=0.0,
            step=1000.0,
        )
        start_date = st.text_input("Start date")
        expected_end_date = st.text_input("Expected end date")

        submitted = st.form_submit_button("Salvar iniciativa")

        if submitted:
            base_to_save = load_data().copy()

            new_row = {
                "initiative_id": current_max_id + 1,
                "title": title,
                "category": category,
                "type": type_,
                "negotiation_lever": negotiation_lever,
                "owner": owner,
                "stage": stage,
                "start_date": start_date,
                "expected_end_date": expected_end_date,
                "baseline_annual_spend": baseline,
                "expected_saving_pct": saving_pct_input / 100,
                "confidence_level": confidence,
                "realized_saving_value": realized_saving_input,
            }

            base_to_save = pd.concat(
                [base_to_save, pd.DataFrame([new_row])],
                ignore_index=True,
            )

            base_to_save.to_csv("initiatives.csv", index=False)

            st.success("Iniciativa criada com sucesso.")
            st.info("Atualize a página para visualizar a nova iniciativa no dashboard.")

# ========================
# Dicionário
# ========================
with tab4:
    st.title("Dicionário de Dados – Compras 2.0")
    st.caption("Definição dos campos e lógica de cálculo utilizada no modelo")

    data_dict = build_dictionary_df()
    st.subheader("Definição dos campos")
    st.dataframe(data_dict, use_container_width=True, hide_index=True)

    st.subheader("Regras de cálculo")
    st.markdown("""
**Expected Saving (R$)**  
= Baseline Annual Spend × Expected Saving %

**Weighted Saving (R$)**  
= Expected Saving × Fator de Confiança

**Capture Rate**  
= Realized Saving Total ÷ Pipeline Total

**Fatores de confiança**
- low → 50%
- medium → 75%
- high → 100%
""")

    st.subheader("Alavancas padrão de negociação")
    lever_df = pd.DataFrame([
        {"Alavanca": "Aggregate Cost", "Definição": "Consolidar gastos e aumentar alavancagem comercial."},
        {"Alavanca": "Change Supplier", "Definição": "Recompor a carteira ou trocar fornecedores."},
        {"Alavanca": "Change Relationship", "Definição": "Redesenhar o modelo de relacionamento comercial."},
        {"Alavanca": "Cost Analysis", "Definição": "Analisar composição de custo, margem e drivers de preço."},
        {"Alavanca": "Change Process", "Definição": "Alterar processo para eliminar custo evitável."},
        {"Alavanca": "Change Specification", "Definição": "Revisar especificação técnica ou escopo."},
        {"Alavanca": "Global Procurement", "Definição": "Aproveitar escala global e harmonização."},
        {"Alavanca": "Avoid Cost", "Definição": "Evitar necessidade de gasto ou consumo."},
        {"Alavanca": "E-Business", "Definição": "Usar ferramentas digitais para ampliar competição e eficiência."},
        {"Alavanca": "Low-cost Countries Sourcing", "Definição": "Buscar fontes em países de menor custo."},
    ])
    st.dataframe(lever_df, use_container_width=True, hide_index=True)

    st.subheader("Boas práticas de preenchimento")
    st.markdown("""
- Usar `expected_saving_pct` como percentual humano: `5 = 5%`
- Não preencher `expected_saving_pct` com valores em R$
- Usar `type` apenas como classificação do saving: Gross ou Net
- Preencher `negotiation_lever` com a principal alavanca de negociação utilizada
- Preencher `realized_saving_value` apenas quando houver captura efetiva
- Manter `baseline_annual_spend` anualizado
- Atualizar `stage` conforme a evolução real da iniciativa
""")

# ========================
# Editar
# ========================
with tab5:
    st.header("Editar iniciativas")

    if df.empty:
        st.warning("Nenhuma iniciativa carregada.")
    else:
        editable_cols = REQUIRED_COLUMNS.copy()
        edit_df = df[editable_cols].copy()

        edit_df["expected_saving_pct"] = edit_df["expected_saving_pct"].apply(
            lambda x: x * 100 if pd.notna(x) else 0
        )

        edited_df = st.data_editor(
            edit_df,
            use_container_width=True,
            num_rows="dynamic",
            key="editor_initiatives",
            column_config={
                "initiative_id": st.column_config.NumberColumn(
                    "initiative_id",
                    disabled=True,
                    help="ID gerado automaticamente",
                ),
                "title": st.column_config.TextColumn("title"),
                "category": st.column_config.TextColumn("category"),
                "type": st.column_config.SelectboxColumn("type", options=TYPE_OPTIONS),
                "negotiation_lever": st.column_config.SelectboxColumn(
                    "negotiation_lever", options=LEVER_OPTIONS
                ),
                "owner": st.column_config.TextColumn("owner"),
                "stage": st.column_config.SelectboxColumn("stage", options=STAGE_OPTIONS),
                "start_date": st.column_config.TextColumn("start_date"),
                "expected_end_date": st.column_config.TextColumn("expected_end_date"),
                "baseline_annual_spend": st.column_config.NumberColumn(
                    "baseline_annual_spend", min_value=0.0, step=1000.0, format="%.2f"
                ),
                "expected_saving_pct": st.column_config.NumberColumn(
                    "expected_saving_pct",
                    min_value=0.0,
                    max_value=100.0,
                    step=0.5,
                    help="Preencher como percentual humano: ex. 5 = 5%",
                ),
                "confidence_level": st.column_config.SelectboxColumn(
                    "confidence_level", options=CONFIDENCE_OPTIONS
                ),
                "realized_saving_value": st.column_config.NumberColumn(
                    "realized_saving_value", min_value=0.0, step=1000.0, format="%.2f"
                ),
            },
        )

        if st.button("Salvar alterações", key="save_edited_initiatives"):
            df_to_save = edited_df.copy()

            df_to_save["baseline_annual_spend"] = pd.to_numeric(
                df_to_save["baseline_annual_spend"], errors="coerce"
            ).fillna(0.0)

            df_to_save["expected_saving_pct"] = pd.to_numeric(
                df_to_save["expected_saving_pct"], errors="coerce"
            ).fillna(0.0)

            df_to_save["realized_saving_value"] = pd.to_numeric(
                df_to_save["realized_saving_value"], errors="coerce"
            ).fillna(0.0)

            df_to_save["expected_saving_pct"] = df_to_save["expected_saving_pct"].apply(
                lambda x: x / 100 if x > 1 else x
            )

            df_to_save.to_csv("initiatives.csv", index=False)
            st.success("Alterações salvas com sucesso.")
            st.info("Atualize a página para refletir os novos dados no dashboard.")

# ========================
# Exportar
# ========================
with tab6:
    st.header("Exportar base")
    st.caption("Baixe a base atual do app para trabalhar offline em Excel")

    export_base_df = build_export_base(df)
    dictionary_df = build_dictionary_df()

    st.subheader("Opções de exportação")

    col1, col2 = st.columns(2)

    with col1:
        initiatives_xlsx = dataframe_to_excel_bytes({
            "initiatives": export_base_df
        })

        st.download_button(
            label="Baixar base de iniciativas (.xlsx)",
            data=initiatives_xlsx,
            file_name="compras_2_0_initiatives.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_initiatives_xlsx",
        )

    with col2:
        full_xlsx = dataframe_to_excel_bytes({
            "initiatives": export_base_df,
            "dictionary": dictionary_df,
        })

        st.download_button(
            label="Baixar base completa + dicionário (.xlsx)",
            data=full_xlsx,
            file_name="compras_2_0_full_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_full_xlsx",
        )

    st.subheader("Preview da base exportada")
    st.dataframe(export_base_df, use_container_width=True, hide_index=True)

    st.info(
        "Na exportação, o campo expected_saving_pct sai em formato humano "
        "(ex.: 5 = 5%) para facilitar o trabalho offline."
    )
