import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
except Exception:
    pdfplumber = None

FINAL_COLUMNS = [
    "Date",
    "Libellé",
    "Compte Client",
    "Montant TTC",
    "Compte HT",
    "Montant HT",
    "Compte TVA",
    "Montant TVA",
    "Compte DT",
    "Montant DT",
]

REQUIRED_SOURCE_FIELDS = [
    "Date",
    "N° facture",
    "Nom du client",
    "Montant HT",
    "TVA",
    "Montant TTC",
    "DT",
]

DEFAULT_ACCOUNTS = {
    "Compte HT": "701000",
    "Compte TVA": "436700",
    "Compte DT": "437000",
}

st.set_page_config(page_title="Générateur écritures factures", layout="wide")


def normalize_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def read_excel_file(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, dtype=str)
    except Exception as exc:
        raise ValueError(f"Impossible de lire le fichier Excel : {exc}")


def read_pdf_file(uploaded_file):
    if pdfplumber is None:
        raise ValueError("La bibliothèque pdfplumber n'est pas installée.")
    rows = []
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for table in tables:
                    if table and len(table) > 1:
                        header = table[0]
                        for row in table[1:]:
                            rows.append(dict(zip(header, row)))
        if not rows:
            raise ValueError("Aucun tableau exploitable n'a été détecté dans le PDF.")
        return pd.DataFrame(rows)
    except Exception as exc:
        raise ValueError(f"Impossible de lire le fichier PDF : {exc}")


def read_source_file(uploaded_file):
    suffix = uploaded_file.name.lower().split(".")[-1]
    if suffix in ["xlsx", "xls"]:
        return read_excel_file(uploaded_file)
    if suffix == "pdf":
        return read_pdf_file(uploaded_file)
    raise ValueError("Format non supporté. Importer un fichier Excel ou PDF.")


def read_plan_comptable(uploaded_file):
    plan = pd.read_excel(uploaded_file, dtype=str)
    plan.columns = [str(c).strip() for c in plan.columns]
    required = {"Compte", "Intitulé"}
    missing = required - set(plan.columns)
    if missing:
        raise ValueError(f"Le plan comptable doit contenir les colonnes : {', '.join(required)}")
    plan = plan[["Compte", "Intitulé"]].copy()
    plan["Compte"] = plan["Compte"].astype(str).str.strip()
    plan["Intitulé"] = plan["Intitulé"].astype(str).str.strip()
    plan = plan[plan["Compte"].str.startswith("411", na=False)]
    plan["client_key"] = plan["Intitulé"].apply(normalize_text)
    return plan


def detect_column_mapping(source_columns):
    synonyms = {
        "Date": ["date", "date facture", "jour"],
        "N° facture": ["n facture", "numero facture", "num facture", "facture", "reference"],
        "Nom du client": ["nom client", "client", "raison sociale", "intitule client", "tiers"],
        "Montant HT": ["montant ht", "ht", "h t", "hors taxe", "base ht"],
        "TVA": ["tva", "montant tva", "taxe"],
        "Montant TTC": ["montant ttc", "ttc", "total ttc", "total"],
        "DT": ["dt", "droit de timbre", "timbre", "montant dt"],
    }
    normalized_source = {normalize_text(col): col for col in source_columns}
    mapping = []
    for target, words in synonyms.items():
        detected = ""
        for norm_col, original_col in normalized_source.items():
            if any(word == norm_col or word in norm_col for word in words):
                detected = original_col
                break
        mapping.append({"Champ final": target, "Colonne source": detected})
    return pd.DataFrame(mapping)


def mapping_to_dict(mapping_df):
    clean = mapping_df.dropna(subset=["Champ final"]).copy()
    return dict(zip(clean["Champ final"].astype(str), clean["Colonne source"].astype(str)))


def to_date(value):
    if pd.isna(value) or str(value).strip() == "":
        return None
    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%d/%m/%Y")


def to_amount(value):
    if pd.isna(value) or str(value).strip() == "":
        return None
    text = str(value).strip()
    text = text.replace(" ", "").replace("\u202f", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    try:
        amount = float(text)
    except Exception:
        return None
    if amount == 0:
        return None
    return amount


def find_client_account(client_name, plan):
    key = normalize_text(client_name)
    if not key:
        return None
    exact = plan.loc[plan["client_key"] == key]
    if not exact.empty:
        return exact.iloc[0]["Compte"]
    contains = plan.loc[plan["client_key"].str.contains(re.escape(key), na=False) | plan["client_key"].apply(lambda x: x in key if x else False)]
    if not contains.empty:
        return contains.iloc[0]["Compte"]
    tokens = [t for t in key.split() if len(t) >= 3]
    if tokens:
        def score(plan_key):
            return sum(1 for t in tokens if t in plan_key)
        scores = plan["client_key"].apply(score)
        best = scores.max() if not scores.empty else 0
        if best >= max(1, min(2, len(tokens))):
def find_client_account(client_name, plan):
    # On imagine que 'scores' est calculé ici dans votre code
    if not scores.empty and scores.max() > 0:
        return plan.iloc[scores.idxmax()]["Compte"]
    else:
        return "41100000"


def add_error(errors, row_num, column, value, reason, suggestion):
    errors.append({
        "Numéro ligne source": row_num,
        "Colonne concernée": column,
        "Valeur détectée": "" if pd.isna(value) else value,
        "Motif de l'erreur": reason,
        "Suggestion de correction": suggestion,
    })


def validate_and_transform(source_df, plan, mapping, accounts):
    generated = []
    errors = []
    invalid_rows = set()

    for idx, row in source_df.iterrows():
        source_line = idx + 2
        values = {}
        for field in REQUIRED_SOURCE_FIELDS:
            col = mapping.get(field, "")
            values[field] = row.get(col, None) if col in source_df.columns else None
            if not col or col not in source_df.columns:
                add_error(errors, source_line, field, "", "Colonne source non mappée", "Choisir une colonne source dans le mapping.")
                invalid_rows.add(idx)

        invoice_no = values["N° facture"]
        client_name = values["Nom du client"]
        date_value = to_date(values["Date"])
        ht = to_amount(values["Montant HT"])
        tva = to_amount(values["TVA"])
        ttc = to_amount(values["Montant TTC"])
        dt_value = 1
        account_client = find_client_account(client_name, plan)

        if not date_value:
            add_error(errors, source_line, "Date", values["Date"], "Date vide ou invalide", "Corriger la date au format JJ/MM/AAAA.")
            invalid_rows.add(idx)
        if pd.isna(invoice_no) or str(invoice_no).strip() == "":
            add_error(errors, source_line, "N° facture", invoice_no, "Numéro de facture vide", "Renseigner le numéro de facture.")
            invalid_rows.add(idx)
        if pd.isna(client_name) or str(client_name).strip() == "":
            add_error(errors, source_line, "Nom du client", client_name, "Nom client vide", "Renseigner le nom du client.")
            invalid_rows.add(idx)
        if not account_client:
            add_error(errors, source_line, "Compte Client", client_name, "Client non trouvé dans le plan comptable 411", "Ajouter ou corriger le client dans le plan comptable.")
            invalid_rows.add(idx)
        for col_name, amount in [("Montant HT", ht), ("TVA", tva), ("Montant TTC", ttc)]:
            if amount is None:
                add_error(errors, source_line, col_name, values[col_name if col_name != "TVA" else "TVA"], "Montant vide, non numérique ou égal à zéro", "Corriger le montant dans le fichier source.")
                invalid_rows.add(idx)

        if idx not in invalid_rows:
            generated.append({
                "Date": date_value,
                "Libellé": f"FACTURE {str(invoice_no).strip()} {str(client_name).strip()}",
                "Compte Client": account_client,
                "Montant TTC": ttc,
                "Compte HT": accounts.get("Compte HT", ""),
                "Montant HT": ht,
                "Compte TVA": accounts.get("Compte TVA", ""),
                "Montant TVA": tva,
                "Compte DT": accounts.get("Compte DT", ""),
                "Montant DT": dt_value,
            })

    generated_df = pd.DataFrame(generated, columns=FINAL_COLUMNS)
    errors_df = pd.DataFrame(errors, columns=[
        "Numéro ligne source", "Colonne concernée", "Valeur détectée", "Motif de l'erreur", "Suggestion de correction"
    ])
    summary = {
        "Nombre total de lignes importées": len(source_df),
        "Nombre de lignes valides": len(generated_df),
        "Nombre de lignes en erreur": len(set(errors_df["Numéro ligne source"])) if not errors_df.empty else 0,
        "Total HT": generated_df["Montant HT"].sum() if not generated_df.empty else 0,
        "Total TVA": generated_df["Montant TVA"].sum() if not generated_df.empty else 0,
        "Total TTC": generated_df["Montant TTC"].sum() if not generated_df.empty else 0,
    }
    return generated_df, errors_df, summary


def build_summary_df(summary, errors_df):
    rows = list(summary.items())
    if errors_df.empty:
        rows.extend([
            ("Clients non trouvés", "Aucun"),
            ("Dates invalides", "Aucune"),
            ("Montants invalides", "Aucun"),
        ])
    else:
        clients = errors_df.loc[errors_df["Colonne concernée"].eq("Compte Client"), "Valeur détectée"].dropna().astype(str).unique()
        dates = errors_df.loc[errors_df["Colonne concernée"].eq("Date"), "Numéro ligne source"].astype(str).unique()
        amounts = errors_df.loc[errors_df["Motif de l'erreur"].str.contains("Montant", case=False, na=False), "Numéro ligne source"].astype(str).unique()
        rows.extend([
            ("Clients non trouvés", ", ".join(clients) if len(clients) else "Aucun"),
            ("Dates invalides", ", ".join(dates) if len(dates) else "Aucune"),
            ("Montants invalides", ", ".join(amounts) if len(amounts) else "Aucun"),
        ])
    return pd.DataFrame(rows, columns=["Indicateur", "Valeur"])


def write_sheet(ws, df, title):
    ws.title = title
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 1), min_col=1, max_col=max(ws.max_column, 1)):
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="center")
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, 12), 45)
    ws.freeze_panes = "A2"


def generate_excel_file(generated_df, errors_df, summary_df):
    output = io.BytesIO()
    wb = Workbook()
    ws1 = wb.active
    write_sheet(ws1, generated_df, "Écritures générées")
    ws2 = wb.create_sheet("Contrôle erreurs")
    write_sheet(ws2, errors_df, "Contrôle erreurs")
    ws3 = wb.create_sheet("Résumé contrôle")
    write_sheet(ws3, summary_df, "Résumé contrôle")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.000'
    wb.save(output)
    output.seek(0)
    return output


def display_kpis(summary):
    cols = st.columns(6)
    keys = list(summary.keys())
    for col, key in zip(cols, keys):
        col.metric(key, f"{summary[key]:,.3f}" if isinstance(summary[key], float) else summary[key])


def main():
    st.title("Générateur automatique d'écritures de factures")
    st.caption("Import Excel/PDF → contrôles → fichier Excel final structuré")

    st.header("1. Télécharger les fichiers")
    source_file = st.file_uploader("Fichier source Excel ou PDF", type=["xlsx", "xls", "pdf"])
    plan_file = st.file_uploader("Plan comptable Excel", type=["xlsx", "xls"])

    if not source_file or not plan_file:
        st.info("Importer le fichier source et le plan comptable pour commencer.")
        return

    try:
        source_df = read_source_file(source_file)
        source_df.columns = [str(c).strip() for c in source_df.columns]
        plan_df = read_plan_comptable(plan_file)
    except Exception as exc:
        st.error(str(exc))
        return

    st.header("2. Prévisualiser les données")
    st.write(f"Nombre total de lignes importées : **{len(source_df)}**")
    st.dataframe(source_df.head(30), use_container_width=True)
    with st.expander("Prévisualiser les comptes clients 411 du plan comptable"):
        st.dataframe(plan_df[["Compte", "Intitulé"]].head(50), use_container_width=True)

    st.header("3. Valider ou modifier le mapping")
    if "mapping_df" not in st.session_state:
        st.session_state.mapping_df = detect_column_mapping(source_df.columns)
    mapping_df = st.data_editor(
        st.session_state.mapping_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Champ final": st.column_config.SelectboxColumn("Champ final", options=REQUIRED_SOURCE_FIELDS, required=True),
            "Colonne source": st.column_config.SelectboxColumn("Colonne source", options=[""] + list(source_df.columns), required=False),
        },
        key="mapping_editor",
    )

    st.subheader("Comptes comptables manuels")
    c1, c2, c3 = st.columns(3)
    accounts = {
        "Compte HT": c1.text_input("Compte HT", value=DEFAULT_ACCOUNTS["Compte HT"]),
        "Compte TVA": c2.text_input("Compte TVA", value=DEFAULT_ACCOUNTS["Compte TVA"]),
        "Compte DT": c3.text_input("Compte DT", value=DEFAULT_ACCOUNTS["Compte DT"]),
    }

    st.header("4. Contrôler les erreurs")
    mapping = mapping_to_dict(mapping_df)
    generated_df, errors_df, summary = validate_and_transform(source_df, plan_df, mapping, accounts)
    summary_df = build_summary_df(summary, errors_df)
    display_kpis(summary)

    c1, c2, c3 = st.columns(3)
    clients_not_found = errors_df.loc[errors_df["Colonne concernée"].eq("Compte Client"), "Valeur détectée"].dropna().astype(str).unique() if not errors_df.empty else []
    invalid_dates = errors_df.loc[errors_df["Colonne concernée"].eq("Date"), "Numéro ligne source"].astype(str).unique() if not errors_df.empty else []
    invalid_amounts = errors_df.loc[errors_df["Motif de l'erreur"].str.contains("Montant", case=False, na=False), "Numéro ligne source"].astype(str).unique() if not errors_df.empty else []
    c1.write("**Clients non trouvés**")
    c1.write(", ".join(clients_not_found) if len(clients_not_found) else "Aucun")
    c2.write("**Dates invalides**")
    c2.write(", ".join(invalid_dates) if len(invalid_dates) else "Aucune")
    c3.write("**Montants invalides**")
    c3.write(", ".join(invalid_amounts) if len(invalid_amounts) else "Aucun")

    if not errors_df.empty:
        st.warning("Des erreurs ont été détectées. Consultez les motifs ci-dessous avant génération définitive.")
        st.dataframe(errors_df, use_container_width=True)
    else:
        st.success("Aucune erreur détectée.")

    st.header("5. Générer le fichier Excel")
    st.dataframe(generated_df, use_container_width=True)
    excel_file = generate_excel_file(generated_df, errors_df, summary_df)

    st.header("6. Télécharger le fichier final")
    st.download_button(
        label="Télécharger le fichier Excel final",
        data=excel_file,
        file_name="ecritures_factures_generees.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=generated_df.empty,
    )


if __name__ == "__main__":
    main()
