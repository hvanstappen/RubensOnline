import pandas as pd


def sheet_to_html(sheet_url, datamodel_sheet, vocab_sheet, columns_to_include):
    base_url = sheet_url.split('/edit')[0]
    export_url = f"{base_url}/export?format=xlsx"

    try:
        df_model = pd.read_excel(export_url, sheet_name=datamodel_sheet)
        df_vocab = pd.read_excel(export_url, sheet_name=vocab_sheet)

        # Voorbereiden Waardenlijsten
        vocab_dict = {}
        for col_idx in range(0, len(df_vocab.columns), 2):
            list_name = df_vocab.columns[col_idx]
            items = df_vocab.iloc[:, [col_idx, col_idx + 1]].dropna(how='all')
            if not items.empty:
                items.columns = ['Label', 'URI']
                vocab_dict[list_name] = items.to_dict('records')

        # Fallbacks voor groepering
        class_col = "Class"
        section_col = "section"
        if class_col not in df_model.columns: df_model[class_col] = "Ongecategoriseerd"
        if section_col not in df_model.columns: df_model[section_col] = "---"

        # Zorg dat lege secties een label hebben
        df_model[section_col] = df_model[section_col].fillna("")

        existing_columns = [col for col in columns_to_include if col in df_model.columns]
        unique_classes = list(dict.fromkeys(df_model[class_col].dropna().tolist()))

        toc_html = ""
        content_html = ""

        # Verwerk Datamodel met Subsecties
        for class_name in unique_classes:
            class_id = str(class_name).replace(" ", "_").lower()
            df_class = df_model[df_model[class_col] == class_name]

            # TOC: Hoofdgroep (Klasse)
            toc_html += f'<details open class="toc-class"><summary><strong>{class_name}</strong></summary>'
            content_html += f'<h2 id="{class_id}" class="class-header">Klasse: {class_name}</h2>'

            # Subgroepering op sectie binnen de klasse
            unique_sections = list(dict.fromkeys(df_class[section_col].tolist()))

            for section_name in unique_sections:
                section_id = f"{class_id}_{str(section_name).replace(' ', '_').lower()}"
                df_section = df_class[df_class[section_col] == section_name]

                # TOC: Subgroep (Sectie)
                toc_html += f'<details open class="toc-section"><summary>{section_name}</summary><ul>'

                # Content: Sectie kop
                content_html += f'<h4 class="section-subheader">{section_name}</h4>'

                for index, row in df_section.iterrows():
                    title_value = row[columns_to_include[1]]
                    if pd.isna(title_value) or str(title_value).strip() == "": continue

                    anchor_id = f"item_{index}"
                    toc_html += f'<li><a href="#{anchor_id}">{title_value}</a></li>'

                    # Eigenschap kaart
                    content_html += f'<section id="{anchor_id}" class="card">'
                    content_html += f'<h3>{title_value} <small style="color:#666; font-weight:normal;">({row[columns_to_include[0]]})</small></h3>'
                    content_html += '<table><tbody>'

                    for col in existing_columns:
                        val = row[col]
                        if pd.notna(val) and str(val).strip() != "":
                            display_content = str(val).strip()
                            if col == "ValueList":
                                list_part = display_content.replace("-->", "").strip()
                                vocab_id = f"vocab_{list_part.replace(' ', '_').lower()}"
                                display_content = f'<a href="#{vocab_id}" class="vocab-link">➔ {list_part}</a>'
                            elif display_content.startswith("http"):
                                display_content = f'<a href="{display_content}" target="_blank">{display_content}</a>'
                            else:
                                display_content = display_content.replace('\n', '<br>')

                            content_html += f'<tr><th>{col}</th><td>{display_content}</td></tr>'

                    content_html += '</tbody></table></section>'

                toc_html += "</ul></details>"

            content_html += f'<div style="text-align: right;"><a href="#top" class="back-to-top">↑ Terug naar boven</a></div>'
            toc_html += "</details>"

        # Verwerk Waardenlijsten
        vocab_section_html = '<h1 id="vocabularies" class="main-header">Waardenlijsten</h1>'
        toc_html += '<details class="toc-class"><summary><strong>Waardenlijsten</strong></summary><ul>'
        for list_name, items in vocab_dict.items():
            vocab_id = f"vocab_{list_name.replace(' ', '_').lower()}"
            toc_html += f'<li><a href="#{vocab_id}">{list_name}</a></li>'
            vocab_section_html += f'<section id="{vocab_id}" class="card vocab-card"><h3 class="class-header-value">Waardenlijst: <strong>{list_name}</strong></h3><table><thead><tr><th>Label</th><th>URI</th></tr></thead><tbody>'
            for item in items[1:]:
                uri_display = f'<a href="{item["URI"]}" target="_blank">{item["URI"]}</a>' if str(
                    item["URI"]).startswith("http") else item["URI"]
                vocab_section_html += f'<tr><td>{item["Label"]}</td><td>{uri_display}</td></tr>'
            vocab_section_html += '</tbody></table></section>'
        toc_html += '</ul></details>'

        return f"""
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <title>Datamodel: {datamodel_sheet}</title>
    <style>
        :root {{ --primary: #2c3e50; --accent: #3498db; --bg: #f4f7f9; --vocab: #e67e22; }}
        body {{ font-family: -apple-system, sans-serif; margin: 0; display: flex; background: var(--bg); scroll-behavior: smooth; }}
        aside {{ width: 320px; height: 100vh; position: sticky; top: 0; background: white; border-right: 1px solid #ddd; overflow-y: auto; padding: 20px; box-sizing: border-box; }}

        /* TOC Stijlen */
        .toc-class {{ margin-bottom: 15px; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
        .toc-section {{ margin-left: 10px; margin-top: 5px; border-left: 2px solid #ddd; padding-left: 5px; }}
        summary {{ cursor: pointer; padding: 5px; color: var(--primary); list-style: none; font-size: 0.95rem; }}
        summary:hover {{ background: #f0f0f0; border-radius: 4px; }}
        aside ul {{ list-style: none; padding-left: 15px; margin: 5px 0; }}
        aside li {{ font-size: 0.8rem; margin-bottom: 3px; }}
        aside a {{ text-decoration: none; color: var(--accent); }}

        main {{ flex: 1; padding: 40px; max-width: 1100px; }}
        .class-header {{ background: var(--primary); color: white; padding: 15px; border-radius: 8px; }}
        .class-header-value {{ background: orange; color: white; padding: 15px; border-radius: 8px; }}
        .section-subheader {{ color: var(--primary); border-bottom: 2px dashed #ccc; padding-bottom: 5px; margin-top: 40px; text-transform: uppercase; font-size: 0.9rem; letter-spacing: 1px; }}
        .card {{ background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.07); margin-bottom: 25px; border-left: 6px solid var(--accent); }}
        .vocab-card {{ border-left: 6px solid var(--vocab); }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 10px; border: 1px solid #eee; text-align: left; vertical-align: top; font-size: 0.9rem; }}
        th {{ background: #f8f9fa; width: 25%; color: #7f8c8d; text-transform: uppercase; font-size: 0.7rem; }}
        .vocab-link {{ color: var(--vocab); font-weight: bold; }}
    </style>
</head>
<body>
    <aside><h2>Inhoud</h2>{toc_html}</aside>
    <main>
        <h1 id="top">Datamodel Documentatie</h1>
        <p><a href="https://docs.google.com/spreadsheets/d/1t_Tt4jtfaT1j-gEBEcGngK_qcSTe3TC07R4XQUFudys/edit?gid=703217171#gid=703217171">Bron</a></p>
        <p>Deze pagina geeft voor elk element uit het datamodel een overzicht van de eigenschappen:</p>
        <ul>
        <li>name: De technische veldnaam zoals gebruikt in de database</li>
        <li>label@nl: Label in het Nederlands</li>
        <li>Definition: Een korte beschrijving van wat er in het veld moet worden ingevuld of wat de waarde representeert.</li>
        <li>Data Type: Het type gegevens (bijv. URI, string, date, number of boolean).</li>
        <li>Encoding Scheme: Het coderingsschema dat wordt gebruikt om de waarde te coderen (bijv. URI, base64, hex).</li>
        <li>ValueList: Verwijst naar een gecontroleerde waardenlijst (zoals AAT, TGN of een specifieke projectlijst) waaruit gekozen moet worden.</li>
        <li>Obligation: Geeft aan of het veld verplicht (Mandatory), optioneel (Optional) of voorwaardelijk (Conditional) is.</li>
        <li>Condition: Geeft aan welke voorwaarde nodig is om het veld te gebruiken (bijvoorbeeld: een andere waarde moet bestaan).</li>
        <li>Repeatable_with: Geeft aan welk veld de waarde moet worden herhaald als het veld Repeatable is ingesteld op True.</li>
        <li>Comments: Extra toelichting voor de gebruikers van het datamodel.</li>
        <li>Internal Note: specifieke opmerkingen voor intern gebruik.</li>
        <li>Example: Een voorbeeldwaarde om te illustreren hoe het veld ingevuld moet worden.</li>
        <li>linkedArt [draft]: De mapping naar het Linked Art conceptuele model (gebaseerd op CIDOC-CRM).</li>
        <li>RKD / Brocade / DAMS / ARCHES: Mapping-informatie naar externe systemen of bronnen zoals het RKD of specifieke interne databases.</li>
        </ul>
        {content_html}
        {vocab_section_html}
    </main>
</body>
</html>"""
    except Exception as e:
        return f"Fout: {e}"


# --- CONFIGURATIE ---
URL = "https://docs.google.com/spreadsheets/d/1t_Tt4jtfaT1j-gEBEcGngK_qcSTe3TC07R4XQUFudys/edit?gid=703217171"
MODEL_TAB = "RO_datamodel_v0-1"
VOCAB_TAB = "Vocabulary"
KOLOMMEN = ["name", "label@nl", "Definition", "Data Type", "ValueList", "Obligation", "Condition", "Repeatability", "RKD", "Comments", "Example", "Internal Note"]

html_result = sheet_to_html(URL, MODEL_TAB, VOCAB_TAB, KOLOMMEN)
with open("Rubensonline_model.html", "w", encoding="utf-8") as f:
    f.write(html_result)

with open("../model/index.html", "w", encoding="utf-8") as f:
    f.write(html_result)