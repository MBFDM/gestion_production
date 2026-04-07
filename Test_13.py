import streamlit as st
import pandas as pd
import io
import tempfile
import os
import subprocess
import shutil
import zipfile
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from datetime import datetime
import locale

# ---------- Configuration des champs ----------
champs_cotes = [
    "N° Assuré", "N° Police", "N° Référence", "Intermédiaire", "Tél", "Tél WhatApps",
    "Nom(s) et Prénoms", "Date de Naissance", "Sexe", "Effet", "Echéance", "Durée (mois)",
    "Fractionnement", "Date de souscription", "Périodicité"
]

champs_dessous = [
    "Garantie", "Capital (FCFA)", "Primes Périodes (FCFA)",
    "Prime nette", "Accessoires", "Prime Totale"
]

champs_attendus = champs_cotes + champs_dessous

# Champs de date à formater
champs_date = ["Date de Naissance", "Effet", "Echéance", "Date de souscription"]

# Champs qui doivent être décalés de deux colonnes à droite
champs_decalage_double = ["N° Référence", "Nom(s) et Prénoms", "Date de Naissance"]
champs_decalage_triple = ["Date de souscription"]

# ---------- Fonction de formatage des dates ----------
def formater_date(valeur):
    """
    Convertit une valeur (datetime, date, str) en format "JJ Mois AAAA" (ex: 12 Mars 2000)
    Si la conversion échoue, retourne la valeur originale sous forme de chaîne.
    """
    if pd.isna(valeur):
        return ""
    
    # Si c'est déjà un objet datetime (pandas Timestamp, datetime.date, datetime.datetime)
    if isinstance(valeur, (pd.Timestamp, datetime)):
        # Utiliser la locale française pour les noms de mois
        try:
            locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')
        except:
            try:
                locale.setlocale(locale.LC_TIME, 'fr_FR')
            except:
                pass
        return valeur.strftime("%d %B %Y")
    
    # Si c'est une chaîne de caractères, tenter de parser
    if isinstance(valeur, str):
        # Essayer différents formats courants
        formats = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d.%m.%Y"]
        for fmt in formats:
            try:
                dt = datetime.strptime(valeur, fmt)
                try:
                    locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')
                except:
                    try:
                        locale.setlocale(locale.LC_TIME, 'fr_FR')
                    except:
                        pass
                return dt.strftime("%d %B %Y")
            except:
                continue
        # Si aucun format ne correspond, retourner la chaîne d'origine
        return valeur
    
    # Pour tout autre type, convertir en chaîne
    return str(valeur)

# ---------- Conversion DOCX -> PDF ----------
def convert_docx_to_pdf(docx_path, pdf_path):
    # LibreOffice
    libreoffice_cmds = ['libreoffice', 'soffice']
    for cmd in libreoffice_cmds:
        if shutil.which(cmd):
            try:
                subprocess.run(
                    [cmd, '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(pdf_path), docx_path],
                    check=True, capture_output=True
                )
                generated_pdf = os.path.join(os.path.dirname(pdf_path),
                                             os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
                if os.path.exists(generated_pdf):
                    os.rename(generated_pdf, pdf_path)
                return True
            except Exception:
                continue
    # docx2pdf (MS Word)
    try:
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        return True
    except Exception:
        return False

# ---------- Remplissage avec distinction côte / dessous et décalage double ----------
def remplir_un_certificat(template_bytes, data_dict, style_config):
    template_stream = BytesIO(template_bytes)
    doc = Document(template_stream)

    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                for champ, valeur in data_dict.items():
                    if champ in cell_text or cell_text == champ:
                        # Déterminer l'emplacement cible
                        if champ in champs_cotes:
                            # À droite : même ligne, colonne suivante (ou +2 pour certains champs)
                            target_row = row_idx
                            if champ in champs_decalage_triple:
                                target_col = col_idx + 3
                            elif champ in champs_decalage_double:
                                target_col = col_idx + 2
                            else:
                                target_col = col_idx + 1
                            # Si la colonne n'existe pas, on ajoute des colonnes
                            while target_col >= len(table.rows[target_row].cells):
                                # Ajouter une colonne à la fin du tableau pour toutes les lignes
                                for r in table.rows:
                                    r.cells.add()
                            target_cell = table.rows[target_row].cells[target_col]
                        elif champ in champs_dessous:
                            # En dessous : même colonne, ligne suivante
                            target_row = row_idx + 1
                            target_col = col_idx
                            # Si la ligne suivante n'existe pas, on l'ajoute
                            while target_row >= len(table.rows):
                                table.add_row()
                            target_cell = table.rows[target_row].cells[target_col]
                        else:
                            continue  # sécurité

                        # Remplir la cellule cible
                        target_cell.text = ""
                        paragraph = target_cell.paragraphs[0]
                        # Alignement
                        if style_config['alignment'] == 'gauche':
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        elif style_config['alignment'] == 'centre':
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        elif style_config['alignment'] == 'droite':
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        # Ajout du texte formaté
                        run = paragraph.add_run(str(valeur))
                        if style_config['font_name']:
                            run.font.name = style_config['font_name']
                        if style_config['font_size']:
                            run.font.size = Pt(style_config['font_size'])
                        if style_config['color_hex']:
                            rgb = RGBColor(
                                int(style_config['color_hex'][1:3], 16),
                                int(style_config['color_hex'][3:5], 16),
                                int(style_config['color_hex'][5:7], 16)
                            )
                            run.font.color.rgb = rgb
                        run.font.bold = style_config['bold']
                        run.font.italic = style_config['italic']
                        break  # champ trouvé

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ---------- Génération de tous les certificats ----------
def generer_tous_certificats(template_bytes, df, style_config):
    certificats = []
    for idx, row in df.iterrows():
        data_dict = {}
        for champ in champs_attendus:
            valeur = row[champ]
            if pd.notna(valeur):
                if champ in champs_date:
                    # Appliquer le formatage de date
                    valeur_formatee = formater_date(valeur)
                else:
                    valeur_formatee = str(valeur)
            else:
                valeur_formatee = ""
            data_dict[champ] = valeur_formatee
        
        identifiant = data_dict.get("Nom(s) et Prénoms", f"ligne_{idx+1}").replace("/", "_")
        docx_bytesio = remplir_un_certificat(template_bytes, data_dict, style_config)
        # Conversion PDF temporaire
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_docx:
            tmp_docx.write(docx_bytesio.getvalue())
            tmp_docx_path = tmp_docx.name
        pdf_path = tmp_docx_path.replace(".docx", ".pdf")
        conversion_ok = convert_docx_to_pdf(tmp_docx_path, pdf_path)
        pdf_bytes = None
        if conversion_ok and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
            os.unlink(pdf_path)
        os.unlink(tmp_docx_path)
        certificats.append((idx, identifiant, docx_bytesio, pdf_bytes))
    return certificats

# ---------- Interface Streamlit ----------
st.set_page_config(page_title="Générateur de certificats personnalisés", layout="wide")
st.title("📄 Générateur de certificats (Word + PDF) personnalisables")
st.markdown("""
Chargez un modèle Word (avec tableaux contenant les libellés) et un fichier Excel.
- Les champs de la liste **côté** (`champs_cotes`) sont insérés **à droite** du libellé.
- Les champs `N° Référence`, `Nom(s) et Prénoms` et `Date de souscription` sont insérés **deux cellules à droite**.
- Les champs de la liste **dessous** (`champs_dessous`) sont insérés **en dessous** du libellé.
- Les dates (`Date de Naissance`, `Effet`, `Echéance`, `Date de souscription`) sont formatées en **JJ Mois AAAA** (ex: 12 Mars 2000).
""")

col1, col2 = st.columns(2)
with col1:
    modele_file = st.file_uploader("📄 Modèle Word (.docx)", type=["docx"])
with col2:
    excel_file = st.file_uploader("📊 Fichier Excel (.xlsx)", type=["xlsx"])

# Personnalisation des styles
st.sidebar.header("🎨 Personnalisation des valeurs insérées")
font_name = st.sidebar.selectbox("Police", ["Arial", "Times New Roman", "Calibri", "Verdana", "Courier New"], index=0)
font_size = st.sidebar.slider("Taille (pt)", 8, 48, 11)
color_hex = st.sidebar.color_picker("Couleur du texte", "#000000")
bold = st.sidebar.checkbox("Gras", value=False)
italic = st.sidebar.checkbox("Italique", value=False)
alignment = st.sidebar.radio("Alignement horizontal", ["gauche", "centre", "droite"], index=0)

style_config = {
    'font_name': font_name,
    'font_size': font_size,
    'color_hex': color_hex,
    'bold': bold,
    'italic': italic,
    'alignment': alignment
}

if modele_file and excel_file:
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        st.success(f"Excel chargé : {df.shape[0]} ligne(s), {df.shape[1]} colonne(s)")
        st.subheader("Aperçu du fichier Excel")
        st.dataframe(df, use_container_width=True)

        colonnes_manquantes = [champ for champ in champs_attendus if champ not in df.columns]
        if colonnes_manquantes:
            st.error(f"❌ Colonnes manquantes : {', '.join(colonnes_manquantes)}")
            st.stop()
        else:
            st.success("✅ Tous les en-têtes requis sont présents.")

        with st.spinner(f"Génération de {df.shape[0]} certificat(s)..."):
            template_bytes = modele_file.read()
            certificats = generer_tous_certificats(template_bytes, df, style_config)

        st.success(f"{len(certificats)} certificat(s) généré(s).")

        if len(certificats) > 0:
            first_docx = certificats[0][2]
            st.download_button(
                label="📄 Télécharger le Modèle Word Final (exemple première ligne)",
                data=first_docx.getvalue(),
                file_name="modele_word_final.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        st.subheader("Certificats générés")
        zip_word = BytesIO()
        zip_pdf = BytesIO()
        with zipfile.ZipFile(zip_word, 'w') as zw:
            with zipfile.ZipFile(zip_pdf, 'w') as zp:
                for idx, ident, docx_bytesio, pdf_bytes in certificats:
                    safe_name = ident.replace(" ", "_").replace("(", "").replace(")", "")
                    docx_name = f"{safe_name}.docx"
                    pdf_name = f"{safe_name}.pdf"
                    zw.writestr(docx_name, docx_bytesio.getvalue())
                    if pdf_bytes:
                        zp.writestr(pdf_name, pdf_bytes)
                    col_a, col_b, col_c, col_d = st.columns([3,1,1,1])
                    col_a.write(f"**{ident}**")
                    col_b.download_button("📄 Word", data=docx_bytesio.getvalue(), file_name=docx_name, key=f"word_{idx}")
                    if pdf_bytes:
                        col_c.download_button("📑 PDF", data=pdf_bytes, file_name=pdf_name, key=f"pdf_{idx}")
                    else:
                        col_c.write("❌ PDF non généré")
        zip_word.seek(0)
        zip_pdf.seek(0)
        st.markdown("---")
        col_zip1, col_zip2 = st.columns(2)
        with col_zip1:
            st.download_button("📦 Tous les Word (ZIP)", data=zip_word, file_name="tous_word.zip", mime="application/zip")
        with col_zip2:
            st.download_button("📦 Tous les PDF (ZIP)", data=zip_pdf, file_name="tous_pdf.zip", mime="application/zip", disabled=(zip_pdf.getbuffer().nbytes == 0))

    except Exception as e:
        st.error(f"Erreur : {str(e)}")
        st.stop()
else:
    st.info("Veuillez charger un modèle Word et un fichier Excel.")