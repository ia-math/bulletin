import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import tempfile
import os
import csv
from collections import defaultdict
import re

# Chargement base INSEE + fréquence
FREQ_FILLE = defaultdict(int)
FREQ_GARCON = defaultdict(int)
PRENOM_CSV_PATH = os.path.join(os.path.dirname(__file__), "prenoms.csv")

with open(PRENOM_CSV_PATH, encoding="utf-8") as f:
    reader = csv.DictReader(f, delimiter=';')
    prenom_col = next((c for c in reader.fieldnames if c.lower() in {"preusuel", "prenoms", "prenom"}), None)
    sexe_col = next((c for c in reader.fieldnames if c.lower() == "sexe"), None)
    nombre_col = next((c for c in reader.fieldnames if c.lower() in {"nombre", "nombres", "nb"}), None)
    for row in reader:
        prenom = row[prenom_col].strip().capitalize()
        sexe = row[sexe_col]
        try:
            nb = int(row[nombre_col])
        except Exception:
            nb = 1
        if prenom.startswith("_") or not prenom:
            continue
        if sexe == "1":
            FREQ_GARCON[prenom] += nb
        elif sexe == "2":
            FREQ_FILLE[prenom] += nb

def extract_nom_prenom(nom_complet):
    """
    Retourne : nom (majuscules), prénom (tout le reste, minuscules/capitales classiques)
    Ex : 'FARES Adam' => nom: 'FARES', prénom: 'Adam'
    """
    if not nom_complet:
        return "", ""
    mots = str(nom_complet).split()
    nom_parts = [part for part in mots if part.isupper()]
    prenom_parts = [part for part in mots if not part.isupper()]
    nom = " ".join(nom_parts)
    prenom = " ".join(prenom_parts)
    return nom, prenom.capitalize()

def detect_genre_majoritaire(prenom):
    if not prenom:
        return "u"
    # Prend le premier prénom (utile en cas de prénoms composés)
    prenom_principal = prenom.split()[0].capitalize()
    nb_f = FREQ_FILLE.get(prenom_principal, 0)
    nb_m = FREQ_GARCON.get(prenom_principal, 0)
    if nb_f > nb_m:
        return "f"
    elif nb_m > nb_f:
        return "m"
    else:
        return "u"

def appreciation_bull(prenom, moyenne, genre_):
    if moyenne is None or moyenne == "Abs":
        return f"Absent ce trimestre."
    try:
        m = float(moyenne)
    except Exception:
        return f"Erreur sur la moyenne."
    serieux = "sérieuse" if genre_ == 'f' else "sérieux"
    applique = "appliquée" if genre_ == 'f' else "appliqué"
    implique = "impliquée" if genre_ == 'f' else "impliqué"
    brillant = "brillante" if genre_ == 'f' else "brillant"
    attentif = "attentive" if genre_ == 'f' else "attentif"

    if m >= 19.5:
        return f"Le bilan trimestriel de {prenom} force les louanges. Félicitations."
    if m >= 19:
        return f"Excellent trimestre pour {prenom}. Élève très {brillant} et volontaire. Félicitations."
    if m >= 18:
        return f"Très bon trimestre. {prenom} est très {attentif} et son travail est de qualité. Félicitations."
    if m >= 17:
        return f"Très bon trimestre pour {prenom}. Élève très {serieux} et {attentif}. Félicitations !"
    if m >= 16:
        return f"Bon trimestre pour {prenom}. Élève très {serieux} et {applique}. Félicitations !"
    elif m >= 15.5:
        return f"Bon trimestre pour {prenom}. L'attitude est sérieuse et le travail est sérieux. Continuer ainsi."
    elif m >= 15:
        return f"Bon trimestre pour {prenom}. Le travail est sérieux et régulier. Bonne participation orale. Continuer ainsi."
    elif m >= 14.5:
        return f"Bon trimestre. {prenom} a fournit un travail sérieux et régulier. Continuer ainsi."
    elif m >= 14:
        return f"Bon trimestre. {prenom} a fournit un travail sérieux et a fait preuve d'un bon état d'esprit. Continuer ainsi."
    elif m >= 13.5:
        return f"Assez bon trimestre. {prenom} a fournit un travail sérieux ce trimestre. Je l'encourage à continuer ainsi afin qu'il progresse encore. "
    elif m >= 13:
        return f"Assez bon trimestre. {prenom} a fourni des efforts et a été {implique}, mais son travail est irrégulier. Je l'encourage à continuer ses efforts."
    elif m >= 12:
        return f"Ensemble assez satisfaisant. {prenom} pourrait sans doute mieux faire avec un travail plus approfondi et régulier."
    elif m >= 11:
        return f"Résultat satisfaisant mais {prenom} peut mieux faire avec plus d'attention et de régularité."
    elif m >= 10:
        return f"Résultats trop justes. {prenom} a fournit un travail trop irrégulier. Il ne faut pas baisser les bras et continuer les efforts pour progresser."
    elif m >= 9:
        return f"Les résultats sont hétérogènes et ils révèlent des difficultés et des lacunes. {prénom} doit travailler régulièrement et sérieusement afin de progresser."
    elif m >= 8:
        return f"Résultats insuffisants en raison de difficultés et de lacunes. {prenom} doit s'investir davantage pour progresser."
    elif m >= 7:
        return f"Résultats insuffisants en raison de difficultés et de lacunes. {prenom} doit réagir en travaillant sérieusement."
    elif m >= 6:
        return f"Résultats très insuffisants en raison de difficultés et d'un manque de travail personnel. {prenom} doit réagir de tout urgence !"
    elif m > 5:
        return f"Résultats très insuffisants en raison d’un manque de travail et de concentration en classe. {prenom} doit accentuer ses efforts pour progresser."
    elif m > 3:
        return f"Les résultats sont inquiétants et ils révèlent des difficultés et des lacunes accentuées par le manque de travail et de motivation."
    elif m > 0:
        return f"Résultats alarmants. {prenom} a des difficultés handicapantes qui nécessiteraient un travail régulier et soutenu.  Il faut réagir en travaillant régulièrement afin de progresser."
    else:
        return f"Donnés manquantes."

def generer_appreciations_excel_selection(input_path, output_path):
    wb_src = openpyxl.load_workbook(input_path)
    ws_src = wb_src.active
    entetes = [str(cell.value).strip().lower() if cell.value else "" for cell in ws_src[1]]
    col_nom = next((i for i, h in enumerate(entetes) if "élève" in h or "nom" in h), 0)
    col_moy = next((i for i, h in enumerate(entetes) if "moyenne" in h), 1)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bulletin"
    ws.append(["Nom", "Prénom", "Moyenne", "Appréciation générale"])

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    for i, cell in enumerate(ws[1]):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    alt_fill = PatternFill(start_color="E8F0F7", end_color="E8F0F7", fill_type="solid")
    row_num = 2
    for src_row in ws_src.iter_rows(min_row=2, max_row=ws_src.max_row):
        nom_complet = src_row[col_nom].value
        moyenne = src_row[col_moy].value
        if nom_complet is None or str(nom_complet).strip() == "":
            continue
        nom, prenom = extract_nom_prenom(nom_complet)
        genre_ = detect_genre_majoritaire(prenom)
        appreciation = appreciation_bull(prenom, moyenne, genre_)
        ws.append([nom, prenom, moyenne, appreciation])

        for col_idx in range(1, 5):
            cell = ws.cell(row=row_num, column=col_idx)
            if row_num % 2 == 0:
                cell.fill = alt_fill
            cell.alignment = Alignment(horizontal="left" if col_idx in [1,2] else "center", vertical="top", wrap_text=True)
            cell.font = Font(size=10)
            cell.border = border
        ws.cell(row=row_num, column=4).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        row_num += 1

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 55
    ws.row_dimensions[1].height = 28
    for i in range(2, ws.max_row + 1):
        ws.row_dimensions[i].height = 40

    wb.save(output_path)

st.set_page_config(page_title="Bulletin Automatique", page_icon="📝", layout="centered")
st.title("📝 Générateur d'appréciations pour bulletins scolaires")
st.header("Outil d'aide pour les professeurs")
st.markdown(
"""
Application créé par M. FARES (Professeur de mathématiques).

- Le programme fournit une appréciation en fonction de la moyenne générale.
- Les appréciations sont accordées selon le genre (Base de données de l'INSEE).
"""
)

uploaded_file = st.file_uploader("📂 Dépose ton fichier Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_in:
            temp_in.write(uploaded_file.read())
            temp_in.flush()
            temp_out = temp_in.name.replace(".xlsx", "_appreciations.xlsx")
            generer_appreciations_excel_selection(temp_in.name, temp_out)
            st.success("✅ Fichier traité avec succès !")
            with open(temp_out, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le bulletin avec appréciations",
                    data=f,
                    file_name="Bulletin_avec_appreciations.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            os.remove(temp_out)
        os.remove(temp_in.name)
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement : {str(e)}")

st.markdown("""---
**💡 Astuces :**
- Extraire les notes de la classe à partir de PRONOTE et les insérer dans un fichier excel.
- Le fichier excel doit contenir le nom de l'élève en MAJUSCULE et le prénom en miniscule.
- Le fichier excel doit contenir sur la première colonne (NOM et prénom) et sur la deuxième colonne (La moyenne).
- L'appréciation prend en compte la moyenne de l'élève, il ne peut pas refléter le niveau de l'élève, c'est au professeur de l'adapter au profil de l'élève.
""")
