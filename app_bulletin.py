import streamlit as st
import openpyxl
import tempfile
import os

def appreciation_bull(nom, moyenne):
    if moyenne is None or moyenne == "":
        return "Absent du trimestre. Peu d'évaluation possible."
    try:
        m = float(moyenne)
    except Exception:
        return "Erreur sur la moyenne."
    if m >= 17:
        return f"Excellent trimestre pour {nom}. Élève très sérieux(se) et appliqué(e). Félicitations !"
    elif m >= 15:
        return f"Très bon trimestre pour {nom}. Travail sérieux et régulier. Bonne participation, continuez ainsi !"
    elif m >= 13:
        return f"Bon trimestre dans l'ensemble. {nom} est impliqué(e), il faut continuer dans cette voie."
    elif m >= 11:
        return f"Résultat satisfaisant mais {nom} peut mieux faire avec plus d'attention et de régularité."
    elif m >= 8:
        return f"Trimestre correct mais le minimum est fait. {nom} doit s'investir davantage pour progresser."
    elif m > 0:
        return f"Trimestre difficile, lacunes importantes. {nom} doit se mobiliser et demander de l'aide pour progresser."
    else:
        return f"Données manquantes ou non valides."

def generer_appreciations_excel(input_path, output_path):
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    # Colonnes : suppose nom en A (0), moyenne en B (1)
    entetes = [str(cell.value).lower() for cell in ws[1]]
    col_nom = next((i for i, h in enumerate(entetes) if "élève" in h or "nom" in h), 0)
    col_moy = next((i for i, h in enumerate(entetes) if "moyenne" in h), 1)
    # Ajout d'en-tête appréciation si absent
    if not any("appréc" in str(cell.value).lower() for cell in ws[1]):
        ws.cell(row=1, column=ws.max_column+1, value="Appréciation")
        col_app = ws.max_column - 1
    else:
        col_app = next(i for i, h in enumerate(entetes) if "appréc" in h)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        nom = row[col_nom].value
        moyenne = row[col_moy].value
        appreciation = appreciation_bull(nom, moyenne)
        ws.cell(row=row[0].row, column=col_app+1, value=appreciation)

    wb.save(output_path)

 
st.title("Générateur d'appréciations pour bulletins scolaires")
st.markdown("""
Bienvenue sur votre espace enseignant !  
 1️⃣ Déposez un fichier .xlsx avec les notes.  
 2️⃣ Générez des appréciations personnalisées.  
 3️⃣ Téléchargez un document prêt à l’emploi.  
""")
uploaded_file = st.file_uploader("Déposez ici votre fichier Excel (.xlsx) contenant le nom des élèves, leurs moyennes, obtenez un bulletin enrichi automatiquement.", type=["xlsx"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_in:
        temp_in.write(uploaded_file.read())
        temp_in.flush()
        temp_out = temp_in.name.replace(".xlsx", "_appreciations.xlsx")
        generer_appreciations_excel(temp_in.name, temp_out)
        st.success("Fichier traité avec succès ! Cliquez ci-dessous pour le télécharger :")
        with open(temp_out, "rb") as f:
            st.download_button(
                label="Télécharger le bulletin avec appréciations",
                data=f,
                file_name="Bulletin_avec_appreciations.xlsx"
            )
        os.remove(temp_out)
    os.remove(temp_in.name)
