import openpyxl

def comparer_fichiers_col1_col4_diff(fichier1, fichier2):
    wb1 = openpyxl.load_workbook(fichier1, data_only=True)
    wb2 = openpyxl.load_workbook(fichier2, data_only=True)

    feuille1 = wb1.active
    feuille2 = wb2.active

    data1 = {}
    data2 = {}

    for row in feuille1.iter_rows(min_row=1, max_col=4):
        nom_cell = row[0]
        tel_cell = row[3] if len(row) >= 4 else None
        if nom_cell.value:
            nom = str(nom_cell.value).strip().lower()
            tel = str(tel_cell.value).strip().replace(" ", "") if tel_cell and tel_cell.value else ""
            data1[nom] = tel

    for row in feuille2.iter_rows(min_row=1, max_col=4):
        nom_cell = row[0]
        tel_cell = row[3] if len(row) >= 4 else None
        if nom_cell.value:
            nom = str(nom_cell.value).strip().lower()
            tel = str(tel_cell.value).strip().replace(" ", "") if tel_cell and tel_cell.value else ""
            data2[nom] = tel

    communs = set(data1.keys()) & set(data2.keys())
    differences = []

    for nom in communs:
        tel1 = data1[nom]
        tel2 = data2[nom]
        if tel1 != tel2:
            differences.append((nom, tel1, tel2))

    if differences:
        print("Différences trouvées pour les numéros de téléphone (colonne 4) :\n")
        for nom, tel1, tel2 in differences:
            print(f"Nom : {nom}")
            print(f" - Fichier 1 : {tel1}")
            print(f" - Fichier 2 : {tel2}\n")
    else:
        print("✅ Aucun numéro de téléphone différent trouvé pour les noms communs.")

if __name__ == "__main__":
    fichier1 = r"C:\Users\gviguerie\Downloads\RELANCES (2)\relance dernière version.xlsm"
    fichier2 = r"C:\Users\gviguerie\Desktop\RELANCES FONDS.xlsx"
    comparer_fichiers_col1_col4_diff(fichier1, fichier2) 
