import os
import json
import streamlit as st
from modules.brief import *
import pandas as pd

st.session_state.file_name0 = "./data/employes-data.json"
st.session_state.file_name1 = ""

#if st.session_state.file_name1:
#    st.write("name1")
#if st.session_state.file_name0:
#    st.write("name0")

def main():
    # Questions du brief
    questions_du_brief = [
        "0.Description",
        "1.Importation et Préparation des Données",
        "2.Calcul des Salaires Mensuels",
        "3.Calcul des Statistiques Salariales",
        "4.Affichage des Résultats dans le terminal",
        "5.Affichage des résultats avec interactions",
        "6.Autres fonctions"
    ]

    # Initialisation de l'état de session
    if "c" not in st.session_state:
        st.session_state.c = 0

    # Crée 3 colonnes
    col1, col2, col3 = st.columns([1, 1, 1])

    # Définir un style CSS
    # seulement pour le titre différent de st.title
    css = """
    <style>
    .large-text {
        font-size: 24px;
        font-weight: bold;
        text-align: center;
        line-height: 1.2;
    }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

    # Colonne 1
    with col1:
        if st.button("Précédent"):
            st.session_state.c = (st.session_state.c - 1) % len(questions_du_brief)

    # Colonne 2
    with col2:
        st.markdown('<div class="large-text">Brief n°1</div>', unsafe_allow_html=True)

    # Colonne 3
    with col3:
        if st.button("Suivant"):
            st.session_state.c = (st.session_state.c + 1) % len(questions_du_brief)

    # Affichage du bouton radio
    choix = st.sidebar.radio("Questions du brief :", questions_du_brief, index=st.session_state.c)
    st.session_state.c = int(choix.split('.')[0])

    # Barre de progression
    progress_value = (st.session_state.c + 1) / len(questions_du_brief)
    st.progress(progress_value)

    # Affichage selon l'étape sélectionnée
    if choix == "0.Description":
        st.subheader("Description du brief")
        st.write("Calcul des Salaires Mensuels dans une Entreprise avec plusieurs filiales.")
        st.image("image.webp", width=300)
        st.markdown("**Contexte du projet**")
        st.write("""
            Vous travaillez avec le service comptable d'une entreprise qui possède plusieurs filiales.
            Chaque filiale emploie entre 15 et 20 employés.
            La direction souhaite automatiser le calcul des salaires mensuels des employés en tenant compte
            des heures supplémentaires et obtenir des statistiques salariales pour optimiser la gestion
            des ressources humaines.

            Les données des employés sont fournies dans un fichier JSON contenant : nom, poste, taux horaire,
            heures travaillées, heures contractuelles.
        """)
        st.write("Le fichier `employes-data.json` se trouve dans le dossier `./data`.")

    elif choix == "1.Importation et Préparation des Données":
        st.subheader("Importation des données")

        def get_file_size(path):
            taille = os.path.getsize(path)
            return f"{taille / 1024:.1f} KB"

      #  file_name = os.path.basename(fichier_par_defaut)
      #  file_size = get_file_size(fichier_par_defaut)

        st.session_state.d0 = importation_des_donnees(st.session_state.file_name0)


        uploaded_file = st.file_uploader("Pour changer de fichier ...", type=["json"])
        if uploaded_file is not None:
            # os.getcwd() est le dossier courant où est lancer app.py
            file_path = os.path.join(os.getcwd(), uploaded_file.name)
            st.session_state.file_name1 = file_path

            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())


            #st.session_state.file_name1 = os.path.abspath(uploaded_file.name)
            st.session_state.d1 =  json.load(uploaded_file)
            st.write(f"Le fichier {st.session_state.file_name1} a été téléchargé !")
        else:
            st.write("Le fichier chargé par défaut est : " + st.session_state.file_name0)

        #st.session_state.file_name = file_name
        #st.session_state.d = file_content

        st.write("Grâce à la fonction `importation_des_donnees` qui retourne le dictionnaire correspondant au fichier json :")
        st.code('''
        def importation_des_donnees(fichier: str = 'employes-data.json') -> dict:
            """Charge un fichier JSON et retourne son contenu sous forme de dictionnaire."""
            with open(fichier, encoding='utf-8') as file:
                return json.load(file)
        ''')
        if st.session_state.file_name1:
            st.session_state.d = st.session_state.d1
            st.session_state.file_name = st.session_state.file_name1
        else:
            st.session_state.d = st.session_state.d0
            st.session_state.file_name = st.session_state.file_name0


        st.session_state.cles = list(st.session_state.d.keys())
        st.session_state.valeurs = list(st.session_state.d.values())

        st.write(f"On charge les informations : `d = importation_des_donnees('{st.session_state.file_name}')`.")
        st.write("`valeurs = list(d.values())`")
        st.write("`cles = list(d.keys())`")

        with st.expander("Clés du dictionnaire :"):
            for cle in st.session_state.cles:
                st.write(cle)

        with st.expander("Valeurs du dictionnaire :"):
            st.write(st.session_state.valeurs)

        st.write("Vérification de la cohérence des employés (mêmes clés) :")
        c0 = st.session_state.cles[0]
        l0 = list(st.session_state.d[c0][0].keys())
        b = all(
            l0 == list(employe.keys())
            for filiale in st.session_state.cles
            for employe in st.session_state.d[filiale]
        )
        st.write("Résultat :", b)

        onglets = st.tabs(["Infos fichier", "Contenu fichier"])
        with onglets[0]:
            st.write("Infos sur le fichier :")
            st.write(f"Nom du fichier : {st.session_state.file_name}")
            #st.write(f"Taille : {get_file_size(st.session_state.file_name)}")
        with onglets[1]:
            st.json(st.session_state.d) #file_content)

    elif choix == "2.Calcul des Salaires Mensuels":
        st.subheader("Calcul des Salaires Mensuels")
        st.write("Présentation de la fonction `calcul_salaire` :")
        st.code('''
                def calcul_salaire(**employe) -> Optional[Tuple[float, int]]:
                """Calcule le salaire mensuel d'un employé et retourne un tuple (salaire, option)
                Où option est égale à : heures_supplementaires = heures_travaillees - heures_contractuelles

                   NB: On a fait le choix de ne payer que les heures travaillées et non les heures prévues.

                   On a considéré que le salaire mensuel était le quadruple du salaire hebdomadaire.
                   """
                if not employe:
                    return None

                taux_horaire = employe['hourly_rate']
                heures_contractuelles = employe['contract_hours']
                heures_travaillees = employe['weekly_hours_worked']
                heures_supplementaires = heures_travaillees - heures_contractuelles

                if  heures_supplementaires > 0:
                    salaire_hebdo = (heures_contractuelles + heures_supplementaires * 1.5) * taux_horaire

                else:
                    salaire_hebdo = heures_travaillees * taux_horaire

                return salaire_hebdo * 4, heures_supplementaires
            ''')

        employe_test = st.session_state.d[st.session_state.cles[0]][0]
        st.write(f"Appliquons `calcul_salaire` à l'employé : {employe_test}")
        st.write(f"Résultat : {calcul_salaire(**employe_test)}")

    elif choix == "3.Calcul des Statistiques Salariales":
        st.subheader("Calcul des Statistiques Salariales")
        st.write("Présentation de la fonction `statistiques_filiales` :")
        st.code('''
                def statistiques_filiales(donnees: dict) -> List[Dict[str, Any]]:
                    """ Génère les statistiques pour chaque filiale à partir d'un dictionnaire
                        représantant une filiale.
                        Pour la filiale entrée dans donnée, on collecte le nom de la filiale puis,
                        en parcourant tous ses employes (i.e. employes) dont on calcul le salaire,
                        on calcul le plus grand salaire (=salaireMax)
                        on calcul le plus petit salaire (=salaireMin)

                        on calcul le nombre d'employes  (=effectif)
                        on calcul la somme de tous les salaires (=somme_salaire)
                        ces deux derniers permettent de calculer le salaire moyen (=salaireMoyen)
                        et seront utiles tout à l'heure pour calculer le salaire moyen de toute
                        l'entreprise.
                        Ces informations sont collectées dans le dictionnaire info_filiale.
                        Enfin on concatène tous ces dictionnaires dans la liste stats,
                        que l'on retourne à la fin.
                        """
                    stats = []
                    for filiale, employes in donnees.items():
                        info_filiale = {
                            "nom": filiale,
                            "salaireMax": float('-inf'),
                            "salaireMin": float('inf'),
                            "effectif": 0,
                            "somme_salaire": 0,
                            "listeEmployes": []
                        }

                        for employe in employes:
                            salaire, option = calcul_salaire(**employe)
                            employe.update({
                                "salary": salaire,
                                "info": option
                            })
                            info_filiale["salaireMax"] = max(info_filiale["salaireMax"], salaire)
                            info_filiale["salaireMin"] = min(info_filiale["salaireMin"], salaire)
                            info_filiale["effectif"] += 1
                            info_filiale["somme_salaire"] += salaire
                            info_filiale["listeEmployes"].append(employe)

                        if info_filiale["effectif"] > 0:
                            info_filiale["salaireMoyen"] = info_filiale["somme_salaire"] / info_filiale["effectif"]
                        else:
                            info_filiale["salaireMoyen"] = 0

                        stats.append(info_filiale)

                    return stats
                ''')

    elif choix == "4.Affichage des Résultats dans le terminal":
        st.subheader("Affichage des Résultats")
        st.write("Présentation de la fonction `affichage` :")
        st.code('''
                def affichage(stats: List[Dict[str, Any]]) -> str:
                """
                A partir de la liste retournée par la fonction statistiques_filiales,
                retourne la chaine de caractères output contenant les statistiques
                par filiale puis à la fin les statistiques de l'entreprise que
                l'on affichera par la suite.
                """
                effectif_total = 0
                somme_totale = 0
                salaires_min = []
                salaires_max = []
                output = ""

                for filiale in stats:
                    effectif_total += filiale["effectif"]
                    somme_totale += filiale["somme_salaire"]
                    salaires_min.append(filiale["salaireMin"])
                    salaires_max.append(filiale["salaireMax"])

                    titre = f"Entreprise : {filiale['nom']}"
                    output += titre + "\\n\\n"

                    for employe in filiale["listeEmployes"]:
                        ligne = f"{employe['name']:<10} | {employe['job']:<15} | {'Salaire mensuel:':<18} {employe['salary']:>10,.2f}"
                        output += ligne + "\\n"

                    output += "\\nStatistiques des salaires pour la filiale : {filiale['nom']}:\\n"
                    output += f"Salaire moyen: {filiale['salaireMoyen'] :,.2f}€\\n"
                    output += f"Salaire le plus élevé : {filiale['salaireMax']:,.2f}€\\n\\n"
                    output += f"Salaire le plus bas : {filiale['salaireMin']:,.2f}€\\n"
                    output += '='*60 + '\\n'

                output += f"\\nStatistiques globales de l'entreprise : {effectif_total}\\n"
                output += f"Salaire moyen : {effectif_total/somme_totale}\\n"
                output += f"Salaire le plus élevé : {max(salaires_max):,.2f}€\\n"
                output += f"Salaire le plus bas : {min(salaires_min):,.2f}€\\n"

                return output
                ''')

        st.write("Pour visualiser la chaîne retournée par affichage nous utiliserons : la fonction lancer_affichage :")
        st.code('''
                def lancer_affichage(fichier):
                    """
                    charge le fichier json dans donnees
                    charge statistiques_filiales(donnees) dans stats
                    affiche la chaine ch_out = affichage(stats) dans le terminal, puis la retourne
                    Args:
                        fichier (str, optional): _description_. Defaults to './../data/employes-data.json'.
                    """
                    donnees = importation_des_donnees(fichier)
                    stats = statistiques_filiales(donnees)
                    ch_out = affichage(stats)
                    print(ch_out)
                    ''')

        ch4 = f"lancer_affichage({st.session_state.file_name})\n donne dans le terminal : \n\n"
        st.write(st.session_state.file_name)
        ch3 = lancer_affichage(st.session_state.file_name)


        st.code(ch4 + ch3, language='text')

    elif choix == "5.Affichage des résultats avec interactions":
        st.subheader("Recherche d'un Employé Spécifique")

        # Sélection de la filiale
        filiales = list(st.session_state.d.keys())
        filiale_choisie = st.selectbox("Choisissez une filiale pour la recherche", filiales)

        # On prend la filiale choisie
        stats_filiale = statistiques_filiales(st.session_state.d)[filiales.index(filiale_choisie)]

        # Création du dataframe
        df_employes = pd.DataFrame(stats_filiale["listeEmployes"])
        # Selection du nom
        recherche_nom = st.text_input("Rechercher un employé par son nom :")

        # On filtre / nom
        if recherche_nom:
            resultat = df_employes[df_employes["name"].str.contains(recherche_nom, case=False, na=False)]
            if not resultat.empty:
                st.table(resultat[["name", "job", "salary"]])
            else:
                st.write("Aucun employé trouvé.")
        else:
            st.table(df_employes[["name", "job", "salary"]])

    elif choix == "6.Autres fonctions":
        st.write("A la racine du projet vous pouvez utiliser menu.py (précédé de 'python3' !) ")
        st.write("Toutes les fonctions sont disponibles dans le répertoire modules dans le fichier brief.py")
        st.write("En bonus :")
        st.code('''
                def export_excel(stats: List[Dict[str, Any]], fichier: str = "statistiques_filiales.xlsx") -> None:
    """Exporte les statistiques au format Excel avec ajustement automatique de la largeur des colonnes."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Statistiques"

    # En-têtes des colonnes
    entetes = ["Filiale", "Nom", "Poste", "Salaire", "Info"]
    ws.append(entetes)


    for filiale in stats:
        for employe in filiale["listeEmployes"]:
            ws.append([
                filiale["nom"],
                employe["name"],
                employe["job"],
                employe["salary"],
                employe["info"]
            ])

    # Pour chaque colonne
    for col in range(1, len(entetes) + 1):
        column = get_column_letter(col)
        max_length = 0
        # On cherche la longueur maximale pour chaque ligne
        for row in ws.iter_rows(min_col=col, max_col=col):
            for cell in row:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))

        adjusted_width = (max_length + 3)
        # On ajuste la largeur avec +3 pour que le texte ne soit pas colé
        ws.column_dimensions[column].width = adjusted_width


    wb.save(fichier)
    print(f"Fichier Excel '{fichier}' créé avec succès.")
                ''')

        st.code('''
                def export_csv(stats: List[Dict[str, Any]], fichier: str = "statistiques_filiales.csv") -> None:
    """Exporte les statistiques au format CSV."""
    with open(fichier, mode='w', newline='', encoding='utf-8') as f_csv:
        writer = csv.writer(f_csv, delimiter=';')
        entetes = ["Filiale", "Nom", "Poste", "Salaire", "Info"]
        writer.writerow(entetes)

        for filiale in stats:
            for employe in filiale["listeEmployes"]:
                writer.writerow([
                    filiale["nom"],
                    employe["name"],
                    employe["job"],
                    employe["salary"],
                    employe["info"]
                ])
    print(f"Fichier CSV '{fichier}' créé avec succès.")
                ''')
        st.code('''
             def export_salaries(lStats: List[Dict[str, Any]], filename: str, fmt: str = "csv", separator: str = ";") -> None:
        """
        Exporte les statistiques salariales sous format CSV ou Excel.

        Cette fonction génère un fichier contenant des informations sur les employés, ainsi que des statistiques globales et par filiale,
        incluant le salaire moyen, minimum, maximum, et la somme totale des salaires.

        Les données peuvent être exportées au format CSV ou Excel, selon la valeur du paramètre `fmt`. Les colonnes sont automatiquement ajustées
        dans le fichier Excel pour s'adapter aux données.

        Paramètres:
        -----------
        lStats : List[Dict[str, Any]]
            Liste des statistiques des filiales, où chaque filiale est représentée par un dictionnaire contenant des informations
            sur les employés (nom, poste, salaire, etc.) et des statistiques liées aux salaires.

        filename : str
            Le chemin et le nom du fichier dans lequel les données seront exportées.

        fmt : str, optionnel, par défaut "csv"
            Le format du fichier à générer, soit "csv" pour un fichier CSV, soit "xlsx" pour un fichier Excel.

        separator : str, optionnel, par défaut ";"
            Le séparateur à utiliser dans le fichier CSV (par exemple, ";" ou ",").

        Levée d'exceptions:
        -------------------
        ValueError :
            Si le format `fmt` n'est pas "csv" ou "xlsx".

        FileNotFoundError :
            Si le chemin du fichier spécifié est inaccessible en écriture.

        Exemple:
        --------
        export_salaries(lStats, "salaires.xlsx", fmt="xlsx")
        export_salaries(lStats, "salaires.csv", fmt="csv", separator=",")
        """
        if fmt not in ("csv", "xlsx"):
            raise ValueError("Format non supporté. Utilisez 'csv' ou 'xlsx'.")

        if not os.access(os.path.dirname(filename) or ".", os.W_OK):
            raise FileNotFoundError(f"Chemin inaccessible : {filename}")

        # --- Détail des employés ---
        employes = []
        total_employes = 0
        total_salaire = 0
        salaire_min = float('inf')
        salaire_max = float('-inf')

        for filiale in lStats:
            nb = filiale["effectif"]
            somme = filiale["somme_salaire"]
            total_employes += nb
            total_salaire += somme

            if filiale["salaireMin"] < salaire_min:
                salaire_min = filiale["salaireMin"]
            if filiale["salaireMax"] > salaire_max:
                salaire_max = filiale["salaireMax"]

            for emp in filiale["listeEmployes"]:
                employes.append({
                    "Filiale": filiale["nom"],
                    "Nom": emp["name"],
                    "Poste": emp["job"],
                    "Salaire mensuel": emp["salary"],
                    "Heures sup": emp["info"]
                })

        salaire_moyen = total_salaire / total_employes if total_employes else 0

        if fmt == "csv":
            with open(filename, mode="w", newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=separator)

                # Données
                writer.writerow(["Filiale", "Nom", "Poste", "Salaire mensuel", "Heures sup"])
                for emp in employes:
                    writer.writerow(emp.values())

                # Statistiques globales
                writer.writerow([])
                writer.writerow(["Statistiques globales"])
                writer.writerow(["Effectif total", total_employes])
                writer.writerow(["Salaire moyen", round(salaire_moyen, 2)])
                writer.writerow(["Salaire minimum", salaire_min])
                writer.writerow(["Salaire maximum", salaire_max])
                writer.writerow(["Somme totale des salaires", total_salaire])

                # Statistiques par filiale
                writer.writerow([])
                writer.writerow(["Statistiques par filiale"])
                writer.writerow(["Filiale", "Effectif", "Salaire moyen", "Salaire min", "Salaire max", "Somme salaires"])
                for f in lStats:
                    writer.writerow([
                        f["nom"],
                        f["effectif"],
                        round(f["salaireMoyen"], 2),
                        f["salaireMin"],
                        f["salaireMax"],
                        f["somme_salaire"]
                    ])

            print(f"Fichier CSV '{filename}' créé avec succès.")

        elif fmt == "xlsx":
            wb = Workbook()

            # Données
            ws1 = wb.active
            ws1.title = "Données"
            entetes = ["Filiale", "Nom", "Poste", "Salaire mensuel", "Heures sup"]
            ws1.append(entetes)
            for cell in ws1[1]:
                cell.font = Font(bold=True)
            for emp in employes:
                ws1.append(list(emp.values()))

            # Statistiques globales
            ws2 = wb.create_sheet("Stats Globales")
            ws2.append(["Statistiques globales"])
            ws2.append(["Effectif total", total_employes])
            ws2.append(["Salaire moyen", round(salaire_moyen, 2)])
            ws2.append(["Salaire minimum", salaire_min])
            ws2.append(["Salaire maximum", salaire_max])
            ws2.append(["Somme totale des salaires", total_salaire])
            for cell in ws2[1]:
                cell.font = Font(bold=True)

            # Statistiques par filiale
            ws3 = wb.create_sheet("Stats par Filiale")
            entetes_stats = ["Filiale", "Effectif", "Salaire moyen", "Salaire min", "Salaire max", "Somme salaires"]
            ws3.append(entetes_stats)
            for cell in ws3[1]:
                cell.font = Font(bold=True)

            for f in lStats:
                ws3.append([
                    f["nom"],
                    f["effectif"],
                    round(f["salaireMoyen"], 2),
                    f["salaireMin"],
                    f["salaireMax"],
                    f["somme_salaire"]
                ])

            # Largeur auto
            for ws in [ws1, ws2, ws3]:
                for col in ws.columns:
                    max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = max_len + 2

            wb.save(filename)
            print(f"Fichier Excel '{filename}' créé avec succès.")
             ''')


if __name__ == "__main__":
    main()
