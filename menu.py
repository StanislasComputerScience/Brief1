from modules.brief import *

def menu():
    fichier = input("Entrez le chemin du fichier JSON des employés, ex : data/employes-data.json : ")
    while True:
        print("\nQue souhaitez-vous faire ?")
        print("1. Afficher les statistiques dans le terminal")
        print("2. Exporter les statistiques en CSV")
        print("3. Exporter les statistiques en Excel")
        print("4. Exporter avec la nouvelle méthode (export_salaries)")
        print("5. Quitter")


        choix = input("Votre choix : ")

        if choix == "1":
            lancer_affichage(fichier)

        elif choix == "2":
            stats = statistiques_filiales(importation_des_donnees(fichier))
            nom_fichier = input("Entrez le nom du fichier CSV à créer (ex: resultats.csv) : ")
            separateur = input("Entrez le séparateur (par défaut ';') : ") or ";"
            export_salaries(stats, nom_fichier, fmt="csv", separator=separateur)

        elif choix == "3":
            stats = statistiques_filiales(importation_des_donnees(fichier))
            nom_fichier = input("Entrez le nom du fichier Excel à créer (ex: resultats.xlsx) : ")
            export_salaries(stats, nom_fichier, fmt="xlsx")

        elif choix == "4":
            stats = statistiques_filiales(importation_des_donnees(fichier))
            nom_fichier = input("Entrez le nom du fichier à créer (ex: resultats.csv ou resultats.xlsx) : ")
            format_export = input("Format ? (csv ou xlsx) [csv par défaut] : ").lower() or "csv"
            if format_export == "csv":
                separateur = input("Entrez le séparateur (par défaut ';') : ") or ";"
                export_salaries(stats, nom_fichier, fmt="csv", separator=separateur)
            else:
                export_salaries(stats, nom_fichier, fmt="xlsx")

        elif choix == "5":
            print("Au revoir !")
            break

        else:
            print("Choix invalide, veuillez réessayer.")

if __name__ == "__main__":
    menu()
