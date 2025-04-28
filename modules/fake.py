"""
fake.py

Charge un dataset JSON d'employés, calcule les salaires annuels et affiche un extrait du dataset.
Peut également générer un nouveau dataset factice si l'option --fake est utilisée.
"""

import argparse
import subprocess
import time
import json

def load_data(filepath):
    """
    Charge les données d'un fichier JSON.
    """
    start = time.time()
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    end = time.time()
    print(f"Chargement des données : {end - start:.2f} secondes")
    return data

def calculate_salaries(data):
    """
    Calcule le salaire annuel de chaque employé.
    """
    start = time.time()
    for company, employees in data.items():
        for employee in employees:
            salary = employee["hourly_rate"] * employee["weekly_hours_worked"] * 52
            employee["annual_salary"] = salary
    end = time.time()
    print(f"Calcul des salaires : {end - start:.2f} secondes")

def main():
    """
    Point d'entrée principal du script.
    Gère l'argument --fake, charge les données, calcule les salaires et affiche un extrait.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("--fake", action="store_true", help="Générer un faux dataset")
    args = parser.parse_args()

    if args.fake:
        print("Génération du dataset factice...")
        subprocess.run(["python3", "gen_fake.py"], check=True)

    print("Chargement du fichier...")
    data = load_data("fake_employees.json")

    print("Calcul des salaires...")
    calculate_salaries(data)

    # Rapport simple
    print(f"\nExtrait du dataset :\n{json.dumps({k: v[:3] for k, v in data.items()}, indent=4, ensure_ascii=False)}")

if __name__ == "__main__":
    main()
