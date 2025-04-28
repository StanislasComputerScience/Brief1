"""
gen_fake.py

Génère un faux dataset d'employés et l'enregistre dans un fichier JSON.
"""

import json
import random
from faker import Faker
from tqdm import tqdm

def generate_employees(num_employees=100000):
    """
    Génère un dictionnaire d'employés factices regroupés par entreprise.
    """
    fake = Faker('fr_FR')
    companies = ["TechCorp", "DesignWorks", "ProjectLead"]
    jobs = ["Développeur", "Designer", "Testeur", "Chef de projet", "Manager", "Analyste"]

    data = {company: [] for company in companies}

    for _ in tqdm(range(num_employees), desc="Génération des employés"):
        company = random.choice(companies)
        employee = {
            "name": fake.first_name(),
            "job": random.choice(jobs),
            "hourly_rate": random.choice([28, 30, 35, 40, 50, 55]),
            "weekly_hours_worked": random.randint(30, 50),
            "contract_hours": random.choice([35, 37, 40])
        }
        data[company].append(employee)

    return data

if __name__ == "__main__":
    employees = generate_employees()
    with open("fake_employees.json", "w", encoding="utf-8") as f:
        json.dump(employees, f, ensure_ascii=False, indent=4)
