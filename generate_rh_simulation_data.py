from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
import json
import posixpath
import random
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "RH_Collecte_BKO_2026.xlsx"
OUTPUT_PATH = ROOT / "RH_Collecte_SIMULATION_2026.xlsx"
PER_FILIALE_DIR = ROOT / "simulation_data"
SUMMARY_PATH = PER_FILIALE_DIR / "simulation_summary.json"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)

SNAPSHOT_DATE = date(2026, 12, 31)
MONTHS_2026 = [date(2026, month, 1) for month in range(1, 13)]
RNG = random.Random(20260413)

ABSENCE_TYPES = [
    "CONGE_PAYE",
    "MALADIE",
    "ABSENCE_INJUSTIFIEE",
    "FORMATION",
    "EVENEMENT_FAMILIAL",
    "AUTRE",
]

MOTIF_RECRUIT = ["CREATION_POSTE", "REMPLACEMENT", "ACCROISSEMENT", "CONTRAT_PROJET"]
MOTIF_DEPART = ["DEMISSION", "ABANDON_POSTE", "LICENCIEMENT", "FIN_CDD", "RETRAITE", "AUTRES"]

TCDP_EXIT_REASON = {
    "DEMISSION": "DEMISSION",
    "ABANDON_POSTE": "ABANDON",
    "LICENCIEMENT": "LICENCIEMENT",
    "FIN_CDD": "FIN_PROGRAMME",
    "RETRAITE": "RETRAITE",
    "AUTRES": "FIN_PROGRAMME",
}

NATIONALITY_BY_FILIALE = {
    "COO": "BENINOISE",
    "OUA": "BURKINABE",
    "ABJ": "IVOIRIENNE",
    "BJL": "GAMBIENNE",
    "CKY": "GUINEENNE",
    "OXB": "GUINEENNE_BISSAU",
    "BKO": "MALIENNE",
    "NKC": "MAURITANIENNE",
    "NIM": "NIGERIENNE",
    "DKR": "SENEGALAISE",
    "LFW": "TOGOLAISE",
}

EXPAT_NATIONALITIES = ["FRANCAISE", "TCHADIENNE", "CAMEROUNAISE", "AUTRES"]

FILIALES = [
    ("COO", "B\u00e9nin", 48, 8, 5, 1.00),
    ("OUA", "Burkina Faso", 62, 11, 7, 1.02),
    ("ABJ", "C\u00f4te d'Ivoire", 74, 12, 8, 1.08),
    ("BJL", "Gambie", 26, 5, 3, 0.92),
    ("CKY", "Guin\u00e9e", 40, 7, 4, 0.95),
    ("OXB", "Guin\u00e9e-Bissau", 24, 4, 3, 0.90),
    ("BKO", "Mali", 88, 15, 10, 1.15),
    ("NKC", "Mauritanie", 30, 6, 4, 0.97),
    ("NIM", "Niger", 32, 6, 4, 0.96),
    ("DKR", "S\u00e9n\u00e9gal", 70, 12, 8, 1.10),
    ("LFW", "Togo", 36, 7, 4, 0.98),
]

FIRST_NAMES_M = ["Amadou", "Moussa", "Ibrahim", "Yao", "Koffi", "Daouda", "Sekou", "Ousmane", "Mamadou", "Jean", "Serge", "Ali", "Karim", "Habib", "Souleymane"]
FIRST_NAMES_F = ["Awa", "Fatou", "Aminata", "Mariam", "Aissata", "Clarisse", "Nadia", "Sonia", "Nafi", "Djeneba", "Salimata", "Esther", "Grace", "Rita", "Yasmine"]
LAST_NAMES = ["Traore", "Diallo", "Konate", "Ouattara", "Sow", "Camara", "Bah", "Kone", "Toure", "Sarr", "Gueye", "Mensah", "Zongo", "Yameogo", "Kabore", "Coulibaly", "Sissoko", "Sanogo", "Adjovi", "Atayi"]

POSITIONS = [
    {"direction": "Direction Generale", "departement": "Direction Generale", "service": "Coordination", "poste": "Directeur de filiale", "fonction": "Direction", "cat": "CADRE", "contracts": ["CDI"], "levels": ["L4", "L5"], "salary": 42000000},
    {"direction": "Operations", "departement": "Exploitation", "service": "Sites", "poste": "Chef de site", "fonction": "Management operationnel", "cat": "CADRE", "contracts": ["CDI", "CDD"], "levels": ["L3C", "L4"], "salary": 26000000},
    {"direction": "Operations", "departement": "Maintenance", "service": "Atelier", "poste": "Ingenieur maintenance", "fonction": "Maintenance", "cat": "CADRE", "contracts": ["CDI", "CDD"], "levels": ["L3A", "L3B", "L3C"], "salary": 21000000},
    {"direction": "Finance", "departement": "Controle de gestion", "service": "Performance", "poste": "Controleur de gestion", "fonction": "Finance", "cat": "CADRE", "contracts": ["CDI"], "levels": ["L3A", "L3B"], "salary": 18500000},
    {"direction": "Ressources Humaines", "departement": "Administration RH", "service": "Paie et carrieres", "poste": "Responsable RH", "fonction": "Ressources humaines", "cat": "CADRE", "contracts": ["CDI"], "levels": ["L3A", "L3B"], "salary": 19000000},
    {"direction": "Commercial", "departement": "Ventes", "service": "Comptes cles", "poste": "Commercial B2B", "fonction": "Developpement commercial", "cat": "PROF_INTER", "contracts": ["CDI", "CDD"], "levels": ["L2B", "L2C", "L3A"], "salary": 14500000},
    {"direction": "Supply Chain", "departement": "Logistique", "service": "Magasin central", "poste": "Coordinateur logistique", "fonction": "Logistique", "cat": "PROF_INTER", "contracts": ["CDI", "CDD"], "levels": ["L2B", "L2C"], "salary": 13200000},
    {"direction": "Digital", "departement": "Systemes d'information", "service": "Support applicatif", "poste": "Analyste BI", "fonction": "Data reporting", "cat": "PROF_INTER", "contracts": ["CDI", "CDD"], "levels": ["L2C", "L3A"], "salary": 17500000},
    {"direction": "Ressources Humaines", "departement": "Talent", "service": "Formation", "poste": "Charge formation", "fonction": "Developpement RH", "cat": "PROF_INTER", "contracts": ["CDI", "CDD"], "levels": ["L2B", "L2C"], "salary": 12800000},
    {"direction": "Operations", "departement": "Production", "service": "Ligne 1", "poste": "Superviseur production", "fonction": "Production", "cat": "PROF_INTER", "contracts": ["CDI", "CDD"], "levels": ["L2A", "L2B"], "salary": 11800000},
    {"direction": "Support", "departement": "Administration", "service": "Services generaux", "poste": "Assistant administratif", "fonction": "Administration", "cat": "EMPLOYE", "contracts": ["CDI", "CDD"], "levels": ["L1", "L2A"], "salary": 8200000},
    {"direction": "Supply Chain", "departement": "Transport", "service": "Fleet", "poste": "Coordinateur transport", "fonction": "Transport", "cat": "EMPLOYE", "contracts": ["CDI", "CDD"], "levels": ["L1", "L2A"], "salary": 7800000},
    {"direction": "Operations", "departement": "Production", "service": "Ligne 2", "poste": "Technicien production", "fonction": "Production", "cat": "EMPLOYE", "contracts": ["CDI", "CDD", "INTERIM"], "levels": ["L1", "L2A"], "salary": 6600000},
    {"direction": "Support", "departement": "Accueil", "service": "Front office", "poste": "Assistant accueil", "fonction": "Accueil", "cat": "EMPLOYE", "contracts": ["CDD", "STAGE"], "levels": ["L1"], "salary": 5400000},
]


@dataclass
class Employee:
    filiale_code: str
    filiale_name: str
    matricule: str
    user_id: str
    nom: str
    prenom: str
    sexe_code: str
    date_naissance: date
    date_embauche: date
    date_entree: date
    direction: str
    departement: str
    service: str
    poste: str
    fonction: str
    code_analytique: str
    affectation_analytique: str
    cat_conv: str
    type_contrat: str
    date_fin_contrat: date | None
    statut_geo: str
    nationalite: str
    niveau_tcdp: str
    base_salary: float
    entry_type: str | None = None
    recruit_reason: str | None = None
    date_depart: date | None = None
    depart_reason: str | None = None
    pays_mobilite: str | None = None


def qname(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def excel_col(index: int) -> str:
    value = index
    letters: list[str] = []
    while value:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def excel_serial(value: date) -> int:
    return (value - date(1899, 12, 30)).days


def month_end(month_start: date) -> date:
    if month_start.month == 12:
        return date(month_start.year, 12, 31)
    return date(month_start.year, month_start.month + 1, 1) - timedelta(days=1)


def add_months(month_start: date, months: int) -> date:
    year = month_start.year + (month_start.month - 1 + months) // 12
    month = (month_start.month - 1 + months) % 12 + 1
    day = min(month_start.day, month_end(date(year, month, 1)).day)
    return date(year, month, day)


def pick_first_name(sexe_code: str) -> str:
    return RNG.choice(FIRST_NAMES_F if sexe_code == "F" else FIRST_NAMES_M)


def build_employee(filiale_code: str, filiale_name: str, sequence: int, entry_date: date, active_at_end: bool) -> Employee:
    sexe_code = RNG.choices(["H", "F"], weights=[72, 28], k=1)[0]
    profile = RNG.choices(
        POSITIONS,
        weights=[2, 5, 7, 3, 3, 6, 4, 2, 3, 6, 5, 4, 11, 3],
        k=1,
    )[0]
    years_old = {"CADRE": RNG.randint(33, 57), "PROF_INTER": RNG.randint(28, 48), "EMPLOYE": RNG.randint(22, 42)}[profile["cat"]]
    birth_month = RNG.randint(1, 12)
    birth_day = min(RNG.randint(1, 28), month_end(date(SNAPSHOT_DATE.year - years_old, birth_month, 1)).day)
    date_naissance = date(SNAPSHOT_DATE.year - years_old, birth_month, birth_day)

    contract = RNG.choices(
        profile["contracts"],
        weights=[8 if c == "CDI" else 3 if c == "CDD" else 2 if c == "INTERIM" else 1 for c in profile["contracts"]],
        k=1,
    )[0]
    date_fin_contrat = None
    if contract in {"CDD", "INTERIM", "STAGE"}:
        duration_months = 6 if contract == "STAGE" else RNG.choice([6, 12, 18])
        reference_date = entry_date if active_at_end else add_months(entry_date, max(duration_months - 1, 1))
        date_fin_contrat = month_end(add_months(reference_date, max(duration_months - 1, 0)))

    statut_geo = RNG.choices(["LOCAL", "EXPAT"], weights=[88, 12], k=1)[0]
    nationalite = NATIONALITY_BY_FILIALE[filiale_code] if statut_geo == "LOCAL" else RNG.choice(EXPAT_NATIONALITIES)
    nom = RNG.choice(LAST_NAMES)
    prenom = pick_first_name(sexe_code)
    niveau_tcdp = RNG.choice(profile["levels"])
    code_analytique = f"{filiale_code}-{profile['direction'][:3].upper()}-{sequence:04d}"
    matricule = f"{filiale_code}{entry_date.year}{sequence:05d}"
    user_id = f"U{sequence:06d}"

    return Employee(
        filiale_code=filiale_code,
        filiale_name=filiale_name,
        matricule=matricule,
        user_id=user_id,
        nom=nom,
        prenom=prenom,
        sexe_code=sexe_code,
        date_naissance=date_naissance,
        date_embauche=entry_date,
        date_entree=entry_date,
        direction=profile["direction"],
        departement=profile["departement"],
        service=profile["service"],
        poste=profile["poste"],
        fonction=profile["fonction"],
        code_analytique=code_analytique,
        affectation_analytique=profile["service"],
        cat_conv=profile["cat"],
        type_contrat=contract,
        date_fin_contrat=date_fin_contrat,
        statut_geo=statut_geo,
        nationalite=nationalite,
        niveau_tcdp=niveau_tcdp,
        base_salary=profile["salary"] * RNG.uniform(0.9, 1.15),
    )


def generate_people() -> dict[str, list[Employee]]:
    by_filiale: dict[str, list[Employee]] = defaultdict(list)
    sequence = 1
    for filiale_code, filiale_name, end_headcount, hires, departures, _ in FILIALES:
        opening_headcount = end_headcount - hires + departures
        opening_employees: list[Employee] = []
        for _ in range(opening_headcount):
            year = RNG.randint(2015, 2025)
            month = RNG.randint(1, 12)
            emp = build_employee(filiale_code, filiale_name, sequence, date(year, month, RNG.randint(1, 28)), active_at_end=True)
            opening_employees.append(emp)
            sequence += 1

        departing_indexes = set(RNG.sample(range(len(opening_employees)), k=departures))
        for idx, emp in enumerate(opening_employees):
            if idx in departing_indexes:
                depart_month = RNG.randint(1, 12)
                emp.date_depart = date(2026, depart_month, RNG.randint(1, month_end(date(2026, depart_month, 1)).day))
                emp.depart_reason = RNG.choices(MOTIF_DEPART, weights=[6, 1, 2, 2, 1, 1], k=1)[0]
                if emp.depart_reason == "AUTRES" and RNG.random() < 0.35:
                    emp.pays_mobilite = RNG.choice([name for code, name, *_ in FILIALES if code != filiale_code])
                if emp.type_contrat in {"CDD", "INTERIM", "STAGE"} and emp.depart_reason == "FIN_CDD":
                    emp.date_fin_contrat = emp.date_depart
            by_filiale[filiale_code].append(emp)

        for _ in range(hires):
            hire_month = RNG.randint(1, 12)
            emp = build_employee(filiale_code, filiale_name, sequence, date(2026, hire_month, RNG.randint(1, 28)), active_at_end=True)
            emp.entry_type = RNG.choices(["RECRUTEMENT_EXTERNE", "MOBILITE"], weights=[85, 15], k=1)[0]
            emp.recruit_reason = RNG.choices(MOTIF_RECRUIT, weights=[4, 3, 3, 1], k=1)[0]
            by_filiale[filiale_code].append(emp)
            sequence += 1
    return by_filiale


def is_active_on(employee: Employee, ref_date: date) -> bool:
    return employee.date_entree <= ref_date and not (employee.date_depart is not None and employee.date_depart <= ref_date)


def people_for_effectif(by_filiale: dict[str, list[Employee]]) -> list[Employee]:
    people: list[Employee] = []
    for employees in by_filiale.values():
        people.extend([emp for emp in employees if is_active_on(emp, SNAPSHOT_DATE)])
    return sorted(people, key=lambda emp: (emp.filiale_code, emp.nom, emp.prenom, emp.user_id))


def people_for_embauches(by_filiale: dict[str, list[Employee]]) -> list[Employee]:
    hires: list[Employee] = []
    for employees in by_filiale.values():
        hires.extend([emp for emp in employees if emp.date_entree.year == 2026])
    return sorted(hires, key=lambda emp: (emp.date_entree, emp.filiale_code, emp.user_id))


def people_for_departs(by_filiale: dict[str, list[Employee]]) -> list[Employee]:
    departures: list[Employee] = []
    for employees in by_filiale.values():
        departures.extend([emp for emp in employees if emp.date_depart is not None and emp.date_depart.year == 2026])
    return sorted(departures, key=lambda emp: (emp.date_depart, emp.filiale_code, emp.user_id))  # type: ignore[arg-type]


def effectif_rows(by_filiale: dict[str, list[Employee]]) -> list[list[object | None]]:
    return [
        [
            emp.filiale_code,
            emp.filiale_name,
            SNAPSHOT_DATE,
            2026,
            emp.matricule,
            emp.user_id,
            emp.nom,
            emp.prenom,
            emp.sexe_code,
            emp.date_naissance,
            emp.date_embauche,
            emp.date_entree,
            emp.direction,
            emp.departement,
            emp.service,
            emp.poste,
            emp.fonction,
            emp.code_analytique,
            emp.affectation_analytique,
            emp.cat_conv,
            emp.type_contrat,
            emp.date_fin_contrat,
            emp.statut_geo,
            emp.nationalite,
            emp.niveau_tcdp,
        ]
        for emp in people_for_effectif(by_filiale)
    ]


def embauches_rows(by_filiale: dict[str, list[Employee]]) -> list[list[object | None]]:
    rows: list[list[object | None]] = []
    for emp in people_for_embauches(by_filiale):
        rows.append(
            [
                emp.filiale_code,
                emp.filiale_name,
                emp.date_entree,
                date(emp.date_entree.year, emp.date_entree.month, 1),
                2026,
                emp.matricule,
                emp.user_id,
                emp.nom,
                emp.prenom,
                emp.sexe_code,
                emp.date_naissance,
                emp.date_embauche,
                emp.date_entree,
                emp.entry_type or "RECRUTEMENT_EXTERNE",
                emp.recruit_reason or "CREATION_POSTE",
                emp.direction,
                emp.departement,
                emp.service,
                emp.poste,
                emp.fonction,
                emp.code_analytique,
                emp.affectation_analytique,
                emp.cat_conv,
                emp.type_contrat,
                emp.date_fin_contrat,
                emp.statut_geo,
                emp.nationalite,
                emp.niveau_tcdp,
            ]
        )
    return rows


def departs_rows(by_filiale: dict[str, list[Employee]]) -> list[list[object | None]]:
    rows: list[list[object | None]] = []
    for emp in people_for_departs(by_filiale):
        assert emp.date_depart is not None
        rows.append(
            [
                emp.filiale_code,
                emp.filiale_name,
                emp.date_depart,
                date(emp.date_depart.year, emp.date_depart.month, 1),
                2026,
                emp.matricule,
                emp.user_id,
                emp.nom,
                emp.prenom,
                emp.sexe_code,
                emp.date_naissance,
                emp.date_embauche,
                emp.date_entree,
                emp.date_depart,
                emp.depart_reason or "AUTRES",
                emp.pays_mobilite,
                emp.direction,
                emp.departement,
                emp.service,
                emp.poste,
                emp.fonction,
                emp.code_analytique,
                emp.affectation_analytique,
                emp.cat_conv,
                emp.type_contrat,
                emp.date_fin_contrat,
                emp.statut_geo,
                emp.nationalite,
                emp.niveau_tcdp,
            ]
        )
    return rows


def headcount_by_month(by_filiale: dict[str, list[Employee]]) -> dict[tuple[str, date], int]:
    counts: dict[tuple[str, date], int] = {}
    for filiale_code, employees in by_filiale.items():
        for month_start in MONTHS_2026:
            counts[(filiale_code, month_start)] = sum(is_active_on(emp, month_end(month_start)) for emp in employees)
    return counts


def hires_by_month(by_filiale: dict[str, list[Employee]]) -> dict[tuple[str, date], list[Employee]]:
    grouped: dict[tuple[str, date], list[Employee]] = defaultdict(list)
    for emp in people_for_embauches(by_filiale):
        grouped[(emp.filiale_code, date(2026, emp.date_entree.month, 1))].append(emp)
    return grouped


def absence_rows(by_filiale: dict[str, list[Employee]]) -> list[list[object | None]]:
    counts = headcount_by_month(by_filiale)
    weights = {"CONGE_PAYE": 0.46, "MALADIE": 0.22, "ABSENCE_INJUSTIFIEE": 0.07, "FORMATION": 0.12, "EVENEMENT_FAMILIAL": 0.06, "AUTRE": 0.07}
    rows: list[list[object | None]] = []
    for filiale_code, filiale_name, *_ in FILIALES:
        carry_balance = counts[(filiale_code, MONTHS_2026[0])] * 2.0
        for month_start in MONTHS_2026:
            headcount = counts[(filiale_code, month_start)]
            opening = round(carry_balance, 1)
            accrued = round(headcount * 1.75, 1)
            leave_taken = round(headcount * RNG.uniform(0.9, 1.6), 1)
            closing = round(max(opening + accrued - leave_taken, 0), 1)
            carry_balance = closing
            total_abs_hours = round(headcount * RNG.uniform(4.0, 8.5), 1)
            for absence_type in ABSENCE_TYPES:
                rows.append(
                    [
                        filiale_code,
                        filiale_name,
                        month_start,
                        2026,
                        opening if absence_type == "CONGE_PAYE" else 0,
                        leave_taken if absence_type == "CONGE_PAYE" else 0,
                        accrued if absence_type == "CONGE_PAYE" else 0,
                        closing if absence_type == "CONGE_PAYE" else 0,
                        round(total_abs_hours * weights[absence_type], 1),
                        absence_type,
                    ]
                )
    return rows


def formation_rows(by_filiale: dict[str, list[Employee]]) -> list[list[object | None]]:
    counts = headcount_by_month(by_filiale)
    rows: list[list[object | None]] = []
    for filiale_code, filiale_name, *_ in FILIALES:
        for month_start in MONTHS_2026:
            headcount = counts[(filiale_code, month_start)]
            potentials = max(int(round(headcount * RNG.uniform(0.35, 0.55))), 4)
            hp = max(int(round(headcount * RNG.uniform(0.10, 0.18))), 1)
            planned = max(int(round(potentials * RNG.uniform(0.25, 0.45))), 2)
            completed = max(min(planned, int(round(planned * RNG.uniform(0.72, 0.96)))), 1)
            trained_people = max(min(headcount, int(round(completed * RNG.uniform(1.4, 2.6)))), completed)
            training_hours = round(trained_people * RNG.uniform(5.0, 12.0), 1)
            budget = round(planned * RNG.uniform(600000, 1200000), 0)
            actual = round(budget * RNG.uniform(0.72, 0.98), 0)
            rows.append([filiale_code, filiale_name, month_start, 2026, potentials, hp, planned, completed, training_hours, trained_people, actual, budget])
    return rows


def recrutement_rows(by_filiale: dict[str, list[Employee]]) -> tuple[list[list[object | None]], list[list[object | None]]]:
    hires_grouped = hires_by_month(by_filiale)
    mensuel_rows: list[list[object | None]] = []
    detail_rows: list[list[object | None]] = []
    request_seq = 1
    for filiale_code, filiale_name, _, _, _, size_factor in FILIALES:
        for month_start in MONTHS_2026:
            hires = hires_grouped[(filiale_code, month_start)]
            budgeted = len(hires) + RNG.randint(0, 2)
            out_of_budget = max(0, len(hires) - budgeted + RNG.randint(0, 1))
            recruitment_budget = round((budgeted + 1) * 2800000 * size_factor * RNG.uniform(0.9, 1.15), 0)
            budget_used = round(recruitment_budget * min(RNG.uniform(0.68, 1.02), 1.05), 0)
            out_budget_cost = round(out_of_budget * 1400000 * size_factor * RNG.uniform(0.9, 1.2), 0)
            mensuel_rows.append([filiale_code, filiale_name, month_start, 2026, len(hires), budgeted, out_of_budget, recruitment_budget, budget_used, out_budget_cost])

            for hire in hires:
                planned_date = hire.date_entree - timedelta(days=RNG.randint(12, 45))
                duration_months: int | None = None
                if hire.type_contrat == "CDD":
                    duration_months = RNG.choice([6, 12, 18])
                elif hire.type_contrat == "INTERIM":
                    duration_months = RNG.choice([3, 6])
                elif hire.type_contrat == "STAGE":
                    duration_months = 6
                detail_rows.append(
                    [
                        filiale_code,
                        filiale_name,
                        f"REQ-{filiale_code}-{request_seq:04d}",
                        month_start,
                        2026,
                        planned_date,
                        hire.recruit_reason or "CREATION_POSTE",
                        hire.poste,
                        1,
                        f"{hire.departement} / {hire.service}",
                        hire.date_entree,
                        hire.type_contrat,
                        duration_months,
                        "YES" if planned_date.year < 2026 else "NO",
                        round(hire.base_salary * RNG.uniform(0.9, 1.15), 0),
                        "YES" if RNG.random() < 0.12 else "NO",
                    ]
                )
                request_seq += 1
    return mensuel_rows, detail_rows


def payroll_rows(by_filiale: dict[str, list[Employee]]) -> list[list[object | None]]:
    rows: list[list[object | None]] = []
    for filiale_code, filiale_name, _, _, _, size_factor in FILIALES:
        employees = by_filiale[filiale_code]
        for month_start in MONTHS_2026:
            active = [emp for emp in employees if is_active_on(emp, month_end(month_start))]
            payroll = round(sum(emp.base_salary for emp in active) / 12.0, 0)
            budget = round(payroll * RNG.uniform(1.03, 1.12), 0)
            payroll_hub = round(payroll * RNG.uniform(0.05, 0.12), 0)
            overtime_hours = round(len(active) * RNG.uniform(2.5, 6.5), 1)
            overtime_cost = round(overtime_hours * 9500 * size_factor, 0)
            leave_balance = round(len(active) * RNG.uniform(12, 18), 1)
            leave_provision = round(leave_balance * 42000 * size_factor, 0)
            rows.append([filiale_code, filiale_name, month_start, 2026, budget, payroll, payroll_hub, overtime_hours, overtime_cost, leave_balance, leave_provision])
    return rows


def tcdp_rows(by_filiale: dict[str, list[Employee]]) -> tuple[list[list[object | None]], list[list[object | None]], list[list[object | None]], list[list[object | None]]]:
    headcount_rows: list[list[object | None]] = []
    entrees_rows: list[list[object | None]] = []
    sorties_rows: list[list[object | None]] = []
    genre_rows: list[list[object | None]] = []
    levels = ["L1", "L2A", "L2B", "L2C", "L3A", "L3B", "L3C", "L4", "L5"]
    for filiale_code, filiale_name, *_ in FILIALES:
        employees = by_filiale[filiale_code]
        for month_start in MONTHS_2026:
            active = [emp for emp in employees if is_active_on(emp, month_end(month_start))]
            for level in levels:
                level_people = [emp for emp in active if emp.niveau_tcdp == level]
                headcount_rows.append([filiale_code, filiale_name, month_start, 2026, level, len(level_people)])
                genre_rows.append([filiale_code, filiale_name, month_start, 2026, level, sum(emp.sexe_code == "H" for emp in level_people), sum(emp.sexe_code == "F" for emp in level_people)])

        for emp in employees:
            if emp.date_entree.year == 2026:
                entrees_rows.append([emp.filiale_code, emp.filiale_name, emp.date_entree, date(2026, emp.date_entree.month, 1), 2026, emp.matricule, emp.user_id, emp.nom, emp.prenom, emp.niveau_tcdp])
            if emp.date_depart is not None and emp.date_depart.year == 2026:
                sorties_rows.append([emp.filiale_code, emp.filiale_name, emp.date_depart, date(2026, emp.date_depart.month, 1), 2026, emp.matricule, emp.user_id, emp.nom, emp.prenom, emp.niveau_tcdp, TCDP_EXIT_REASON.get(emp.depart_reason or "AUTRES", "FIN_PROGRAMME")])
    headcount_rows.sort(key=lambda row: (row[0], row[2], row[4]))
    entrees_rows.sort(key=lambda row: (row[0], row[2], row[6]))
    sorties_rows.sort(key=lambda row: (row[0], row[2], row[6]))
    genre_rows.sort(key=lambda row: (row[0], row[2], row[4]))
    return headcount_rows, entrees_rows, sorties_rows, genre_rows


def build_datasets() -> dict[str, list[list[object | None]]]:
    by_filiale = generate_people()
    recrutement_mensuel, recrutement_detail = recrutement_rows(by_filiale)
    tcdp_headcount, tcdp_entrees, tcdp_sorties, tcdp_genre = tcdp_rows(by_filiale)
    return {
        "Effectif": effectif_rows(by_filiale),
        "Embauches": embauches_rows(by_filiale),
        "Departs": departs_rows(by_filiale),
        "AbsenceMensuelle": absence_rows(by_filiale),
        "FormationMensuelle": formation_rows(by_filiale),
        "RecrutementMensuel": recrutement_mensuel,
        "RecrutementDetail": recrutement_detail,
        "MasseSalarialeMensuelle": payroll_rows(by_filiale),
        "TCDP_Headcount": tcdp_headcount,
        "TCDP_Entrees": tcdp_entrees,
        "TCDP_Sorties": tcdp_sorties,
        "TCDP_Genre": tcdp_genre,
    }


def load_template_parts() -> dict[str, bytes]:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
    with zipfile.ZipFile(TEMPLATE_PATH) as source_zip:
        return {name: source_zip.read(name) for name in source_zip.namelist()}


def workbook_sheet_paths(file_map: dict[str, bytes]) -> dict[str, str]:
    rels_root = ET.fromstring(file_map["xl/_rels/workbook.xml.rels"])
    rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root.findall(qname(REL_NS, "Relationship"))}
    workbook_root = ET.fromstring(file_map["xl/workbook.xml"])
    sheets = workbook_root.find(qname(MAIN_NS, "sheets"))
    if sheets is None:
        raise ValueError("Workbook XML is missing sheets.")
    result: dict[str, str] = {}
    for sheet in sheets:
        result[sheet.attrib["name"]] = "xl/" + rid_to_target[sheet.attrib[qname(DOC_REL_NS, "id")]]
    return result


def table_path_for_sheet(file_map: dict[str, bytes], sheet_path: str) -> str:
    rels_path = f"xl/worksheets/_rels/{Path(sheet_path).name}.rels"
    rels_root = ET.fromstring(file_map[rels_path])
    for rel in rels_root.findall(qname(REL_NS, "Relationship")):
        if rel.attrib.get("Type", "").endswith("/table"):
            return posixpath.normpath(f"{Path(sheet_path).parent.as_posix()}/{rel.attrib['Target']}")
    raise ValueError(f"No table relationship found for {sheet_path}")


def template_rows_and_cells(sheet_root: ET.Element) -> tuple[ET.Element, ET.Element, ET.Element, dict[int, ET.Element], dict[int, ET.Element]]:
    sheet_data = sheet_root.find(qname(MAIN_NS, "sheetData"))
    if sheet_data is None:
        raise ValueError("Worksheet is missing sheetData.")
    rows = sheet_data.findall(qname(MAIN_NS, "row"))
    header = next(row for row in rows if row.attrib.get("r") == "1")
    stripe_a = next(row for row in rows if row.attrib.get("r") == "2")
    stripe_b = next(row for row in rows if row.attrib.get("r") == "3")

    def cell_map(row: ET.Element) -> dict[int, ET.Element]:
        mapping: dict[int, ET.Element] = {}
        for cell in row.findall(qname(MAIN_NS, "c")):
            letters = "".join(ch for ch in cell.attrib["r"] if ch.isalpha())
            idx = 0
            for char in letters:
                idx = idx * 26 + ord(char) - 64
            mapping[idx] = cell
        return mapping

    return header, stripe_a, stripe_b, cell_map(stripe_a), cell_map(stripe_b)


def clear_children(element: ET.Element) -> None:
    for child in list(element):
        element.remove(child)


def set_cell_value(cell: ET.Element, value: object | None) -> None:
    clear_children(cell)
    if value is None or value == "":
        cell.attrib.pop("t", None)
    elif isinstance(value, date):
        cell.attrib.pop("t", None)
        ET.SubElement(cell, qname(MAIN_NS, "v")).text = str(excel_serial(value))
    elif isinstance(value, bool):
        cell.attrib.pop("t", None)
        ET.SubElement(cell, qname(MAIN_NS, "v")).text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        cell.attrib.pop("t", None)
        ET.SubElement(cell, qname(MAIN_NS, "v")).text = f"{value}"
    else:
        cell.attrib["t"] = "inlineStr"
        inline = ET.SubElement(cell, qname(MAIN_NS, "is"))
        ET.SubElement(inline, qname(MAIN_NS, "t")).text = str(value)


def build_row(row_template: ET.Element, template_cells: dict[int, ET.Element], row_number: int, values: list[object | None]) -> ET.Element:
    row = deepcopy(row_template)
    clear_children(row)
    row.attrib["r"] = str(row_number)
    row.attrib["spans"] = f"1:{len(values)}"
    for idx, value in enumerate(values, start=1):
        cell = deepcopy(template_cells[idx]) if idx in template_cells else ET.Element(qname(MAIN_NS, "c"))
        clear_children(cell)
        cell.attrib["r"] = f"{excel_col(idx)}{row_number}"
        set_cell_value(cell, value)
        row.append(cell)
    return row


def replace_sheet_rows(file_map: dict[str, bytes], sheet_name: str, sheet_path: str, table_path: str, rows_data: list[list[object | None]]) -> None:
    sheet_root = ET.fromstring(file_map[sheet_path])
    table_root = ET.fromstring(file_map[table_path])
    header, stripe_a, stripe_b, stripe_a_cells, stripe_b_cells = template_rows_and_cells(sheet_root)
    sheet_data = sheet_root.find(qname(MAIN_NS, "sheetData"))
    if sheet_data is None:
        raise ValueError(f"{sheet_name}: missing sheetData.")

    for existing_row in list(sheet_data):
        sheet_data.remove(existing_row)
    header.attrib["r"] = "1"
    sheet_data.append(header)

    for index, values in enumerate(rows_data, start=2):
        template_row = stripe_a if index % 2 == 0 else stripe_b
        template_cells = stripe_a_cells if index % 2 == 0 else stripe_b_cells
        sheet_data.append(build_row(template_row, template_cells, index, values))

    last_col = excel_col(len(rows_data[0])) if rows_data else "A"
    last_row = len(rows_data) + 1 if rows_data else 1
    dimension = sheet_root.find(qname(MAIN_NS, "dimension"))
    if dimension is not None:
        dimension.attrib["ref"] = f"A1:{last_col}{last_row}"
    table_root.attrib["ref"] = f"A1:{last_col}{last_row}"
    auto_filter = table_root.find(qname(MAIN_NS, "autoFilter"))
    if auto_filter is not None:
        auto_filter.attrib["ref"] = f"A1:{last_col}{last_row}"

    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
    file_map[table_path] = ET.tostring(table_root, encoding="utf-8", xml_declaration=True)


def write_workbook(target_path: Path, datasets: dict[str, list[list[object | None]]]) -> None:
    file_map = load_template_parts()
    sheet_paths = workbook_sheet_paths(file_map)
    for sheet_name, rows_data in datasets.items():
        sheet_path = sheet_paths[sheet_name]
        table_path = table_path_for_sheet(file_map, sheet_path)
        replace_sheet_rows(file_map, sheet_name, sheet_path, table_path, rows_data)

    target_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target_path, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
        for name, payload in file_map.items():
            output_zip.writestr(name, payload)


def split_by_filiale(datasets: dict[str, list[list[object | None]]]) -> dict[str, dict[str, list[list[object | None]]]]:
    per_filiale: dict[str, dict[str, list[list[object | None]]]] = {}
    for filiale_code, _, *_ in FILIALES:
        per_filiale[filiale_code] = {}
        for sheet_name, rows in datasets.items():
            per_filiale[filiale_code][sheet_name] = [row for row in rows if row and row[0] == filiale_code]
    return per_filiale


def build_summary(datasets: dict[str, list[list[object | None]]]) -> dict[str, object]:
    per_filiale = split_by_filiale(datasets)
    summary: dict[str, object] = {
        "snapshot_date": SNAPSHOT_DATE.isoformat(),
        "source_template": TEMPLATE_PATH.name,
        "consolidated_workbook": OUTPUT_PATH.name,
        "tables": {sheet: len(rows) for sheet, rows in datasets.items()},
        "filiales": {},
    }
    filiales_summary = summary["filiales"]
    assert isinstance(filiales_summary, dict)
    for filiale_code, filiale_name, *_ in FILIALES:
        filiales_summary[filiale_code] = {
            "name": filiale_name,
            "tables": {sheet: len(rows) for sheet, rows in per_filiale[filiale_code].items()},
        }
    return summary


def main() -> None:
    datasets = build_datasets()
    write_workbook(OUTPUT_PATH, datasets)

    per_filiale = split_by_filiale(datasets)
    PER_FILIALE_DIR.mkdir(parents=True, exist_ok=True)
    for filiale_code, _, *_ in FILIALES:
        write_workbook(PER_FILIALE_DIR / f"RH_Collecte_{filiale_code}_2026.xlsx", per_filiale[filiale_code])

    SUMMARY_PATH.write_text(json.dumps(build_summary(datasets), indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Generated consolidated workbook: {OUTPUT_PATH}")
    print(f"Generated per-filiale workbooks in: {PER_FILIALE_DIR}")
    print(f"Summary written to: {SUMMARY_PATH}")


if __name__ == "__main__":
    main()
