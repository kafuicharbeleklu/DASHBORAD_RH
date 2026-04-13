from __future__ import annotations

from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile


ROOT = Path(__file__).resolve().parent
OUTPUT_PATH = ROOT / "MODELE_COLLECTE_RH_MULTI_FILIALES.xlsx"


SHEET_ORDER = [
    "README",
    "Effectif",
    "Embauches",
    "Departs",
    "AbsenceMensuelle",
    "FormationMensuelle",
    "RecrutementMensuel",
    "RecrutementDetail",
    "MasseSalarialeMensuelle",
    "TCDP_Headcount",
    "TCDP_Entrees",
    "TCDP_Sorties",
    "TCDP_Genre",
    "Referentiels",
]


DATA_SHEETS = OrderedDict(
    [
        (
            "Effectif",
            {
                "table_name": "tbl_Effectif",
                "headers": [
                    "FilialeCode", "FilialeName", "SnapshotDate", "Matricule", "UserID", "Nom", "Prenom",
                    "SexeCode", "DateNaissance", "DateEmbauche", "DateEntree", "Direction", "Departement",
                    "Service", "PosteHarmonise", "Fonction", "CodeAnalytique", "AffectationAnalytique",
                    "CategorieConventionnelleCode", "TypeContratCode", "DateFinContrat", "StatutGeoCode",
                    "NationaliteCode", "NiveauTCDPCode",
                ],
            },
        ),
        (
            "Embauches",
            {
                "table_name": "tbl_Embauches",
                "headers": [
                    "FilialeCode", "FilialeName", "EventDate", "MonthStartDate", "Matricule", "UserID", "Nom",
                    "Prenom", "SexeCode", "DateNaissance", "DateEmbauche", "DateEntree", "TypeEntreeSocieteCode",
                    "MotifRecrutementCode", "Direction", "Departement", "Service", "PosteHarmonise", "Fonction",
                    "CodeAnalytique", "AffectationAnalytique", "CategorieConventionnelleCode", "TypeContratCode",
                    "DateFinContrat", "StatutGeoCode", "NationaliteCode", "NiveauTCDPCode",
                ],
            },
        ),
        (
            "Departs",
            {
                "table_name": "tbl_Departs",
                "headers": [
                    "FilialeCode", "FilialeName", "EventDate", "MonthStartDate", "Matricule", "UserID", "Nom",
                    "Prenom", "SexeCode", "DateNaissance", "DateEmbauche", "DateEntree", "DateDepart",
                    "MotifDepartCode", "PaysMobilite", "Direction", "Departement", "Service", "PosteHarmonise",
                    "Fonction", "CodeAnalytique", "AffectationAnalytique", "CategorieConventionnelleCode",
                    "TypeContratCode", "DateFinContrat", "StatutGeoCode", "NationaliteCode", "NiveauTCDPCode",
                ],
            },
        ),
        (
            "AbsenceMensuelle",
            {
                "table_name": "tbl_AbsenceMensuelle",
                "headers": [
                    "FilialeCode", "FilialeName", "MonthStartDate", "OpeningLeaveDays", "LeaveDaysTaken",
                    "LeaveDaysAccrued", "ClosingLeaveDays", "AbsenceHours", "AbsenceRate",
                ],
            },
        ),
        (
            "FormationMensuelle",
            {
                "table_name": "tbl_FormationMensuelle",
                "headers": [
                    "FilialeCode", "FilialeName", "MonthStartDate", "NombrePotentiels", "NombreHP",
                    "PlannedTrainings", "CompletedTrainings", "TrainingHours", "TrainedPeople",
                    "TrainingCostActual", "TrainingBudget",
                ],
            },
        ),
        (
            "RecrutementMensuel",
            {
                "table_name": "tbl_RecrutementMensuel",
                "headers": [
                    "FilialeCode", "FilialeName", "MonthStartDate", "Recruitments", "RecruitmentsBudget",
                    "RecruitmentsOutOfBudget", "RecruitmentBudget", "BudgetUsed", "OutOfBudgetCost",
                ],
            },
        ),
        (
            "RecrutementDetail",
            {
                "table_name": "tbl_RecrutementDetail",
                "headers": [
                    "FilialeCode", "FilialeName", "RecruitmentRequestID", "MonthStartDate",
                    "PlannedRecruitmentDate", "RecruitmentReasonCode", "PositionTitle",
                    "HarmonizedPositionCount", "DepartmentServiceLOB", "EffectiveEntryDate", "TypeContratCode",
                    "ContractDurationMonths", "PriorYearCarryOverStatus", "EstimatedAnnualCostBudget",
                    "IsOutOfBudget",
                ],
            },
        ),
        (
            "MasseSalarialeMensuelle",
            {
                "table_name": "tbl_MasseSalarialeMensuelle",
                "headers": [
                    "FilialeCode", "FilialeName", "MonthStartDate", "PayrollBudget", "Payroll", "PayrollHUB",
                    "OvertimeHours", "OvertimeCost", "LeaveBalance", "LeaveProvision",
                ],
            },
        ),
        (
            "TCDP_Headcount",
            {
                "table_name": "tbl_TCDP_Headcount",
                "headers": ["FilialeCode", "FilialeName", "MonthStartDate", "TCDPLevelCode", "Headcount"],
            },
        ),
        (
            "TCDP_Entrees",
            {
                "table_name": "tbl_TCDP_Entrees",
                "headers": ["FilialeCode", "FilialeName", "EventDate", "MonthStartDate", "UserID", "Nom", "Prenom", "TCDPLevelCode"],
            },
        ),
        (
            "TCDP_Sorties",
            {
                "table_name": "tbl_TCDP_Sorties",
                "headers": [
                    "FilialeCode", "FilialeName", "EventDate", "MonthStartDate", "UserID", "Nom", "Prenom",
                    "TCDPLevelCode", "TCDPExitReasonCode",
                ],
            },
        ),
        (
            "TCDP_Genre",
            {
                "table_name": "tbl_TCDP_Genre",
                "headers": ["FilialeCode", "FilialeName", "MonthStartDate", "TCDPLevelCode", "MenCount", "WomenCount"],
            },
        ),
    ]
)


REFERENCE_LISTS = OrderedDict(
    [
        ("Filiale", {"code_name": "List_FilialeCode", "label_name": "List_FilialeName", "title": "Filiales", "values": [
            ("BEN", "Bénin"), ("BFA", "Burkina Faso"), ("CIV", "Côte d'Ivoire"), ("ERS", "Ersum"),
            ("FRA", "France"), ("FRZ", "Free Zone"), ("GUI", "Guinée"), ("MLI", "Mali"),
            ("MRT", "Mauritania"), ("NI", "NI"), ("NER", "Niger"), ("SEN", "Sénégal"),
            ("TGO", "Togo"), ("UEI", "UEI"),
        ]}),
        ("Sexe", {"code_name": "List_SexeCode", "label_name": "List_SexeLabel", "title": "Sexe", "values": [
            ("H", "Masculin"), ("F", "Féminin"),
        ]}),
        ("TypeContrat", {"code_name": "List_TypeContratCode", "label_name": "List_TypeContratLabel", "title": "Type de contrat", "values": [
            ("CDI", "CDI"), ("CDD", "CDD"), ("CJD", "CJD"), ("STAGE", "Stage"), ("INTERIM", "Intérim"), ("AUTRE", "Autre"),
        ]}),
        ("StatutGeo", {"code_name": "List_StatutGeoCode", "label_name": "List_StatutGeoLabel", "title": "Statut géographique", "values": [
            ("LOCAL", "Local"), ("EXPAT", "Expatrié"),
        ]}),
        ("CategorieConventionnelle", {"code_name": "List_CategorieConventionnelleCode", "label_name": "List_CategorieConventionnelleLabel", "title": "Catégorie conventionnelle", "values": [
            ("CADRE", "Cadre"), ("PROF_INTER", "Profession intermédiaire"), ("AGENT_MAITRISE", "Agent de maitrise"),
            ("EMPLOYE", "Employé"), ("STAGIAIRE", "Stagiaire"), ("AUTRE", "Autre"),
        ]}),
        ("TypeEntreeSociete", {"code_name": "List_TypeEntreeSocieteCode", "label_name": "List_TypeEntreeSocieteLabel", "title": "Type d'entrée dans la société", "values": [
            ("RECRUTEMENT_EXTERNE", "Recrutement externe"), ("MOBILITE", "Mobilité"),
            ("PROGRAMME_TCDP", "Programme TCDP"), ("AUTRE", "Autre"),
        ]}),
        ("MotifRecrutement", {"code_name": "List_MotifRecrutementCode", "label_name": "List_MotifRecrutementLabel", "title": "Motifs de recrutement", "values": [
            ("CONTRAT_PROJET", "Contrat de projet"),
            ("ACCROISSEMENT_TEMPORAIRE", "Besoin lié à un accroissement temporaire d'activité"),
            ("ACCROISSEMENT_SAISONNIER", "Besoin lié à un accroissement saisonnier d'activité"),
            ("REMPLACEMENT_PERMANENT", "Remplacement d'un employé sur un emploi permanent"),
            ("CREATION_POSTE", "Création d'un nouveau poste"),
            ("PROGRAMME_TCDP", "Programme TCDP"),
        ]}),
        ("MotifDepart", {"code_name": "List_MotifDepartCode", "label_name": "List_MotifDepartLabel", "title": "Motifs de départ", "values": [
            ("ABANDON_POSTE", "Abandon de poste"), ("REAFFECTATION", "Réaffectation"), ("AUTRES", "Autres"),
            ("DECES", "Décès"), ("DEMISSION_RAISON_PERSO", "Démission/Raisons personnelles"),
            ("DEMISSION_HIERARCHIE", "Démission/Hiérarchie"), ("DEMISSION_MOBILITE_RESEAU", "Démission/mobilité réseau"),
            ("DEMISSION_SALAIRE", "Démission/Salaire"), ("RETRAITE_INIT_EMPLOYEUR", "Retraite à l'initiative de l'employeur"),
            ("FIN_CONTRAT_ENTREPRISE", "Fin de contrat entreprise"), ("FIN_CONTRAT_APPRENTISSAGE", "Fin de contrat d'apprentissage"),
            ("FIN_CDD", "Fin de CDD"), ("FIN_CDD_SANS_RENOUVELLEMENT", "Fin de CDD - Pas de renouvellement proposé"),
            ("FIN_PERIODE_ESSAI_ENTREPRISE", "Fin de période d'essai à l'initiative de l'entreprise"),
            ("FIN_PERIODE_ESSAI_SALARIE", "Fin de période d'essai à l'initiative du salarié"),
            ("FIN_STAGE", "Fin de Stage"), ("LICENCIEMENT_AUTRE", "Licenciement pour autre motif"),
            ("LICENCIEMENT_ECONOMIQUE", "Licenciement économique"), ("LICENCIEMENT_FAUTE_LOURDE", "Licenciement pour faute lourde"),
            ("DISPONIBILITE_PERSO", "Mis en Disponibilité/motif personnel"),
            ("DISPONIBILITE_MOBILITE_RESEAU", "Mis en Disponibilité/mobilité réseau"),
            ("FIN_CHANTIER", "Fin de chantier"), ("RUPTURE_ANTICIPEE_CDD_EMPLOYEUR", "Rupture anticipée CDD sur initiative employeur"),
            ("RUPTURE_ANTICIPEE_CDD_SALARIE", "Rupture anticipée CDD ou apprenti à l'initiative salarié"),
            ("RUPTURE_CONVENTIONNELLE", "Rupture conventionnelle"), ("RETRAITE", "Retraite"),
        ]}),
        ("TCDPLevel", {"code_name": "List_TCDPLevelCode", "label_name": "List_TCDPLevelLabel", "title": "Niveaux TCDP", "values": [
            ("L1", "L1"), ("L2A", "L2A"), ("L2B", "L2B"), ("L2C", "L2C"), ("L3A", "L3A"), ("L3B", "L3B"), ("L3C", "L3C"), ("L4", "L4"), ("L5", "L5"),
        ]}),
        ("Nationalite", {"code_name": "List_NationaliteCode", "label_name": "List_NationaliteLabel", "title": "Nationalités", "values": [
            ("MALIENNE", "Malienne"), ("BURKINABE", "Burkinabè"), ("IVOIRIENNE", "Ivoirienne"),
            ("NIGERIENNE", "Nigérienne"), ("GHANEENNE", "Ghanéenne"), ("INDONESIENNE", "Indonésienne"),
            ("SUD_AFRICAINE", "Sud africaine"), ("GUINEENNE", "Guinéenne"), ("PHILIPPINE", "Philippine"),
            ("TUNISIENNE", "Tunisienne"), ("BRESILIENNE", "Brésilienne"), ("MALGACHE", "Malgache"),
            ("CHINOISE", "Chinoise"), ("TANZANIENNE", "Tanzanienne"), ("BENINOISE", "Beninoise"),
            ("TOGOLAISE", "Togolaise"), ("CAMEROUNAISE", "Camerounaise"), ("COLOMBIENNE", "Colombienne"),
            ("FRANCAISE", "Française"), ("MAURITANIENNE", "Mauritanienne"), ("AUTRE", "Autre"),
        ]}),
        ("TCDPExitReason", {"code_name": "List_TCDPExitReasonCode", "label_name": "List_TCDPExitReasonLabel", "title": "Motifs de sortie TCDP", "values": [
            ("DEMISSION", "Démission"), ("ABANDON", "Abandon"), ("FIN_CONTRAT", "Fin contrat"),
            ("RETRAITE", "Retirer du programme pour départ à la retraite"),
            ("LACUNES_VALIDATION", "Retirer du programme pour lacunes dans les validations"),
            ("AUTRE", "Autre"),
        ]}),
        ("YesNo", {"code_name": "List_YesNoCode", "label_name": "List_YesNoLabel", "title": "Oui/Non", "values": [
            ("YES", "Yes"), ("NO", "No"),
        ]}),
    ]
)


VALIDATION_TARGETS = {
    "FilialeCode": "List_FilialeCode",
    "FilialeName": "List_FilialeName",
    "SexeCode": "List_SexeLabel",
    "TypeContratCode": "List_TypeContratLabel",
    "StatutGeoCode": "List_StatutGeoLabel",
    "CategorieConventionnelleCode": "List_CategorieConventionnelleLabel",
    "TypeEntreeSocieteCode": "List_TypeEntreeSocieteLabel",
    "MotifRecrutementCode": "List_MotifRecrutementLabel",
    "MotifDepartCode": "List_MotifDepartLabel",
    "TCDPLevelCode": "List_TCDPLevelCode",
    "NationaliteCode": "List_NationaliteLabel",
    "TCDPExitReasonCode": "List_TCDPExitReasonLabel",
    "IsOutOfBudget": "List_YesNoLabel",
}


README_ROWS = [
    ("MODELE DE COLLECTE RH MULTI-FILIALES", 3),
    ("Version", 1),
    ("1.0", 0),
    ("Objectif", 1),
    ("Ce classeur est le gabarit standard de collecte RH a dupliquer pour chaque filiale.", 2),
    ("Regles de remplissage", 1),
    ("1. Ne pas renommer les onglets ni les en-tetes.", 2),
    ("2. Saisir les donnees uniquement dans les tableaux Excel de chaque feuille.", 2),
    ("3. Ne pas inserer de lignes de titre, de totaux ou de sous-tableaux dans les feuilles de saisie.", 2),
    ("4. Utiliser les listes deroulantes pour les champs controles.", 2),
    ("5. Les dates mensuelles doivent etre renseignees au premier jour du mois.", 2),
    ("6. La feuille Referentiels est cachee et reservee a l'administration du modele.", 2),
    ("Onglets de collecte", 1),
    ("Effectif: photo des employes actifs a une date donnee.", 2),
    ("Embauches: mouvements d'entree dans la filiale.", 2),
    ("Departs: mouvements de sortie de la filiale.", 2),
    ("AbsenceMensuelle: indicateurs mensuels d'absence et de conges.", 2),
    ("FormationMensuelle: indicateurs mensuels de formation.", 2),
    ("RecrutementMensuel: indicateurs de suivi mensuel du recrutement.", 2),
    ("RecrutementDetail: detail des besoins ou recrutements individuels.", 2),
    ("MasseSalarialeMensuelle: indicateurs mensuels de masse salariale.", 2),
    ("TCDP_Headcount, TCDP_Entrees, TCDP_Sorties, TCDP_Genre: suivi du programme TCDP.", 2),
    ("Notes importantes", 1),
    ("Les colonnes suffixees par Code utilisent des listes de valeurs controlees.", 2),
    ("Les calculs d'age, d'anciennete, de taux et de ratios seront realises dans Power BI et non dans Excel.", 2),
    ("Le futur filtre par filiale dans le dashboard depend de la qualite des champs FilialeCode et FilialeName.", 2),
]


def excel_col(index: int) -> str:
    value = index
    letters = []
    while value:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def cell_ref(row: int, col: int) -> str:
    return f"{excel_col(col)}{row}"


def xml_text(value: str) -> str:
    value = escape(value)
    if value.startswith(" ") or value.endswith(" "):
        return f'<t xml:space="preserve">{value}</t>'
    return f"<t>{value}</t>"


def inline_string_cell(row: int, col: int, value: str, style: int = 0) -> str:
    return f'<c r="{cell_ref(row, col)}" t="inlineStr" s="{style}"><is>{xml_text(value)}</is></c>'


def build_styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF1F2937"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0B5CAB"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF3F4F6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F2937"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""


def build_content_types_xml(sheet_count: int, table_count: int) -> str:
    overrides = [
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    for index in range(1, table_count + 1):
        overrides.append(f'<Override PartName="/xl/tables/table{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    )


def build_root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def build_app_xml() -> str:
    parts = "".join(f"<vt:lpstr>{escape(name)}</vt:lpstr>" for name in SHEET_ORDER)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>OpenAI Codex</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(SHEET_ORDER)}</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="{len(SHEET_ORDER)}" baseType="lpstr">{parts}</vt:vector></TitlesOfParts>
</Properties>
"""


def build_core_xml() -> str:
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Modèle de collecte RH multi-filiales</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>
</cp:coreProperties>
"""


def build_workbook_xml(named_ranges: list[tuple[str, str]], referentiels_sheet_id: int) -> str:
    sheets_xml = []
    for index, name in enumerate(SHEET_ORDER, start=1):
        hidden = ' state="hidden"' if index == referentiels_sheet_id else ""
        sheets_xml.append(f'<sheet name="{escape(name)}" sheetId="{index}"{hidden} r:id="rId{index}"/>')
    defined_names = "".join(f'<definedName name="{name}">{formula}</definedName>' for name, formula in named_ranges)
    defined_xml = f"<definedNames>{defined_names}</definedNames>" if defined_names else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="32000" windowHeight="18000" activeTab="0"/></bookViews>
  <sheets>{''.join(sheets_xml)}</sheets>
  {defined_xml}
  <calcPr calcId="191029"/>
</workbook>
"""


def build_workbook_rels_xml() -> str:
    rels = []
    for index in range(1, len(SHEET_ORDER) + 1):
        rels.append(f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>')
    rels.append(f'<Relationship Id="rId{len(SHEET_ORDER) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(rels) + "</Relationships>"


def build_cols_xml(width_specs: list[tuple[int, int, float]]) -> str:
    if not width_specs:
        return ""
    cols = [f'<col min="{min_col}" max="{max_col}" width="{width}" customWidth="1"/>' for min_col, max_col, width in width_specs]
    return "<cols>" + "".join(cols) + "</cols>"


def build_page_margins() -> str:
    return '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'


def column_widths(headers: list[str]) -> list[tuple[int, int, float]]:
    return [(idx, idx, min(max(len(header) + 2, 12), 28)) for idx, header in enumerate(headers, start=1)]


def build_sheet_header_row(headers: list[str], row_number: int = 1) -> str:
    cells = [inline_string_cell(row_number, idx, header, 1) for idx, header in enumerate(headers, start=1)]
    return f'<row r="{row_number}" ht="24" customHeight="1">{"".join(cells)}</row>'


def build_blank_data_row(headers: list[str], row_number: int = 2) -> str:
    cells = [inline_string_cell(row_number, idx, "", 0) for idx, _ in enumerate(headers, start=1)]
    return f'<row r="{row_number}">{"".join(cells)}</row>'


def build_data_validations(headers: list[str], start_row: int = 2, end_row: int = 5000) -> str:
    validations = []
    for idx, header in enumerate(headers, start=1):
        list_name = VALIDATION_TARGETS.get(header)
        if not list_name:
            continue
        col = excel_col(idx)
        validations.append(
            f'<dataValidation type="list" allowBlank="1" showDropDown="0" showInputMessage="1" showErrorMessage="1" sqref="{col}{start_row}:{col}{end_row}"><formula1>{list_name}</formula1></dataValidation>'
        )
    if not validations:
        return ""
    return f'<dataValidations count="{len(validations)}">{"".join(validations)}</dataValidations>'


def build_table_xml(table_id: int, table_name: str, headers: list[str]) -> str:
    last_col = excel_col(len(headers))
    columns = "".join(f'<tableColumn id="{idx}" name="{escape(header)}"/>' for idx, header in enumerate(headers, start=1))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{table_id}" name="{table_name}" displayName="{table_name}" ref="A1:{last_col}2" headerRowCount="1" totalsRowShown="0">
  <autoFilter ref="A1:{last_col}2"/>
  <tableColumns count="{len(headers)}">{columns}</tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
"""


def build_sheet_rels_xml(table_id: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table{table_id}.xml"/>
</Relationships>
"""


def build_data_sheet_xml(headers: list[str]) -> str:
    last_col = excel_col(len(headers))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:{last_col}2"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  {build_cols_xml(column_widths(headers))}
  <sheetData>{build_sheet_header_row(headers)}{build_blank_data_row(headers)}</sheetData>
  {build_data_validations(headers)}
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
  {build_page_margins()}
</worksheet>
"""


def build_readme_sheet_xml() -> str:
    rows = []
    for row_idx, (value, style) in enumerate(README_ROWS, start=1):
        col = 1 if style in (1, 3) else 2
        row_xml = inline_string_cell(row_idx, col, value, style if style else 2)
        if style == 3:
            rows.append(f'<row r="{row_idx}" ht="28" customHeight="1">{row_xml}</row>')
        elif style == 1:
            rows.append(f'<row r="{row_idx}" ht="22" customHeight="1">{row_xml}</row>')
        else:
            rows.append(f'<row r="{row_idx}">{row_xml}</row>')
    cols_xml = '<cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="110" customWidth="1"/></cols>'
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B{len(README_ROWS)}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  {cols_xml}
  <sheetData>{''.join(rows)}</sheetData>
  {build_page_margins()}
</worksheet>
"""


def build_referentiels_sheet() -> tuple[list[tuple[str, str]], str]:
    rows: dict[int, list[str]] = {}
    named_ranges: list[tuple[str, str]] = []
    width_specs: list[tuple[int, int, float]] = []
    current_col = 1
    for ref in REFERENCE_LISTS.values():
        code_col = current_col
        label_col = current_col + 1
        rows.setdefault(1, []).append(inline_string_cell(1, code_col, ref["title"], 1))
        rows.setdefault(2, []).append(inline_string_cell(2, code_col, "Code", 1))
        rows.setdefault(2, []).append(inline_string_cell(2, label_col, "Label", 1))
        width_specs.extend([(code_col, code_col, 24), (label_col, label_col, 42)])
        start_row = 3
        for offset, (code, label) in enumerate(ref["values"]):
            row = start_row + offset
            rows.setdefault(row, []).append(inline_string_cell(row, code_col, code, 2))
            rows.setdefault(row, []).append(inline_string_cell(row, label_col, label, 2))
        end_row = start_row + len(ref["values"]) - 1
        named_ranges.append((ref["code_name"], f"Referentiels!${excel_col(code_col)}${start_row}:${excel_col(code_col)}${end_row}"))
        named_ranges.append((ref["label_name"], f"Referentiels!${excel_col(label_col)}${start_row}:${excel_col(label_col)}${end_row}"))
        current_col += 3
    max_row = max(rows)
    row_xml = "".join(f'<row r="{row_idx}">{"".join(rows.get(row_idx, []))}</row>' for row_idx in range(1, max_row + 1))
    sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:{excel_col(current_col - 1)}{max_row}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A3" sqref="A3"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  {build_cols_xml(width_specs)}
  <sheetData>{row_xml}</sheetData>
  {build_page_margins()}
</worksheet>
"""
    return named_ranges, sheet_xml


def write_workbook() -> None:
    referentiels_sheet_id = SHEET_ORDER.index("Referentiels") + 1
    named_ranges, referentiels_xml = build_referentiels_sheet()
    with zipfile.ZipFile(OUTPUT_PATH, "w", compression=zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", build_content_types_xml(len(SHEET_ORDER), len(DATA_SHEETS)))
        xlsx.writestr("_rels/.rels", build_root_rels_xml())
        xlsx.writestr("docProps/app.xml", build_app_xml())
        xlsx.writestr("docProps/core.xml", build_core_xml())
        xlsx.writestr("xl/workbook.xml", build_workbook_xml(named_ranges, referentiels_sheet_id))
        xlsx.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels_xml())
        xlsx.writestr("xl/styles.xml", build_styles_xml())

        table_id = 1
        for sheet_index, sheet_name in enumerate(SHEET_ORDER, start=1):
            if sheet_name == "README":
                xlsx.writestr(f"xl/worksheets/sheet{sheet_index}.xml", build_readme_sheet_xml())
                continue
            if sheet_name == "Referentiels":
                xlsx.writestr(f"xl/worksheets/sheet{sheet_index}.xml", referentiels_xml)
                continue
            headers = DATA_SHEETS[sheet_name]["headers"]
            xlsx.writestr(f"xl/worksheets/sheet{sheet_index}.xml", build_data_sheet_xml(headers))
            xlsx.writestr(f"xl/worksheets/_rels/sheet{sheet_index}.xml.rels", build_sheet_rels_xml(table_id))
            xlsx.writestr(f"xl/tables/table{table_id}.xml", build_table_xml(table_id, DATA_SHEETS[sheet_name]["table_name"], headers))
            table_id += 1


if __name__ == "__main__":
    write_workbook()
    print(f"Template generated: {OUTPUT_PATH}")
