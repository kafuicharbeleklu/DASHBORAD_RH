from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parent
TEMPLATE_DIR = ROOT / "MODELE_COLLECTE_RH_NEEMBA"
SOURCE_PATH = TEMPLATE_DIR / "RH_Collecte.xlsx"
FALLBACK_PATH = TEMPLATE_DIR / "RH_Collecte_CORRIGE.xlsx"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)


DATA_TABLES = {
    "Effectif": "tbl_Effectif",
    "Embauches": "tbl_Embauches",
    "Departs": "tbl_Departs",
    "AbsenceMensuelle": "tbl_AbsenceMensuelle",
    "FormationMensuelle": "tbl_FormationMensuelle",
    "RecrutementMensuel": "tbl_RecrutementMensuel",
    "RecrutementDetail": "tbl_RecrutementDetail",
    "MasseSalarialeMensuelle": "tbl_MasseSalarialeMensuelle",
    "TCDP_Headcount": "tbl_TCDP_Headcount",
    "TCDP_Entrees": "tbl_TCDP_Entrees",
    "TCDP_Sorties": "tbl_TCDP_Sorties",
    "TCDP_Genre": "tbl_TCDP_Genre",
}


def qname(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def excel_col(index: int) -> str:
    value = index
    letters: list[str] = []
    while value:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def last_col_from_ref(ref: str) -> str:
    match = re.match(r"[A-Z]+[0-9]+:([A-Z]+)[0-9]+$", ref)
    if not match:
        raise ValueError(f"Unexpected ref format: {ref}")
    return match.group(1)


def build_blank_row_xml(last_col: str, column_count: int) -> ET.Element:
    row = ET.Element(qname(MAIN_NS, "row"), {"r": "2"})
    for col_index in range(1, column_count + 1):
        cell = ET.SubElement(
            row,
            qname(MAIN_NS, "c"),
            {"r": f"{excel_col(col_index)}2", "t": "inlineStr"},
        )
        inline = ET.SubElement(cell, qname(MAIN_NS, "is"))
        ET.SubElement(inline, qname(MAIN_NS, "t")).text = ""
    return row


def build_table_xml(table_id: int, table_name: str, headers: list[str]) -> bytes:
    last_col = excel_col(len(headers))
    table = ET.Element(
        qname(MAIN_NS, "table"),
        {
            "id": str(table_id),
            "name": table_name,
            "displayName": table_name,
            "ref": f"A1:{last_col}2",
            "headerRowCount": "1",
            "totalsRowCount": "0",
            "totalsRowShown": "0",
        },
    )
    ET.SubElement(table, qname(MAIN_NS, "autoFilter"), {"ref": f"A1:{last_col}2"})
    columns = ET.SubElement(table, qname(MAIN_NS, "tableColumns"), {"count": str(len(headers))})
    for idx, header in enumerate(headers, start=1):
        ET.SubElement(columns, qname(MAIN_NS, "tableColumn"), {"id": str(idx), "name": header})
    ET.SubElement(
        table,
        qname(MAIN_NS, "tableStyleInfo"),
        {
            "name": "TableStyleMedium2",
            "showFirstColumn": "0",
            "showLastColumn": "0",
            "showRowStripes": "1",
            "showColumnStripes": "0",
        },
    )
    return ET.tostring(table, encoding="utf-8", xml_declaration=True)


def build_sheet_rels_xml(table_id: int) -> bytes:
    relationships = ET.Element(qname(REL_NS, "Relationships"))
    ET.SubElement(
        relationships,
        qname(REL_NS, "Relationship"),
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
            "Target": f"../tables/table{table_id}.xml",
        },
    )
    return ET.tostring(relationships, encoding="utf-8", xml_declaration=True)


def update_content_types(xml_bytes: bytes, table_count: int) -> bytes:
    root = ET.fromstring(xml_bytes)
    existing = {
        override.get("PartName")
        for override in root.findall(qname(CONTENT_TYPES_NS, "Override"))
    }
    for table_id in range(1, table_count + 1):
        part_name = f"/xl/tables/table{table_id}.xml"
        if part_name in existing:
            continue
        ET.SubElement(
            root,
            qname(CONTENT_TYPES_NS, "Override"),
            {
                "PartName": part_name,
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
            },
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def read_header_names(table_root: ET.Element) -> list[str]:
    table_columns = table_root.find(qname(MAIN_NS, "tableColumns"))
    if table_columns is None:
        raise ValueError("Table XML has no tableColumns element.")
    return [column.get("name", "") for column in table_columns.findall(qname(MAIN_NS, "tableColumn"))]


def repair_workbook() -> Path:
    if not SOURCE_PATH.exists():
        raise FileNotFoundError(f"Template not found: {SOURCE_PATH}")

    with zipfile.ZipFile(SOURCE_PATH) as source_zip:
        file_map = {name: source_zip.read(name) for name in source_zip.namelist()}

    workbook_root = ET.fromstring(file_map["xl/workbook.xml"])
    workbook_rels_root = ET.fromstring(file_map["xl/_rels/workbook.xml.rels"])
    workbook_rel_map = {
        rel.get("Id"): rel.get("Target")
        for rel in workbook_rels_root.findall(qname(REL_NS, "Relationship"))
    }

    table_id = 1
    for sheet in workbook_root.find(qname(MAIN_NS, "sheets")):
        sheet_name = sheet.get("name")
        if sheet_name not in DATA_TABLES:
            continue

        rel_id = sheet.get(qname(DOC_REL_NS, "id"))
        target = workbook_rel_map[rel_id]
        sheet_path = Path("xl") / target

        worksheet_root = ET.fromstring(file_map[str(sheet_path).replace("\\", "/")])
        sheet_data = worksheet_root.find(qname(MAIN_NS, "sheetData"))
        if sheet_data is None:
            raise ValueError(f"No sheetData found in {sheet_name}")

        header_row = None
        for row in sheet_data.findall(qname(MAIN_NS, "row")):
            if row.get("r") == "1":
                header_row = deepcopy(row)
                break
        if header_row is None:
            raise ValueError(f"No header row found in {sheet_name}")

        last_col = "A"
        table_xml_path = f"xl/tables/table{table_id}.xml"
        if table_xml_path in file_map:
            table_root = ET.fromstring(file_map[table_xml_path])
            headers = read_header_names(table_root)
            last_col = last_col_from_ref(table_root.get("ref", "A1:A2"))
        else:
            header_cells = header_row.findall(qname(MAIN_NS, "c"))
            headers = [""] * len(header_cells)
            last_col = excel_col(len(headers))

        for row in list(sheet_data):
            sheet_data.remove(row)
        header_row.set("r", "1")
        sheet_data.append(header_row)
        sheet_data.append(build_blank_row_xml(last_col, len(headers)))

        dimension = worksheet_root.find(qname(MAIN_NS, "dimension"))
        if dimension is not None:
            dimension.set("ref", f"A1:{last_col}2")

        table_parts = worksheet_root.find(qname(MAIN_NS, "tableParts"))
        if table_parts is None:
            table_parts = ET.SubElement(worksheet_root, qname(MAIN_NS, "tableParts"), {"count": "1"})
        table_parts.set("count", "1")
        for child in list(table_parts):
            table_parts.remove(child)
        ET.SubElement(table_parts, qname(MAIN_NS, "tablePart"), {qname(DOC_REL_NS, "id"): "rId1"})

        file_map[str(sheet_path).replace("\\", "/")] = ET.tostring(worksheet_root, encoding="utf-8", xml_declaration=True)
        file_map[f"xl/worksheets/_rels/{sheet_path.name}.rels"] = build_sheet_rels_xml(table_id)
        file_map[table_xml_path] = build_table_xml(table_id, DATA_TABLES[sheet_name], headers)
        table_id += 1

    file_map["[Content_Types].xml"] = update_content_types(file_map["[Content_Types].xml"], len(DATA_TABLES))

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=TEMPLATE_DIR) as tmp_file:
        tmp_path = Path(tmp_file.name)

    with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
        for name, payload in file_map.items():
            output_zip.writestr(name, payload)

    try:
        tmp_path.replace(SOURCE_PATH)
        return SOURCE_PATH
    except PermissionError:
        tmp_path.replace(FALLBACK_PATH)
        return FALLBACK_PATH


if __name__ == "__main__":
    target = repair_workbook()
    print(f"Corrected workbook written to: {target}")
