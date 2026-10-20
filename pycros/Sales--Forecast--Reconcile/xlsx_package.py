from __future__ import annotations

import os
import posixpath
import re
import tempfile
import zipfile
from copy import deepcopy
from decimal import Decimal
from pathlib import Path
from xml.dom import Node, minidom
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_EXTENDED = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
MAIN = f"{{{NS_MAIN}}}"
WORKSHEET_RELATIONSHIP = f"{NS_REL}/worksheet"
WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


def _reconciliation_styles(archive: zipfile.ZipFile):
    styles = minidom.parseString(archive.read("xl/styles.xml"))
    changed = False
    fonts = styles.getElementsByTagNameNS(NS_MAIN, "fonts")[0]
    font_nodes = [node for node in fonts.childNodes if node.nodeType == Node.ELEMENT_NODE]
    bold_font = font_nodes[0].cloneNode(deep=True)
    if not bold_font.getElementsByTagNameNS(NS_MAIN, "b"):
        bold_font.insertBefore(styles.createElementNS(NS_MAIN, "b"), bold_font.firstChild)
    for font_id, font in enumerate(font_nodes):
        if font.toxml() == bold_font.toxml():
            break
    else:
        font_id = len(font_nodes)
        fonts.appendChild(bold_font)
        fonts.setAttribute("count", str(font_id + 1))
        changed = True

    cell_xfs = styles.getElementsByTagNameNS(NS_MAIN, "cellXfs")[0]
    xfs = [node for node in cell_xfs.childNodes if node.nodeType == Node.ELEMENT_NODE]
    style_ids = []
    for num_fmt, style_font in ((0, font_id), (3, 0), (4, 0)):
        for style_id, xf in enumerate(xfs):
            if (
                xf.getAttribute("numFmtId") == str(num_fmt)
                and xf.getAttribute("fontId") == str(style_font)
                and xf.getAttribute("fillId") == "0"
                and xf.getAttribute("borderId") == "0"
            ):
                break
        else:
            style_id = len(xfs)
            xf = styles.createElementNS(NS_MAIN, "xf")
            for attr, value in (
                ("numFmtId", num_fmt),
                ("fontId", style_font),
                ("fillId", 0),
                ("borderId", 0),
                ("xfId", 0),
            ):
                xf.setAttribute(attr, str(value))
            xf.setAttribute("applyFont" if num_fmt == 0 else "applyNumberFormat", "1")
            cell_xfs.appendChild(xf)
            xfs.append(xf)
            cell_xfs.setAttribute("count", str(len(xfs)))
            changed = True
        style_ids.append(style_id)

    return (styles.toxml(encoding="UTF-8") if changed else None), tuple(style_ids)


def _reconciliation_sheet(sheet_name, rows, header_labels, style_ids, original=None):
    bold_style, qty_style, amount_style = style_ids
    root = ET.Element(f"{MAIN}worksheet")
    last_row = max(4, len(rows) + 3)
    ET.SubElement(root, f"{MAIN}dimension", {"ref": f"A1:N{last_row}"})
    views = ET.SubElement(root, f"{MAIN}sheetViews")
    view = ET.SubElement(views, f"{MAIN}sheetView", {"workbookViewId": "0"})
    ET.SubElement(view, f"{MAIN}pane", {"ySplit": "3", "topLeftCell": "A4", "activePane": "bottomLeft", "state": "frozen"})
    ET.SubElement(view, f"{MAIN}selection", {"pane": "bottomLeft", "activeCell": "A4", "sqref": "A4"})
    ET.SubElement(root, f"{MAIN}sheetFormatPr", {"defaultRowHeight": "15"})
    columns = ET.SubElement(root, f"{MAIN}cols")
    for first, last, width in ((1, 1, 22), (2, 7, 16), (8, 8, 47), (9, 12, 12), (13, 13, 18), (14, 14, 55)):
        ET.SubElement(columns, f"{MAIN}col", {"min": str(first), "max": str(last), "width": str(width), "customWidth": "1"})

    sheet_data = ET.SubElement(root, f"{MAIN}sheetData")
    totals = {key: Decimal(0) for key in ("new_qty", "new_amount", "old_qty", "old_amount")}
    for entry in rows:
        for key in totals:
            if entry[key] is not None:
                totals[key] += Decimal(str(entry[key]))

    for row_index in range(1, last_row + 1):
        values = []
        if row_index == 1:
            values = [("B", header_labels[0], "text", bold_style), ("D", header_labels[1], "text", bold_style), ("F", "Diff", "text", bold_style)]
        elif row_index == 2:
            values = [
                ("A", "Job", "text", bold_style),
                ("B", "Qty", "text", bold_style),
                ("C", "Amt", "text", bold_style),
                ("D", "Qty", "text", bold_style),
                ("E", "Amt", "text", bold_style),
                ("F", "Qty", "text", bold_style),
                ("G", "Amt", "text", bold_style),
                ("L", "Remark - SIV Comment", "text", bold_style),
                ("M", "GAC CODE", "text", bold_style),
                ("N", "SIG Finance Comment", "text", bold_style),
            ]
        elif row_index == 3:
            values = [("A", "Job", "text", bold_style)]
            for column, key, style in (("B", "new_qty", qty_style), ("C", "new_amount", amount_style), ("D", "old_qty", qty_style), ("E", "old_amount", amount_style)):
                values.append((column, (f"SUM({column}4:{column}{last_row})", totals[key]), "formula", style))
            values.extend((
                ("F", ("B3-D3", totals["new_qty"] - totals["old_qty"]), "formula", qty_style),
                ("G", ("C3-E3", totals["new_amount"] - totals["old_amount"]), "formula", amount_style),
                ("H", "Remark", "text", bold_style),
                ("L", "Remark - SIV Comment", "text", bold_style),
                ("M", "GAC CODE", "text", bold_style),
                ("N", "SIG Finance Comment", "text", bold_style),
            ))
        elif row_index <= len(rows) + 3:
            entry = rows[row_index - 4]
            new_qty = Decimal(str(entry["new_qty"] or 0))
            new_amount = Decimal(str(entry["new_amount"] or 0))
            old_qty = Decimal(str(entry["old_qty"] or 0))
            old_amount = Decimal(str(entry["old_amount"] or 0))
            values = [("A", entry["job"], "text", 0)]
            for column, key, style in (("B", "new_qty", qty_style), ("C", "new_amount", amount_style), ("D", "old_qty", qty_style), ("E", "old_amount", amount_style)):
                if entry[key] is not None:
                    values.append((column, entry[key], "number", style))
            values.extend((
                ("F", (f"B{row_index}-D{row_index}", new_qty - old_qty), "formula", qty_style),
                ("G", (f"C{row_index}-E{row_index}", new_amount - old_amount), "formula", amount_style),
            ))
            if entry["remark"]:
                values.append(("H", entry["remark"], "text", 0))
            if entry["gac"] is not None and str(entry["gac"]).strip():
                gac = entry["gac"]
                if isinstance(gac, (int, float, Decimal)) or re.fullmatch(r"\d+(?:\.0+)?", str(gac).strip()):
                    values.append(("M", gac, "number", 0))
                else:
                    values.append(("M", gac, "text", 0))
            formula = f'IF(M{row_index}="","",IFERROR(VLOOKUP(M{row_index},\'Reason code\'!$A:$C,3,FALSE),IFERROR(VLOOKUP(VALUE(M{row_index}),\'Reason code\'!$A:$C,3,FALSE),"")))'
            values.append(("N", (formula, None), "formula", 0))

        xml_row = ET.SubElement(sheet_data, f"{MAIN}row", {"r": str(row_index)})
        for column, value, kind, style in values:
            cell = ET.SubElement(xml_row, f"{MAIN}c", {"r": f"{column}{row_index}"})
            if style:
                cell.set("s", str(style))
            if kind == "text":
                cell.set("t", "inlineStr")
                inline = ET.SubElement(cell, f"{MAIN}is")
                txt = ET.SubElement(inline, f"{MAIN}t")
                txt.text = str(value)
                if txt.text != txt.text.strip():
                    txt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            elif kind == "number":
                number = Decimal(str(value))
                if not number.is_finite():
                    raise ValueError(f"Non-finite number in {sheet_name}!{column}{row_index}")
                ET.SubElement(cell, f"{MAIN}v").text = format(number, "f")
            else:
                formula, cached = value
                ET.SubElement(cell, f"{MAIN}f").text = formula
                if cached is not None:
                    if not cached.is_finite():
                        raise ValueError(f"Non-finite number in {sheet_name}!{column}{row_index}")
                    ET.SubElement(cell, f"{MAIN}v").text = format(cached, "f")

    ET.SubElement(root, f"{MAIN}autoFilter", {"ref": f"A3:N{last_row}"})
    merges = ET.SubElement(root, f"{MAIN}mergeCells", {"count": "3"})
    for ref in ("B1:C1", "D1:E1", "F1:G1"):
        ET.SubElement(merges, f"{MAIN}mergeCell", {"ref": ref})
    if original is not None:
        old_root = ET.fromstring(original)
        for tag in ("printOptions", "pageMargins", "pageSetup", "headerFooter"):
            child = old_root.find(f"{MAIN}{tag}")
            if child is not None:
                root.append(deepcopy(child))
    else:
        ET.SubElement(root, f"{MAIN}pageMargins", {"left": "0.7", "right": "0.7", "top": "0.75", "bottom": "0.75", "header": "0.3", "footer": "0.3"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _update_variance_values(original, updates):
    document = minidom.parseString(original)
    sheet_data = document.getElementsByTagNameNS(NS_MAIN, "sheetData")[0]
    for ref, value in updates.items():
        match = re.fullmatch(r"([A-Z]+)([1-9]\d*)", ref)
        if not match:
            raise ValueError(f"Invalid variance cell reference: {ref}")
        numeric = Decimal(str(value))
        if not numeric.is_finite():
            raise ValueError(f"Non-finite variance value for {ref}")
        row_number = int(match.group(2))
        row = next((item for item in sheet_data.childNodes if item.nodeType == Node.ELEMENT_NODE and item.localName == "row" and item.getAttribute("r") == str(row_number)), None)
        if row is None:
            row = document.createElementNS(NS_MAIN, "row")
            row.setAttribute("r", str(row_number))
            later = next((item for item in sheet_data.childNodes if item.nodeType == Node.ELEMENT_NODE and item.localName == "row" and int(item.getAttribute("r")) > row_number), None)
            sheet_data.insertBefore(row, later)
        cell = next((item for item in row.childNodes if item.nodeType == Node.ELEMENT_NODE and item.localName == "c" and item.getAttribute("r") == ref), None)
        if cell is None:
            cell = document.createElementNS(NS_MAIN, "c")
            cell.setAttribute("r", ref)
            target_column = (len(match.group(1)), match.group(1))
            later = next((item for item in row.childNodes if item.nodeType == Node.ELEMENT_NODE and item.localName == "c" and (len(re.match(r"[A-Z]+", item.getAttribute("r")).group()), re.match(r"[A-Z]+", item.getAttribute("r")).group()) > target_column), None)
            row.insertBefore(cell, later)
        for child in list(cell.childNodes):
            cell.removeChild(child)
        if cell.hasAttribute("t"):
            cell.removeAttribute("t")
        numeric_value = document.createElementNS(NS_MAIN, "v")
        numeric_value.appendChild(document.createTextNode(format(numeric, "f")))
        cell.appendChild(numeric_value)
    return document.toxml(encoding="UTF-8")


def write_reconciliation_workbook(source_path, output_path, recon_sheets, variance_updates, header_labels):
    source_path = Path(source_path)
    output_path = Path(output_path)
    if source_path.resolve() == output_path.resolve():
        raise ValueError("Output path must be different from the source workbook")
    if output_path.exists():
        raise FileExistsError(output_path)

    replacements = {}
    additions = {}
    removed = set()
    regenerated_filters = {}
    with zipfile.ZipFile(source_path) as source:
        members = set(source.namelist())
        workbook = minidom.parseString(source.read("xl/workbook.xml"))
        workbook.documentElement.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", NS_REL)
        relationships = minidom.parseString(source.read("xl/_rels/workbook.xml.rels"))
        content_types = minidom.parseString(source.read("[Content_Types].xml"))
        sheets_node = workbook.getElementsByTagNameNS(NS_MAIN, "sheets")[0]
        sheets = [node for node in sheets_node.childNodes if node.nodeType == Node.ELEMENT_NODE]
        relationship_nodes = relationships.getElementsByTagNameNS(NS_PKG_REL, "Relationship")
        targets = {node.getAttribute("Id"): node.getAttribute("Target") for node in relationship_nodes}
        sheet_paths = {
            node.getAttribute("name"): posixpath.normpath(posixpath.join("xl", targets[node.getAttributeNS(NS_REL, "id")].lstrip("/")))
            if not targets[node.getAttributeNS(NS_REL, "id")].startswith("/")
            else targets[node.getAttributeNS(NS_REL, "id")].lstrip("/")
            for node in sheets
        }
        variance_sheets = [name for name in sheet_paths if re.search(r"\bVariance FY\d+\b", name, re.IGNORECASE)]
        if len(variance_sheets) != 1:
            raise ValueError("Sales Report must have one Variance FY sheet")
        if "Reason code" not in sheet_paths:
            raise ValueError("Sales Report is missing the Reason code sheet")

        style_xml, style_ids = _reconciliation_styles(source)
        if style_xml is not None:
            replacements["xl/styles.xml"] = style_xml
        views = workbook.getElementsByTagNameNS(NS_MAIN, "workbookView")
        active_sheets = []
        for view in views:
            active_index = int(view.getAttribute("activeTab") or 0)
            first_index = int(view.getAttribute("firstSheet") or 0)
            active_sheets.append((view, sheets[active_index] if active_index < len(sheets) else sheets[0], sheets[first_index] if first_index < len(sheets) else sheets[0]))
        next_sheet_id = max(int(node.getAttribute("sheetId")) for node in sheets) + 1
        next_part_id = max(int(match.group(1)) for name in members if (match := re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", name))) + 1
        next_rel_id = max((int(match.group(1)) for node in relationship_nodes if (match := re.fullmatch(r"rId(\d+)", node.getAttribute("Id")))), default=0) + 1

        for sheet_name, rows in recon_sheets.items():
            if not sheet_name.startswith("Recon "):
                raise ValueError(f"Recon sheet name must start with 'Recon ': {sheet_name}")
            month_name = sheet_name[6:].strip()
            matching_months = [node for node in sheets if node.getAttribute("name").strip().casefold() == month_name.casefold()]
            if len(matching_months) != 1:
                raise ValueError(f"Expected one month sheet for {sheet_name}, found {len(matching_months)}")
            month_sheet = matching_months[0]
            existing = [node for node in sheets if re.fullmatch(rf"Recon\s+{re.escape(month_name)}(?:\s+ok)?\s*", node.getAttribute("name"), re.IGNORECASE)]
            if len(existing) > 1:
                raise ValueError(f"Multiple existing Recon sheets for {month_name}")
            if existing:
                recon_sheet = existing[0]
                old_name = recon_sheet.getAttribute("name")
                part = sheet_paths[old_name]
                sheets_node.removeChild(recon_sheet)
                sheets_node.insertBefore(recon_sheet, month_sheet)
                original = source.read(part)
                regenerated_filters[recon_sheet] = (old_name, max(4, len(rows) + 3))
            else:
                part = f"xl/worksheets/sheet{next_part_id}.xml"
                next_part_id += 1
                recon_sheet = workbook.createElementNS(NS_MAIN, "sheet")
                recon_sheet.setAttribute("name", sheet_name)
                recon_sheet.setAttribute("sheetId", str(next_sheet_id))
                recon_sheet.setAttributeNS(NS_REL, "r:id", f"rId{next_rel_id}")
                next_sheet_id += 1
                sheets_node.insertBefore(recon_sheet, month_sheet)

                relation = relationships.createElementNS(NS_PKG_REL, "Relationship")
                relation.setAttribute("Id", f"rId{next_rel_id}")
                relation.setAttribute("Type", WORKSHEET_RELATIONSHIP)
                relation.setAttribute("Target", f"worksheets/{posixpath.basename(part)}")
                relationships.documentElement.appendChild(relation)
                next_rel_id += 1

                override = content_types.createElementNS(NS_CONTENT, "Override")
                override.setAttribute("PartName", f"/{part}")
                override.setAttribute("ContentType", WORKSHEET_CONTENT_TYPE)
                content_types.documentElement.appendChild(override)
                original = None

            generated = _reconciliation_sheet(sheet_name, rows, header_labels, style_ids, original)
            if part in members:
                replacements[part] = generated
            else:
                additions[part] = generated

        ordered_sheets = [node for node in sheets_node.childNodes if node.nodeType == Node.ELEMENT_NODE]
        old_titles = [node.getAttribute("name") for node in sheets]
        new_titles = [node.getAttribute("name") for node in ordered_sheets]
        if new_titles != old_titles and "docProps/app.xml" in members:
            app = minidom.parseString(source.read("docProps/app.xml"))
            heading = app.getElementsByTagNameNS(NS_EXTENDED, "HeadingPairs")
            titles = app.getElementsByTagNameNS(NS_EXTENDED, "TitlesOfParts")
            if heading and titles:
                variants = heading[0].getElementsByTagNameNS(NS_VT, "variant")
                for index in range(0, len(variants) - 1, 2):
                    label = variants[index].getElementsByTagNameNS(NS_VT, "lpstr")
                    if label and label[0].firstChild and label[0].firstChild.data == "Worksheets":
                        count = variants[index + 1].getElementsByTagNameNS(NS_VT, "i4")[0]
                        count.firstChild.data = str(len(new_titles))
                        break
                title_vector = titles[0].getElementsByTagNameNS(NS_VT, "vector")[0]
                parts = [node for node in title_vector.childNodes if node.nodeType == Node.ELEMENT_NODE]
                if len(parts) < len(old_titles):
                    raise ValueError("Sales Report worksheet titles are incomplete")
                for part in parts[:len(old_titles)]:
                    title_vector.removeChild(part)
                next_part = parts[len(old_titles)] if len(parts) > len(old_titles) else None
                for name in new_titles:
                    part = app.createElementNS(NS_VT, "vt:lpstr")
                    part.appendChild(app.createTextNode(name))
                    title_vector.insertBefore(part, next_part)
                title_vector.setAttribute("size", str(len(new_titles) + len(parts) - len(old_titles)))
                replacements["docProps/app.xml"] = app.toxml(encoding="UTF-8")
        for view, active_sheet, first_sheet in active_sheets:
            view.setAttribute("activeTab", str(ordered_sheets.index(active_sheet)))
            if view.hasAttribute("firstSheet"):
                view.setAttribute("firstSheet", str(ordered_sheets.index(first_sheet)))
        for defined_name in workbook.getElementsByTagNameNS(NS_MAIN, "definedName"):
            if defined_name.hasAttribute("localSheetId"):
                original_sheet = sheets[int(defined_name.getAttribute("localSheetId"))]
                defined_name.setAttribute("localSheetId", str(ordered_sheets.index(original_sheet)))
                if original_sheet in regenerated_filters:
                    name, last_row = regenerated_filters[original_sheet]
                    quoted_name = name.replace("'", "''")
                    if defined_name.getAttribute("name") == "_xlnm._FilterDatabase":
                        defined_name.firstChild.data = f"'{quoted_name}'!$A$3:$N${last_row}"
                    elif defined_name.getAttribute("name") == "_xlnm.Print_Area":
                        area = defined_name.firstChild.data
                        if re.search(r":\$[A-Z]+\$\d+$", area):
                            defined_name.firstChild.data = re.sub(r":\$[A-Z]+\$\d+$", f":$N${last_row}", area)
        for relation in list(relationships.getElementsByTagNameNS(NS_PKG_REL, "Relationship")):
            if relation.getAttribute("Type") == f"{NS_REL}/calcChain":
                target = relation.getAttribute("Target")
                part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
                removed.add(part)
                relationships.documentElement.removeChild(relation)
        for override in list(content_types.getElementsByTagNameNS(NS_CONTENT, "Override")):
            if override.getAttribute("PartName").lstrip("/") in removed:
                content_types.documentElement.removeChild(override)
        calculation = workbook.getElementsByTagNameNS(NS_MAIN, "calcPr")
        if calculation:
            calculation[0].setAttribute("fullCalcOnLoad", "1")
            calculation[0].setAttribute("forceFullCalc", "1")
        else:
            calc = workbook.createElementNS(NS_MAIN, "calcPr")
            calc.setAttribute("fullCalcOnLoad", "1")
            calc.setAttribute("forceFullCalc", "1")
            workbook.documentElement.appendChild(calc)
        replacements["xl/workbook.xml"] = workbook.toxml(encoding="UTF-8")
        replacements["xl/_rels/workbook.xml.rels"] = relationships.toxml(encoding="UTF-8")
        replacements["[Content_Types].xml"] = content_types.toxml(encoding="UTF-8")
        variance_path = sheet_paths[variance_sheets[0]]
        if variance_updates:
            replacements[variance_path] = _update_variance_values(source.read(variance_path), variance_updates)

        handle, temporary = tempfile.mkstemp(prefix=f".{output_path.stem}.", suffix=".xlsx", dir=output_path.parent)
        os.close(handle)
        try:
            with zipfile.ZipFile(temporary, "w") as output:
                for info in source.infolist():
                    if info.filename in removed:
                        continue
                    data = replacements.pop(info.filename) if info.filename in replacements else source.read(info.filename)
                    output.writestr(info, data)
                for name, data in additions.items():
                    output.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
            if replacements:
                raise ValueError(f"Missing source workbook parts: {', '.join(replacements)}")
            os.replace(temporary, output_path)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)
