"""Divestiture Separation Plan generator (Solution Architecture perspective).

Combines:
  1) An IT Application Inventory workbook (one row per application, with a
     business-set disposition per app -- e.g. Carve Out / Direct Transfer /
     Lift & Shift -- plus a disposition method and Shared vs Dedicated flag).
  2) An Application Integration catalog workbook (one sheet per application,
     each row an Initiator/Terminator interface, plus a Summary sheet listing
     every application sheet and its total interface count).

...and produces:
  - An Excel workbook: full separation plan, one row per in-scope
    application, with SA guidance tailored to the disposition/method/
    shared-vs-dedicated combination, plus integration counterpart detail.
  - A PowerPoint executive summary: an overview slide, an overview table of
    every application that has an integration partner also present in the
    inventory, and one fully self-contained detail slide per such
    application (every partner application listed in-slide -- nothing is
    deferred to the Excel workbook).

IMPORTANT SCOPE RULE: integrations are only ever reported between two
applications that both appear in the Inventory workbook. An interface whose
other end is not in the inventory is simply excluded -- no assumption is
made about whether that system is retained, moving, or just not yet
catalogued.
"""
from __future__ import annotations

import argparse
import collections
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ---------------------------------------------------------------------------
# Colors / shared styling
# ---------------------------------------------------------------------------
NAVY_HEX = "264653"
NAVY = RGBColor(0x26, 0x46, 0x53)
TEAL = RGBColor(0x1C, 0x8C, 0x72)
DARK = RGBColor(0x24, 0x32, 0x3D)
GRAY = RGBColor(0x5C, 0x6B, 0x73)
BG = RGBColor(0xF6, 0xF3, 0xEE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
AMBER = RGBColor(0xE0, 0x8A, 0x3E)
RED = RGBColor(0xC2, 0x4B, 0x3D)
LIGHT_HEX = "F1F5F4"

# ---------------------------------------------------------------------------
# Inventory column detection (header text is matched fuzzily, not by a fixed
# column index, so minor export-to-export layout changes do not break this).
# ---------------------------------------------------------------------------
# Each field is a list of (must-not-contain, must-contain-all) rules tested in
# order against the normalized header text; first match wins.
FIELD_RULES: dict[str, list[tuple[list[str], list[str]]]] = {
    "name": ([], ["product name"]),
    "appid": ([], ["appid"]),
    "desc": ([], ["description"]),
    "op_status": ([], ["status"]),
    "capability": ([], ["capability"]),
    "track": ([["sub"]], ["application track"]),
    "subtrack": ([], ["sub track"]),
    "countries": ([], ["countries"]),
    "domain": ([], ["domain"]),
    "legal_entity": ([], ["legal entity"]),
    "hosting": ([], ["hosting type"]),
    "location": ([], ["location"]),
    "sbd_owner": ([["buyer"]], ["sbd it owner"]),
    "buyer_owner": ([], ["buyer it owner"]),
    "eol": ([], ["end of life"]),
    "tsa_support": ([], ["tsa support"]),
    "term_date": ([], ["termination", "disposition date"]),
    "method": ([], ["disposition method"]),
    "disposition": ([["method"]], ["disposition"]),
    "comingled": ([], ["co-mingled"]),
    "dep_yn": ([["provide"]], ["dependency on another app"]),
    "dep_on": ([], ["dependent on"]),
    "license": ([], ["license type"]),
    "contract_status": ([], ["contract status"]),
    "user_count": ([], ["user count"]),
    "shared_dedicated": ([], ["shared", "dedicated"]),
    "vendor": ([], ["vendor"]),
    "tsa_line_item": ([], ["tsa line item"]),
    "buyer_comments": ([], ["buyer comments"]),
    "comments": ([["buyer"]], ["comments"]),
    "sites": ([], ["site"]),
}


def normalize_header(value) -> str:
    text = str(value or "")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def detect_inventory_sheet(wb) -> str:
    for name in wb.sheetnames:
        if "it app inventory" in name.lower() and "do not use" not in name.lower():
            return name
    return wb.sheetnames[0]


def detect_header_row(ws, max_scan_rows: int = 25) -> tuple[int, dict[str, int]]:
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan_rows, values_only=True), 1):
        normalized = [normalize_header(v) for v in row]
        if any("product name" in n for n in normalized) and any("appid" in n for n in normalized):
            header_map = {n: i for i, n in enumerate(normalized) if n}
            return row_idx, header_map
    raise ValueError("Could not find an inventory header row (expected a row with 'Product Name' and 'AppID').")


def resolve_field_columns(header_map: dict[str, int]) -> dict[str, int]:
    resolved: dict[str, int] = {}
    used_columns: set[int] = set()
    for field, rules in FIELD_RULES.items():
        exclude_groups, require_all = rules
        candidates = []
        for header_text, col_idx in header_map.items():
            if col_idx in used_columns:
                continue
            if not all(term in header_text for term in require_all):
                continue
            if any(all(term in header_text for term in group) for group in exclude_groups):
                continue
            candidates.append((len(header_text), col_idx))
        if candidates:
            candidates.sort()
            resolved[field] = candidates[0][1]
            used_columns.add(candidates[0][1])
    missing = [f for f in ("name", "appid", "disposition") if f not in resolved]
    if missing:
        raise ValueError(f"Could not locate required inventory column(s): {missing}")
    return resolved


def load_inventory(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet_name = detect_inventory_sheet(wb)
    ws = wb[sheet_name]
    header_row, header_map = detect_header_row(ws)
    field_cols = resolve_field_columns(header_map)

    apps = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        def get(field):
            idx = field_cols.get(field)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        name = get("name")
        if not name:
            continue
        app = {field: get(field) for field in FIELD_RULES}
        app["shared_dedicated"] = (str(app["shared_dedicated"]).strip().title()
                                    if app["shared_dedicated"] else None)
        for text_field in ("disposition", "method", "comingled", "tsa_support"):
            if app[text_field] is not None:
                app[text_field] = str(app[text_field]).strip()
        apps.append(app)
    return apps


# ---------------------------------------------------------------------------
# Integration workbook loading (Summary sheet + one sheet per application)
# ---------------------------------------------------------------------------
def load_integrations(path: Path, by_appid: dict[str, dict]) -> dict[str, dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "Summary" not in wb.sheetnames:
        raise ValueError("Integrations workbook must contain a 'Summary' sheet.")
    summary_ws = wb["Summary"]
    summary_rows = list(summary_ws.iter_rows(min_row=1, values_only=True))
    header = [normalize_header(h) for h in summary_rows[0]]

    def col(*names):
        # Exact match first, then shortest-header substring match, so e.g.
        # "terminator" never accidentally resolves to "terminator app id".
        for n in names:
            for i, h in enumerate(header):
                if h == n:
                    return i
        candidates = [(len(h), i) for i, h in enumerate(header) if any(n in h for n in names)]
        return min(candidates)[1] if candidates else None

    name_col = col("application name")
    count_col = col("integration count")
    sheet_col = col("sheet name")
    if name_col is None or sheet_col is None:
        raise ValueError("Summary sheet must have 'Application Name' and 'Sheet Name' columns.")

    sheet_info = {}
    for row in summary_rows[1:]:
        if row[sheet_col]:
            sheet_info[row[sheet_col]] = (row[name_col], row[count_col] if count_col is not None else None)

    integration_data: dict[str, dict] = {}
    for sheet_name, (home_name, _total_count) in sheet_info.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(min_row=1, values_only=True))
        if not rows:
            continue
        iface_header = [normalize_header(h) for h in rows[0]]

        def icol(*names):
            for n in names:
                for i, h in enumerate(iface_header):
                    if h == n:
                        return i
            candidates = [(len(h), i) for i, h in enumerate(iface_header) if any(n in h for n in names)]
            return min(candidates)[1] if candidates else None

        data_obj_c = icol("data object")
        display_c = icol("display text")
        init_id_c = icol("initiator app id")
        init_name_c = icol("initiator app")
        iface_c = icol("interfaces")
        term_id_c = icol("terminator app id")
        term_name_c = icol("terminator")
        if init_id_c is None or term_id_c is None:
            continue

        data_rows = rows[1:]
        home_appid = None
        for row in data_rows:
            if init_name_c is not None and row[init_name_c] and str(row[init_name_c]).strip() == str(home_name).strip():
                home_appid = row[init_id_c]
            if term_name_c is not None and row[term_name_c] and str(row[term_name_c]).strip() == str(home_name).strip():
                home_appid = row[term_id_c]
        if home_appid is None:
            continue

        matched_rows = []
        partner_appids = {}
        for row in data_rows:
            init_id = row[init_id_c]
            term_id = row[term_id_c]
            init_name = row[init_name_c] if init_name_c is not None else None
            term_name = row[term_name_c] if term_name_c is not None else None
            data_object = row[data_obj_c] if data_obj_c is not None else None
            display_text = row[display_c] if display_c is not None else None
            if init_id == home_appid and term_id and term_id != home_appid and term_id in by_appid:
                matched_rows.append({
                    "direction": "Outbound (this app \u2192 partner)", "data_object": data_object,
                    "display_text": display_text, "partner_appid": term_id, "partner_name": term_name,
                })
                partner_appids[term_id] = term_name
            if term_id == home_appid and init_id and init_id != home_appid and init_id in by_appid:
                matched_rows.append({
                    "direction": "Inbound (partner \u2192 this app)", "data_object": data_object,
                    "display_text": display_text, "partner_appid": init_id, "partner_name": init_name,
                })
                partner_appids[init_id] = init_name
        integration_data[home_appid] = {
            "home_name": home_name, "rows": matched_rows, "partner_appids": partner_appids,
        }
    return integration_data


# ---------------------------------------------------------------------------
# Glossary + SA guidance
# ---------------------------------------------------------------------------
GLOSSARY = [
    ("SBD Proposed Disposition", "Carve Out from SBD Environment",
     "The application STAYS with SBD. It is shared infrastructure also used by retained SBD "
     "business units, so it cannot be transferred or decommissioned. Separation is achieved "
     "purely through the Disposition Method -- when that method is Data Extract, only the "
     "divesting business's data is extracted for the buyer."),
    ("SBD Proposed Disposition", "Direct Transfer",
     "Ownership/contract and access move to the buyer as-is, with no technical migration "
     "expected. Normally applies only to Dedicated applications (used solely by the divesting "
     "business)."),
    ("SBD Proposed Disposition", "Lift & Shift",
     "The application (infrastructure and data) physically relocates into buyer-owned hosting. "
     "Intended for Dedicated applications that can be removed without impacting other SBD "
     "consumers; a Shared app marked Lift & Shift is a conflict that must be resolved with the "
     "business before planning proceeds."),
    ("SBD Proposed Disposition", "(Not yet determined)",
     "No disposition has been confirmed by the business yet. Technical separation planning "
     "cannot start until this is set."),
    ("Disposition Method", "Data Extract",
     "A one-time or periodic extract of only the divesting business's data is produced and "
     "delivered to the buyer. The source application, its hosting, and its integrations are "
     "NOT modified, migrated, or decommissioned -- the application itself is not going to the "
     "buyer."),
    ("Shared vs Dedicated", "Shared",
     "The application serves the divesting business AND other retained SBD business units. It "
     "cannot be fully transferred, moved, or decommissioned as part of this divestiture; any "
     "separation here is data-only."),
    ("Shared vs Dedicated", "Dedicated",
     "The application serves only the divesting business. It is a genuine candidate for "
     "ownership transfer, migration, or eventual decommission, and its integrations need real "
     "cutover planning."),
    ("Integration Scope Rule", "In-inventory integrations only",
     "This plan only reports integrations where BOTH applications (the app and its partner) "
     "appear in the IT Application Inventory. Interface rows pointing to a system that is not "
     "in the inventory are excluded entirely -- no assumption is made about whether that system "
     "stays with SBD, moves, or is simply not yet catalogued."),
]


def build_guidance(disp, method, shared_dedicated):
    if disp is None:
        return ("Disposition not yet confirmed by the business \u2014 treat as an open item. Do "
                "not plan technical separation activity until SBD Proposed Disposition is set.",
                "not_applicable")
    if disp == "Carve Out from SBD Environment" and method == "Data Extract" and shared_dedicated == "Shared":
        return (
            "Application remains shared SBD infrastructure after close \u2014 no migration, no "
            "decommission. Produce a one-time/periodic Data Extract of only the divesting "
            "business's data for the buyer; confirm extract scope and trace data lineage across "
            "any in-inventory partner applications. Do NOT rebuild, re-point, or retire any of "
            "this application's integrations \u2014 they remain required for continuing SBD "
            "operations.", "data_extract_only"
        )
    if disp == "Lift & Shift" and shared_dedicated == "Shared":
        return (
            "CONFLICT: marked Lift & Shift (implying it leaves SBD) but also flagged Shared "
            "(implying other retained SBD business units depend on it). Resolve with the "
            "business before planning.", "conflict_needs_clarification"
        )
    if disp == "Lift & Shift" and shared_dedicated == "Dedicated":
        return (
            "Dedicated application that physically relocates to buyer-owned hosting "
            "(infrastructure + data as-is). Re-point every integration to an in-inventory "
            "partner application to the new hosting location; any partner on a different "
            "disposition path must be rebuilt, replaced, or formally retired before cutover.",
            "real_migration"
        )
    if disp == "Direct Transfer":
        return (
            "Dedicated application; contract/ownership and access transfer directly to the "
            "buyer with no expected technical migration. Verify every in-inventory integration "
            "partner's own disposition \u2014 a link to a partner staying with SBD will break at "
            "cutover unless explicitly addressed.", "real_migration"
        )
    return (
        f"Disposition '{disp}' / Method '{method}' / {shared_dedicated or 'unspecified'} \u2014 "
        "review manually.", "review_manually"
    )


# ---------------------------------------------------------------------------
# Assemble plan rows
# ---------------------------------------------------------------------------
def build_plan(apps, by_appid, integration_data):
    plan_rows = []
    integration_detail_rows = []
    open_items = []
    conflicts = []
    apps_with_integrations = []
    movers = []

    for a in apps:
        appid, disp, method, shared_dedicated = a["appid"], a["disposition"], a["method"], a["shared_dedicated"]
        guidance, impact_type = build_guidance(disp, method, shared_dedicated)
        integ = integration_data.get(appid)

        partner_summaries, partner_detail = [], []
        same_disp_count = diff_disp_count = 0
        if integ:
            for p_appid, p_name in integ["partner_appids"].items():
                partner = by_appid.get(p_appid)
                p_disp = partner["disposition"] if partner else None
                if p_disp == disp:
                    same_disp_count += 1
                    tag = f"Same disposition ({p_disp})"
                else:
                    diff_disp_count += 1
                    tag = f"Different disposition ({p_disp or '(Not yet determined)'})"
                partner_summaries.append(f"{p_name} [{tag}]")
                partner_detail.append({"partner_name": p_name, "partner_appid": p_appid,
                                        "partner_disposition": p_disp or "(Not yet determined)", "tag": tag})
            for row in integ["rows"]:
                integration_detail_rows.append({
                    "App Name": a["name"], "App ID": appid, "Disposition": disp or "(TBD)",
                    "Direction": row["direction"], "Data Object": row["data_object"],
                    "Interface": row["display_text"], "Partner App": row["partner_name"],
                    "Partner App ID": row["partner_appid"],
                    "Partner Disposition": (by_appid.get(row["partner_appid"]) or {}).get("disposition") or "(Not yet determined)",
                })

        interface_count = len(integ["rows"]) if integ else 0
        distinct_partners = len(integ["partner_appids"]) if integ else 0

        if distinct_partners > 0:
            apps_with_integrations.append({
                "name": a["name"], "appid": appid, "disposition": disp or "(Not yet determined)",
                "method": method or "", "shared_dedicated": shared_dedicated or "",
                "guidance": guidance, "impact_type": impact_type,
                "interface_count": interface_count, "distinct_partners": distinct_partners,
                "same_disp_count": same_disp_count, "diff_disp_count": diff_disp_count,
                "partners": partner_detail,
            })

        if impact_type == "data_extract_only":
            integration_impact = (f"No action needed \u2014 application stays with SBD. "
                                   f"{distinct_partners} in-inventory partner application(s) continue "
                                   "integrating unchanged." if distinct_partners else
                                   "No in-inventory integration partners found.")
        elif impact_type == "real_migration":
            integration_impact = (
                f"{distinct_partners} in-inventory partner application(s); {diff_disp_count} on a "
                "different disposition path must be rebuilt, replaced, or formally retired before "
                "cutover." if distinct_partners else "No in-inventory integration partners found.")
        elif impact_type == "conflict_needs_clarification":
            integration_impact = (f"{distinct_partners} in-inventory partner application(s) \u2014 "
                                   "impact cannot be assessed until the conflict above is resolved.")
        else:
            integration_impact = "Not applicable."

        risk_flags = []
        if disp is None:
            risk_flags.append("No disposition set")
        if impact_type == "conflict_needs_clarification":
            risk_flags.append("Disposition vs Shared/Dedicated conflict \u2014 needs business clarification")
            conflicts.append(a["name"])
        if a["comingled"] == "Yes":
            risk_flags.append("Co-mingled data \u2014 confirm extract isolates only divesting business data")
        if impact_type == "real_migration" and diff_disp_count > 0:
            risk_flags.append(f"{diff_disp_count} in-inventory partner(s) on a different disposition path")
        if a["tsa_support"] != "Y" and impact_type == "real_migration":
            risk_flags.append("No TSA support flagged \u2014 confirm cutover coverage")

        if impact_type == "real_migration":
            movers.append({
                "name": a["name"], "appid": appid or "(none)", "disposition": disp,
                "method": method or "(none)", "shared_dedicated": shared_dedicated or "",
                "distinct_partners": distinct_partners, "interface_count": interface_count,
                "guidance": guidance, "integration_impact": integration_impact,
                "risk_flags": list(risk_flags),
            })

        plan_rows.append({
            "App ID": appid or "(none)", "Application Name": a["name"], "Description": a["desc"],
            "Capability": a["capability"], "Application Track": a["track"],
            "SBD Proposed Disposition": disp or "(Not yet determined)",
            "Disposition Method": method or "", "Shared vs Dedicated": shared_dedicated or "",
            "Does Application Move?": {
                "data_extract_only": "No \u2014 data only",
                "real_migration": "Yes \u2014 application relocates/transfers",
                "conflict_needs_clarification": "Unclear \u2014 conflicting flags",
                "not_applicable": "TBD", "review_manually": "Review manually",
            }[impact_type],
            "SA Recommended Approach": guidance, "Integration Impact": integration_impact,
            "Hosting Type": a["hosting"] or "", "Location": a["location"] or "", "Sites": a["sites"] or "",
            "SBD IT Owner": a["sbd_owner"] or "", "Buyer IT Owner": a["buyer_owner"] or "",
            "TSA Support (Y/N)": a["tsa_support"] or "", "Co-mingled Data": a["comingled"] or "",
            "In-Inventory Integration Interfaces": interface_count,
            "In-Inventory Partner Applications": distinct_partners,
            "Partners on Different Disposition Path": diff_disp_count,
            "Partner Applications (detail)": "; ".join(partner_summaries) if partner_summaries else "None",
            "Risk / Attention Flags": "; ".join(risk_flags) if risk_flags else "None identified",
            "Vendor": a["vendor"] or "", "License Type": a["license"] or "",
            "Contract Status": a["contract_status"] or "", "Comments": a["comments"] or "",
        })
        if disp is None:
            open_items.append({"App ID": appid or "(none)", "Application Name": a["name"],
                                "Description": a["desc"], "SBD Owner": a["sbd_owner"] or "",
                                "Comments": a["comments"] or ""})

    plan_rows.sort(key=lambda r: (r["Does Application Move?"] != "Yes \u2014 application relocates/transfers",
                                   r["SBD Proposed Disposition"], -r["In-Inventory Integration Interfaces"]))
    apps_with_integrations.sort(key=lambda r: -r["interface_count"])
    movers.sort(key=lambda r: r["name"])
    return plan_rows, integration_detail_rows, open_items, conflicts, apps_with_integrations, movers


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11, name="Segoe UI")
HEADER_FILL = PatternFill("solid", fgColor=NAVY_HEX)
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(bottom=Side(style="thin", color="D8DFDC"))


def write_table(ws, headers, rows, col_widths=None):
    ws.freeze_panes = "A2"
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for ri, row in enumerate(rows, 2):
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=ri, column=ci, value=row.get(h, ""))
            cell.alignment = WRAP
            cell.border = THIN
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=LIGHT_HEX)
    if col_widths:
        for ci, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"


def write_excel(plan_rows, integration_detail_rows, open_items, conflicts, apps_with_integrations, movers, output_path: Path):
    wb = openpyxl.Workbook()
    plan_ws = wb.active
    plan_ws.title = "Separation Plan"
    plan_widths = [10, 30, 34, 16, 18, 26, 16, 16, 28, 44, 50, 14, 16, 16, 18, 18, 14, 12, 16, 16,
                   18, 44, 36, 14, 16, 16, 30]
    write_table(plan_ws, list(plan_rows[0].keys()), plan_rows, plan_widths)

    glossary_ws = wb.create_sheet("Glossary", 1)
    glossary_ws["A1"], glossary_ws["B1"], glossary_ws["C1"] = "Category", "Value", "Meaning"
    for ci in range(1, 4):
        glossary_ws.cell(row=1, column=ci).font = HEADER_FONT
        glossary_ws.cell(row=1, column=ci).fill = HEADER_FILL
    for ri, (cat, val, meaning) in enumerate(GLOSSARY, 2):
        glossary_ws.cell(row=ri, column=1, value=cat).alignment = WRAP
        glossary_ws.cell(row=ri, column=2, value=val).alignment = WRAP
        glossary_ws.cell(row=ri, column=3, value=meaning).alignment = WRAP
        if ri % 2 == 0:
            for ci in range(1, 4):
                glossary_ws.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor=LIGHT_HEX)
    glossary_ws.column_dimensions["A"].width = 22
    glossary_ws.column_dimensions["B"].width = 30
    glossary_ws.column_dimensions["C"].width = 100
    glossary_ws.freeze_panes = "A2"

    open_ws = wb.create_sheet("Open Items - No Disposition")
    if open_items:
        write_table(open_ws, list(open_items[0].keys()), open_items, [12, 34, 44, 20, 30])
    else:
        open_ws["A1"] = "No open items."

    movers_ws = wb.create_sheet("Fully Transferring Applications", 2)
    movers_ws["A1"] = "Applications That Actually Relocate / Transfer Ownership"
    movers_ws["A1"].font = Font(bold=True, size=14, color=NAVY_HEX, name="Segoe UI")
    movers_ws["A2"] = ("Direct Transfer, or Lift & Shift + Dedicated \u2014 these are the only "
                        "dispositions where the application itself physically moves or changes "
                        "ownership; every other disposition in this plan stays with SBD.")
    movers_ws["A2"].font = Font(italic=True, color="5C6B73", size=10, name="Segoe UI")
    movers_headers = ["Application", "App ID", "Disposition", "Method", "Shared vs Dedicated",
                       "In-Inventory Partner Apps", "In-Inventory Interfaces", "Integration Impact",
                       "SA Recommended Approach", "Risk / Attention Flags"]
    for c, h in enumerate(movers_headers, 1):
        movers_ws.cell(row=4, column=c, value=h).font = HEADER_FONT
        movers_ws.cell(row=4, column=c).fill = HEADER_FILL
    r = 5
    for row in movers:
        note = row["integration_impact"]
        if row["distinct_partners"] == 0:
            note = ("No interfaces found for this application in the Integration catalog "
                    "(either it has no AppID match or no sheet exists for it there) \u2014 "
                    "confirm integration scope directly with the application owner before cutover.")
        values = [row["name"], row["appid"], row["disposition"], row["method"], row["shared_dedicated"],
                  row["distinct_partners"], row["interface_count"], note, row["guidance"],
                  "; ".join(row["risk_flags"]) if row["risk_flags"] else "None identified"]
        for c, v in enumerate(values, 1):
            cell = movers_ws.cell(row=r, column=c, value=v)
            cell.alignment = WRAP
            if r % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=LIGHT_HEX)
        r += 1
    for col, width in zip("ABCDEFGHIJ", [30, 12, 26, 16, 16, 14, 14, 50, 50, 36]):
        movers_ws.column_dimensions[col].width = width
    movers_ws.freeze_panes = "A5"

    detail_ws = wb.create_sheet("Integration Detail (In-Inventory)")
    if integration_detail_rows:
        write_table(detail_ws, list(integration_detail_rows[0].keys()), integration_detail_rows,
                    [30, 12, 26, 26, 20, 40, 30, 14, 22])

    summary_ws = wb.create_sheet("Summary", 0)
    disp_counts = collections.Counter(r["SBD Proposed Disposition"] for r in plan_rows)
    summary_ws["A1"] = "Divestiture \u2014 Solution Architecture Separation Plan"
    summary_ws["A1"].font = Font(bold=True, size=16, color=NAVY_HEX, name="Segoe UI")
    summary_ws["A2"] = ("Note: integrations are reported ONLY where both the application and its "
                         "partner appear in the IT Application Inventory. See the Glossary tab for "
                         "disposition/method definitions.")
    summary_ws["A2"].font = Font(italic=True, color="5C6B73", size=10, name="Segoe UI")
    summary_ws["A4"], summary_ws["B4"] = "Disposition", "App Count"
    summary_ws["A4"].font = summary_ws["B4"].font = HEADER_FONT
    summary_ws["A4"].fill = summary_ws["B4"].fill = HEADER_FILL
    r = 5
    for disp, cnt in disp_counts.most_common():
        summary_ws.cell(row=r, column=1, value=disp)
        summary_ws.cell(row=r, column=2, value=cnt)
        r += 1
    summary_ws.column_dimensions["A"].width = 34
    summary_ws.column_dimensions["B"].width = 14

    r += 1
    summary_ws.cell(row=r, column=1,
                     value=f"Applications that actually relocate / transfer ownership: {len(movers)} "
                           "\u2014 see 'Fully Transferring Applications' tab").font = Font(bold=True, color="C24B3D", name="Segoe UI")
    r += 2
    summary_ws.cell(row=r, column=1, value="Applications with in-inventory integration partners").font = Font(bold=True, color=NAVY_HEX, name="Segoe UI")
    r += 1
    headers2 = ["Application", "Disposition", "Partner Apps", "Interfaces", "Different Disposition Path"]
    for c, h in enumerate(headers2, 1):
        summary_ws.cell(row=r, column=c, value=h).font = HEADER_FONT
        summary_ws.cell(row=r, column=c).fill = HEADER_FILL
    r += 1
    for row in apps_with_integrations:
        summary_ws.cell(row=r, column=1, value=row["name"])
        summary_ws.cell(row=r, column=2, value=row["disposition"])
        summary_ws.cell(row=r, column=3, value=row["distinct_partners"])
        summary_ws.cell(row=r, column=4, value=row["interface_count"])
        summary_ws.cell(row=r, column=5, value=row["diff_disp_count"])
        r += 1

    r += 2
    summary_ws.cell(row=r, column=1, value="Conflicts needing business clarification").font = Font(bold=True, color="C24B3D", name="Segoe UI")
    r += 1
    for name in conflicts:
        summary_ws.cell(row=r, column=1, value=name)
        r += 1
    if not conflicts:
        summary_ws.cell(row=r, column=1, value="None")
    summary_ws.column_dimensions["C"].width = 16
    summary_ws.column_dimensions["D"].width = 14
    summary_ws.column_dimensions["E"].width = 24

    wb.save(output_path)


# ---------------------------------------------------------------------------
# PowerPoint output
# ---------------------------------------------------------------------------
def add_background(prs, slide):
    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG
    bg.line.fill.background()
    bg.shadow.inherit = False
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)


def add_header(prs, slide, title_text, size=25):
    header = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(1.0))
    header.fill.solid()
    header.fill.fore_color.rgb = NAVY
    header.line.fill.background()
    header.shadow.inherit = False
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.14), Inches(12.3), Inches(0.75))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = WHITE
    run.font.name = "Segoe UI"


def add_card(slide, x, y, w, h, label, value, color):
    card = slide.shapes.add_shape(1, x, y, w, h)
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = RGBColor(0xD9, 0xE2, 0xE5)
    card.line.width = Pt(1)
    card.shadow.inherit = False
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.12)
    tf.margin_top = Inches(0.08)
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = label
    run.font.size = Pt(12)
    run.font.color.rgb = GRAY
    run.font.name = "Segoe UI"
    p2 = tf.add_paragraph()
    run2 = p2.add_run()
    run2.text = str(value)
    run2.font.size = Pt(26)
    run2.font.bold = True
    run2.font.color.rgb = color
    run2.font.name = "Segoe UI"


def set_fixed_row_heights(table, header_height_in, data_height_in, n_data_rows):
    """Explicitly size every row so python-pptx never stretches a 1-2 row
    table to fill the full allotted table height (its default behavior)."""
    table.rows[0].height = Inches(header_height_in)
    for i in range(1, n_data_rows + 1):
        table.rows[i].height = Inches(data_height_in)


def add_footer(slide, prs):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(7.1), Inches(12.3), Inches(0.3))
    p = box.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = ("Source: IT Application Inventory + Application Integration catalog \u2014 "
                "integrations shown only where both applications are in the inventory")
    run.font.size = Pt(9.5)
    run.font.italic = True
    run.font.color.rgb = GRAY
    run.font.name = "Segoe UI"


def build_ppt(plan_rows, open_items, conflicts, apps_with_integrations, movers, output_path: Path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    disp_counts = collections.Counter(r["SBD Proposed Disposition"] for r in plan_rows)

    # --- Slide 1: Executive overview ---------------------------------------
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(prs, slide)
    add_header(prs, slide, "Divestiture \u2014 Solution Architecture Separation Plan")

    data_extract_count = sum(1 for r in plan_rows if r["Does Application Move?"] == "No \u2014 data only")
    moving_count = sum(1 for r in plan_rows if r["Does Application Move?"] == "Yes \u2014 application relocates/transfers")
    card_specs = [
        ("In-Scope Applications", len(plan_rows), NAVY),
        ("Data Extract Only (stays with SBD)", data_extract_count, TEAL),
        ("Actually Relocate / Transfer Ownership", moving_count, AMBER),
        ("Disposition Not Yet Determined", len(open_items), AMBER),
        ("Apps with In-Inventory Integration Partners", len(apps_with_integrations), RED),
    ]
    card_w, card_h = Inches(2.35), Inches(1.35)
    start_x, start_y, gap_x = Inches(0.5), Inches(1.3), Inches(0.2)
    for i, (label, value, color) in enumerate(card_specs):
        add_card(slide, start_x + i * (card_w + gap_x), start_y, card_w, card_h, label, value, color)

    note_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.85), Inches(12.3), Inches(0.9))
    tf = note_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = ("Scope rule: integrations are reported ONLY where both the application and its "
                "integration partner appear in the IT Application Inventory. Systems outside the "
                "inventory are excluded, not labeled as \u201cretained\u201d or \u201cexternal.\u201d")
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = TEAL
    run.font.name = "Segoe UI"
    p.line_spacing = 1.15

    goal_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.85), Inches(12.3), Inches(0.9))
    tf = goal_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "Goal"
    run.font.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = NAVY
    run.font.name = "Segoe UI"
    p2 = tf.add_paragraph()
    run2 = p2.add_run()
    run2.text = ("Translate each application's business-approved disposition into a concrete SA "
                 "plan: does the application actually move, and for apps that integrate with "
                 "another inventory application, what needs to happen to that interface before "
                 "TSA exit.")
    run2.font.size = Pt(13.5)
    run2.font.color.rgb = DARK
    run2.font.name = "Segoe UI"
    p2.line_spacing = 1.1

    risk_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.85), Inches(12.3), Inches(2.1))
    tf = risk_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "Key SA Observations"
    run.font.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RED
    run.font.name = "Segoe UI"
    conflict_note = (f"{len(conflicts)} disposition conflict(s) ({', '.join(conflicts)}) need business "
                      "clarification before planning." if conflicts else "No disposition conflicts detected.")
    bullets = [
        f"{moving_count} application(s) actually relocate or transfer ownership (Direct Transfer, "
        "or Lift & Shift + Dedicated) \u2014 see the dedicated slide that follows for every one of "
        "them and their integration status.",
        f"{len(apps_with_integrations)} application(s) have at least one integration partner that "
        "is also in the inventory \u2014 see the detail slides that follow for every partner and "
        "disposition alignment.",
        conflict_note,
        f"{len(open_items)} application(s) have no disposition yet \u2014 no technical planning can "
        "start for these.",
    ]
    for b in bullets:
        pp = tf.add_paragraph()
        run = pp.add_run()
        run.text = f"\u2022  {b}"
        run.font.size = Pt(13)
        run.font.color.rgb = DARK
        run.font.name = "Segoe UI"
        pp.line_spacing = 1.1
        pp.space_after = Pt(6)

    # --- Slide 2: Fully transferring applications ---------------------------
    if movers:
        slide_m = prs.slides.add_slide(prs.slide_layouts[6])
        add_background(prs, slide_m)
        add_header(prs, slide_m, "Applications That Actually Relocate / Transfer Ownership")

        note_box = slide_m.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12.3), Inches(0.5))
        tf = note_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = ("Direct Transfer, or Lift & Shift + Dedicated \u2014 the only dispositions where "
                    "the application itself physically moves or changes ownership.")
        run.font.size = Pt(13)
        run.font.italic = True
        run.font.color.rgb = GRAY
        run.font.name = "Segoe UI"

        n = len(movers)
        headers_m = ["Application", "Disposition", "Method", "Shared/\nDedicated", "In-Inv.\nPartners",
                     "Integration Status"]
        col_widths = [2.9, 2.1, 1.6, 1.1, 1.0, 3.6]
        table_shape = slide_m.shapes.add_table(n + 1, len(headers_m), Inches(0.5), Inches(1.7),
                                                Inches(sum(col_widths)), Inches(0.5))
        table = table_shape.table
        for c, w in enumerate(col_widths):
            table.columns[c].width = Inches(w)
        for c, h in enumerate(headers_m):
            cell = table.cell(0, c)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(12)
            p.font.color.rgb = WHITE
            p.font.name = "Segoe UI"
        for r, row in enumerate(movers, 1):
            if row["distinct_partners"] == 0:
                status = "No interfaces found in Integration catalog \u2014 confirm with app owner"
            else:
                status = f"{row['distinct_partners']} in-inventory partner app(s), {row['interface_count']} interface(s)"
            values = [row["name"], row["disposition"], row["method"], row["shared_dedicated"],
                      str(row["distinct_partners"]), status]
            for c, v in enumerate(values):
                cell = table.cell(r, c)
                cell.text = v
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF1, 0xF5, 0xF4) if r % 2 == 0 else WHITE
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(11)
                p.font.color.rgb = DARK
                p.font.name = "Segoe UI"
        available_height_in = 5.3
        row_height_in = min(0.5, max(0.3, available_height_in / (n + 1)))
        set_fixed_row_heights(table, row_height_in, row_height_in, n)
        add_footer(slide_m, prs)

    # --- Slide 3: Overview table --------------------------------------------
    if apps_with_integrations:
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        add_background(prs, slide2)
        add_header(prs, slide2, "Applications with In-Inventory Integration Partners \u2014 Overview")

        n = len(apps_with_integrations)
        table_shape = slide2.shapes.add_table(n + 1, 4, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5))
        table = table_shape.table
        table.columns[0].width = Inches(5.5)
        table.columns[1].width = Inches(3.8)
        table.columns[2].width = Inches(1.5)
        table.columns[3].width = Inches(1.5)
        headers = ["Application", "Disposition", "Partner Apps", "Interfaces"]
        for c, h in enumerate(headers):
            cell = table.cell(0, c)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(13)
            p.font.color.rgb = WHITE
            p.font.name = "Segoe UI"
        for r, row in enumerate(apps_with_integrations, 1):
            values = [row["name"], row["disposition"], str(row["distinct_partners"]), str(row["interface_count"])]
            for c, v in enumerate(values):
                cell = table.cell(r, c)
                cell.text = v
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF1, 0xF5, 0xF4) if r % 2 == 0 else WHITE
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(11.5)
                p.font.color.rgb = DARK
                p.font.name = "Segoe UI"
        available_height_in = 5.8
        row_height_in = min(0.5, max(0.3, available_height_in / (n + 1)))
        set_fixed_row_heights(table, row_height_in, row_height_in, n)
        add_footer(slide2, prs)

    # --- One fully self-contained detail slide per application --------------
    for app in apps_with_integrations:
        s = prs.slides.add_slide(prs.slide_layouts[6])
        add_background(prs, s)
        add_header(prs, s, app["name"], size=24)

        meta_box = s.shapes.add_textbox(Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4))
        tf = meta_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        meta_parts = [f"App ID: {app['appid'] or 'n/a'}", f"Disposition: {app['disposition']}",
                      f"Method: {app['method'] or 'n/a'}", f"Shared/Dedicated: {app['shared_dedicated'] or 'n/a'}"]
        run = p.add_run()
        run.text = "   |   ".join(meta_parts)
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = TEAL
        run.font.name = "Segoe UI"

        guidance_box = s.shapes.add_textbox(Inches(0.5), Inches(1.6), Inches(12.3), Inches(1.1))
        tf = guidance_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "SA Recommended Approach: "
        run.font.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = NAVY
        run.font.name = "Segoe UI"
        run2 = p.add_run()
        run2.text = app["guidance"]
        run2.font.size = Pt(12)
        run2.font.color.rgb = DARK
        run2.font.name = "Segoe UI"
        p.line_spacing = 1.05

        stat_y = Inches(2.75)
        card_w2, card_h2 = Inches(2.9), Inches(0.9)
        stat_specs = [
            ("Partner Applications", app["distinct_partners"], NAVY),
            ("Total Interfaces", app["interface_count"], TEAL),
            ("Same Disposition Path", app["same_disp_count"], TEAL),
            ("Different Disposition Path", app["diff_disp_count"], RED if app["diff_disp_count"] else GRAY),
        ]
        for i, (label, value, color) in enumerate(stat_specs):
            x = Inches(0.5) + i * (card_w2 + Inches(0.15))
            add_card(s, x, stat_y, card_w2, card_h2, label, value, color)

        # Partner table -- every partner is listed, none deferred to Excel.
        partners = app["partners"]
        table_top_in = 3.85
        table_bottom_in = 7.05
        available_height_in = table_bottom_in - table_top_in
        n_data_rows = len(partners)
        tshape = s.shapes.add_table(n_data_rows + 1, 3, Inches(0.5), Inches(table_top_in),
                                     Inches(12.3), Inches(available_height_in))
        tbl = tshape.table
        tbl.columns[0].width = Inches(5.5)
        tbl.columns[1].width = Inches(4.0)
        tbl.columns[2].width = Inches(2.8)
        phead = ["Partner Application", "Partner Disposition", "Relationship"]
        for c, h in enumerate(phead):
            cell = tbl.cell(0, c)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(12)
            p.font.color.rgb = WHITE
            p.font.name = "Segoe UI"

        # Row height scales with row count so a 2-row table isn't stretched
        # into oversized rows, and a 12-row table doesn't overflow the slide.
        row_height_in = min(0.5, max(0.26, available_height_in / (n_data_rows + 1)))
        font_size = 11 if row_height_in >= 0.32 else 9.5
        set_fixed_row_heights(tbl, row_height_in, row_height_in, n_data_rows)

        for r, partner in enumerate(partners, 1):
            relationship = "Same disposition" if "Same disposition" in partner["tag"] else "Different disposition"
            values = [partner["partner_name"], partner["partner_disposition"], relationship]
            for c, v in enumerate(values):
                cell = tbl.cell(r, c)
                cell.text = v
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF1, 0xF5, 0xF4) if r % 2 == 0 else WHITE
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(font_size)
                p.font.color.rgb = DARK
                p.font.name = "Segoe UI"
                cell.margin_top = Pt(2)
                cell.margin_bottom = Pt(2)
        add_footer(s, prs)

    prs.save(output_path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args() -> argparse.Namespace:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
    from sa_paths import output_dir_for

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--inventory", type=Path, required=True,
                         help="Path to the IT Application Inventory workbook (e.g. dropped in SA\\Inputs).")
    parser.add_argument("--integrations", type=Path, required=True,
                         help="Path to the Application Integration catalog workbook (e.g. dropped in SA\\Inputs).")
    parser.add_argument("--output-dir", type=Path, default=output_dir_for("divestiture-separation-plan"),
                         help="Defaults to the shared SA\\Outputs\\divestiture-separation-plan folder.")
    parser.add_argument("--excel-output", type=Path,
                         default=Path("Divestiture-SA-Separation-Plan.xlsx"))
    parser.add_argument("--ppt-output", type=Path,
                         default=Path("Divestiture-SA-Separation-Plan-Executive-Summary.pptx"))
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.inventory.exists():
        raise SystemExit(f"Inventory file not found: {args.inventory}")
    if not args.integrations.exists():
        raise SystemExit(f"Integrations file not found: {args.integrations}")

    apps = load_inventory(args.inventory)
    if not apps:
        raise SystemExit("No named applications found in the inventory workbook.")
    by_appid = {a["appid"]: a for a in apps if a["appid"]}
    integration_data = load_integrations(args.integrations, by_appid)

    plan_rows, integration_detail_rows, open_items, conflicts, apps_with_integrations, movers = build_plan(
        apps, by_appid, integration_data)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    excel_path = args.excel_output if args.excel_output.is_absolute() else args.output_dir / args.excel_output.name
    ppt_path = args.ppt_output if args.ppt_output.is_absolute() else args.output_dir / args.ppt_output.name

    write_excel(plan_rows, integration_detail_rows, open_items, conflicts, apps_with_integrations, movers, excel_path)
    build_ppt(plan_rows, open_items, conflicts, apps_with_integrations, movers, ppt_path)

    print(f"In-scope applications: {len(plan_rows)}")
    print(f"Applications that actually relocate/transfer: {len(movers)}")
    print(f"Applications with in-inventory integration partners: {len(apps_with_integrations)}")
    print(f"Open items (no disposition): {len(open_items)}")
    print(f"Conflicts needing clarification: {conflicts}")
    print(f"Created {excel_path}")
    print(f"Created {ppt_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
