# -*- coding: utf-8 -*-
from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any

from workbook_catalog import (
    best_act_lot,
    filename_lot,
    index_dce_files,
    inspect_workbook,
    norm_lot,
)

EXCLUDED = {"DOCUMENT", "RECAP", "ANNEXE"}


def _meta(mapping, original):
    return mapping.get(str(original), {}) if isinstance(mapping, dict) else {}


def _enterprise_name(original, enterprises):
    item = _meta(enterprises, original)
    value = str(item.get("enterprise_name") or "").strip()
    return value or Path(str(original)).stem


def _fallback_lot(original, lots, enterprises):
    for source in (enterprises, lots):
        item = _meta(source, original)
        value = norm_lot(item.get("lot_number"))
        if value:
            return value
    return ""


def _entries(path, role):
    return [
        entry for entry in inspect_workbook(path, role)
        if entry.lot and entry.category not in EXCLUDED
    ]


def _dce_priority(original, path, entry):
    explicit_file_lot = filename_lot(original) or filename_lot(path)
    # Un DCE unitaire explicitement nommé passe avant un classeur général.
    return (
        1 if explicit_file_lot == entry.lot else 0,
        entry.score,
        -len(Path(str(original)).name),
    )


def _dce_catalog(dces):
    catalog = defaultdict(list)
    for original, path in dces:
        for entry in _entries(path, "DCE"):
            catalog[(entry.lot, entry.category)].append({
                "original": original,
                "path": path,
                "sheet": entry.sheet,
                "entry": entry,
            })
    for key, items in catalog.items():
        items.sort(
            key=lambda item: _dce_priority(item["original"], item["path"], item["entry"]),
            reverse=True,
        )
    return catalog


def _legacy_jobs(dces, acts, lots, enterprises):
    index, warnings = index_dce_files(dces)
    jobs = []
    diagnostics = []
    for original, path in acts:
        fallback = _fallback_lot(original, lots, enterprises)
        lot, entries = best_act_lot(path, fallback)
        if not lot:
            diagnostics.append(f"{original} : aucun lot exploitable détecté")
            continue
        references = index.get(lot, [])
        if not references:
            diagnostics.append(f"Lot {lot} : aucune référence DCE disponible pour {original}")
            continue
        jobs.append({
            "lot": lot,
            "enterprise": _enterprise_name(original, enterprises),
            "dce_path": references[0]["path"],
            "act_path": path,
        })
    return jobs, {
        "mode": "HISTORIQUE",
        "warnings": list(warnings) + diagnostics,
        "lots_ready": sorted({job["lot"] for job in jobs}),
        "jobs": len(jobs),
    }


def build_case_jobs(dces, acts, lots=None, enterprises=None):
    """Construit des jobs historiques ou des jobs TCE pointant vers des feuilles.

    Le mode TCE ne s'active que si au moins un ACT contient plusieurs lots métier
    distincts. Un dossier mono-lot conserve donc exactement le pipeline historique.
    """
    lots = lots or {}
    enterprises = enterprises or {}
    act_catalogs = []
    multi_lot = False
    for original, path in acts:
        entries = _entries(path, "ACT")
        distinct_lots = {entry.lot for entry in entries}
        if len(distinct_lots) >= 2:
            multi_lot = True
        act_catalogs.append((original, path, entries))

    if not multi_lot:
        return _legacy_jobs(dces, acts, lots, enterprises)

    dce_catalog = _dce_catalog(dces)
    jobs = []
    warnings = []
    coverage = defaultdict(lambda: {
        "dce": False,
        "companies": [],
        "status": "INCOMPLET",
    })

    for original, path, entries in act_catalogs:
        enterprise = _enterprise_name(original, enterprises)
        seen = set()
        for entry in entries:
            key = (entry.lot, entry.category, entry.sheet)
            if key in seen:
                continue
            seen.add(key)
            refs = dce_catalog.get((entry.lot, entry.category), [])
            coverage_key = f"{entry.lot} {entry.category}"
            coverage[coverage_key]["companies"].append(enterprise)
            if not refs:
                warnings.append(
                    f"Lot {entry.lot} {entry.category} : aucune référence DCE, "
                    f"feuille {entry.sheet} de {enterprise} non produite."
                )
                continue
            ref = refs[0]
            coverage[coverage_key]["dce"] = True
            coverage[coverage_key]["status"] = "PRET"
            jobs.append({
                "lot": entry.lot,
                "category": entry.category,
                "enterprise": enterprise,
                "dce_path": ref["path"],
                "dce_sheet": ref["sheet"],
                "act_path": path,
                "act_sheet": entry.sheet,
                "virtual_sheet_job": True,
            })

    # Déduplique les entreprises dans le diagnostic sans modifier les jobs.
    for value in coverage.values():
        value["companies"] = sorted(set(value["companies"]), key=str.casefold)

    return jobs, {
        "mode": "TCE_MULTI_LOTS",
        "warnings": warnings,
        "lots_ready": sorted({job["lot"] for job in jobs}),
        "jobs": len(jobs),
        "coverage": dict(sorted(coverage.items())),
    }
