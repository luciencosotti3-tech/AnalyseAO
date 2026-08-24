# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
import json
import re
import shutil
import uuid

from flask import Flask, jsonify, request, send_from_directory
from werkzeug.utils import secure_filename

from side_by_side import build_side_by_side
from multi_company import build_multi_company_workbook, normalize_lot
from workbook_catalog import index_dce_files, best_act_lot
from case_matrix import build_case_jobs
from offer_analysis import analyse_workbook

ROOT = Path(__file__).resolve().parent.parent
FRONTEND = ROOT / "frontend"
RUNTIME = Path(__file__).resolve().parent / "runtime"
UPLOADS = RUNTIME / "uploads"
OUTPUTS = RUNTIME / "outputs"
UPLOADS.mkdir(parents=True, exist_ok=True)
OUTPUTS.mkdir(parents=True, exist_ok=True)
ALLOWED = {".xlsx", ".xlsm"}
app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024


def response(status="OK", http=200, **payload):
    return jsonify({"status": status, **payload}), http


def json_list(field):
    raw = request.form.get(field, "")
    if not raw:
        return []
    value = json.loads(raw)
    if not isinstance(value, list):
        raise ValueError(f"{field} doit contenir une liste JSON.")
    return value


def by_name(items):
    return {str(item.get("original_name")): item for item in items if isinstance(item, dict) and item.get("original_name")}


def save_excel(storage, folder):
    # TCE_VIRTUAL_SHEETS_V22
    original = storage.filename or ""
    safe = secure_filename(original)
    if not safe or Path(safe).suffix.lower() not in ALLOWED:
        raise ValueError(f"{original or 'Fichier sans nom'} : seuls .xlsx et .xlsm sont acceptés.")
    target = folder / safe
    index = 2
    stem = Path(safe).stem
    suffix = Path(safe).suffix
    while target.exists():
        target = folder / f"{stem}_{index}{suffix}"
        index += 1
    storage.save(target)
    return original, target


@app.get("/api/health")
def health():
    return response(message="AnalyseAO disponible", port=8008)


@app.post("/api/run/files")
def run_files():
    dce_uploads = request.files.getlist("dce_files")
    act_uploads = request.files.getlist("act_files")
    if not dce_uploads or not act_uploads:
        return response("ERROR", 400, errors=["Au moins un DCE et un ACT sont obligatoires."])
    try:
        lots = by_name(json_list("dropped_file_lots"))
        enterprises = by_name(json_list("act_enterprise_names"))
        run_id = uuid.uuid4().hex
        work = UPLOADS / run_id
        work.mkdir(parents=True)
        dces = [save_excel(item, work) for item in dce_uploads]
        acts = [save_excel(item, work) for item in act_uploads]

        jobs, matrix_report = build_case_jobs(dces, acts, lots, enterprises)
        if not jobs:
            details = matrix_report.get("warnings") or ["Aucun couple DCE/ACT compatible détecté."]
            return response("ERROR", 400, errors=details, case_matrix=matrix_report)

        output_name = f"{run_id}_tableau_restructure.xlsx"
        output_path = OUTPUTS / output_name
        build_multi_company_workbook(jobs, output_path, work)
        warnings = matrix_report.get("warnings", [])
        return response(
            workbook_path=output_path.name,
            download_url=f"/api/download/{output_path.name}",
            case_matrix=matrix_report,
            warnings=warnings,
            errors=[],
        )
    except Exception as exc:
        app.logger.exception("Echec de la restructuration")
        return response("ERROR", 500, errors=[f"Erreur de traitement : {exc}"])



@app.post("/api/run/offer-analysis")
def run_offer_analysis():
    upload = request.files.get("workbook")
    if upload is None or not upload.filename:
        return response("ERROR", 400, errors=["Classeur restructuré manquant."])
    if Path(upload.filename).suffix.lower() != ".xlsx":
        return response("ERROR", 400, errors=["Le classeur d'analyse doit être au format .xlsx."])
    try:
        run_id = uuid.uuid4().hex
        work = UPLOADS / run_id
        work.mkdir(parents=True)
        _, source_path = save_excel(upload, work)
        output_name = f"analyse_offres_{Path(upload.filename).stem}.xlsx"
        output_path = OUTPUTS / f"{run_id}_{output_name}"
        _, report = analyse_workbook(source_path, output_path)
        return response(
            workbook_path=output_path.name,
            download_url=f"/api/download/{output_path.name}",
            analysis_report=report,
            warnings=[], errors=[],
        )
    except Exception as exc:
        app.logger.exception("Echec de l'analyse d'offres")
        return response("ERROR", 500, errors=[f"Erreur d'analyse : {exc}"])


@app.get("/api/download/<path:filename>")
def download(filename):
    # Le fichier reste stocké avec un identifiant interne unique, mais le nom
    # présenté au navigateur vient du pop-up frontend.
    requested_name = str(request.args.get("download_name", "")).strip()
    if requested_name:
        requested_name = secure_filename(requested_name)
        if not requested_name.lower().endswith(".xlsx"):
            requested_name += ".xlsx"
    else:
        # Ne jamais exposer le préfixe UUID interne dans le nom téléchargé.
        requested_name = re.sub(r"^[0-9a-f]{32}_", "", Path(filename).name, flags=re.I)
    return send_from_directory(
        OUTPUTS,
        filename,
        as_attachment=True,
        download_name=requested_name,
        max_age=0,
    )


@app.get("/")
def index():
    return send_from_directory(FRONTEND, "index.html", max_age=0)


@app.get("/<path:filename>")
def frontend_file(filename):
    # Les routes /api inexistantes ne doivent jamais retomber sur le frontend.
    if filename.startswith("api/"):
        return response("ERROR", 404, errors=["Route API introuvable."])
    return send_from_directory(FRONTEND, filename, max_age=0)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8008, debug=False, use_reloader=False)
