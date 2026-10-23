"""Profiler page routes."""

import os
import zipfile
from datetime import datetime
from flask import Blueprint, jsonify, render_template, request
import pandas as pd
from xlrd.biffh import XLRDError
from .services import (
    MAX_YEAR,
    MIN_YEAR,
    PROFILE_EXCEL_PATHS,
    PROFILER_DATA,
    VARIABLES,
    build_rows,
    build_total_display,
    get_display_calculations,
    get_profile_type,
    get_years_for_range,
    leap_indicator,
    load_existing_data,
    results_to_display,
    save_form_data,
    update_profiler_snapshot,
)

profiler_blueprint = Blueprint("profiler", __name__)


@profiler_blueprint.route("/profiler", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        start_year = request.form.get("start_year", MIN_YEAR)
        end_year = request.form.get("end_year", MAX_YEAR)
        raw_profile_type = request.form.get("profile_type", "complete")
    else:
        start_year = request.args.get("start_year", MIN_YEAR)
        end_year = request.args.get("end_year", MAX_YEAR)
        raw_profile_type = request.args.get("profile_type", "complete")

    try:
        profile_type = get_profile_type(raw_profile_type)
    except ValueError as error:
        return str(error), 400
    years = get_years_for_range(start_year, end_year)
    workbook_error = None
    try:
        existing = load_existing_data(years=years, profile_type=profile_type)
    except (zipfile.BadZipFile, XLRDError):
        existing = {}
        workbook_error = (
            "The saved profiler workbook is damaged and cannot be read. "
            "Your saved workbook has not been changed. To continue, use "
            "'Back up damaged workbook and start fresh' below; the damaged "
            "file will be preserved as a backup."
        )
    calc_display = get_display_calculations(existing, years=years)
    rows = build_rows()
    totals = build_total_display(rows, years, existing, calc_display)
    return render_template(
        "index.html",
        rows=rows,
        years=years,
        leap_indicators={year: leap_indicator(year) for year in years},
        existing=existing,
        calc_display=calc_display,
        totals=totals,
        workbook_error=workbook_error,
        start_year=years[0],
        end_year=years[-1],
        profile_type=profile_type,
        page_title=(
            f"iCON Profiler - {profile_type.title()} "
            f"({years[0]} - {years[-1]})"
        ),
    )

@profiler_blueprint.route("/reset", methods=["POST"])
def reset():
    """Back up the current workbook, then clear the selected input profile."""
    try:
        profile_type = get_profile_type(
            request.form.get("profile_type", request.args.get("profile_type", "complete"))
        )
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    path = PROFILE_EXCEL_PATHS[profile_type]
    if os.path.exists(path):
        directory, filename = os.path.split(path)
        stem, extension = os.path.splitext(filename)
        backup_name = (
            f"{stem}.backup-{datetime.now().strftime('%Y%m%d-%H%M%S-%f')}"
            f"{extension}"
        )
        backup_path = os.path.join(directory, backup_name)
        try:
            os.replace(path, backup_path)
        except OSError as error:
            return jsonify({
                "error": f"Could not back up the existing workbook: {error}"
            }), 500
    else:
        backup_path = None
    for values in PROFILER_DATA[profile_type].values():
        values.clear()
    return jsonify({
        "status": "ok",
        "backup": os.path.basename(backup_path) if backup_path else None,
    })

@profiler_blueprint.route("/profiler-state", methods=["POST"])
def profiler_state():
    years = get_years_for_range(request.form.get("start_year", MIN_YEAR),
                                request.form.get("end_year", MAX_YEAR))
    try:
        profile_type = get_profile_type(request.form.get("profile_type", "complete"))
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    update_profiler_snapshot(request.form, years, profile_type=profile_type)
    return jsonify({"status": "ok"})

@profiler_blueprint.route("/autosave", methods=["POST"])
def autosave():
    """
    Called automatically (debounced) whenever an input changes. Saves the
    whole current grid to database.xlsx and returns the freshly calculated
    values so the page can update the read-only calc rows without a reload.
    """
    start_year = request.form.get("start_year", MIN_YEAR)
    end_year = request.form.get("end_year", MAX_YEAR)
    years = get_years_for_range(start_year, end_year)
    try:
        profile_type = get_profile_type(request.form.get("profile_type", "complete"))
    except ValueError as error:
        return jsonify({"error": str(error)}), 400

    try:
        raw_df, results = save_form_data(
            request.form, years=years, profile_type=profile_type
        )
    except (zipfile.BadZipFile, XLRDError):
        return jsonify({
            "error": (
                "The saved profiler workbook is damaged. Reset the profile to "
                "back it up and start fresh before saving."
            )
        }), 409
    except PermissionError:
        return jsonify({
            "error": (
                "Could not save the profiler workbook because access was denied. "
                "Close it in Excel and wait for OneDrive to finish syncing, then try again."
            )
        }), 423
    update_profiler_snapshot(request.form, years, profile_type=profile_type)

    existing = {}
    for key, label in VARIABLES:
        existing[key] = {}
        for year in years:
            if year not in raw_df.columns:
                continue
            try:
                raw_value = raw_df.loc[label, year]
            except Exception:
                raw_value = ""
            existing[key][int(year)] = "" if pd.isna(raw_value) else raw_value

    calc_display = results_to_display(results, years=years)
    rows = build_rows()
    totals = build_total_display(rows, years, existing, calc_display)
    return jsonify({
        "status": "ok",
        "calculations": calc_display,
        "totals": totals,
    })
