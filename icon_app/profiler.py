"""Profiler page routes."""

from io import BytesIO
import math
import os
import zipfile
from datetime import datetime
from flask import Blueprint, jsonify, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.exceptions import InvalidFileException
import pandas as pd
from xlrd.biffh import XLRDError
from .services import (
    MAX_YEAR,
    MIN_YEAR,
    PROFILER_DATA,
    VARIABLES,
    build_rows,
    build_total_display,
    get_display_calculations,
    get_profile_type,
    get_profile_workbook_path,
    get_years_for_range,
    leap_indicator,
    load_existing_data,
    load_project_name,
    results_to_display,
    save_form_data,
    update_profiler_snapshot,
)

profiler_blueprint = Blueprint("profiler", __name__)


def _render_profiler(years, profile_type, imported=None, project_name=None):
    workbook_error = None
    try:
        saved_data = load_existing_data(years=years, profile_type=profile_type)
        if project_name is None:
            project_name = load_project_name(profile_type)
    except (zipfile.BadZipFile, InvalidFileException, OSError, XLRDError):
        saved_data = {}
        project_name = project_name or ""
        workbook_error = (
            "The saved profiler workbook is damaged and cannot be read. "
            "Your saved workbook has not been changed. To continue, use "
            "'Back up damaged workbook and start fresh' below; the damaged "
            "file will be preserved as a backup."
        )

    if imported is None:
        existing = saved_data
    else:
        existing = {
            key: {
                year: imported.get(key, {}).get(year, "")
                for year in years
            }
            for key, _ in VARIABLES
        }

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
        project_name=project_name or "",
        imported=imported is not None,
        page_title=(
            f"iCON Profiler - {profile_type.title()} "
            f"({years[0]} - {years[-1]})"
        ),
    )


def _parse_template_year(value):
    if isinstance(value, bool) or value in (None, ""):
        return None
    try:
        numeric_year = float(value)
    except (TypeError, ValueError):
        return None
    if not numeric_year.is_integer():
        return None
    year = int(numeric_year)
    if not MIN_YEAR <= year <= MAX_YEAR:
        raise ValueError(
            f"Template year {year} is outside the supported range "
            f"({MIN_YEAR}-{MAX_YEAR})."
        )
    return year


def _read_profiler_template(file_stream):
    try:
        workbook = load_workbook(file_stream, read_only=True, data_only=False)
    except (zipfile.BadZipFile, InvalidFileException, OSError, ValueError) as error:
        raise ValueError("The uploaded file is not a readable .xlsx workbook.") from error

    try:
        worksheet = workbook.active
        if worksheet.max_row > 250 or worksheet.max_column > 100:
            raise ValueError("The uploaded worksheet is larger than the profiler template.")

        year_row = None
        year_columns = {}
        for row_number, row in enumerate(
            worksheet.iter_rows(values_only=True), start=1
        ):
            if not row or str(row[0] or "").strip().casefold() != "year":
                continue
            year_row = row_number
            for column, cell_value in enumerate(row):
                if column < 2:
                    continue
                year = _parse_template_year(cell_value)
                if year is None:
                    continue
                if year in year_columns.values():
                    raise ValueError(f"Year {year} appears more than once in the header.")
                year_columns[column] = year
            break

        if year_row is None or not year_columns:
            raise ValueError(
                "The workbook must contain a 'Year' header row with supported "
                f"year columns ({MIN_YEAR}-{MAX_YEAR}). Download the profiler "
                "template to get the correct format."
            )

        input_rows = [row for row in build_rows() if row["type"] == "input"]
        labels_to_keys = {
            row["display_label"]: row["key"] for row in input_rows
        }
        labels_to_keys.update({
            label: key
            for key, label in VARIABLES
            if key in labels_to_keys.values()
        })
        labels_to_keys.update({key: key for key, _ in VARIABLES})
        imported = {key: {} for key, _ in VARIABLES}
        seen_labels = set()
        populated_years = set()
        for row in worksheet.iter_rows(min_row=year_row + 1, values_only=True):
            if not row:
                continue
            label = str(row[0] or "").strip()
            key = labels_to_keys.get(label)
            if key is None:
                continue
            if key in seen_labels:
                raise ValueError(f"Profiler row '{label}' appears more than once.")
            seen_labels.add(key)
            for column, year in year_columns.items():
                if column >= len(row):
                    continue
                value = row[column]
                if value is None or (isinstance(value, str) and not value.strip()):
                    continue
                if isinstance(value, bool):
                    raise ValueError(f"'{label}' for {year} must be a number.")
                try:
                    numeric_value = float(value)
                except (TypeError, ValueError) as error:
                    raise ValueError(
                        f"'{label}' for {year} must be a number."
                    ) from error
                if not math.isfinite(numeric_value):
                    raise ValueError(f"'{label}' for {year} must be a finite number.")
                imported[key][year] = numeric_value
                populated_years.add(year)

        if not seen_labels:
            raise ValueError(
                "No profiler input rows were found. Keep the variable labels "
                "from the downloaded template unchanged."
            )
        if not populated_years:
            raise ValueError(
                "No data was found in the profiler input rows. Enter at least "
                "one value before uploading."
            )
        start_year = min(populated_years)
        end_year = max(populated_years)
        years = list(range(start_year, end_year + 1))
        return imported, years
    finally:
        workbook.close()


def _build_profiler_template():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Profiler"

    years = range(MIN_YEAR, MAX_YEAR + 1)
    worksheet.append(["Year", "Unit"] + list(years))
    header_fill = PatternFill("solid", fgColor="E2E8F0")
    for row in worksheet.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.font = Font(bold=True, color="475462")
            cell.fill = header_fill

    for row in build_rows():
        if row["type"] != "input":
            continue
        worksheet.append(
            [row["display_label"], row["unit"]] + [None] * len(years)
        )

    worksheet.freeze_panes = "C2"
    worksheet.column_dimensions["A"].width = 38
    worksheet.column_dimensions["B"].width = 18
    for column in range(3, 3 + len(years)):
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=column).column_letter
        ].width = 12
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


@profiler_blueprint.route("/profiler", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        project_name = None
        start_year = request.form.get("start_year", MIN_YEAR)
        end_year = request.form.get("end_year", MAX_YEAR)
        raw_profile_type = request.form.get("profile_type", "complete")
    else:
        project_name = request.args.get("project_name")
        start_year = request.args.get("start_year", MIN_YEAR)
        end_year = request.args.get("end_year", MAX_YEAR)
        raw_profile_type = request.args.get("profile_type", "complete")

    try:
        profile_type = get_profile_type(raw_profile_type)
    except ValueError as error:
        return str(error), 400
    years = get_years_for_range(start_year, end_year)
    return _render_profiler(years, profile_type, project_name=project_name)


@profiler_blueprint.route("/profiler/import", methods=["POST"])
def import_profiler():
    uploaded_file = request.files.get("workbook")
    if uploaded_file is None or not uploaded_file.filename:
        return _render_setup_with_import_error("Choose an .xlsx profiler workbook to upload.")
    if not uploaded_file.filename.lower().endswith(".xlsx"):
        return _render_setup_with_import_error("Only .xlsx profiler workbooks are supported.")

    try:
        imported, years = _read_profiler_template(uploaded_file.stream)
    except ValueError as error:
        return _render_setup_with_import_error(str(error))

    return _render_profiler(years, "complete", imported=imported)


def _render_setup_with_import_error(error, project_name=""):
    return render_template(
        "setup.html",
        min_year=MIN_YEAR,
        max_year=MAX_YEAR,
        page_title="Main Page",
        import_error=error,
        project_name=project_name,
    )


@profiler_blueprint.route("/profiler/template.xlsx")
def download_profiler_template():
    return send_file(
        _build_profiler_template(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="iCON_Profiler_Template.xlsx",
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
    path = get_profile_workbook_path(profile_type)
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
    whole current grid to its project-named database workbook and returns calculations
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
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
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
    except OSError as error:
        return jsonify({"error": f"Could not save the profiler workbook: {error}"}), 500
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
