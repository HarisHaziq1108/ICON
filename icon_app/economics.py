"""Economics page routes."""

import math
from flask import Blueprint, jsonify, render_template, request
from .services import (
    DEFAULT_CUSTOM_ESCALATION_RATE,
    DEFAULT_CUSTOM_REFERENCE_YEAR,
    DEFAULT_WACC,
    MAX_YEAR,
    MIN_YEAR,
    build_dashboard_data,
    get_profile_type,
    get_years_for_range,
    load_price_scenarios,
)

economics_blueprint = Blueprint("economics", __name__)


@economics_blueprint.route("/dashboard")
def dashboard():
    years = get_years_for_range(request.args.get("start_year", MIN_YEAR),
                                request.args.get("end_year", MAX_YEAR))
    try:
        profile_type = get_profile_type(request.args.get("profile_type", "complete"))
        price_data = load_price_scenarios()
        default_mrp = price_data["default_mrp"]
        default_custom_liquid_price, default_custom_gas_price = (
            price_data["scenarios"]["base"][DEFAULT_CUSTOM_REFERENCE_YEAR]
        )
    except (OSError, ValueError, KeyError, TypeError) as error:
        return str(error), 500
    return render_template(
        "dashboard.html",
        years=years,
        start_year=years[0],
        end_year=years[-1],
        profile_type=profile_type,
        default_mrp=default_mrp,
        default_custom_liquid_price=default_custom_liquid_price,
        default_custom_gas_price=default_custom_gas_price,
        default_custom_reference_year=DEFAULT_CUSTOM_REFERENCE_YEAR,
        default_custom_escalation_rate=DEFAULT_CUSTOM_ESCALATION_RATE * 100,
        page_title=f"iCON Economics Dashboard ({years[0]} - {years[-1]})",
    )

@economics_blueprint.route("/dashboard-data")
def dashboard_data():
    years = get_years_for_range(request.args.get("start_year", MIN_YEAR),
                                request.args.get("end_year", MAX_YEAR))
    try:
        wacc_percent = float(request.args.get("wacc", DEFAULT_WACC * 100))
    except (TypeError, ValueError):
        return jsonify({"error": "WACC must be a number between 0 and 100."}), 400
    if not math.isfinite(wacc_percent) or not 0 <= wacc_percent <= 100:
        return jsonify({"error": "WACC must be a number between 0 and 100."}), 400
    try:
        source = load_price_scenarios()
    except (OSError, ValueError, KeyError, TypeError) as error:
        return jsonify({"error": str(error)}), 500
    try:
        profile_type = get_profile_type(request.args.get("profile_type", "complete"))
        price_scenario = str(request.args.get("price_scenario", "base")).lower()
        cash_flow_mode = str(request.args.get("cash_flow_mode", "workbook")).lower()
        mrp = float(request.args.get("mrp", source["default_mrp"]))
        custom_reference_year = int(request.args.get(
            "custom_reference_year", DEFAULT_CUSTOM_REFERENCE_YEAR
        ))
        if not MIN_YEAR <= custom_reference_year <= MAX_YEAR:
            return jsonify({
                "error": f"Custom reference year must be between {MIN_YEAR} and {MAX_YEAR}."
            }), 400
        default_custom_liquid_price, default_custom_gas_price = source[
            "scenarios"
        ]["base"][custom_reference_year]
        custom_liquid_price = float(request.args.get(
            "custom_liquid_price", default_custom_liquid_price
        ))
        custom_gas_price = float(request.args.get(
            "custom_gas_price", default_custom_gas_price
        ))
        custom_escalation_rate = (
            float(request.args.get("custom_escalation_rate", 2)) / 100
        )
        sensitivities = {}
        for name in ("production", "capex", "opex", "abex"):
            percent = float(request.args.get(f"{name}_sensitivity", 100))
            if not math.isfinite(percent) or not 0 <= percent <= 1000:
                return jsonify({
                    "error": f"{name.title()} sensitivity must be between 0 and 1000%."
                }), 400
            sensitivities[name] = percent / 100
        if not math.isfinite(mrp) or not 0 <= mrp <= 100:
            return jsonify({"error": "MRP must be between 0 and 100."}), 400
        if not math.isfinite(custom_liquid_price) or custom_liquid_price < 0:
            return jsonify({"error": "Custom liquid price must be zero or greater."}), 400
        if not math.isfinite(custom_gas_price) or custom_gas_price < 0:
            return jsonify({"error": "Custom gas price must be zero or greater."}), 400
        if not 0 <= custom_escalation_rate <= 1:
            return jsonify({"error": "Escalation rate must be between 0 and 100%."}), 400
        data = build_dashboard_data(
            years,
            wacc=wacc_percent / 100,
            profile_type=profile_type,
            price_scenario=price_scenario,
            mrp=mrp,
            custom_liquid_price=custom_liquid_price,
            custom_gas_price=custom_gas_price,
            custom_reference_year=custom_reference_year,
            custom_escalation_rate=custom_escalation_rate,
            sensitivities=sensitivities,
            cash_flow_mode=cash_flow_mode,
        )
    except (TypeError, ValueError) as error:
        return jsonify({"error": str(error)}), 400
    except (OSError, FileNotFoundError) as error:
        return jsonify({"error": str(error)}), 500
    return jsonify(data)
