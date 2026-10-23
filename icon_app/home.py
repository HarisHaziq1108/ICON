"""Home page routes."""


from flask import Blueprint, render_template
from .services import MAX_YEAR, MIN_YEAR

home_blueprint = Blueprint("home", __name__)


@home_blueprint.route("/")
def range_setup():
    return render_template(
        "setup.html",
        min_year=MIN_YEAR,
        max_year=MAX_YEAR,
        page_title="Main Page",
    )
