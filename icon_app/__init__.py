"""Application factory for the iCON profiler and economics toolkit."""

from flask import Flask


def create_app():
    app = Flask(__name__, template_folder="../templates", static_folder="../static")

    from .economics import economics_blueprint
    from .home import home_blueprint
    from .profiler import profiler_blueprint

    app.register_blueprint(home_blueprint)
    app.register_blueprint(profiler_blueprint)
    app.register_blueprint(economics_blueprint)
    return app
