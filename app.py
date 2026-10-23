"""Run the iCON Profiler and Economics Flask application."""

from icon_app import create_app

app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
