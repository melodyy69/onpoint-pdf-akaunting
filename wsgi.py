"""WSGI entry point for alwaysdata, cPanel, PythonAnywhere and Gunicorn."""

from web_app import create_app


application = create_app()
app = application
