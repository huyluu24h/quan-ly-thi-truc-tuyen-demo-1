"""WSGI entry point used by Gunicorn in production."""

from app import app, init_db


init_db()

