"""
Gunicorn entrypoint.

Usage:
    gunicorn wsgi:app
"""
from app import app

# Expose `app` for Gunicorn
