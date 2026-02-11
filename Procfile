web: pip install --upgrade setuptools && pip install -r requirements.txt && gunicorn app:app
web: gunicorn --bind 0.0.0.0:$PORT app:app
