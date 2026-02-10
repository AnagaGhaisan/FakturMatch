web: uvicorn app:app --host 0.0.0.0 --port 8000
web: gunicorn -w 4 -b 0.0.0.0:$PORT app:app
