# Naver Keyword Analyzer API

Flask API for manual Naver Blog URL analysis. It collects the selected public post URLs with Selenium and extracts Korean nouns with Kiwi.

## Render settings

- Runtime: Python
- Build command: `pip install -r requirements.txt && python -m playwright install chromium`
- Start command: `gunicorn --workers 1 --timeout 120 --bind 0.0.0.0:$PORT app:app`
