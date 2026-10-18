FRESMART FRESH PRICE PORTAL

Render Build Command:
pip install -r requirements.txt

Render Start Command:
gunicorn app:app --bind 0.0.0.0:$PORT --timeout 120

Environment variable for printer bridge:
PRINT_BRIDGE_KEY = set-your-secret-key

Manager PIN: 2026 (change MANAGER_PIN inside app.py if required)

Files required on GitHub: app.py and requirements.txt
store_prices.db is created automatically.
