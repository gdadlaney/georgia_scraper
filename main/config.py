import json
import os

# Path to the service account key file
SERVICE_ACCOUNT_KEY_PATH = os.path.join(os.path.dirname(__file__), 'service-account-for-gsheets-key.json')

with open(SERVICE_ACCOUNT_KEY_PATH, 'r') as f:
    SHEETS_KEY = json.load(f)