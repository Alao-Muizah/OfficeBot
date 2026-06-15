import os

API_KEY = os.environ.get("MISTRAL_API_KEY", "YOUR_API_KEY_HERE")
MODEL = "mistral-small-latest"

DEFAULT_WORKDIR = os.path.join(os.path.expanduser("~"), "officebot_files")

