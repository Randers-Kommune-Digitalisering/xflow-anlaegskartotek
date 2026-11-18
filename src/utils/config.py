import os
from dotenv import load_dotenv


# loads .env file, will not overide already set enviroment variables (will do nothing when testing, building and deploying)
load_dotenv()


DEBUG = os.getenv('DEBUG', 'False') in ['True', 'true']
PORT = os.getenv('PORT', '8080')
POD_NAME = os.getenv('POD_NAME', 'pod_name_not_set')
SMTP_SERVER = os.getenv('SMTP_SERVER', '').strip()
SMTP_PORT = os.getenv('SMTP_PORT', '25').strip()
SMTP_SENDER = os.getenv('SMTP_SENDER', '').strip()
RECIPIENTS = os.getenv('RECIPIENTS', '').strip()
