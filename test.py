import os
from dotenv import load_dotenv

load_dotenv('.env')
MAIL_USERNAME = os.getenv("MAIL_USERNAME")
print(type(MAIL_USERNAME))