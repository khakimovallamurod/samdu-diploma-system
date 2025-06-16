from dotenv import load_dotenv
import os

load_dotenv()
def get_token():
    TOKEN = os.getenv("token")
    if TOKEN is None:
        raise ValueError("TOKEN not found.")
    return TOKEN

def get_url():
    URL = os.getenv("url")
    if URL is None:
        raise ValueError("URL not found.")
    return URL

