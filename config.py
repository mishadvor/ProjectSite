import os
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

client = OpenAI(
    api_key=os.getenv('AITUNNEL_API_KEY'),
    base_url='https://api.aitunnel.ru/v1/'
)
