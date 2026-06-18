from dotenv import load_dotenv
load_dotenv()
import os
from google import genai

client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
model = os.getenv("GEMINI_MODEL", "models/gemini-2.5-flash")

r = client.models.generate_content(model=model, contents="안녕하세요! 한 문장으로 답해주세요.")
print("✓ API 연결 성공!")
print(f"  모델: {model}")
print(f"  응답: {r.text.strip()}")
