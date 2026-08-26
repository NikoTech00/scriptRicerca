import os

from dotenv import load_dotenv
from google import genai


load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")

client = genai.Client(api_key=api_key)

try:
    interaction = client.interactions.create(
        model="gemini-3.6-flash",
        input="Chi è l'attuale presidente della Repubblica Italiana?",
        tools=[
            {
                "type": "google_search"
            }
        ],
    )

    print("SUCCESSO")
    print(interaction.output_text)

except Exception as exc:
    print("ERRORE")
    print(type(exc).__name__)
    print(exc)