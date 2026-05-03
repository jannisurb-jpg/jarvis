import requests
from datetime import datetime
import xml.etree.ElementTree as ET
import json

messages = [
            {"role": "system", "content": "Du bist Jarvis, ein KI Assistent. Antworte kurz und präzise in max. 50 Wörtern."}
        ]

def tagesschau_nachrichten():
    url = "https://www.tagesschau.de/xml/rss2/"

    response = requests.get(url)
    response.raise_for_status()

    root = ET.fromstring(response.content)

    nachrichten = []

    # RSS Struktur: channel -> item -> title
    for item in root.findall(".//item"):
        title = item.find("title").text
        nachrichten.append(title)

    return nachrichten


def get_news_command():
    # Beispiel Nutzung
    news = tagesschau_nachrichten()

    articles = news[:10]

    prompt_content = (
        "Du bist ein Nachrichten-Redakteur.\n"
        "Erstelle aus Stichpunkten einen kurzen, neutralen Nachrichtentext auf Deutsch.\n"
        "Schreibe nicht in Stichpunkten, sondern in ganzen Sätzen.\n"
        "Beschränke dich auf die wichtigsten Informationen und fasse sie prägnant zusammen. Nutze maximal 150 Wörter\n"
        "Nur Fakten, keine Meinung.\n\n"
        "Stichpunkte:\n"
    )

    for article in articles:
        print(article)
        print("-" * 50)
        prompt_content += article + "\n"

    print("Prompt Content:")
    print(prompt_content)

    starting_time = datetime.now()
    with requests.post(
        "http://localhost:11434/api/generate",
        json={
            "model": "phi3",
            "prompt": prompt_content,
            "stream": True
        },
        stream=True
    ) as response:

        for line in response.iter_lines():
            if line:
                chunk = json.loads(line.decode("utf-8"))
                yield chunk.get("response", "")
    print(f"Zeit für die Generierung: {datetime.now() - starting_time}")

session = requests.Session()

def createAnswer(cmd):
    messages.append({"role": "user", "content": cmd})

    starting_time = datetime.now()

    response = session.post(
        "http://localhost:11434/api/chat",
        json={
            "model": "phi3:mini",
            "messages": messages,
            "stream": False,
            "keep_alive": "5m",
            "options": {
                "temperature": 0.2
            }
        }
    )

    assistant_msg = response.json()["message"]
    messages.append(assistant_msg)

    print(f"Zeit: {datetime.now() - starting_time}")
    return assistant_msg["content"]