"""import subprocess

def ai(prompt: str, model: str = "llama3") -> str:
    result = subprocess.run(
        ["ollama", "run", model, prompt],
        capture_output=True,
        text=True
    )
    return result.stdout.strip()"""

from groq import Groq

client = Groq(api_key=GROQ_API_KEY)

def ai(prompt: str, model: str = "llama-3.1-8b-instant") -> str:
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content.strip()

