import openai
import json
from dotenv import load_dotenv
load_dotenv()

api_key = os.environ.get("OPENAI_API_KEY")
if not api_key:
    raise RuntimeError("OPENAI_API_KEY not found in environment variables!")
client = openai.OpenAI(api_key=api_key)

# Load PDF text
WORD_JSON_FILE = "word_red_data.json"
PDF_TEXT_FILE = "pdf_all_text_full.txt"
OUTPUT_FILE = "updated_word_data1.json"

# --- Load files ---
with open(WORD_JSON_FILE, "r", encoding="utf-8") as f:
    word_json = f.read()
with open(PDF_TEXT_FILE, "r", encoding="utf-8") as f:
    pdf_txt = f.read()

# --- Build prompt ---
user_prompt = f"""
Here is a JSON template. It contains only the fields that need updating:
{word_json}

Here is the extracted text from a PDF:
{pdf_txt}

Instructions:
- ONLY update the fields present in the JSON template, using information from the PDF text.
- DO NOT add any extra fields, and do not change the JSON structure.
- Output ONLY the updated JSON, as raw JSON (no markdown, no extra text, no greetings).
- Make sure the JSON is valid and ready to use.
"""

# --- Call OpenAI API (no env var needed) ---
client = openai.OpenAI(api_key=OPENAI_API_KEY)
response = client.chat.completions.create(
    model="gpt-4o",
    messages=[
        {"role": "system", "content": "You are a data extraction assistant. Only reply with valid JSON. Do not add any extra text or formatting. Do NOT use markdown/code blocks, just output JSON."},
        {"role": "user", "content": user_prompt}
    ],
    max_tokens=4096,
    temperature=0
)

updated_json_str = response.choices[0].message.content.strip()

# --- Try to parse as JSON ---
try:
    parsed = json.loads(updated_json_str)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(parsed, f, indent=2, ensure_ascii=False)
    print("✅ JSON updated and saved to", OUTPUT_FILE)
except Exception as e:
    print("⚠️ Model did not return valid JSON. Raw output below:\n")
    print(updated_json_str)
    print("\n❌ Failed to parse updated JSON:", e)