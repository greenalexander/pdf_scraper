import os
import json
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv() # Loads OPENAI_API_KEY from your .env file
client = OpenAI()

def extract_indicators(md_file_path):
    with open(md_file_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Load your prompt from the prompts folder
    with open("prompts/system_prompt.txt", "r") as f:
        system_instructions = f.read()

    response = client.chat.completions.create(
        model="gpt-4o", # Use gpt-4o for complex reports
        messages=[
            {"role": "system", "content": system_instructions},
            {"role": "user", "content": f"Extract indicators from this report: \n\n {content}"}
        ],
        response_format={"type": "json_object"}
    )
    
    return json.loads(response.choices[0].message.content)

# Logic to loop through data/markdown/ and save to data/output/
if __name__ == "__main__":
    md_dir = "data/markdown"
    out_dir = "data/output"
    os.makedirs(out_dir, exist_ok=True)

    for md_file in os.listdir(md_dir):
        data = extract_indicators(os.path.join(md_dir, md_file))
        with open(os.path.join(out_dir, md_file.replace(".md", ".json")), "w") as f:
            json.dump(data, f)

"""
import os
import json
import ollama  # This replaces the openai library

def extract_indicators(md_file_path):
    with open(md_file_path, "r", encoding="utf-8") as f:
        content = f.read()

    with open("prompts/system_prompt.txt", "r") as f:
        system_instructions = f.read()

    # We use llama3 or mistral here - both are great at Italian
    response = ollama.chat(
        model='llama3',
        format='json', # Ollama also supports forced JSON output
        messages=[
            {'role': 'system', 'content': system_instructions},
            {'role': 'user', 'content': f"Extract from this report: {content}"},
        ]
    )
    
    return json.loads(response['message']['content'])

if __name__ == "__main__":
    md_dir = "data/markdown"
    out_dir = "data/output"
    os.makedirs(out_dir, exist_ok=True)

    for md_file in os.listdir(md_dir):
        if md_file.endswith(".md"):
            print(f"Extracting indicators from {md_file}...")
            data = extract_indicators(os.path.join(md_dir, md_file))
            with open(os.path.join(out_dir, md_file.replace(".md", ".json")), "w") as f:
                json.dump(data, f)
"""