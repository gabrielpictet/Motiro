import os

api_key = os.getenv('OPENAI_API_KEY')

if api_key:
    print(f"API key found: {api_key}")
else:
    print("API key not found")
