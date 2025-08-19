# excelDocify
Excel to Document conversion with the help of openAI

Setup Guide: Streamlit Excel + AI Insights App
1. Install Required Tools
1. Install Python (3.9+) → https://www.python.org/downloads/
2. Install VS Code → https://code.visualstudio.com/
3. (Optional) Install Cursor AI IDE → https://cursor.com/
4. Install Git (optional but recommended) → https://git-scm.com/

2. Create Project Folder
Open a terminal (VS Code → View → Terminal or Windows CMD/PowerShell) and run:
mkdir excel_ai_app && cd excel_ai_app

3. Create Virtual Environment
python -m venv venv
venv\Scripts\activate   # On Windows
source venv/bin/activate  # On Mac/Linux

4. Install Python Libraries
pip install streamlit pandas openpyxl requests python-dotenv

5. Get API Key
Option A: OpenAI
- Sign up at https://platform.openai.com/
- Get API key from 'View API Keys'
- Note: May require adding billing; free trial credits may be unavailable in some regions.

Option B: OpenRouter (recommended if avoiding billing)
- Sign up at https://openrouter.ai/
- Get API key from https://openrouter.ai/keys
- Choose models at https://openrouter.ai/models

6. Save API Key in .env File
Create a file named `.env` in your project folder:
For OpenRouter:
OPENROUTER_API_KEY=sk-or-xxxxxxxxxxxxxxxx

For OpenAI:
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxx

7. Create Streamlit App (app.py)
import streamlit as st
import pandas as pd
import requests
from dotenv import load_dotenv
import os

# Load API key
load_dotenv()
API_KEY = os.getenv("OPENROUTER_API_KEY")

BASE_URL = "https://openrouter.ai/api/v1/chat/completions"

st.title("📊 Excel to AI Insights (OpenRouter)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.subheader("Preview of Excel Data")
    st.dataframe(df)

    data_str = df.to_csv(index=False)

    if st.button("Ask AI for Insights"):
        with st.spinner("Thinking..."):
            headers = {
                "Authorization": f"Bearer {API_KEY}",
                "HTTP-Referer": "http://localhost",
                "X-Title": "Excel AI Insights"
            }
            payload = {
                "model": "mistralai/mixtral-8x7b-instruct",
                "messages": [
                    {"role": "system", "content": "You are an expert data analyst."},
                    {"role": "user", "content": f"Analyze the following Excel data and provide insights:\n\n{data_str}"}
                ]
            }
            response = requests.post(BASE_URL, headers=headers, json=payload)
            if response.status_code == 200:
                result = response.json()
                answer = result["choices"][0]["message"]["content"]
                st.subheader("AI Response")
                st.write(answer)
            else:
                st.error(f"Error {response.status_code}: {response.text}")

8. Run the App
venv\Scripts\activate
streamlit run app.py

9. Switching Between OpenRouter and OpenAI
- Change API key in `.env`
- Change BASE_URL and headers in `app.py` to match service
- For OpenAI, you would use their official client instead of requests

