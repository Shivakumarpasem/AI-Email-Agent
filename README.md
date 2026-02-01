# Resume Agent (Streamlit)

A local Streamlit app that:
- Scores resume fit for a job description (as a hiring manager)
- Suggests resume tweaks
- Generates a cold email and a LinkedIn connection note
- Logs outreach to a CSV and exports to Excel

## Setup

1) Create a `.env` file next to `app.py` (do NOT commit it):

GEMINI_API_KEY / OPENAI_API_KEY=your_key_here


2) Install dependencies:

pip install -r requirements.txt

3) Run the app:

streamlit run app.py


Optional: on Windows you can run:

run_app.bat


## Notes
- Do not commit `.env`, `saved_resumes/`, or your outreach logs.
- The app is intended to run locally so file links work.
