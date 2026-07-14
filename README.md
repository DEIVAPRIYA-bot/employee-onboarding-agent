# KBC Onboarding Assistant - Starter implementation

This repository contains a starter implementation of the KBC Onboarding Assistant: a simple Python-based onboarding agent that uses SharePoint (via Microsoft Graph) as the primary source of truth.

Capabilities included in this starter:
- Answer onboarding questions by retrieving SharePoint pages, lists, and documents
- Provide onboarding status for a specific employee
- Show pending and completed tasks (from SharePoint lists)
- Generate simple summaries (optionally using OpenAI / OAI if configured)
- Provide onboarding metrics and insights (basic counts)

This is intended as a scaffold you can extend for production use.

Quick start
1. Copy `config.example.env` to `.env` and fill in the SharePoint / Microsoft Graph credentials and site identifiers.
2. Create a Python virtual environment and install dependencies:

   python -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt

3. Run the app:

   export FLASK_APP=app.py
   flask run --host=0.0.0.0 --port=8080

Endpoints
- POST /query  - body: {"user": "alice@example.com", "question":"Where's the benefits doc?"}
- GET /status/<employee_email> - returns onboarding status summary

Security & production
- This scaffold uses client credentials for Microsoft Graph. For production, follow secure secret storage, least-privilege app registrations, and proper authentication flows.

Files added
- app.py: Flask HTTP endpoints
- onboarding_agent.py: KBCOnboardingAssistant implementation (business logic)
- sharepoint_client.py: wrapper for Microsoft Graph SharePoint calls
- config.example.env: example environment variables
- requirements.txt
- .gitignore
- README.md (this file)

Notes
- The SharePoint client methods are minimal and meant to be adjusted for your site structure (site IDs, list names, document libraries).
- If OpenAI/Microsoft OAI keys are provided, the agent will use the configured model for improved summaries; otherwise it falls back to simple heuristics.
