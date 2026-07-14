from flask import Flask, request, jsonify
from onboarding_agent import KBCOnboardingAssistant
import os

app = Flask(__name__)

# Initialize assistant with a SharePoint-backed client
assistant = KBCOnboardingAssistant.from_env()

@app.route('/query', methods=['POST'])
def query():
    payload = request.get_json() or {}
    user = payload.get('user')
    question = payload.get('question')
    if not question:
        return jsonify({"error": "question is required"}), 400
    answer = assistant.answer_question(user=user, question=question)
    return jsonify({"answer": answer})

@app.route('/status/<employee_email>', methods=['GET'])
def status(employee_email):
    status = assistant.get_onboarding_status(employee_email)
    return jsonify(status)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
