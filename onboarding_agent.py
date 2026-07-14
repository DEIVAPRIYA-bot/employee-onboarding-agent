import os
from typing import Optional, Dict, Any, List
from sharepoint_client import SharePointClient
from dotenv import load_dotenv

try:
    load_dotenv()
except Exception:
    pass

class KBCOnboardingAssistant:
    """KBC Onboarding Assistant — uses SharePoint as the source of truth.

    Methods return concise strings or structured dicts suitable for UI responses.
    If source data is missing, methods return clear "not found in SharePoint" messages.
    """

    def __init__(self, sp_client: SharePointClient, summarizer=None):
        self.sp = sp_client
        self.summarizer = summarizer

    @classmethod
    def from_env(cls):
        sp = SharePointClient.from_env()
        # Optional summarizer using OpenAI/OAI if configured
        summarizer = None
        oai_key = os.environ.get('OPENAI_API_KEY') or os.environ.get('OAI_API_KEY')
        if oai_key:
            try:
                from functools import partial
                # lightweight summarizer wrapper — real LLM integration should be added
                summarizer = partial(cls._simple_summarize)
            except Exception:
                summarizer = None
        return cls(sp_client=sp, summarizer=summarizer)

    def answer_question(self, user: Optional[str], question: str) -> str:
        """Answer an onboarding-related question by searching SharePoint pages, lists, and docs.
        Returns a concise answer or a not-found message.
        """
        # 1) Try search on SharePoint pages
        pages = self.sp.search_site_pages(query=question, top=5)
        if pages:
            # return top result's excerpt if available
            page = pages[0]
            title = page.get('title') or page.get('name')
            url = page.get('webUrl') or page.get('link')
            return f"Found SharePoint page: {title} — {url}"

        # 2) Try search in lists (tasks, onboarding checklist, policies)
        lists = self.sp.search_lists(query=question)
        if lists:
            item = lists[0]
            list_name = item.get('list')
            title = item.get('title') or item.get('Name')
            return f"Found list item in '{list_name}': {title}"

        # 3) Try searching documents
        docs = self.sp.search_documents(query=question)
        if docs:
            d = docs[0]
            name = d.get('name')
            link = d.get('webUrl')
            return f"Found document: {name} — {link}"

        return "I couldn't find information matching that question in the configured SharePoint source."

    def get_onboarding_status(self, employee_email: str) -> Dict[str, Any]:
        """Return onboarding status for an employee using a probable 'Onboarding' list.
        """
        checklist = self.sp.get_list_items(list_name='Onboarding', filter_query=f"Email eq '{employee_email}'")
        if not checklist:
            return {"employee": employee_email, "status": "No onboarding record found in SharePoint (list 'Onboarding' not found or empty)."}

        # Assume checklist returns items with fields Task and Completed (boolean)
        tasks = []
        completed = 0
        for it in checklist:
            t = it.get('Task') or it.get('Title') or 'Unnamed task'
            done = it.get('Completed') or it.get('completed') or False
            tasks.append({"task": t, "completed": bool(done)})
            if done:
                completed += 1
        total = len(tasks)
        percent = round((completed / total) * 100) if total else 0
        return {
            "employee": employee_email,
            "completed": completed,
            "total": total,
            "percent_complete": percent,
            "tasks": tasks
        }

    def list_tasks(self, employee_email: str) -> Dict[str, Any]:
        return self.get_onboarding_status(employee_email)

    def generate_summary(self, employee_email: str) -> str:
        """Generate a concise onboarding summary for the employee.
        Uses summarizer if configured, otherwise returns a short structured text.
        """
        status = self.get_onboarding_status(employee_email)
        if isinstance(status.get('status'), str) and 'No onboarding record' in status.get('status'):
            return status.get('status')

        # Build short summary
        s = f"Onboarding for {employee_email}: {status['completed']}/{status['total']} tasks completed ({status['percent_complete']}%)."
        if self.summarizer:
            try:
                return self.summarizer(s)
            except Exception:
                return s
        return s

    def metrics(self) -> Dict[str, Any]:
        """Return basic onboarding metrics across the Onboarding list.
        """
        all_items = self.sp.get_list_items(list_name='Onboarding') or []
        if not all_items:
            return {"message": "No onboarding data found in SharePoint list 'Onboarding'"}
        by_employee = {}
        for it in all_items:
            email = it.get('Email') or it.get('email') or it.get('AssignedTo') or 'unknown'
            done = bool(it.get('Completed') or it.get('completed') or False)
            by_employee.setdefault(email, {'total':0, 'completed':0})
            by_employee[email]['total'] += 1
            if done:
                by_employee[email]['completed'] += 1
        # compute simple averages
        totals = [{'employee':k, 'completed':v['completed'], 'total':v['total']} for k,v in by_employee.items()]
        return {"employees": totals, "count": len(totals)}

    @staticmethod
    def _simple_summarize(text: str) -> str:
        # placeholder summarizer — in real deployment call an LLM
        if len(text) < 200:
            return text
        return text[:200].rsplit('.', 1)[0] + '.'
