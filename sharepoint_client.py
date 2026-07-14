import os
import requests
from typing import Optional, List, Dict, Any
from dotenv import load_dotenv

try:
    load_dotenv()
except Exception:
    pass

GRAPH_TOKEN_URL = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'
GRAPH_BASE = 'https://graph.microsoft.com/v1.0'

class SharePointClient:
    """Minimal Microsoft Graph-based SharePoint client using client credentials flow.

    Requires MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, and MS_SITE_ID or MS_SITE_HOSTNAME.
    """

    def __init__(self, tenant_id: str, client_id: str, client_secret: str, site_id: Optional[str]=None, site_hostname: Optional[str]=None):
        self.tenant = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_id = site_id
        self.site_hostname = site_hostname
        self._token = None

    @classmethod
    def from_env(cls):
        tenant = os.environ.get('MS_TENANT_ID')
        client_id = os.environ.get('MS_CLIENT_ID')
        client_secret = os.environ.get('MS_CLIENT_SECRET')
        site_id = os.environ.get('MS_SITE_ID')
        site_hostname = os.environ.get('MS_SITE_HOSTNAME')
        return cls(tenant, client_id, client_secret, site_id, site_hostname)

    def _get_token(self) -> str:
        if self._token:
            return self._token
        url = GRAPH_TOKEN_URL.format(tenant=self.tenant)
        data = {
            'client_id': self.client_id,
            'scope': 'https://graph.microsoft.com/.default',
            'client_secret': self.client_secret,
            'grant_type': 'client_credentials'
        }
        r = requests.post(url, data=data)
        r.raise_for_status()
        self._token = r.json().get('access_token')
        return self._token

    def _get_headers(self) -> Dict[str, str]:
        return {'Authorization': f'Bearer {self._get_token()}'}

    def _site_path(self) -> str:
        if self.site_id:
            return f"/sites/{self.site_id}"
        if self.site_hostname:
            return f"/sites/{self.site_hostname}"
        raise RuntimeError('No MS_SITE_ID or MS_SITE_HOSTNAME configured')

    def search_site_pages(self, query: str, top: int=5) -> List[Dict[str, Any]]:
        # This uses the SharePoint search API via Graph search endpoint
        try:
            url = f"{GRAPH_BASE}{self._site_path()}/pages?search={requests.utils.quote(query)}&$top={top}"
            headers = self._get_headers()
            r = requests.get(url, headers=headers)
            if r.status_code == 200:
                return r.json().get('value', [])
            return []
        except Exception:
            return []

    def search_documents(self, query: str, top: int=5) -> List[Dict[str, Any]]:
        try:
            url = f"{GRAPH_BASE}/sites/{self.site_id}/drive/root/search(q='{requests.utils.quote(query)}')"
            headers = self._get_headers()
            r = requests.get(url, headers=headers)
            if r.status_code == 200:
                return r.json().get('value', [])
            return []
        except Exception:
            return []

    def search_lists(self, query: str, top: int=5) -> List[Dict[str, Any]]:
        # Very minimal; attempts to query known lists named 'Onboarding' or 'Tasks'
        results = []
        for list_name in ['Onboarding', 'Tasks', 'Checklist']:
            items = self.get_list_items(list_name=list_name, top=top)
            for it in items or []:
                # naive search in item fields
                if any(query.lower() in str(v).lower() for v in it.values() if v):
                    it_copy = it.copy()
                    it_copy['list'] = list_name
                    results.append(it_copy)
        return results

    def get_list_items(self, list_name: str, filter_query: Optional[str]=None, top: int=200) -> Optional[List[Dict[str, Any]]]:
        try:
            site = self._site_path()
            url = f"{GRAPH_BASE}{site}/lists/{requests.utils.quote(list_name)}/items?expand=fields&$top={top}"
            if filter_query:
                url += f"&$filter={requests.utils.quote(filter_query)}"
            headers = self._get_headers()
            r = requests.get(url, headers=headers)
            if r.status_code == 200:
                values = r.json().get('value', [])
                return [v.get('fields', {}) for v in values]
            return []
        except Exception:
            return None
