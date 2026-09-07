import json
from urllib.parse import quote

import azure.functions as func
import requests


GRAPH_PACKAGES_URL = "https://graph.microsoft.com/v1.0/copilot/admin/catalog/packages"
AGENTS_FILTER = "supportedHosts/any(host:host eq 'Copilot')"
MAX_PAGES = 100


def response(payload: object, status_code: int = 200) -> func.HttpResponse:
    return func.HttpResponse(json.dumps(payload), status_code=status_code, mimetype="application/json")


def access_token(request: func.HttpRequest) -> str | None:
    token = request.headers.get("X-Graph-Access-Token", "")
    return token.strip() or None


def graph_get(url: str, token: str, params: dict[str, str] | None = None) -> requests.Response:
    return requests.get(
        url,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
        params=params,
        timeout=30,
    )


def graph_error(graph_response: requests.Response) -> func.HttpResponse:
    try:
        body = graph_response.json()
    except ValueError:
        body = {"error": {"message": "Microsoft Graph returned an unreadable response."}}
    return response(body, graph_response.status_code)


def package_url(package_id: str) -> str:
    return f"{GRAPH_PACKAGES_URL}/{quote(package_id, safe='')}"