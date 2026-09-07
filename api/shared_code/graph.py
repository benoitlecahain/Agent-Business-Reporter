import json
import time
from urllib.parse import quote

import azure.functions as func
import requests


GRAPH_PACKAGES_URL = "https://graph.microsoft.com/v1.0/copilot/admin/catalog/packages"
AGENTS_FILTER = "supportedHosts/any(host:host eq 'Copilot')"
MAX_PAGES = 100
GRAPH_ATTEMPTS = 3
RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}


def response(payload: object, status_code: int = 200) -> func.HttpResponse:
    return func.HttpResponse(json.dumps(payload), status_code=status_code, mimetype="application/json")


def access_token(request: func.HttpRequest) -> str | None:
    token = request.headers.get("X-Graph-Access-Token", "")
    return token.strip() or None


def graph_get(url: str, token: str, params: dict[str, str] | None = None) -> requests.Response:
    for attempt in range(GRAPH_ATTEMPTS):
        try:
            graph_response = requests.get(
                url,
                headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
                params=params,
                timeout=30,
            )
        except requests.RequestException:
            if attempt == GRAPH_ATTEMPTS - 1:
                graph_response = requests.Response()
                graph_response.status_code = 503
                graph_response._content = b'{"error":{"message":"Microsoft Graph could not be reached."}}'
                return graph_response
            time.sleep(2 ** attempt)
            continue
        if graph_response.status_code not in RETRYABLE_STATUS_CODES or attempt == GRAPH_ATTEMPTS - 1:
            return graph_response
        retry_after = graph_response.headers.get("Retry-After", "")
        delay = int(retry_after) if retry_after.isdigit() else 2 ** attempt
        time.sleep(min(delay, 10))

    return graph_response


def graph_error(graph_response: requests.Response) -> func.HttpResponse:
    try:
        body = graph_response.json()
    except ValueError:
        body = {"error": {"message": "Microsoft Graph returned an unreadable response."}}
    return response(body, graph_response.status_code)


def package_url(package_id: str) -> str:
    return f"{GRAPH_PACKAGES_URL}/{quote(package_id, safe='')}"