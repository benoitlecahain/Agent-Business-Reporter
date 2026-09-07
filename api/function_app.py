import json
from urllib.parse import quote

import azure.functions as func
import requests


app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)
GRAPH_PACKAGES_URL = "https://graph.microsoft.com/v1.0/copilot/admin/catalog/packages"
AGENTS_FILTER = "supportedHosts/any(host:host eq 'Copilot')"
MAX_PAGES = 100


def _response(payload: object, status_code: int = 200) -> func.HttpResponse:
    return func.HttpResponse(
        json.dumps(payload),
        status_code=status_code,
        mimetype="application/json",
    )


def _access_token(request: func.HttpRequest) -> str | None:
    authorization = request.headers.get("Authorization", "")
    scheme, _, token = authorization.partition(" ")
    if scheme.lower() != "bearer" or not token.strip():
        return None
    return token.strip()


def _graph_get(url: str, access_token: str, params: dict[str, str] | None = None) -> requests.Response:
    return requests.get(
        url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        },
        params=params,
        timeout=30,
    )


def _graph_error(response: requests.Response) -> func.HttpResponse:
    try:
        body = response.json()
    except ValueError:
        body = {"error": {"message": "Microsoft Graph returned an unreadable response."}}
    return _response(body, response.status_code)


@app.route(route="packages", methods=["GET"])
def list_packages(request: func.HttpRequest) -> func.HttpResponse:
    access_token = _access_token(request)
    if not access_token:
        return _response({"error": {"message": "A Microsoft Graph bearer token is required."}}, 401)

    packages: list[dict] = []
    next_url: str | None = GRAPH_PACKAGES_URL
    params: dict[str, str] | None = {"$filter": AGENTS_FILTER}

    for _ in range(MAX_PAGES):
        response = _graph_get(next_url, access_token, params)
        if not response.ok:
            return _graph_error(response)

        payload = response.json()
        packages.extend(payload.get("value", []))
        next_url = payload.get("@odata.nextLink")
        params = None
        if not next_url:
            return _response({"value": packages, "count": len(packages)})

        if not next_url.startswith(GRAPH_PACKAGES_URL):
            return _response({"error": {"message": "Microsoft Graph returned an invalid continuation URL."}}, 502)

    return _response({"error": {"message": "Microsoft Graph pagination exceeded the safety limit."}}, 502)


@app.route(route="packages/{package_id}", methods=["GET"])
def get_package(request: func.HttpRequest) -> func.HttpResponse:
    access_token = _access_token(request)
    if not access_token:
        return _response({"error": {"message": "A Microsoft Graph bearer token is required."}}, 401)

    package_id = request.route_params.get("package_id", "").strip()
    if not package_id:
        return _response({"error": {"message": "A package ID is required."}}, 400)

    response = _graph_get(f"{GRAPH_PACKAGES_URL}/{quote(package_id, safe='')}", access_token)
    if not response.ok:
        return _graph_error(response)
    return _response(response.json())