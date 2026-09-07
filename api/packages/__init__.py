import azure.functions as func

from shared_code.graph import (
    AGENTS_FILTER,
    GRAPH_PACKAGES_URL,
    MAX_PAGES,
    access_token,
    graph_error,
    graph_get,
    response,
)


def main(request: func.HttpRequest) -> func.HttpResponse:
    token = access_token(request)
    if not token:
        return response({"error": {"message": "A Microsoft Graph bearer token is required."}}, 401)

    packages: list[dict] = []
    next_url: str | None = GRAPH_PACKAGES_URL
    params: dict[str, str] | None = {"$filter": AGENTS_FILTER}

    for _ in range(MAX_PAGES):
        graph_response = graph_get(next_url, token, params)
        if not graph_response.ok:
            return graph_error(graph_response)

        payload = graph_response.json()
        packages.extend(payload.get("value", []))
        next_url = payload.get("@odata.nextLink")
        params = None
        if not next_url:
            return response({"value": packages, "count": len(packages)})
        if not next_url.startswith(GRAPH_PACKAGES_URL):
            return response({"error": {"message": "Microsoft Graph returned an invalid continuation URL."}}, 502)

    return response({"error": {"message": "Microsoft Graph pagination exceeded the safety limit."}}, 502)