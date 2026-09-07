from concurrent.futures import ThreadPoolExecutor

import azure.functions as func

from shared_code.graph import (
    AGENTS_FILTER,
    GRAPH_PACKAGES_URL,
    MAX_PAGES,
    access_token,
    graph_error,
    graph_get,
    package_url,
    response,
)


DETAIL_WORKERS = 8


def enrich_package(package: dict, token: str) -> dict:
    package_id = package.get("id")
    if not package_id:
        return package

    detail_response = graph_get(package_url(package_id), token)
    if not detail_response.ok:
        return package

    detail = detail_response.json()
    return detail if isinstance(detail, dict) else package


def enrich_packages(packages: list[dict], token: str) -> list[dict]:
    if not packages:
        return packages

    worker_count = min(DETAIL_WORKERS, len(packages))
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        return list(executor.map(lambda package: enrich_package(package, token), packages))


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
            enriched_packages = enrich_packages(packages, token)
            return response({"value": enriched_packages, "count": len(enriched_packages)})
        if not next_url.startswith(GRAPH_PACKAGES_URL):
            return response({"error": {"message": "Microsoft Graph returned an invalid continuation URL."}}, 502)

    return response({"error": {"message": "Microsoft Graph pagination exceeded the safety limit."}}, 502)