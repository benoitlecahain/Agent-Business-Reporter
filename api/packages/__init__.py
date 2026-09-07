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


DETAIL_WORKERS = 4
USAGE_DETAIL_ATTEMPTS = 3


def was_used(package: dict) -> bool:
    return bool(
        package.get("lastUsedDateTime")
        or (package.get("activeUsers") or 0) > 0
        or (package.get("totalSessions") or 0) > 0
        or (package.get("totalRunTimeInHours") or 0) > 0
    )


def enrich_package(package: dict, token: str) -> tuple[dict, dict | None]:
    package_id = package.get("id")
    if not package_id:
        return package, {"id": None, "status": 400}

    enriched = package
    for attempt in range(USAGE_DETAIL_ATTEMPTS):
        detail_response = graph_get(package_url(package_id), token)
        if not detail_response.ok:
            return package, {"id": package_id, "status": detail_response.status_code}

        try:
            detail = detail_response.json()
        except ValueError:
            return package, {"id": package_id, "status": 502}
        if not isinstance(detail, dict):
            return package, {"id": package_id, "status": 502}

        enriched = {**package, **detail}
        if not was_used(enriched) or enriched.get("activeUsers") is not None:
            break

    return enriched, None


def enrich_packages(packages: list[dict], token: str) -> tuple[list[dict], list[dict]]:
    if not packages:
        return packages, []

    worker_count = min(DETAIL_WORKERS, len(packages))
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        results = list(executor.map(lambda package: enrich_package(package, token), packages))

    enriched = [package for package, _ in results]
    failures = [failure for _, failure in results if failure]
    return enriched, failures


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
            enriched_packages, failures = enrich_packages(packages, token)
            if failures:
                return response({
                    "error": {
                        "message": "The import was not saved because some package details could not be retrieved.",
                        "failedPackages": failures,
                    }
                }, 502)
            used_without_active_users = sum(
                1 for package in enriched_packages
                if was_used(package) and package.get("activeUsers") is None
            )
            return response({
                "value": enriched_packages,
                "count": len(enriched_packages),
                "importSummary": {
                    "detailsRetrieved": len(enriched_packages),
                    "usedAgentsMissingActiveUsers": used_without_active_users,
                },
            })
        if not next_url.startswith(GRAPH_PACKAGES_URL):
            return response({"error": {"message": "Microsoft Graph returned an invalid continuation URL."}}, 502)

    return response({"error": {"message": "Microsoft Graph pagination exceeded the safety limit."}}, 502)