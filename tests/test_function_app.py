import json
import sys
from pathlib import Path
from unittest.mock import Mock, patch

import azure.functions as func

sys.path.insert(0, str(Path(__file__).parents[1] / "api"))

from package_detail import main as get_package
from packages import main as list_packages
from shared_code.graph import AGENTS_FILTER, GRAPH_PACKAGES_URL


def make_request(url: str, route_params: dict[str, str] | None = None) -> func.HttpRequest:
    return func.HttpRequest(
        method="GET",
        url=url,
        headers={"X-MS-Graph-Token": "graph-token"},
        params={},
        route_params=route_params or {},
        body=None,
    )


@patch("shared_code.graph.requests.get")
def test_list_packages_filters_agents_and_forwards_token(graph_get: Mock) -> None:
    graph_get.return_value = Mock(
        ok=True,
        json=lambda: {"value": [{"id": "package-1", "displayName": "HR Agent"}]},
    )

    response = list_packages(make_request("https://localhost/api/packages"))

    assert response.status_code == 200
    assert json.loads(response.get_body()) == {
        "value": [{"id": "package-1", "displayName": "HR Agent"}],
        "count": 1,
    }
    graph_get.assert_called_once_with(
        GRAPH_PACKAGES_URL,
        headers={"Authorization": "Bearer graph-token", "Accept": "application/json"},
        params={"$filter": AGENTS_FILTER},
        timeout=30,
    )


@patch("shared_code.graph.requests.get")
def test_get_package_encodes_id(graph_get: Mock) -> None:
    graph_get.return_value = Mock(ok=True, json=lambda: {"id": "package/one"})

    response = get_package(
        make_request(
            "https://localhost/api/packages/package%2Fone",
            {"package_id": "package/one"},
        )
    )

    assert response.status_code == 200
    assert graph_get.call_args.args[0] == f"{GRAPH_PACKAGES_URL}/package%2Fone"


def test_list_packages_requires_bearer_token() -> None:
    request = func.HttpRequest(
        method="GET",
        url="https://localhost/api/packages",
        headers={},
        params={},
        route_params={},
        body=None,
    )

    response = list_packages(request)

    assert response.status_code == 401