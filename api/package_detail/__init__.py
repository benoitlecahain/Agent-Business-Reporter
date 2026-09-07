import azure.functions as func

from shared_code.graph import access_token, graph_error, graph_get, package_url, response


def main(request: func.HttpRequest) -> func.HttpResponse:
    token = access_token(request)
    if not token:
        return response({"error": {"message": "A Microsoft Graph bearer token is required."}}, 401)

    package_id = request.route_params.get("package_id", "").strip()
    if not package_id:
        return response({"error": {"message": "A package ID is required."}}, 400)

    graph_response = graph_get(package_url(package_id), token)
    if not graph_response.ok:
        return graph_error(graph_response)
    return response(graph_response.json())