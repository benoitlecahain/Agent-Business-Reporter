@@
-# Agent-Business-Reporter
-This repo provides source code to the Agent-Business-Reporter apps, hosted on Azure.
+# Agent Business Reporter
+
+Agent Business Reporter is a multitenant Microsoft Entra web application for viewing the Agent 365 catalog in any Microsoft 365 organization. It uses the Microsoft Graph Package Management API to list agents and retrieve package details.
+
+## Architecture
+
+- `app/`: static HTML, CSS, and JavaScript client using MSAL Browser.
+- `api/`: Python Azure Functions proxy for Microsoft Graph.
+- Microsoft Graph delegated permission: `CopilotPackages.Read.All`.
+- Azure Static Web Apps deployment through GitHub Actions.
+
+The browser obtains a delegated Microsoft Graph token for the signed-in tenant. The token is sent only to the same-origin Python API, which forwards it to the two read-only Graph endpoints. Tokens and tenant data are not persisted by the application.
+
+## Entra app registration
+
+Application (client) ID: `5f8259cb-61e1-402e-8d38-f38fe1a2db31`
+
+1. Set **Supported account types** to **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant)**.
+2. Under **Authentication**, add a **Single-page application** redirect URI for the Azure Static Web Apps production URL. For local development, also add `http://localhost:4280`.
+3. Under **API permissions**, add Microsoft Graph delegated permission `CopilotPackages.Read.All` and grant admin consent in the home tenant.
+4. Other tenants consent when an administrator signs in. Each tenant also requires a Microsoft Agent 365 license to use the Package Management API.
+
+No client secret is required or supported by this browser-based delegated flow.
+
+## Local development
+
+Install the Azure Static Web Apps CLI and Azure Functions Core Tools, then run:
+
+```powershell
+python -m venv .venv
+.\.venv\Scripts\Activate.ps1
+pip install -r requirements-dev.txt
+swa start app --api-location api
+```
+
+Open `http://localhost:4280`.
+
+## Tests
+
+```powershell
+python -m pytest -q
+```
+
+## Deployment
+
+Pushes to `main` deploy through `.github/workflows/azure-static-web-apps-orange-sand-058c15d03.yml` to the linked Azure Static Web App. The repository secret `AZURE_STATIC_WEB_APPS_API_TOKEN_ORANGE_SAND_058C15D03` must remain configured.
# Agent-Business-Reporter
This repo provides source code to the Agent-Business-Reporter apps, hosted on Azure.
