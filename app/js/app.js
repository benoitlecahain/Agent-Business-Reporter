const CLIENT_ID = "5f8259cb-61e1-402e-8d38-f38fe1a2db31";
const GRAPH_SCOPES = ["https://graph.microsoft.com/CopilotPackages.Read.All"];
const elements = Object.fromEntries([
  "signedOutView", "dashboardView", "signInButton", "signOutButton", "accountBlock", "accountName", "tenantName", "refreshButton", "retryButton", "searchInput", "statusFilter", "loadingState", "errorState", "errorMessage", "emptyState", "tableWrap", "agentRows", "resultCount", "totalMetric", "availableMetric", "blockedMetric", "publisherMetric", "drawerBackdrop", "drawerTitle", "drawerContent", "closeDrawerButton", "themeButton", "toast"
].map((id) => [id, document.querySelector(`#${id}`)]));
elements.drawer = document.querySelector("#detailDrawer");
let msalClient;
let account;
let packages = [];
let lastFocusedElement;

function createElement(tag, className, text) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (text !== undefined && text !== null) element.textContent = String(text);
  return element;
}
function icon(name) { const element = createElement("i"); element.dataset.lucide = name; element.setAttribute("aria-hidden", "true"); return element; }
function refreshIcons() { if (window.lucide) window.lucide.createIcons(); }
function showToast(message) { elements.toast.textContent = message; elements.toast.hidden = false; window.setTimeout(() => { elements.toast.hidden = true; }, 3500); }
function setView(isSignedIn) {
  elements.signedOutView.hidden = isSignedIn;
  elements.dashboardView.hidden = !isSignedIn;
  elements.signOutButton.hidden = !isSignedIn;
  elements.accountBlock.hidden = !isSignedIn;
  if (isSignedIn) { elements.accountName.textContent = account.name || account.username; elements.tenantName.textContent = account.tenantId; }
}
async function getAccessToken(interactive = true) {
  const request = { account, scopes: GRAPH_SCOPES };
  try { return (await msalClient.acquireTokenSilent(request)).accessToken; }
  catch (error) { if (!interactive) throw error; return (await msalClient.acquireTokenPopup(request)).accessToken; }
}
async function apiRequest(path) {
  const token = await getAccessToken();
  const response = await fetch(path, { headers: { "X-Graph-Access-Token": token } });
  let payload;
  try { payload = await response.json(); } catch { payload = {}; }
  if (!response.ok) { const error = new Error(payload.error?.message || `Request failed with status ${response.status}.`); error.status = response.status; throw error; }
  return payload;
}
function showLoading() { elements.loadingState.hidden = false; elements.errorState.hidden = true; elements.emptyState.hidden = true; elements.tableWrap.hidden = true; elements.refreshButton.disabled = true; }
function showError(error) {
  elements.loadingState.hidden = true; elements.tableWrap.hidden = true; elements.emptyState.hidden = true; elements.errorState.hidden = false; elements.refreshButton.disabled = false;
  elements.errorMessage.textContent = error.status === 403 ? "Your tenant may need admin consent for CopilotPackages.Read.All or an Agent 365 license." : error.message;
  refreshIcons();
}
function updateMetrics() {
  const blocked = packages.filter((item) => item.isBlocked).length;
  elements.totalMetric.textContent = packages.length;
  elements.availableMetric.textContent = packages.length - blocked;
  elements.blockedMetric.textContent = blocked;
  elements.publisherMetric.textContent = new Set(packages.map((item) => item.publisher).filter(Boolean)).size;
}
function filteredPackages() {
  const query = elements.searchInput.value.trim().toLocaleLowerCase();
  const status = elements.statusFilter.value;
  return packages.filter((item) => {
    const searchable = [item.displayName, item.publisher, item.shortDescription, item.platform].filter(Boolean).join(" ").toLocaleLowerCase();
    return (!query || searchable.includes(query)) && (status === "all" || (status === "blocked" ? item.isBlocked : !item.isBlocked));
  });
}
function appendChips(cell, values) {
  const wrapper = createElement("div", "chips");
  (values?.length ? values : ["Not specified"]).forEach((value) => wrapper.append(createElement("span", "chip", value)));
  cell.append(wrapper);
}
function formatDate(value) {
  if (!value) return "Never";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "Not available" : date.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}
function distributionLabels(item) {
  const labels = [];
  const sharedCount = item.sharedWithUsersAndGroups?.length || 0;
  const installedCount = item.acquireUsersAndGroups?.length || 0;
  if (sharedCount) labels.push(`Shared with ${sharedCount}`);
  else if (item.type === "shared") labels.push("Shared via link");
  if (installedCount) labels.push(`Installed by ${installedCount}`);
  if (item.deployedTo && item.deployedTo !== "none") labels.push(`Published to ${item.deployedTo}`);
  if (!labels.length && item.availableTo && item.availableTo !== "none") labels.push(`Available to ${item.availableTo}`);
  return labels.length ? labels : ["Not distributed"];
}
function renderTable() {
  const visible = filteredPackages();
  elements.agentRows.replaceChildren();
  elements.resultCount.textContent = `${visible.length} ${visible.length === 1 ? "result" : "results"}`;
  elements.emptyState.hidden = visible.length !== 0;
  elements.tableWrap.hidden = visible.length === 0;
  visible.forEach((item) => {
    const row = document.createElement("tr");
    const nameCell = document.createElement("td");
    const name = createElement("div", "agent-name");
    name.append(createElement("strong", "", item.displayName || "Unnamed agent"), createElement("span", "", item.shortDescription || item.id));
    nameCell.append(name);
    const hostsCell = document.createElement("td"); appendChips(hostsCell, item.supportedHosts);
    const usageCell = createElement("td", "usage-value", item.activeUsers == null ? "Not available" : Number(item.activeUsers).toLocaleString());
    const distributionCell = document.createElement("td"); appendChips(distributionCell, distributionLabels(item));
    const statusCell = document.createElement("td"); statusCell.append(createElement("span", `status ${item.isBlocked ? "blocked" : "available"}`, item.isBlocked ? "Blocked" : "Available"));
    const actionCell = document.createElement("td");
    const action = createElement("button", "icon-button row-button"); action.type = "button"; action.title = `Open ${item.displayName || "agent"} details`; action.setAttribute("aria-label", action.title); action.append(icon("chevron-right")); action.addEventListener("click", () => openDetails(item)); actionCell.append(action);
    row.append(nameCell, createElement("td", "", item.publisher || "Unknown"), createElement("td", "", item.platform || "Not specified"), hostsCell, usageCell, createElement("td", "date-value", formatDate(item.lastUsedDateTime)), distributionCell, statusCell, actionCell);
    elements.agentRows.append(row);
  });
  refreshIcons();
}
async function loadPackages() {
  showLoading();
  try { const payload = await apiRequest("/api/packages"); packages = Array.isArray(payload.value) ? payload.value : []; updateMetrics(); elements.loadingState.hidden = true; elements.refreshButton.disabled = false; renderTable(); }
  catch (error) { showError(error); }
}
function detailItem(label, value) { const wrapper = createElement("div", "detail-item"); wrapper.append(createElement("dt", "", label), createElement("dd", "", value ?? "Not specified")); return wrapper; }
function renderDetails(item) {
  elements.drawerTitle.textContent = item.displayName || "Unnamed agent";
  elements.drawerContent.replaceChildren();
  if (item.longDescription || item.shortDescription) elements.drawerContent.append(createElement("p", "", item.longDescription || item.shortDescription));
  const overview = createElement("section", "detail-section"); overview.append(createElement("h3", "", "Overview"));
  const grid = createElement("dl", "detail-grid");
  grid.append(detailItem("Publisher", item.publisher), detailItem("Status", item.isBlocked ? "Blocked" : "Available"), detailItem("Type", item.type), detailItem("Platform", item.platform), detailItem("Version", item.version), detailItem("Last modified", item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleString() : null), detailItem("Active users", item.activeUsers == null ? null : Number(item.activeUsers).toLocaleString()), detailItem("Last used", item.lastUsedDateTime ? new Date(item.lastUsedDateTime).toLocaleString() : "Never"), detailItem("Total sessions", item.totalSessions == null ? null : Number(item.totalSessions).toLocaleString()), detailItem("Runtime", item.totalRunTimeInHours == null ? null : `${Number(item.totalRunTimeInHours).toLocaleString()} hours`), detailItem("Available to", item.availableTo), detailItem("Deployed to", item.deployedTo), detailItem("Shared with", item.sharedWithUsersAndGroups?.length), detailItem("Installed by", item.acquireUsersAndGroups?.length), detailItem("Package ID", item.id), detailItem("Application ID", item.appId));
  overview.append(grid); elements.drawerContent.append(overview);
  const capabilities = createElement("section", "detail-section"); capabilities.append(createElement("h3", "", "Capabilities"));
  const chips = createElement("div", "chips"); [...(item.supportedHosts || []), ...(item.elementTypes || []), ...(item.categories || [])].forEach((value) => chips.append(createElement("span", "chip", value))); capabilities.append(chips); elements.drawerContent.append(capabilities);
  if (item.elementDetails?.length) { const definitions = createElement("section", "detail-section"); definitions.append(createElement("h3", "", "Element definitions"), createElement("pre", "code-block", JSON.stringify(item.elementDetails, null, 2))); elements.drawerContent.append(definitions); }
}
async function openDetails(summary) {
  lastFocusedElement = document.activeElement; elements.drawerTitle.textContent = summary.displayName || "Agent"; elements.drawerContent.replaceChildren(createElement("p", "muted", "Loading details…")); elements.drawerBackdrop.hidden = false; elements.drawer.classList.add("open"); elements.drawer.setAttribute("aria-hidden", "false"); elements.closeDrawerButton.focus();
  try { renderDetails(await apiRequest(`/api/packages/${encodeURIComponent(summary.id)}`)); } catch (error) { elements.drawerContent.replaceChildren(createElement("p", "", error.message)); }
}
function closeDetails() { elements.drawer.classList.remove("open"); elements.drawer.setAttribute("aria-hidden", "true"); window.setTimeout(() => { elements.drawerBackdrop.hidden = true; }, 220); lastFocusedElement?.focus(); }
async function signIn() {
  try { const result = await msalClient.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" }); account = result.account; msalClient.setActiveAccount(account); setView(true); await loadPackages(); }
  catch (error) { showToast(error.message || "Sign-in was not completed."); }
}
async function signOut() { await msalClient.logoutRedirect({ account, postLogoutRedirectUri: window.location.origin }); }
function toggleTheme() { document.documentElement.dataset.theme = document.documentElement.dataset.theme === "dark" ? "light" : "dark"; }
async function initialize() {
  refreshIcons();
  msalClient = new msal.PublicClientApplication({ auth: { clientId: CLIENT_ID, authority: "https://login.microsoftonline.com/organizations", redirectUri: window.location.origin }, cache: { cacheLocation: "sessionStorage" } });
  await msalClient.initialize();
  const redirectResult = await msalClient.handleRedirectPromise();
  account = redirectResult?.account || msalClient.getActiveAccount() || msalClient.getAllAccounts()[0];
  if (account) { msalClient.setActiveAccount(account); setView(true); await loadPackages(); } else setView(false);
}
elements.signInButton.addEventListener("click", signIn);
elements.signOutButton.addEventListener("click", signOut);
elements.refreshButton.addEventListener("click", loadPackages);
elements.retryButton.addEventListener("click", loadPackages);
elements.searchInput.addEventListener("input", renderTable);
elements.statusFilter.addEventListener("change", renderTable);
elements.closeDrawerButton.addEventListener("click", closeDetails);
elements.drawerBackdrop.addEventListener("click", closeDetails);
elements.themeButton.addEventListener("click", toggleTheme);
document.addEventListener("keydown", (event) => { if (event.key === "Escape" && elements.drawer.classList.contains("open")) closeDetails(); });
initialize().catch((error) => { setView(false); showToast(error.message || "The application could not start."); });