/* AffiAds Mail — Microsoft Graph via MSAL public client (OAuth + PKCE).
   IMPORTANT: replace the placeholder with the Application (client) ID from Microsoft Entra.
   Client IDs are public identifiers; never add a client secret to this static site. */
const MICROSOFT_CLIENT_ID = "REMPLACER_PAR_ID_APPLICATION_MICROSOFT";
const REDIRECT_URI = window.location.origin + window.location.pathname;
const SCOPES = ["User.Read", "Mail.Read", "Mail.Send"];
const GRAPH = "https://graph.microsoft.com/v1.0";
let msalApp = null;
let activeAccount = null;
let accessToken = null;

const $ = (id) => document.getElementById(id);
const connectBtn = $("connectBtn");
const disconnectBtn = $("disconnectBtn");
const authStatus = $("authStatus");
const authBadge = $("authBadge");
const mailWorkspace = $("mailWorkspace");
const messageList = $("messageList");

function status(message, isError = false) {
  authStatus.textContent = message;
  authStatus.classList.toggle("error-text", isError);
}
function escapeHtml(value = "") {
  return String(value).replace(/[&<>"']/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
}
function formatDate(value) {
  if (!value) return "";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toLocaleString("fr-FR", {dateStyle:"short",timeStyle:"short"});
}
function updateConnectedUI() {
  const connected = Boolean(activeAccount);
  authBadge.textContent = connected ? "Connecté" : "Non connecté";
  connectBtn.hidden = connected;
  disconnectBtn.hidden = !connected;
  mailWorkspace.hidden = !connected;
  if (connected) status("Connecté en tant que " + (activeAccount.username || activeAccount.name || "compte Microsoft") + ".");
}
function ensureConfigured() {
  if (!window.msal || MICROSOFT_CLIENT_ID.startsWith("REMPLACER_")) {
    status("Configuration à terminer : ajoute l’identifiant public de l’application Microsoft dans app.js après l’avoir créée dans le portail Microsoft.", true);
    return false;
  }
  return true;
}
async function initializeMsal() {
  if (!ensureConfigured()) return false;
  if (!msalApp) {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: MICROSOFT_CLIENT_ID,
        authority: "https://login.microsoftonline.com/consumers",
        redirectUri: REDIRECT_URI,
        postLogoutRedirectUri: REDIRECT_URI
      },
      cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
    });
    await msalApp.initialize();
    const result = await msalApp.handleRedirectPromise();
    if (result?.account) msalApp.setActiveAccount(result.account);
    activeAccount = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0] || null;
    if (activeAccount) msalApp.setActiveAccount(activeAccount);
  }
  return true;
}
async function getToken(interactive = false) {
  if (!await initializeMsal()) throw new Error("La configuration Microsoft n’est pas terminée.");
  if (!activeAccount) {
    if (!interactive) throw new Error("Connecte-toi à ton compte Outlook.");
    const login = await msalApp.loginPopup({scopes: SCOPES, prompt:"select_account"});
    activeAccount = login.account;
    msalApp.setActiveAccount(activeAccount);
  }
  try {
    const result = await msalApp.acquireTokenSilent({scopes:SCOPES, account:activeAccount});
    accessToken = result.accessToken;
    return accessToken;
  } catch (error) {
    if (!interactive) throw error;
    const result = await msalApp.acquireTokenPopup({scopes:SCOPES, account:activeAccount});
    accessToken = result.accessToken;
    return accessToken;
  }
}
async function graph(path, options = {}) {
  const token = await getToken(false);
  const response = await fetch(GRAPH + path, {
    ...options,
    headers: {...(options.headers || {}), Authorization:"Bearer " + token, "Content-Type":"application/json"}
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error("Microsoft Graph (" + response.status + ") : " + body.slice(0, 240));
  }
  if (response.status === 204) return null;
  return response.json();
}
function renderMessages(messages = []) {
  if (!messages.length) {
    messageList.innerHTML = '<p class="small-muted">Aucun message trouvé.</p>';
    return;
  }
  messageList.innerHTML = messages.map((m) => {
    const sender = m.from?.emailAddress || {};
    return '<button type="button" class="message-item" data-message-id="' + escapeHtml(m.id) + '">' +
      '<strong>' + escapeHtml(sender.name || sender.address || "Expéditeur inconnu") + '</strong>' +
      '<span>' + escapeHtml(m.subject || "(sans objet)") + '</span>' +
      '<small>' + escapeHtml(formatDate(m.receivedDateTime)) + '</small></button>';
  }).join("");
  messageList.querySelectorAll("[data-message-id]").forEach((button) => {
    button.addEventListener("click", () => openMessage(button.dataset.messageId));
  });
}
async function loadMessages(query = "") {
  status("Chargement des messages…");
  const params = new URLSearchParams({
    "$top":"25",
    "$select":"id,subject,from,receivedDateTime,bodyPreview,isRead",
    "$orderby":"receivedDateTime desc"
  });
  let path = "/me/mailFolders/inbox/messages?" + params.toString();
  if (query.trim()) {
    const filter = new URLSearchParams({
      "$top":"25",
      "$select":"id,subject,from,receivedDateTime,bodyPreview,isRead",
      "$search": '"' + query.trim().replace(/"/g, '\"') + '"'
    });
    path = "/me/messages?" + filter.toString();
  }
  const data = await graph(path);
  renderMessages(data.value || []);
  status((data.value || []).length + " message(s) affiché(s).");
}
async function openMessage(id) {
  try {
    status("Ouverture du message…");
    const msg = await graph("/me/messages/" + encodeURIComponent(id) + "?$select=subject,from,toRecipients,receivedDateTime,body,isRead");
    $("messageTitle").textContent = msg.subject || "(sans objet)";
    $("messageMeta").textContent = (msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "") + " · " + formatDate(msg.receivedDateTime);
    const body = $("messageBody");
    body.replaceChildren();
    if (msg.body?.contentType === "html") {
      const frame = document.createElement("iframe");
      frame.setAttribute("sandbox", "");
      frame.title = "Contenu du message";
      frame.className = "message-frame";
      frame.srcdoc = "<!doctype html><meta charset='utf-8'><meta http-equiv='Content-Security-Policy' content=\"default-src 'none'; img-src https: data:; style-src 'unsafe-inline'\"><body>" + (msg.body.content || "") + "</body>";
      body.appendChild(frame);
    } else {
      body.textContent = msg.body?.content || "Message vide.";
    }
    status("Message ouvert.");
  } catch (error) { status(error.message || "Impossible d’ouvrir ce message.", true); }
}
async function connect() {
  try {
    if (!await initializeMsal()) return;
    await getToken(true);
    updateConnectedUI();
    await loadMessages();
  } catch (error) {
    status(error.message || "Connexion Microsoft impossible. Vérifie la configuration de l’application.", true);
  }
}
async function disconnect() {
  if (msalApp && activeAccount) {
    const account = activeAccount;
    activeAccount = null; accessToken = null;
    updateConnectedUI();
    try { await msalApp.logoutPopup({account, postLogoutRedirectUri:REDIRECT_URI}); } catch { /* local session is cleared */ }
  }
  status("Déconnecté. Aucun mot de passe n’a été enregistré dans AffiAds Mail.");
  messageList.replaceChildren();
  $("messageTitle").textContent = "Lire un message";
  $("messageMeta").textContent = "";
  $("messageBody").textContent = "Sélectionne un message dans la liste.";
}
connectBtn?.addEventListener("click", connect);
disconnectBtn?.addEventListener("click", disconnect);
$("refreshBtn")?.addEventListener("click", () => loadMessages().catch(error => status(error.message, true)));
$("searchBtn")?.addEventListener("click", () => loadMessages($("searchInput").value).catch(error => status(error.message, true)));
$("searchInput")?.addEventListener("keydown", event => { if (event.key === "Enter") { event.preventDefault(); $("searchBtn").click(); } });
$("sendForm")?.addEventListener("submit", async event => {
  event.preventDefault();
  const to = $("toInput").value.trim();
  const subject = $("subjectInput").value.trim();
  const content = $("bodyInput").value;
  if (!to || !subject || !content.trim()) return;
  if (!window.confirm("Envoyer cet e-mail à " + to + " ?")) return;
  try {
    status("Envoi en cours…");
    await graph("/me/sendMail", {method:"POST", body:JSON.stringify({message:{
      subject,
      body:{contentType:"Text",content},
      toRecipients:[{emailAddress:{address:to}}]
    }, saveToSentItems:true})});
    $("sendForm").reset();
    status("E-mail envoyé avec succès.");
  } catch (error) { status(error.message || "Échec de l’envoi.", true); }
});

const composeBtn = $("composeBtn");
const menuBtn = $("menuBtn");
const sidebar = $("sidebar");
composeBtn?.addEventListener("click", () => {
  document.querySelector("#sendForm")?.scrollIntoView({behavior:"smooth",block:"center"});
  $("toInput")?.focus({preventScroll:true});
  if (!activeAccount) status("Connecte-toi à Outlook pour rédiger et envoyer un message.", true);
});
menuBtn?.addEventListener("click", () => sidebar?.classList.toggle("open"));
document.querySelectorAll(".nav-item").forEach(link => link.addEventListener("click", () => {
  document.querySelectorAll(".nav-item").forEach(item => item.classList.remove("active"));
  link.classList.add("active");
  sidebar?.classList.remove("open");
}));
if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
  window.addEventListener("load", () => navigator.serviceWorker.register("./sw.js").catch(err => console.info("Service worker non enregistré", err)));
}
window.addEventListener("load", () => {
  if (!ensureConfigured()) return;
  initializeMsal().then(() => {
    if (activeAccount) { updateConnectedUI(); getToken(false).then(loadMessages).catch(() => status("Session Microsoft expirée. Reconnecte-toi.")); }
  }).catch(error => status(error.message, true));
});
