/* global Office, msal */

interface FollowupItem {
  subject: string;
  to: string;
  sent: string; // ISO string
}

const SPA_CLIENT_ID = "a82e535e-8e90-4d9b-9765-9b8ca00769a5";
const AUTHORITY = "https://login.microsoftonline.com/fe62ff8e-1750-452e-b2ff-2d788a3db229";
const GRAPH_SCOPES = ["https://graph.microsoft.com/Mail.Read"];

// ===== MSAL setup =====
const msalApp = new (window as any).msal.PublicClientApplication({
  auth: { clientId: SPA_CLIENT_ID, authority: AUTHORITY },
  cache: { cacheLocation: "sessionStorage" }
});

async function getGraphAccessToken(): Promise<string> {
  const request = { scopes: GRAPH_SCOPES };
  let account = msalApp.getAllAccounts()[0];
  if (!account) {
    const loginResult = await msalApp.loginPopup(request);
    account = loginResult.account;
  }
  const silentReq = { ...request, account };
  try {
    const resp = await msalApp.acquireTokenSilent(silentReq);
    return resp.accessToken;
  } catch {
    const resp = await msalApp.acquireTokenPopup(request);
    return resp.accessToken;
  }
}

// ===== Your existing UI code (trimmed to the important bits) =====
let currentData: FollowupItem[] = [];

Office.onReady(() => {
  (document.getElementById("app-body") as HTMLElement).style.display = "block";

  const btn = document.getElementById("fetchUnreplied") as HTMLButtonElement;
  const tbody = document.querySelector("#resultsTable tbody") as HTMLElement;
  const emptyState = document.getElementById("emptyState") as HTMLElement;
  const errorEl = document.getElementById("error") as HTMLElement;
  const search = document.getElementById("search") as HTMLInputElement;
  const daysSel = document.getElementById("days") as HTMLSelectElement;

  function escapeHtml(s: string) {
    return s.replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]!));
  }

  function render(rows: FollowupItem[]) {
    tbody.innerHTML = "";
    if (!rows.length) { emptyState.style.display = "block"; return; }
    emptyState.style.display = "none";
    for (const item of rows) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td title="${escapeHtml(item.subject)}">${escapeHtml(item.subject)}</td>
        <td title="${escapeHtml(item.to)}">${escapeHtml(item.to)}</td>
        <td>${new Date(item.sent).toLocaleString()}</td>`;
      tbody.appendChild(tr);
    }
  }
  function filterAndRender() {
    const q = (search.value || "").toLowerCase();
    render(currentData.filter(i => i.subject.toLowerCase().includes(q) || i.to.toLowerCase().includes(q)));
  }
  search.addEventListener("input", filterAndRender);

  btn.addEventListener("click", async () => {
    btn.disabled = true; errorEl.style.display = "none"; emptyState.style.display = "none"; tbody.innerHTML = "";
    try {
      const token = await getGraphAccessToken();
      const days = daysSel?.value || "30";
      // UPDATE THIS to your Render API base:
      const apiBase = "https://outlook-followups-backend.onrender.com";
      const resp = await fetch(`${apiBase}/unreplied?days=${encodeURIComponent(days)}`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (!resp.ok) throw new Error(`API ${resp.status}`);
      currentData = await resp.json();
      filterAndRender();
    } catch (e: any) {
      errorEl.textContent = `Error: ${e?.message || String(e)}`;
      errorEl.style.display = "block";
    } finally {
      btn.disabled = false;
    }
  });
});
