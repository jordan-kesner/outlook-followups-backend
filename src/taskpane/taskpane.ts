/* global Office */
interface FollowupItem {
  subject: string;
  to: string;
  sent: string; // ISO string
}

let currentData: FollowupItem[] = [];
let sortKey: "subject" | "to" | "sent" = "sent";
let sortAsc = false;

Office.onReady(() => {
  // Show the app body immediately once Office is ready
  const appBody = document.getElementById("app-body") as HTMLElement;
  if (appBody) appBody.style.display = "block";

  const btn = document.getElementById("fetchUnreplied") as HTMLButtonElement;
  const tbody = document.querySelector("#resultsTable tbody") as HTMLElement;
  const emptyState = document.getElementById("emptyState") as HTMLElement;
  const errorEl = document.getElementById("error") as HTMLElement;
  const search = document.getElementById("search") as HTMLInputElement;
  const daysSel = document.getElementById("days") as HTMLSelectElement;

  function sortData(rows: FollowupItem[]) {
    return [...rows].sort((a, b) => {
      const A = a[sortKey] ?? "", B = b[sortKey] ?? "";
      if (sortKey === "sent") {
        const n = new Date(A).getTime() - new Date(B).getTime();
        return sortAsc ? n : -n;
      }
      const n = String(A).localeCompare(String(B));
      return sortAsc ? n : -n;
    });
  }

  function setSortIndicators() {
    document.querySelectorAll<HTMLTableCellElement>("th[data-key]").forEach(th => {
      const key = th.getAttribute("data-key") as typeof sortKey;
      const span = th.querySelector(".sort") as HTMLElement | null;
      if (!span) return;
      span.textContent = key === sortKey ? (sortAsc ? "▲" : "▼") : "";
    });
  }

  function render(rows: FollowupItem[]) {
    tbody.innerHTML = "";
    if (!rows.length) {
      emptyState.style.display = "block";
      return;
    }
    emptyState.style.display = "none";
    for (const item of rows) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td title="${escapeHtml(item.subject)}">${escapeHtml(item.subject)}</td>
        <td title="${escapeHtml(item.to)}">${escapeHtml(item.to)}</td>
        <td>${new Date(item.sent).toLocaleString()}</td>
      `;
      tbody.appendChild(tr);
    }
  }

  function filterAndRender() {
    const q = (search.value || "").toLowerCase();
    const filtered = currentData.filter(
      i => i.subject.toLowerCase().includes(q) || i.to.toLowerCase().includes(q)
    );
    render(sortData(filtered));
  }

  // Live search
  search.addEventListener("input", filterAndRender);

  // Clickable sortable headers
  document.querySelectorAll<HTMLTableCellElement>("th[data-key]").forEach(th => {
    th.style.cursor = "pointer";
    th.addEventListener("click", () => {
      const key = th.getAttribute("data-key") as typeof sortKey;
      if (key === sortKey) {
        sortAsc = !sortAsc;
      } else {
        sortKey = key;
        sortAsc = key !== "sent"; // default asc for text, desc for date
      }
      setSortIndicators();
      filterAndRender();
    });
  });
  setSortIndicators();

  // Fetch handler with button spinner + days selector
  btn.addEventListener("click", async () => {
    // spinner lives inside the button (see HTML .btn-spinner)
    const spinner = btn.querySelector(".btn-spinner") as HTMLElement | null;

    btn.disabled = true;
    if (spinner) spinner.style.display = "inline-block";
    errorEl.style.display = "none";
    emptyState.style.display = "none";
    tbody.innerHTML = "";

    try {
      // IMPORTANT: use HTTPS if your page is served over HTTPS
      // Change this to your hosted API when you deploy
      const days = daysSel?.value || "30";
      const resp = await fetch(`http://localhost:8000/unreplied?days=${encodeURIComponent(days)}`);
      if (!resp.ok) throw new Error(`API ${resp.status}`);
      currentData = await resp.json(); // [{ subject, to, sent }]
      filterAndRender();
    } catch (err: any) {
      errorEl.textContent = `Error: ${err?.message || String(err)}`;
      errorEl.style.display = "block";
    } finally {
      btn.disabled = false;
      if (spinner) spinner.style.display = "none";
    }
  });
});

// Simple HTML escape to prevent weird characters rendering
function escapeHtml(s: string) {
  return s.replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]!));
}
