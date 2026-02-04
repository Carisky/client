import "bootstrap/dist/css/bootstrap.min.css";
import "./index.css";
import { RAPORT_COLUMNS } from "./raportColumns";

type TabName = "import" | "preview" | "dashboard" | "validation" | "export" | "settings";

const els = {
  tabImportBtn: document.getElementById("tab-btn-import") as HTMLButtonElement,
  tabPreviewBtn: document.getElementById(
    "tab-btn-preview",
  ) as HTMLButtonElement,
  tabDashboardBtn: document.getElementById(
    "tab-btn-dashboard",
  ) as HTMLButtonElement,
  tabValidationBtn: document.getElementById(
    "tab-btn-validation",
  ) as HTMLButtonElement,
  tabExportBtn: document.getElementById("tab-btn-export") as HTMLButtonElement,
  tabSettingsBtn: document.getElementById(
    "tab-btn-settings",
  ) as HTMLButtonElement,

  tabImport: document.getElementById("tab-import") as HTMLElement,
  tabPreview: document.getElementById("tab-preview") as HTMLElement,
  tabDashboard: document.getElementById("tab-dashboard") as HTMLElement,
  tabValidation: document.getElementById("tab-validation") as HTMLElement,
  tabExport: document.getElementById("tab-export") as HTMLElement,
  tabSettings: document.getElementById("tab-settings") as HTMLElement,

  importBtn: document.getElementById("btn-import") as HTMLButtonElement,
  importProgress: document.getElementById(
    "import-progress",
  ) as HTMLProgressElement,
  importProgressText: document.getElementById(
    "import-progress-text",
  ) as HTMLElement,
  importStatus: document.getElementById("import-status") as HTMLElement,
  previewStatus: document.getElementById("preview-status") as HTMLElement,
  dashboardStatus: document.getElementById("dashboard-status") as HTMLElement,
  settingsStatus: document.getElementById("settings-status") as HTMLElement,

  meta: document.getElementById("meta") as HTMLElement,
  appVersion: document.getElementById("app-version") as HTMLElement,

  btnPrev: document.getElementById("btn-prev") as HTMLButtonElement,
  btnNext: document.getElementById("btn-next") as HTMLButtonElement,
  btnRefresh: document.getElementById("btn-refresh") as HTMLButtonElement,
  pageInfo: document.getElementById("page-info") as HTMLElement,
  pageSize: document.getElementById("page-size") as HTMLSelectElement,
  tableHead: document.querySelector(
    "#data-table thead",
  ) as HTMLTableSectionElement,
  tableBody: document.querySelector(
    "#data-table tbody",
  ) as HTMLTableSectionElement,

  btnMrnRebuild: document.getElementById(
    "btn-mrn-rebuild",
  ) as HTMLButtonElement,
  btnMrnRefresh: document.getElementById(
    "btn-mrn-refresh",
  ) as HTMLButtonElement,
  mrnMeta: document.getElementById("mrn-meta") as HTMLElement,
  mrnGroups: document.getElementById("mrn-groups") as HTMLElement,

  validationMonth: document.getElementById(
    "validation-month",
  ) as HTMLInputElement,
  validationPeriod: document.getElementById(
    "validation-period",
  ) as HTMLSelectElement,
  validationYear: document.getElementById(
    "validation-year",
  ) as HTMLInputElement,
  validationGrouping: document.getElementById(
    "validation-grouping",
  ) as HTMLSelectElement,
  validationMrnFilter: document.getElementById(
    "validation-mrn-filter",
  ) as HTMLInputElement,
  btnValidationMrnClear: document.getElementById(
    "btn-validation-mrn-clear",
  ) as HTMLButtonElement,
  btnValidationRefresh: document.getElementById(
    "btn-validation-refresh",
  ) as HTMLButtonElement,
  validationMeta: document.getElementById("validation-meta") as HTMLElement,
  validationGroups: document.getElementById("validation-groups") as HTMLElement,
  validationStatus: document.getElementById("validation-status") as HTMLElement,

  exportPeriod: document.getElementById("export-period") as HTMLSelectElement,
  exportMonth: document.getElementById("export-month") as HTMLInputElement,
  exportYear: document.getElementById("export-year") as HTMLInputElement,
  exportGrouping: document.getElementById("export-grouping") as HTMLSelectElement,
  exportMrnFilter: document.getElementById("export-mrn-filter") as HTMLInputElement,
  btnExportRefresh: document.getElementById("btn-export-refresh") as HTMLButtonElement,
  btnExportDo: document.getElementById("btn-export-do") as HTMLButtonElement,
  exportMeta: document.getElementById("export-meta") as HTMLElement,
  exportPreview: document.getElementById("export-preview") as HTMLElement,
  exportStatus: document.getElementById("export-status") as HTMLElement,

  exportFilterImporter: document.getElementById("export-filter-importer") as HTMLInputElement,
  exportAgentBtn: document.getElementById("export-filter-agent-btn") as HTMLButtonElement,
  exportAgentPopover: document.getElementById("export-filter-agent-popover") as HTMLElement,
  exportAgentSearch: document.getElementById("export-filter-agent-search") as HTMLInputElement,
  exportAgentList: document.getElementById("export-filter-agent-list") as HTMLElement,
  btnExportAgentClear: document.getElementById("export-filter-agent-clear") as HTMLButtonElement,
  exportFilterDzial: document.getElementById("export-filter-dzial") as HTMLInputElement,
  btnExportFiltersClear: document.getElementById("btn-export-filters-clear") as HTMLButtonElement,

  dbPath: document.getElementById("db-path") as HTMLElement,
  btnShowDb: document.getElementById("btn-show-db") as HTMLButtonElement,
  btnClear: document.getElementById("btn-clear") as HTMLButtonElement,

  agentDzialMeta: document.getElementById("agent-dzial-meta") as HTMLElement,
  agentDzialStatus: document.getElementById(
    "agent-dzial-status",
  ) as HTMLElement,
  btnAgentDzialShow: document.getElementById(
    "btn-agent-dzial-show",
  ) as HTMLButtonElement,
  btnAgentDzialClear: document.getElementById(
    "btn-agent-dzial-clear",
  ) as HTMLButtonElement,
};

const state = {
  tab: "import" as TabName,
  page: 1,
  pageSize: 250,
  total: 0,
  columns: [] as Array<{ field: string; label: string }>,
};

let busyCount = 0;
function setBusy(isBusy: boolean) {
  if (isBusy) busyCount += 1;
  else busyCount = Math.max(0, busyCount - 1);
  document.body.classList.toggle("is-busy", busyCount > 0);
}

let lastUpdateCheck: Awaited<
  ReturnType<typeof window.api.checkForUpdates>
> | null = null;
let lastUpdateStatus: unknown | null = null;
let updateStatusUnsub: (() => void) | null = null;
let updateFallbackUrl: string | null = null;
let updatePrimaryBtn: HTMLButtonElement | null = null;
let updateFallbackBtn: HTMLButtonElement | null = null;

function renderUpdateStatus(status: unknown) {
  if (!document.getElementById("update-backdrop")) return;

  const textEl = document.getElementById("update-status-text");
  const barWrap = document.getElementById("update-progress-wrap");
  const barEl = document.getElementById(
    "update-progress-bar",
  ) as HTMLDivElement | null;
  if (!textEl) return;

  const s = (status ?? {}) as {
    state?: unknown;
    message?: unknown;
    percent?: unknown;
  };
  const state = typeof s.state === "string" ? s.state : "idle";
  const msg = typeof s.message === "string" ? s.message : "";
  const pct =
    typeof s.percent === "number" && Number.isFinite(s.percent)
      ? Math.max(0, Math.min(100, s.percent))
      : null;
const label =
  state === "checking"
    ? "Sprawdzam aktualizacje…"
    : state === "available"
      ? "Znaleziono aktualizację."
      : state === "not-available"
        ? "Brak dostępnych aktualizacji."
        : state === "downloading"
          ? `Pobieranie…${pct == null ? "" : ` ${pct.toFixed(0)}%`}`
          : state === "downloaded"
            ? "Pobrano. Instaluję…"
            : state === "installing"
              ? "Instaluję…"
              : state === "error"
                ? `Błąd aktualizacji: ${msg || "unknown"}`
                : msg || "";

  textEl.textContent = label;

  if (barWrap && barEl) {
    if (pct == null || state !== "downloading") {
      barWrap.classList.add("d-none");
      barEl.style.width = "0%";
    } else {
      barWrap.classList.remove("d-none");
      barEl.style.width = `${pct}%`;
    }
  }

  if (updateFallbackBtn) {
    const showFallback = state === "error" || state === "not-available";
    updateFallbackBtn.classList.toggle(
      "d-none",
      !showFallback || !updateFallbackUrl,
    );
  }
  if (updatePrimaryBtn) {
    const enable =
      state === "idle" || state === "error" || state === "not-available";
    updatePrimaryBtn.disabled = !enable;
  }
}

function renderUpdateBlock(
  latestVersion: string,
  downloadUrl: string,
  currentVersion: string,
  feedUrl: string | null,
) {
  if (document.getElementById("update-backdrop")) return;
  document.body.classList.add("update-required");

  const el = document.createElement("div");
  el.id = "update-backdrop";
  el.className = "update-backdrop";
  el.innerHTML = `
  <div class="update-card" role="dialog" aria-modal="true" aria-label="Wymagana aktualizacja">
    <div class="update-card-header">
      <div>
        <div class="update-title">Wymagana aktualizacja</div>
        <div class="update-sub">Aktualna wersja: ${escapeHtml(currentVersion)} • Dostępna: ${escapeHtml(latestVersion)}</div>
      </div>
      <div class="update-badge">
        <svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true" style="margin:0">
          <path d="M12 3v10m0 0l4-4m-4 4l-4-4" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" />
          <path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" />
        </svg>
        v${escapeHtml(latestVersion)}
      </div>
    </div>
    <div class="update-card-body">
      <div class="muted">Aby kontynuować, zainstaluj najnowszą wersję aplikacji.</div>
      <div id="update-status-text" class="muted" style="margin-top:8px"></div>
      <div id="update-progress-wrap" class="progress d-none" style="height:8px; margin-top:8px">
        <div id="update-progress-bar" class="progress-bar" role="progressbar" style="width:0%"></div>
      </div>
      <div class="update-actions">
        <button id="btn-update-download" class="btn btn-outline-light d-none">
          Pobierz instalator
        </button>
        <button id="btn-update-now" class="btn btn-primary">
          Zaktualizuj do wersji ${escapeHtml(latestVersion)}
        </button>
      </div>
    </div>
  </div>
`;
  document.body.appendChild(el);

  updateFallbackUrl = downloadUrl;
  updatePrimaryBtn = el.querySelector(
    "#btn-update-now",
  ) as HTMLButtonElement | null;
  updateFallbackBtn = el.querySelector(
    "#btn-update-download",
  ) as HTMLButtonElement | null;

  updateFallbackBtn?.addEventListener("click", async () => {
    if (!updateFallbackUrl) return;
    updateFallbackBtn.disabled = true;
    try {
      await window.api.openExternal(updateFallbackUrl);
    } finally {
      void window.api.quitApp();
    }
  });

  updatePrimaryBtn?.addEventListener("click", async () => {
    updatePrimaryBtn.disabled = true;
    try {
      if (feedUrl) {
        const started = await window.api.downloadAndInstallUpdate(feedUrl);
        if (!started.ok) {
          renderUpdateStatus({
            state: "error",
            message: started.error ?? "Failed to start updater",
          });
          updateFallbackBtn?.classList.remove("d-none");
        } else {
          renderUpdateStatus({ state: "checking" });
        }
      } else {
        updateFallbackBtn?.classList.remove("d-none");
        renderUpdateStatus({
          state: "error",
          message: "Auto-update not available.",
        });
      }
    } finally {
      // autoUpdater will quit+install; fallback button can quit the app
    }
  });

  renderUpdateStatus(lastUpdateStatus);

  const stop = (e: Event) => e.preventDefault();
  window.addEventListener("keydown", stop, { capture: true });
}

async function checkForUpdatesAndBlock() {
  try {
    const res = await window.api.checkForUpdates();
    lastUpdateCheck = res;
    if (!res.supported) return;
    if (!res.updateAvailable) return;
    if (!res.latestVersion || !res.downloadUrl) return;
    renderUpdateBlock(
      res.latestVersion,
      res.downloadUrl,
      res.currentVersion,
      res.squirrelFeedUrl,
    );
  } catch {
    // ignore
  }
}

function setStatus(el: HTMLElement, text: string) {
  el.textContent = text;
}

function errorMessage(e: unknown): string {
  if (e instanceof Error) return e.message;
  try {
    return String(e);
  } catch {
    return "Nieznany błąd";
  }
}

function setTab(name: TabName) {
  state.tab = name;

  els.tabImportBtn.classList.toggle("active", name === "import");
  els.tabPreviewBtn.classList.toggle("active", name === "preview");
  els.tabDashboardBtn.classList.toggle("active", name === "dashboard");
  els.tabValidationBtn.classList.toggle("active", name === "validation");
  els.tabExportBtn.classList.toggle("active", name === "export");
  els.tabSettingsBtn.classList.toggle("active", name === "settings");

  els.tabImportBtn.setAttribute("aria-selected", String(name === "import"));
  els.tabPreviewBtn.setAttribute("aria-selected", String(name === "preview"));
  els.tabDashboardBtn.setAttribute(
    "aria-selected",
    String(name === "dashboard"),
  );
  els.tabValidationBtn.setAttribute(
    "aria-selected",
    String(name === "validation"),
  );
  els.tabExportBtn.setAttribute("aria-selected", String(name === "export"));
  els.tabSettingsBtn.setAttribute("aria-selected", String(name === "settings"));

  els.tabImport.classList.toggle("hidden", name !== "import");
  els.tabPreview.classList.toggle("hidden", name !== "preview");
  els.tabDashboard.classList.toggle("hidden", name !== "dashboard");
  els.tabValidation.classList.toggle("hidden", name !== "validation");
  els.tabExport.classList.toggle("hidden", name !== "export");
  els.tabSettings.classList.toggle("hidden", name !== "settings");

  if (name === "preview") void refreshPreview();
  if (name === "dashboard") void refreshDashboard();
  if (name === "validation") void refreshValidation();
  if (name === "export") void refreshExportPreview();
  if (name === "settings") void refreshSettings();
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function renderTable(page: {
  columns: Array<{ field: string; label: string }>;
  rows: Array<Record<string, string | null>>;
}) {
  const headers = page.columns;
  const thead = `<tr>${headers.map((c) => `<th title="${escapeHtml(c.label)}">${escapeHtml(c.label)}</th>`).join("")}</tr>`;
  els.tableHead.innerHTML = thead;

  const body = page.rows
    .map((row) => {
      const tds = headers
        .map((c) => {
          const v = row[c.field];
          const text = v == null ? "" : String(v);
          return `<td title="${escapeHtml(text)}">${escapeHtml(text)}</td>`;
        })
        .join("");
      return `<tr>${tds}</tr>`;
    })
    .join("");
  els.tableBody.innerHTML = body;
}

async function refreshMeta() {
  try {
    const meta = await window.api.getRaportMeta();
    if (!meta.importedAt || meta.rowCount === 0) {
      els.meta.textContent = "Brak zaimportowanych danych";
      return;
    }
    const when = new Date(meta.importedAt).toLocaleString("pl-PL");
    const file = meta.sourceFile ? ` • Plik: ${meta.sourceFile}` : "";
    els.meta.textContent = `Zaimportowano: ${when} • Wiersze: ${meta.rowCount}${file}`;
  } catch {
    els.meta.textContent = "Nie udało się odczytać informacji o imporcie";
  }
}

function updatePagination() {
  const totalPages = Math.max(1, Math.ceil(state.total / state.pageSize));
  if (state.page > totalPages) state.page = totalPages;

  els.btnPrev.disabled = state.page <= 1;
  els.btnNext.disabled = state.page >= totalPages;
  els.pageInfo.textContent = `Strona ${state.page} z ${totalPages} (łącznie: ${state.total})`;
}

async function refreshPreview() {
  setStatus(els.previewStatus, "Ładowanie danych…");
  setBusy(true);
  await refreshMeta();

  try {
    const page = await window.api.getRaportPage(state.page, state.pageSize);
    state.total = page.total;
    state.columns = page.columns;
    renderTable(page);
    updatePagination();
    setStatus(
      els.previewStatus,
      page.total === 0 ? "Brak danych do wyświetlenia." : "",
    );
  } catch (e: unknown) {
    setStatus(els.previewStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    setBusy(false);
  }
}

function renderAgentDzialMetaText(info: {
  filePath: string;
  exists: boolean;
  rowCount: number;
  modifiedAt: string | null;
  error?: string;
} | null): string {
  if (!info) return "—";
  if (info.error) return `Błąd: ${info.error}`;
  const rows = info.rowCount ?? 0;
  const mod = info.modifiedAt ? ` • ${new Date(info.modifiedAt).toLocaleString("pl-PL")}` : "";
  return `${rows} pozycji • ${info.filePath}${mod}`;
}

async function refreshAgentDzialUi(): Promise<void> {
  setStatus(els.agentDzialStatus, "");
  try {
    const info = await window.api.getAgentDzialInfo();
    els.agentDzialMeta.textContent = renderAgentDzialMetaText(info);
    els.btnAgentDzialClear.disabled = !info || info.rowCount === 0;
  } catch (e: unknown) {
    els.agentDzialMeta.textContent = "—";
    els.btnAgentDzialClear.disabled = true;
    setStatus(els.agentDzialStatus, `Błąd: ${errorMessage(e)}`);
  }
}

async function refreshSettings() {
  setStatus(els.settingsStatus, "");
  setStatus(els.agentDzialStatus, "");
  setBusy(true);
  try {
    const db = await window.api.getDbInfo();
    els.dbPath.textContent =
      db.filePath + (db.exists ? "" : " (nie utworzono)");
    await refreshAgentDzialUi();
  } catch (e: unknown) {
    els.dbPath.textContent = "—";
    setStatus(els.settingsStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    const showUpdateDebug = localStorage.getItem("showUpdateDebug") === "1";
    if (showUpdateDebug && !els.settingsStatus.textContent && lastUpdateCheck) {
      const u = lastUpdateCheck;
      const manifest = u.manifestUrl ? ` (${u.manifestUrl})` : "";
      const base = `Updates: current=${u.currentVersion} latest=${u.latestVersion ?? "—"} available=${u.updateAvailable}`;
      if (!u.supported)
        setStatus(els.settingsStatus, `${base} disabled${manifest}`);
      else if (u.error)
        setStatus(els.settingsStatus, `${base} error=${u.error}${manifest}`);
      else setStatus(els.settingsStatus, `${base}${manifest}`);
    }
    setBusy(false);
  }
}

async function refreshAppVersion() {
  if (!els.appVersion) return;
  try {
    const res = await window.api.getAppVersion();
    const v = String(res?.version ?? "").trim();
    els.appVersion.textContent = v || "—";
  } catch {
    els.appVersion.textContent = "—";
  }
}

const LABELS: Record<string, string> = {
  id: "ID",
  rowNumber: "Row (Excel)",
  ...Object.fromEntries(RAPORT_COLUMNS.map((c) => [c.field, c.label])),
};

function formatMaybeDate(s: string): string {
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) return s;
  return d.toLocaleString("pl-PL");
}

function uniqValues(
  rows: Array<Record<string, string | null>>,
  field: string,
): string[] {
  const out = new Set<string>();
  for (const r of rows) {
    const v = r[field];
    if (v == null) continue;
    const s = String(v).trim();
    if (s) out.add(s);
  }
  return Array.from(out);
}

function renderPills(values: string[], emptyText = "-") {
  if (values.length === 0)
    return `<div class="pills"><span class="pill">${escapeHtml(emptyText)}</span></div>`;
  const slice = values.slice(0, 30);
  const rest = values.length - slice.length;
  return `<div class="pills">${slice
    .map(
      (v) =>
        `<span class="pill" title="${escapeHtml(v)}">${escapeHtml(v)}</span>`,
    )
    .join("")}${rest > 0 ? `<span class="pill">+${rest}</span>` : ""}</div>`;
}

const ROW_DETAIL_KEYS: ReadonlyArray<string> = [
  "id",
  "rowNumber",
  "nr_sad",
  "numer_pozycji",
  "rodzaj_sad",
  "typ_sad_u",
  "data_sad_p_54",
  "data_mrn",
  "stan",
  "numer_mrn",
];

function renderRowKv(row: Record<string, string | null>): string {
  const keys: ReadonlyArray<string> = ROW_DETAIL_KEYS;
  const parts: string[] = [];

  for (const k of keys) {
    const v = row[k];
    if (v == null) continue;
    const s = String(v).trim();
    if (!s) continue;
    const label = LABELS[k] ?? k;
    parts.push(
      `<div class="kv-key" title="${escapeHtml(k)}">${escapeHtml(label)}</div><div class="kv-val">${escapeHtml(s)}</div>`,
    );
  }

  if (parts.length === 0) return `<div class="muted">Brak danych.</div>`;
  return `<div class="kv-grid">${parts.join("")}</div>`;
}

async function loadMrnGroupDetails(detailsEl: HTMLDetailsElement) {
  const numerMrn = detailsEl.dataset.mrn ?? "";
  if (!numerMrn) return;
  if (detailsEl.dataset.loaded === "1") return;
  if (detailsEl.dataset.loading === "1") return;

  const body = detailsEl.querySelector(".accordion-body") as HTMLElement | null;
  if (!body) return;

  detailsEl.dataset.loading = "1";
  body.innerHTML = `<div class="muted">Ładowanie...</div>`;

  try {
    const res = await window.api.getMrnBatchRows(numerMrn);
    const rows = res.rows ?? [];

    const nrSad = uniqValues(rows, "nr_sad");
    const dataMrn = uniqValues(rows, "data_mrn");

    const rowsHtml =
      rows.length === 0
        ? `<div class="muted">Brak wierszy.</div>`
        : rows
            .map((row) => {
              const rowNumber = row.rowNumber ? String(row.rowNumber) : "-";
              const nr = row.nr_sad ? String(row.nr_sad) : "-";
              const dt = row.data_mrn ? String(row.data_mrn) : "-";
              const st = row.stan ? String(row.stan) : "-";

              return `
                <details class="row-accordion">
                  <summary>
                    <div class="row-summary">
                      <div class="muted" title="rowNumber">${escapeHtml(rowNumber)}</div>
                      <div class="muted" title="nr_sad">${escapeHtml(nr)}</div>
                      <div class="muted" title="data_mrn">${escapeHtml(dt)}</div>
                      <div class="muted" title="stan">${escapeHtml(st)}</div>
                    </div>
                  </summary>
                  ${renderRowKv(row)}
                </details>
              `;
            })
            .join("");

    body.innerHTML = `
      <div class="muted">Nr SAD</div>
      ${renderPills(nrSad)}
      <div class="muted">Data MRN</div>
      ${renderPills(dataMrn)}
      ${rowsHtml}
    `;
    detailsEl.dataset.loaded = "1";
  } catch (e: unknown) {
    body.innerHTML = `<div class="muted">Błąd: ${escapeHtml(errorMessage(e))}</div>`;
  } finally {
    detailsEl.dataset.loading = "0";
  }
}

async function refreshDashboard() {
  setStatus(els.dashboardStatus, "Ładowanie...");
  els.mrnGroups.innerHTML = "";
  els.mrnMeta.textContent = "";
  setBusy(true);

  try {
    const [meta, groups] = await Promise.all([
      window.api.getMrnBatchMeta(),
      window.api.getMrnBatchGroups(1000),
    ]);

    const scanned = meta.scannedAt ? formatMaybeDate(meta.scannedAt) : "-";
    els.mrnMeta.textContent = `Skan: ${scanned} | Grupy: ${meta.groups} | Wiersze: ${meta.rows}`;

    if (groups.length === 0) {
      els.mrnGroups.innerHTML = `<div class="muted" style="padding:10px 12px;">Brak duplikatów. Kliknij „Skanuj”, aby zbudować mrn_batch.</div>`;
      setStatus(els.dashboardStatus, "");
      return;
    }

    els.mrnGroups.innerHTML = groups
      .map(
        (g) => `
          <details class="accordion" data-mrn="${escapeHtml(g.numer_mrn)}">
            <summary>
              <span class="mrn-code" title="${escapeHtml(g.numer_mrn)}">${escapeHtml(g.numer_mrn)}</span>
              <span class="badge rounded-pill badge-count">${g.count}</span>
            </summary>
            <div class="accordion-body">
              <div class="muted">Otwórz, aby załadować szczegóły.</div>
            </div>
          </details>
        `,
      )
      .join("");

    for (const el of Array.from(
      els.mrnGroups.querySelectorAll("details.accordion"),
    )) {
      const d = el as HTMLDetailsElement;
      d.addEventListener("toggle", () => {
        if (d.open) void loadMrnGroupDetails(d);
      });
    }

    setStatus(els.dashboardStatus, "");
  } catch (e: unknown) {
    setStatus(els.dashboardStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    setBusy(false);
  }
}

type ValidationGroupKey = {
  odbiorca: string;
  kraj_wysylki: string;
  warunki_dostawy: string;
  waluta: string;
  kurs_waluty: string;
  transport_na_granicy_rodzaj: string;
  kod_towaru: string;
};

type ValidationDayFilter = "all" | "outliersHigh" | "outliersLow" | "singles";
type ValidationPeriodMode = "month" | "year";

type ValidationOutlierError = Awaited<
  ReturnType<typeof window.api.getValidationOutlierErrors>
>["items"][number];

function encodeKey(key: unknown): string {
  try {
    return btoa(unescape(encodeURIComponent(JSON.stringify(key))));
  } catch {
    return "";
  }
}

function decodeKey<T>(key: string): T | null {
  try {
    const s = decodeURIComponent(escape(atob(key)));
    return JSON.parse(s) as T;
  } catch {
    return null;
  }
}

function prettyKeyLine(label: string, value: string): string {
  const v = value && value.trim().length > 0 ? value : "-";
  return `<div class="kv-key">${escapeHtml(label)}</div><div class="kv-val">${escapeHtml(v)}</div>`;
}

function renderCoefArrow(side: "low" | "high" | null): string {
  if (side === "low")
    return '<span class="coef-arrow low" title="Below lower IQR bound">&darr;</span>';
  if (side === "high")
    return '<span class="coef-arrow high" title="Above upper IQR bound">&uarr;</span>';
  return "";
}

function getValidationPeriodMode(): ValidationPeriodMode {
  const v = String(els.validationPeriod?.value ?? "month").trim();
  return v === "year" ? "year" : "month";
}

function getValidationPeriodValue(): string {
  const mode = getValidationPeriodMode();
  if (mode === "year") {
    const raw = String(els.validationYear.value ?? "").trim();
    const y = Number.parseInt(raw, 10);
    if (!Number.isFinite(y) || y < 1900 || y > 2200) return "";
    return String(y).padStart(4, "0");
  }
  return String(els.validationMonth.value ?? "").trim();
}

function updateValidationPeriodUi(): void {
  const mode = getValidationPeriodMode();
  els.validationMonth.classList.toggle("hidden", mode !== "month");
  els.validationYear.classList.toggle("hidden", mode !== "year");
}

function getExportPeriodMode(): ValidationPeriodMode {
  const v = String(els.exportPeriod?.value ?? "month").trim();
  return v === "year" ? "year" : "month";
}

function getExportPeriodValue(): string {
  const mode = getExportPeriodMode();
  if (mode === "year") {
    const raw = String(els.exportYear.value ?? "").trim();
    const y = Number.parseInt(raw, 10);
    if (!Number.isFinite(y) || y < 1900 || y > 2200) return "";
    return String(y).padStart(4, "0");
  }
  return String(els.exportMonth.value ?? "").trim();
}

function updateExportPeriodUi(): void {
  const mode = getExportPeriodMode();
  els.exportMonth.classList.toggle("hidden", mode !== "month");
  els.exportYear.classList.toggle("hidden", mode !== "year");
}

type ValidationDateGrouping =
  | "day"
  | "days2"
  | "days3"
  | "week"
  | "month"
  | "months2";

const VALIDATION_GROUPING_STORAGE_KEY = "validationGrouping";
const EXPORT_GROUPING_STORAGE_KEY = "exportGrouping";

function normalizeValidationDateGrouping(value: unknown): ValidationDateGrouping {
  const v = String(value ?? "").trim();
  if (v === "days2") return "days2";
  if (v === "days3") return "days3";
  if (v === "week") return "week";
  if (v === "month") return "month";
  if (v === "months2") return "months2";
  return "day";
}

function getValidationGroupingOptions(): { grouping: ValidationDateGrouping } {
  const grouping = normalizeValidationDateGrouping(els.validationGrouping?.value);
  return { grouping };
}

function getExportGroupingOptions(): { grouping: ValidationDateGrouping } {
  const grouping = normalizeValidationDateGrouping(els.exportGrouping?.value);
  return { grouping };
}

function setValidationGroupingValue(value: ValidationDateGrouping): void {
  els.validationGrouping.value = value;
  try {
    localStorage.setItem(VALIDATION_GROUPING_STORAGE_KEY, value);
  } catch {
    // ignore
  }
}

function setExportGroupingValue(value: ValidationDateGrouping): void {
  els.exportGrouping.value = value;
  try {
    localStorage.setItem(EXPORT_GROUPING_STORAGE_KEY, value);
  } catch {
    // ignore
  }
}

function formatBucketLabel(start: string, end: string): string {
  if (!start) return "-";
  if (!end || end === start) return start;
  return `${start}–${end}`;
}

const VALIDATION_MRN_FILTER_STORAGE_KEY = "validationMrnFilter";
const EXPORT_MRN_FILTER_STORAGE_KEY = "exportMrnFilter";

function getValidationMrnFilterValue(): string {
  return String(els.validationMrnFilter?.value ?? "").trim();
}

function setValidationMrnFilterValue(value: string): void {
  els.validationMrnFilter.value = value;
  try {
    localStorage.setItem(VALIDATION_MRN_FILTER_STORAGE_KEY, value);
  } catch {
    // ignore
  }
  updateValidationMrnFilterUi();
}

function updateValidationMrnFilterUi(): void {
  const v = getValidationMrnFilterValue();
  els.btnValidationMrnClear.classList.toggle("d-none", v.length === 0);
}

function getExportMrnFilterValue(): string {
  return String(els.exportMrnFilter?.value ?? "").trim();
}

type ExportFilters = {
  importer?: string;
  agent?: string[];
  dzial?: string;
};

function normalizeAgentKey(value: unknown): string {
  return String(value ?? "").trim().toLowerCase();
}

type ExportAgentOption = { key: string; name: string };
let exportAgentOptions: ExportAgentOption[] = [];
let exportAgentSelectedKeys = new Set<string>();

let exportAgentOutsideClickUnsub: (() => void) | null = null;
let exportAgentKeydownUnsub: (() => void) | null = null;

function getSelectedExportAgents(): string[] {
  if (!exportAgentSelectedKeys.size || !exportAgentOptions.length) return [];
  const map = new Map(exportAgentOptions.map((o) => [o.key, o.name] as const));
  const out: string[] = [];
  for (const k of exportAgentSelectedKeys) {
    const v = map.get(k);
    if (v) out.push(v);
  }
  out.sort((a, b) => a.localeCompare(b));
  return out;
}

function updateExportAgentUi(): void {
  const selected = getSelectedExportAgents();
  els.btnExportAgentClear.disabled = selected.length === 0;

  if (!exportAgentOptions.length) {
    els.exportAgentBtn.textContent = "Agent celny: Ładowanie…";
    els.exportAgentBtn.disabled = true;
    return;
  }

  els.exportAgentBtn.disabled = false;
  if (selected.length === 0) {
    els.exportAgentBtn.textContent = "Agent celny: wszyscy";
  } else if (selected.length === 1) {
    els.exportAgentBtn.textContent = `Agent celny: ${selected[0]}`;
  } else {
    els.exportAgentBtn.textContent = `Agent celny: ${selected.length} wybranych`;
  }
}

function renderExportAgentList(): void {
  const q = String(els.exportAgentSearch.value ?? "").trim().toLowerCase();
  const items = q
    ? exportAgentOptions.filter((o) => o.name.toLowerCase().includes(q))
    : exportAgentOptions;

  els.exportAgentList.innerHTML = "";
  if (!items.length) {
    els.exportAgentList.innerHTML = `<div class="agent-filter-empty">Brak wyników.</div>`;
    return;
  }

  const frag = document.createDocumentFragment();
  for (const o of items) {
    const label = document.createElement("label");
    label.className = "agent-filter-item";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.className = "form-check-input";
    input.dataset.key = o.key;
    input.checked = exportAgentSelectedKeys.has(o.key);

    const span = document.createElement("span");
    span.textContent = o.name;

    label.appendChild(input);
    label.appendChild(span);
    frag.appendChild(label);
  }
  els.exportAgentList.appendChild(frag);
}

function setExportAgentPopoverOpen(open: boolean): void {
  els.exportAgentPopover.classList.toggle("hidden", !open);
  if (open) {
    renderExportAgentList();
    els.exportAgentSearch.focus();
    els.exportAgentSearch.select();

    if (!exportAgentOutsideClickUnsub) {
      const onDown = (e: MouseEvent) => {
        const target = e.target as Node | null;
        if (!target) return;
        if (els.exportAgentPopover.contains(target)) return;
        if (els.exportAgentBtn.contains(target)) return;
        setExportAgentPopoverOpen(false);
      };
      document.addEventListener("mousedown", onDown, true);
      exportAgentOutsideClickUnsub = () =>
        document.removeEventListener("mousedown", onDown, true);
    }

    if (!exportAgentKeydownUnsub) {
      const onKey = (e: KeyboardEvent) => {
        if (e.key === "Escape") setExportAgentPopoverOpen(false);
      };
      document.addEventListener("keydown", onKey, true);
      exportAgentKeydownUnsub = () =>
        document.removeEventListener("keydown", onKey, true);
    }
  } else {
    exportAgentOutsideClickUnsub?.();
    exportAgentOutsideClickUnsub = null;
    exportAgentKeydownUnsub?.();
    exportAgentKeydownUnsub = null;
  }
}

function setAvailableExportAgents(agents: unknown): void {
  const input = Array.isArray(agents) ? agents : [];
  const map = new Map<string, string>();
  for (const a of input) {
    const name = String(a ?? "").trim();
    if (!name) continue;
    const key = normalizeAgentKey(name);
    if (!key || map.has(key)) continue;
    map.set(key, name);
  }
  exportAgentOptions = Array.from(map.entries())
    .map(([key, name]) => ({ key, name }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const availableKeys = new Set(exportAgentOptions.map((o) => o.key));
  exportAgentSelectedKeys = new Set(
    Array.from(exportAgentSelectedKeys).filter((k) => availableKeys.has(k)),
  );

  updateExportAgentUi();
  renderExportAgentList();
}

function getExportFilters(): ExportFilters {
  const importer = String(els.exportFilterImporter?.value ?? "").trim();
  const agent = getSelectedExportAgents();
  const dzial = String(els.exportFilterDzial?.value ?? "").trim();

  const out: ExportFilters = {};
  if (importer) out.importer = importer;
  if (agent.length) out.agent = agent;
  if (dzial) out.dzial = dzial;
  return out;
}

function updateExportFiltersUi(): void {
  const f = getExportFilters();
  const any = Boolean(f.importer || (f.agent && f.agent.length) || f.dzial);
  els.btnExportFiltersClear.disabled = !any;
}

function clearExportFilters(): void {
  els.exportFilterImporter.value = "";
  exportAgentSelectedKeys.clear();
  els.exportFilterDzial.value = "";
  updateExportAgentUi();
  updateExportFiltersUi();
}

function formatPct(pct: number | null): string {
  if (pct == null || !Number.isFinite(pct)) return "-";
  if (pct < 0.05) return "<0.1%";
  return `${pct.toFixed(1)}%`;
}

function formatNum(n: number | null): string {
  if (n == null || !Number.isFinite(n)) return "-";
  return n.toFixed(4);
}

function renderValidationErrorsByAgent(items: ValidationOutlierError[]): string {
  if (!items || items.length === 0) {
    return `<div class="muted" style="padding:6px 2px 10px;">Brak błędów (odchyleń poza limitem).</div>`;
  }

  const byAgent = new Map<string, ValidationOutlierError[]>();
  for (const it of items) {
    const raw = String(it.agent_celny ?? "").trim();
    const agent = raw.length > 0 ? raw : "—";
    const arr = byAgent.get(agent);
    if (arr) arr.push(it);
    else byAgent.set(agent, [it]);
  }

  const agents = Array.from(byAgent.entries()).sort(
    (a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]),
  );

  return agents
    .map(([agent, list]) => {
      list.sort(
        (a, b) =>
          (b.discrepancyPct ?? -1) - (a.discrepancyPct ?? -1) ||
          String(a.data_mrn ?? "").localeCompare(String(b.data_mrn ?? "")) ||
          String(a.numer_mrn ?? "").localeCompare(String(b.numer_mrn ?? "")),
      );

      const rows = list
        .map((e) => {
          const groupParts = [
            e.key?.odbiorca || "-",
            e.key?.kod_towaru ? `kod:${e.key.kod_towaru}` : "",
            e.key?.waluta ? `wal:${e.key.waluta}` : "",
          ].filter(Boolean);
          const groupTitle = groupParts.join(" | ");
          const arrow = renderCoefArrow(e.outlierSide);
          const opisParts = [
            e.data_mrn ?? "-",
            e.nr_sad ? `SAD: ${e.nr_sad}` : "",
            groupTitle,
            `coef=${formatNum(e.coef)}`,
            `limit=${formatNum(e.limit)}`,
          ].filter(Boolean);
          const opis = opisParts.join(" • ");
          return `<tr class="outlier">
            <td class="mono">${escapeHtml(e.numer_mrn ?? "-")}</td>
            <td title="${escapeHtml(opis)}">${escapeHtml(groupTitle || "-")}</td>
            <td class="mono" title="${escapeHtml(opis)}">${arrow}<span style="margin-left:6px;">${escapeHtml(formatPct(e.discrepancyPct))}</span></td>
            <td>${escapeHtml(agent)}</td>
          </tr>`;
        })
        .join("");

      return `
        <details class="accordion validation-agent">
          <summary>
            <span class="mrn-code" title="${escapeHtml(agent)}">${escapeHtml(agent)}</span>
            <span class="badge rounded-pill badge-count">${list.length}</span>
          </summary>
          <div class="accordion-body">
            <table class="table table-sm table-dark table-hover mini-table" style="margin:0">
              <thead>
                <tr>
                  <th>Lista błędów (MRN)</th>
                  <th>Opis</th>
                  <th>Rozbieżność %</th>
                  <th>Agent celny</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </details>
      `;
    })
    .join("");
}

async function refreshValidationDashboardCounts() {
  const period = getValidationPeriodValue();
  if (!period) return;
  const UP = "\u2191";
  const DOWN = "\u2193";
  try {
    const mrn = getValidationMrnFilterValue();
    const dash = await window.api.getValidationDashboard(
      period,
      mrn || undefined,
      getValidationGroupingOptions(),
    );

    const statHigh = els.validationGroups.querySelector(
      "#validation-stat-high",
    ) as HTMLElement | null;
    const statLow = els.validationGroups.querySelector(
      "#validation-stat-low",
    ) as HTMLElement | null;
    const statSingles = els.validationGroups.querySelector(
      "#validation-stat-singles",
    ) as HTMLElement | null;
    const statManual = els.validationGroups.querySelector(
      "#validation-stat-manual",
    ) as HTMLElement | null;
    if (statHigh) statHigh.textContent = String(dash.stats.outliersHigh);
    if (statLow) statLow.textContent = String(dash.stats.outliersLow);
    if (statSingles) statSingles.textContent = String(dash.stats.singles);
    if (statManual) statManual.textContent = String(dash.stats.verifiedManual);

    const byDate = new Map(dash.days.map((d) => [d.date, d] as const));
    for (const el of Array.from(
      els.validationGroups.querySelectorAll(".day-count"),
    )) {
      const span = el as HTMLElement;
      const date = span.dataset.date ?? "";
      const d = byDate.get(date);
      span.textContent = d ? String(d.total) : "0";
    }

    for (const el of Array.from(
      els.validationGroups.querySelectorAll(".day-badge"),
    )) {
      const span = el as HTMLElement;
      const date = span.dataset.date ?? "";
      const field = span.dataset.field ?? "";
      const d = byDate.get(date);
      if (!d) {
        if (field === "high") span.textContent = `${UP} 0`;
        else if (field === "low") span.textContent = `${DOWN} 0`;
        else if (field === "singles") span.textContent = `1x 0`;
        if (field !== "total") span.dataset.zero = "1";
        continue;
      }
      if (field === "high") {
        span.textContent = `${UP} ${d.outliersHigh}`;
        span.dataset.zero = d.outliersHigh ? "0" : "1";
      } else if (field === "low") {
        span.textContent = `${DOWN} ${d.outliersLow}`;
        span.dataset.zero = d.outliersLow ? "0" : "1";
      } else if (field === "singles") {
        span.textContent = `1x ${d.singles}`;
        span.dataset.zero = d.singles ? "0" : "1";
      }
    }
  } catch {
    // ignore
  }
}

async function loadValidationGroupDetails(detailsEl: HTMLDetailsElement) {
  const period = getValidationPeriodValue();
  const keyEncoded = detailsEl.dataset.key ?? "";
  if (!keyEncoded || !period) return;
  if (detailsEl.dataset.loaded === "1") return;
  if (detailsEl.dataset.loading === "1") return;

  const key = decodeKey<ValidationGroupKey>(keyEncoded);
  if (!key) return;

  const body = detailsEl.querySelector(".accordion-body") as HTMLElement | null;
  if (!body) return;

  detailsEl.dataset.loading = "1";
  body.innerHTML = `<div class="muted">Ładowanie...</div>`;

  try {
    const mrn = getValidationMrnFilterValue();
    const res = await window.api.getValidationItems(
      period,
      key,
      mrn || undefined,
      getValidationGroupingOptions(),
    );
    const items = res.items ?? [];

    const keyGrid = `
      <div class="kv-grid">
        ${prettyKeyLine("Odbiorca", res.key.odbiorca)}
        ${prettyKeyLine("Kraj wysylki", res.key.kraj_wysylki)}
        ${prettyKeyLine("Warunki dostawy", res.key.warunki_dostawy)}
        ${prettyKeyLine("Waluta", res.key.waluta)}
        ${prettyKeyLine("Kurs waluty", res.key.kurs_waluty)}
        ${prettyKeyLine("Transport (rodzaj)", res.key.transport_na_granicy_rodzaj)}
        ${prettyKeyLine("Kod towaru", res.key.kod_towaru)}
      </div>
    `;

    const rows =
      items.length === 0
        ? `<div class="muted">Brak pozycji.</div>`
        : `
          <table class="mini-table">
            <thead>
              <tr>
                <th>Data MRN</th>
                <th>Odbiorca</th>
                <th>Numer MRN</th>
                <th class="coef-col">Wspolczynnik</th>
                <th class="manual-col" title="Verified manually">Manual</th>
              </tr>
            </thead>
            <tbody>
              ${items
                .map((it) => {
                  const coef =
                    it.coef == null ? "-" : Number(it.coef).toFixed(6);
                  const rowClass = it.outlier ? ' class="outlier"' : "";
                  const arrow = renderCoefArrow(it.outlierSide);
                  const isVerifiable =
                    it.rowId &&
                    Number.isFinite(it.rowId) &&
                    it.coef != null &&
                    Number.isFinite(it.coef);
                  const disabledAttr = isVerifiable ? "" : " disabled";
                  const checkedAttr = it.verifiedManual ? " checked" : "";
                  const singleTag =
                    !it.verifiedManual &&
                    it.coef != null &&
                    Number.isFinite(it.coef) &&
                    !it.checkable
                      ? '<span class="single-tag" title="Not enough items for IQR on this day">single</span>'
                      : "";
                  return `<tr${rowClass}>
                    <td class="mono">${escapeHtml(it.data_mrn ?? "-")}</td>
                    <td title="${escapeHtml(it.odbiorca ?? "")}">${escapeHtml(it.odbiorca ?? "-")}</td>
                    <td class="mono">${escapeHtml(it.numer_mrn ?? "-")}</td>
                    <td class="mono coef-cell coef-col"><span class="coef-value">${escapeHtml(coef)}</span><span class="coef-meta">${singleTag}${arrow}</span></td>
                    <td class="mono manual-col">
                      <input class="manual-verify" type="checkbox" data-rowid="${escapeHtml(String(it.rowId ?? ""))}"${checkedAttr}${disabledAttr} />
                    </td>
                  </tr>`;
                })
                .join("")}
            </tbody>
          </table>
        `;

    body.innerHTML = `${keyGrid}${rows}`;
    detailsEl.dataset.loaded = "1";

    for (const el of Array.from(body.querySelectorAll("input.manual-verify"))) {
      const input = el as HTMLInputElement;
      input.addEventListener("change", async () => {
        const rowId = Number(input.dataset.rowid);
        if (!Number.isFinite(rowId) || rowId <= 0) return;
        input.disabled = true;
        try {
          await window.api.setValidationManualVerified(rowId, input.checked);
          detailsEl.dataset.loaded = "0";
          await refreshValidationDashboardCounts();
          await loadValidationGroupDetails(detailsEl);
        } catch {
          // ignore
        }
      });
    }
  } catch (e: unknown) {
    body.innerHTML = `<div class="muted">Błąd: ${escapeHtml(errorMessage(e))}</div>`;
  } finally {
    detailsEl.dataset.loading = "0";
  }
}

async function loadValidationDayDetails(detailsEl: HTMLDetailsElement) {
  const period = getValidationPeriodValue();
  const date = detailsEl.dataset.date ?? "";
  if (!date || !period) return;
  if (detailsEl.dataset.loading === "1") return;
  if (detailsEl.dataset.loaded === "1") return;

  const body = detailsEl.querySelector(".accordion-body") as HTMLElement | null;
  if (!body) return;

  const filter = (detailsEl.dataset.filter as ValidationDayFilter) ?? "all";

  detailsEl.dataset.loading = "1";
  body.innerHTML = `<div class="muted">Ładowanie...</div>`;

  try {
    const mrn = getValidationMrnFilterValue();
    const res = await window.api.getValidationDayItems(
      period,
      date,
      filter,
      mrn || undefined,
      getValidationGroupingOptions(),
    );

    const btn = (f: ValidationDayFilter, label: string) =>
      `<button class="btn btn-outline-light btn-sm btn-filter${filter === f ? " active" : ""}" data-filter="${escapeHtml(f)}">${escapeHtml(label)}</button>`;

    const filters = `
      <div class="day-filters">
        ${btn("all", `Wszystkie (${res.totals.all})`)}
        ${btn("outliersHigh", `↑ (${res.totals.outliersHigh})`)}
        ${btn("outliersLow", `↓ (${res.totals.outliersLow})`)}
        ${btn("singles", `single (${res.totals.singles})`)}
      </div>
    `;

    const rows =
      res.items.length === 0
        ? `<div class="muted">Brak pozycji.</div>`
        : `
          <table class="mini-table">
            <thead>
              <tr>
                <th>Data MRN</th>
                <th>Grupa</th>
                <th>Numer MRN</th>
                <th class="coef-col">Wspolczynnik</th>
                <th class="manual-col" title="Verified manually">Manual</th>
              </tr>
            </thead>
            <tbody>
              ${res.items
                .map((it) => {
                  const coef =
                    it.coef == null ? "-" : Number(it.coef).toFixed(6);
                  const rowClass = it.outlier ? ' class="outlier"' : "";
                  const arrow = renderCoefArrow(it.outlierSide);
                  const isVerifiable =
                    it.rowId &&
                    Number.isFinite(it.rowId) &&
                    it.coef != null &&
                    Number.isFinite(it.coef);
                  const disabledAttr = isVerifiable ? "" : " disabled";
                  const checkedAttr = it.verifiedManual ? " checked" : "";
                  const groupParts = [
                    it.key.odbiorca || "-",
                    it.key.kod_towaru ? `kod:${it.key.kod_towaru}` : "",
                    it.key.waluta ? `wal:${it.key.waluta}` : "",
                  ].filter(Boolean);
                  const groupTitle = groupParts.join(" | ");
                  const singleTag =
                    !it.verifiedManual &&
                    it.coef != null &&
                    Number.isFinite(it.coef) &&
                    !it.checkable
                      ? '<span class="single-tag" title="Not enough items for IQR in this group/day">single</span>'
                      : "";
                  return `<tr${rowClass}>
                    <td class="mono">${escapeHtml(it.data_mrn ?? "-")}</td>
                    <td title="${escapeHtml(groupTitle)}">${escapeHtml(groupTitle)}</td>
                    <td class="mono">${escapeHtml(it.numer_mrn ?? "-")}</td>
                    <td class="mono coef-cell coef-col"><span class="coef-value">${escapeHtml(coef)}</span><span class="coef-meta">${singleTag}${arrow}</span></td>
                    <td class="mono manual-col">
                      <input class="manual-verify" type="checkbox" data-rowid="${escapeHtml(String(it.rowId ?? ""))}"${checkedAttr}${disabledAttr} />
                    </td>
                  </tr>`;
                })
                .join("")}
            </tbody>
          </table>
        `;

    body.innerHTML = `${filters}${rows}`;
    detailsEl.dataset.loaded = "1";

    for (const el of Array.from(body.querySelectorAll("button[data-filter]"))) {
      const b = el as HTMLButtonElement;
      b.addEventListener("click", () => {
        const f = (b.dataset.filter as ValidationDayFilter) ?? "all";
        detailsEl.dataset.filter = f;
        detailsEl.dataset.loaded = "0";
        void loadValidationDayDetails(detailsEl);
      });
    }

    for (const el of Array.from(body.querySelectorAll("input.manual-verify"))) {
      const input = el as HTMLInputElement;
      input.addEventListener("change", async () => {
        const rowId = Number(input.dataset.rowid);
        if (!Number.isFinite(rowId) || rowId <= 0) return;
        input.disabled = true;
        try {
          await window.api.setValidationManualVerified(rowId, input.checked);
          detailsEl.dataset.loaded = "0";
          await refreshValidationDashboardCounts();
          await loadValidationDayDetails(detailsEl);
        } catch {
          // ignore
        }
      });
    }
  } catch (e: unknown) {
    body.innerHTML = `<div class="muted">Błąd: ${escapeHtml(errorMessage(e))}</div>`;
  } finally {
    detailsEl.dataset.loading = "0";
  }
}

async function ensureValidationDefaults() {
  if (els.validationMonth.value && els.validationYear.value) return;
  try {
    const res = await window.api.getValidationDefaultMonth();
    if (res.month && !els.validationMonth.value) els.validationMonth.value = res.month;
    if (res.month && !els.validationYear.value) els.validationYear.value = String(res.month).slice(0, 4);
  } catch {
    // ignore
  }
}

async function ensureExportDefaults() {
  if (els.exportMonth.value && els.exportYear.value) return;
  try {
    const res = await window.api.getValidationDefaultMonth();
    if (res.month && !els.exportMonth.value) els.exportMonth.value = res.month;
    if (res.month && !els.exportYear.value) els.exportYear.value = String(res.month).slice(0, 4);
  } catch {
    // ignore
  }
}

async function refreshValidation() {
  setStatus(els.validationStatus, "");
  updateValidationPeriodUi();
  await ensureValidationDefaults();
  const period = getValidationPeriodValue();
  if (!period) {
    setStatus(els.validationStatus, "Wybierz miesiąc lub rok.");
    return;
  }

  const openDayDates = Array.from(
    els.validationGroups.querySelectorAll("details.validation-day[open]"),
  )
    .map((d) => (d as HTMLDetailsElement).dataset.date ?? "")
    .filter(Boolean);
  const openGroupKeys = Array.from(
    els.validationGroups.querySelectorAll("details.validation-group[open]"),
  )
    .map((d) => (d as HTMLDetailsElement).dataset.key ?? "")
    .filter(Boolean);

  setStatus(els.validationStatus, "Ładowanie...");
  els.validationGroups.innerHTML = "";
  els.validationMeta.textContent = "";
  setBusy(true);

  try {
    const [dash, res, outliers] = await Promise.all([
      window.api.getValidationDashboard(
        period,
        getValidationMrnFilterValue() || undefined,
        getValidationGroupingOptions(),
      ),
      window.api.getValidationGroups(
        period,
        getValidationMrnFilterValue() || undefined,
        getValidationGroupingOptions(),
      ),
      window.api.getValidationOutlierErrors(
        period,
        getValidationMrnFilterValue() || undefined,
        getValidationGroupingOptions(),
      ),
    ]);

    const wynikHtml = renderValidationErrorsByAgent(outliers.items ?? []);

    const dashboardHtml = `
      <div class="section-title">Podsumowanie</div>
      <div class="dash-summary">
        <div class="dash-card">
          <div class="dash-label">Odchylenia ↑</div>
          <div id="validation-stat-high" class="dash-value">${dash.stats.outliersHigh}</div>
        </div>
        <div class="dash-card">
          <div class="dash-label">Odchylenia ↓</div>
          <div id="validation-stat-low" class="dash-value">${dash.stats.outliersLow}</div>
        </div>
        <div class="dash-card">
          <div class="dash-label">Single</div>
          <div id="validation-stat-singles" class="dash-value">${dash.stats.singles}</div>
        </div>
        <div class="dash-card">
          <div class="dash-label">Ręcznie</div>
          <div id="validation-stat-manual" class="dash-value">${dash.stats.verifiedManual}</div>
        </div>
      </div>
      <div class="section-title">Wynik</div>
      ${wynikHtml}
      <div class="section-title">Dni</div>
      ${
        dash.days.length === 0
          ? `<div class="muted" style="padding:6px 2px 10px;">Brak dni.</div>`
          : dash.days
              .slice(0, 5000)
              .map(
                (d) => `
                  <details class="accordion validation-day" data-date="${escapeHtml(d.date)}">
                    <summary>
                      <span class="day-count" data-date="${escapeHtml(d.date)}">${d.total}</span>
                      <span class="mrn-code">${escapeHtml(formatBucketLabel(d.date, d.end))}</span>
                      <span class="badge rounded-pill badge-orange day-badge" data-date="${escapeHtml(d.date)}" data-field="low" data-zero="${d.outliersLow ? "0" : "1"}">↓ ${d.outliersLow}</span>
                      <span class="badge rounded-pill badge-orange day-badge" data-date="${escapeHtml(d.date)}" data-field="high" data-zero="${d.outliersHigh ? "0" : "1"}">↑ ${d.outliersHigh}</span>
                      <span class="badge rounded-pill badge-slate day-badge" data-date="${escapeHtml(d.date)}" data-field="singles" data-zero="${d.singles ? "0" : "1"}">1x ${d.singles}</span>
                    </summary>
                    <div class="accordion-body">
                      <div class="muted">Otwórz, aby załadować paczkę.</div>
                    </div>
                  </details>
                `,
              )
              .join("")
      }
      <div class="section-title">Grupy</div>
    `;

    const groupsHtml =
      res.groups.length === 0
        ? `<div class="muted" style="padding:6px 2px 10px;">Brak grup w tym okresie.</div>`
        : res.groups
            .slice(0, 1000)
            .map((g) => {
              const titleParts = [
                g.key.odbiorca || "-",
                g.key.kod_towaru ? `kod:${g.key.kod_towaru}` : "",
                g.key.waluta ? `wal:${g.key.waluta}` : "",
              ].filter(Boolean);
              const title = titleParts.join(" | ");
              const keyEncoded = encodeKey(g.key);
              return `
                <details class="accordion validation-group" data-key="${escapeHtml(keyEncoded)}">
                  <summary>
                    <span class="mrn-code" title="${escapeHtml(title)}">${escapeHtml(title || "-")}</span>
                    <span class="badge rounded-pill badge-count">${g.count}</span>
                  </summary>
                  <div class="accordion-body">
                    <div class="muted">Otwórz, aby załadować szczegóły.</div>
                  </div>
                </details>
              `;
            })
            .join("");

    els.validationGroups.innerHTML = `${dashboardHtml}${groupsHtml}`;

    for (const el of Array.from(
      els.validationGroups.querySelectorAll("details.accordion"),
    )) {
      const d = el as HTMLDetailsElement;
      d.addEventListener("toggle", () => {
        if (!d.open) return;
        if (d.classList.contains("validation-day")) void loadValidationDayDetails(d);
        else if (d.classList.contains("validation-group")) void loadValidationGroupDetails(d);
      });
    }

    for (const el of Array.from(
      els.validationGroups.querySelectorAll(".day-badge"),
    )) {
      const badge = el as HTMLElement;
      badge.addEventListener("click", (e) => {
        const details = badge.closest("details") as HTMLDetailsElement | null;
        if (!details) return;
        const field = badge.dataset.field ?? "";
        let filter: ValidationDayFilter = "all";
        if (field === "high") filter = "outliersHigh";
        else if (field === "low") filter = "outliersLow";
        else if (field === "singles") filter = "singles";
        details.dataset.filter = filter;
        details.dataset.loaded = "0";
        if (details.open) {
          e.stopPropagation();
          void loadValidationDayDetails(details);
        }
      });
    }

    const openDaySet = new Set(openDayDates);
    for (const el of Array.from(
      els.validationGroups.querySelectorAll("details.validation-day"),
    )) {
      const d = el as HTMLDetailsElement;
      if (!openDaySet.has(d.dataset.date ?? "")) continue;
      d.open = true;
      void loadValidationDayDetails(d);
    }

    const openGroupSet = new Set(openGroupKeys);
    for (const el of Array.from(
      els.validationGroups.querySelectorAll("details.validation-group"),
    )) {
      const d = el as HTMLDetailsElement;
      if (!openGroupSet.has(d.dataset.key ?? "")) continue;
      d.open = true;
      void loadValidationGroupDetails(d);
    }

    setStatus(els.validationStatus, "");
  } catch (e: unknown) {
    setStatus(els.validationStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    setBusy(false);
  }
}

function renderPreviewTable(rows: Array<Record<string, unknown>>): string {
  if (!rows.length) return `<div class="muted">Brak danych.</div>`;
  const cols = Object.keys(rows[0] ?? {});
  const thead = `<tr>${cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("")}</tr>`;
  const body = rows
    .map((r) => {
      const tds = cols
        .map((c) => {
          const v = (r as Record<string, unknown>)[c];
          const text = v == null ? "" : String(v);
          return `<td title="${escapeHtml(text)}">${escapeHtml(text)}</td>`;
        })
        .join("");
      return `<tr>${tds}</tr>`;
    })
    .join("");

  return `
    <div class="table-wrap" style="height:auto; max-height: 360px;">
      <table class="table table-sm table-dark table-hover mini-table" style="margin:0">
        <thead>${thead}</thead>
        <tbody>${body}</tbody>
      </table>
    </div>
  `;
}

async function refreshExportPreview() {
  setStatus(els.exportStatus, "");
  updateExportPeriodUi();
  updateExportFiltersUi();
  await ensureExportDefaults();
  const period = getExportPeriodValue();
  if (!period) {
    setStatus(els.exportStatus, "Wybierz miesi?c lub rok.");
    return;
  }

  const mrn = getExportMrnFilterValue() || undefined;
  els.exportMeta.innerHTML = "";
  els.exportPreview.innerHTML = "";
  setStatus(els.exportStatus, "?adowanie podgl?du...");
  setBusy(true);

  try {
    const res = await window.api.previewValidationExport(
      period,
      mrn,
      getExportGroupingOptions(),
      getExportFilters(),
    );

    if (!res?.ok) {
      const msg = String(res?.error ?? "unknown");
      setStatus(els.exportStatus, `B??d: ${msg}`);
      return;
    }

    const p = res.preview;
    setAvailableExportAgents(p.availableAgents);
    const sheetHtml = (p.sheets ?? [])
      .map((s) => {
        const sectionsHtml = (s.sections ?? [])
          .map((sec) => {
            const meta =
              sec.truncated && sec.totalRows
                ? `<div class="muted" style="margin-top:6px">Pokazano ${sec.rows.length} z ${sec.totalRows} wierszy.</div>`
                : "";
            return `
              <div class="section-title">${escapeHtml(sec.title)}</div>
              ${meta}
              ${renderPreviewTable(sec.rows ?? [])}
            `;
          })
          .join("");

        return `
          <details class="accordion export-sheet" open>
            <summary>
              <span class="mrn-code">${escapeHtml(s.name)}</span>
            </summary>
            <div class="accordion-body">${sectionsHtml}</div>
          </details>
        `;
      })
      .join("");

    const metaRows = (p.meta ?? []).map((kv) => ({
      Key: kv.key,
      Value: kv.value,
    }));
    const metaHtml = `
      <details class="accordion export-sheet">
        <summary><span class="mrn-code">Meta</span></summary>
        <div class="accordion-body">${renderPreviewTable(metaRows)}</div>
      </details>
    `;

    els.exportPreview.innerHTML = `${metaHtml}${sheetHtml}`;
    const groupingLabel =
      Array.from(els.exportGrouping.options).find(
        (o) => String(o.value) === String(p.grouping ?? ""),
      )?.textContent ?? String(p.grouping ?? "");
    const agents = getSelectedExportAgents();
    const agentLabel =
      agents.length === 0
        ? "wszyscy"
        : agents.length === 1
          ? agents[0]
          : `${agents.length} wybranych`;
    els.exportMeta.innerHTML = `
      <div class="meta-lines">
        <div><span class="muted">Okres:</span> <span class="mono">${escapeHtml(p.period)}</span></div>
        <div><span class="muted">Zakres:</span> <span class="mono">${escapeHtml(p.range.start)}?${escapeHtml(p.range.end)}</span></div>
        <div><span class="muted">IQR:</span> ${escapeHtml(String(groupingLabel).trim())} <span class="muted">? Agent:</span> ${escapeHtml(agentLabel)}</div>
      </div>
    `;
    setStatus(els.exportStatus, "");
  } catch (e: unknown) {
    setStatus(els.exportStatus, `B??d: ${errorMessage(e)}`);
  } finally {
    setBusy(false);
  }
}

async function importRaport() {
  els.importBtn.disabled = true;
  setStatus(els.importStatus, "Importowanie…");
  setStatus(els.previewStatus, "");
  els.importProgress.value = 0;
  setStatus(els.importProgressText, "");
  setBusy(true);

  const unsubscribe = window.api.onImportProgress((p) => {
    const percent =
      p.total > 0
        ? Math.max(0, Math.min(100, Math.round((p.current / p.total) * 100)))
        : 0;
    els.importProgress.value = p.stage === "done" ? 100 : percent;
    const counter = p.total > 0 ? ` (${p.current}/${p.total})` : "";
    setStatus(els.importProgressText, `${p.message}${counter}`);
  });

  try {
    const res = await window.api.importRaport();
    if (!res.sourceFile) {
      setStatus(els.importStatus, "Anulowano wybór pliku.");
      return;
    }
    setStatus(
      els.importStatus,
      `OK: zaimportowano ${res.rowCount} wierszy z pliku: ${res.sourceFile}`,
    );
    state.page = 1;
    await refreshMeta();
  } catch (e: unknown) {
    setStatus(els.importStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    unsubscribe();
    els.importBtn.disabled = false;
    setBusy(false);
  }
}

async function clearData() {
  const ok = window.confirm(
    "Na pewno wyczyścić wszystkie dane z bazy? Tej operacji nie można cofnąć.",
  );
  if (!ok) return;

  els.btnClear.disabled = true;
  setStatus(els.settingsStatus, "Czyszczenie danych…");
  setBusy(true);

  try {
    await window.api.clearRaport();
    setStatus(els.settingsStatus, "OK: dane zostały wyczyszczone.");
    state.page = 1;
    state.total = 0;
    await refreshMeta();
    if (state.tab === "preview") await refreshPreview();
  } catch (e: unknown) {
    setStatus(els.settingsStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    els.btnClear.disabled = false;
    setBusy(false);
  }
}

els.tabImportBtn.addEventListener("click", () => setTab("import"));
els.tabPreviewBtn.addEventListener("click", () => setTab("preview"));
els.tabDashboardBtn.addEventListener("click", () => setTab("dashboard"));
els.tabValidationBtn.addEventListener("click", () => setTab("validation"));
els.tabExportBtn.addEventListener("click", () => setTab("export"));
els.tabSettingsBtn.addEventListener("click", () => setTab("settings"));

els.importBtn.addEventListener("click", () => void importRaport());

els.btnPrev.addEventListener("click", () => {
  state.page = Math.max(1, state.page - 1);
  void refreshPreview();
});
els.btnNext.addEventListener("click", () => {
  state.page = state.page + 1;
  void refreshPreview();
});
els.btnRefresh.addEventListener("click", () => void refreshPreview());
els.pageSize.addEventListener("change", () => {
  const v = Number(els.pageSize.value);
  state.pageSize = Number.isFinite(v) && v > 0 ? v : 250;
  state.page = 1;
  void refreshPreview();
});

els.btnShowDb.addEventListener(
  "click",
  () => void window.api.showDbInFolder().catch(() => {}),
);
els.btnClear.addEventListener("click", () => void clearData());

els.btnAgentDzialShow.addEventListener(
  "click",
  () => void window.api.showAgentDzialInFolder().catch(() => {}),
);
els.btnAgentDzialClear.addEventListener("click", async () => {
  const ok = window.confirm("Zresetować słownik agentów (Agent->Dział) do pustego JSON?");
  if (!ok) return;
  els.btnAgentDzialClear.disabled = true;
  setStatus(els.agentDzialStatus, "Resetowanie słownika…");
  setBusy(true);
  try {
    await window.api.clearAgentDzialMap();
    setStatus(els.agentDzialStatus, "OK: słownik zresetowany.");
    await refreshAgentDzialUi();
  } catch (e: unknown) {
    setStatus(els.agentDzialStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    setBusy(false);
    els.btnAgentDzialClear.disabled = false;
  }
});

els.btnMrnRefresh.addEventListener("click", () => void refreshDashboard());
els.btnMrnRebuild.addEventListener("click", async () => {
  els.btnMrnRebuild.disabled = true;
  setStatus(els.dashboardStatus, "Skanowanie...");
  setBusy(true);
  try {
    await window.api.rebuildMrnBatch();
    await refreshDashboard();
    setStatus(els.dashboardStatus, "Gotowe.");
  } catch (e: unknown) {
    setStatus(els.dashboardStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    els.btnMrnRebuild.disabled = false;
    setBusy(false);
  }
});

els.btnValidationRefresh.addEventListener(
  "click",
  () => void refreshValidation(),
);
/*
  const period = getValidationPeriodValue();
  if (!period) {
    setStatus(els.validationStatus, "Wybierz miesi?c lub rok.");
    return;
  }

  els.btnValidationExport.disabled = true;
  setStatus(els.validationStatus, "Eksportowanie do Excel...");
  setBusy(true);
  try {
    const mrn = getValidationMrnFilterValue() || undefined;
    const filters = getValidationExportFilters();
    const res = await window.api.exportValidationXlsx(
      period,
      mrn,
      getValidationGroupingOptions(),
      filters,
    );
    if (res?.ok) {
      const fp = res.filePath ? ` ${res.filePath}` : "";
      setStatus(els.validationStatus, `Zapisano.${fp}`);
    } else if (res?.canceled) {
      setStatus(els.validationStatus, "Anulowano.");
    } else {
      setStatus(
        els.validationStatus,
        `B??d eksportu: ${String(res?.error ?? "unknown")}`,
      );
    }
  } catch (e: unknown) {
    setStatus(els.validationStatus, `B??d eksportu: ${errorMessage(e)}`);
  } finally {
    els.btnValidationExport.disabled = false;
    setBusy(false);
  }
});
*/
els.validationPeriod.addEventListener("change", () => {
  updateValidationPeriodUi();
  void refreshValidation();
});
els.validationMonth.addEventListener("change", () => void refreshValidation());
els.validationYear.addEventListener("change", () => void refreshValidation());
updateValidationPeriodUi();

try {
  const savedGrouping = localStorage.getItem(VALIDATION_GROUPING_STORAGE_KEY);
  if (savedGrouping) {
    els.validationGrouping.value = normalizeValidationDateGrouping(savedGrouping);
  }
} catch {
  // ignore
}

els.validationGrouping.addEventListener("change", () => {
  const v = normalizeValidationDateGrouping(els.validationGrouping.value);
  setValidationGroupingValue(v);
  void refreshValidation();
});

try {
  const saved = localStorage.getItem(VALIDATION_MRN_FILTER_STORAGE_KEY) ?? "";
  if (!els.validationMrnFilter.value) els.validationMrnFilter.value = saved;
} catch {
  // ignore
}
updateValidationMrnFilterUi();

let validationMrnDebounce: number | null = null;
els.validationMrnFilter.addEventListener("input", () => {
  updateValidationMrnFilterUi();
  const v = getValidationMrnFilterValue();
  try {
    localStorage.setItem(VALIDATION_MRN_FILTER_STORAGE_KEY, v);
  } catch {
    // ignore
  }
  if (validationMrnDebounce != null) window.clearTimeout(validationMrnDebounce);
  validationMrnDebounce = window.setTimeout(() => {
    validationMrnDebounce = null;
    void refreshValidation();
  }, 280);
});
els.btnValidationMrnClear.addEventListener("click", () => {
  setValidationMrnFilterValue("");
  void refreshValidation();
});

let exportPreviewDebounce: number | null = null;
const scheduleExportPreviewRefresh = () => {
  if (exportPreviewDebounce != null) window.clearTimeout(exportPreviewDebounce);
  exportPreviewDebounce = window.setTimeout(() => {
    exportPreviewDebounce = null;
    if (state.tab === "export") void refreshExportPreview();
  }, 280);
};

els.btnExportRefresh.addEventListener("click", () => void refreshExportPreview());
els.btnExportDo.addEventListener("click", async () => {
  setStatus(els.exportStatus, "");
  updateExportPeriodUi();
  await ensureExportDefaults();
  const period = getExportPeriodValue();
  if (!period) {
    setStatus(els.exportStatus, "Wybierz miesi?c lub rok.");
    return;
  }

  els.btnExportDo.disabled = true;
  setStatus(els.exportStatus, "Eksportowanie do Excel...");
  setBusy(true);
  try {
    const mrn = getExportMrnFilterValue() || undefined;
    const res = await window.api.exportValidationXlsx(
      period,
      mrn,
      getExportGroupingOptions(),
      getExportFilters(),
    );
    if (res?.ok) {
      const fp = res.filePath ? ` ${res.filePath}` : "";
      setStatus(els.exportStatus, `Zapisano.${fp}`);
    } else if (res?.canceled) {
      setStatus(els.exportStatus, "Anulowano.");
    } else {
      setStatus(
        els.exportStatus,
        `B??d eksportu: ${String(res?.error ?? "unknown")}`,
      );
    }
  } catch (e: unknown) {
    setStatus(els.exportStatus, `B??d eksportu: ${errorMessage(e)}`);
  } finally {
    els.btnExportDo.disabled = false;
    setBusy(false);
  }
});

els.exportPeriod.addEventListener("change", () => {
  updateExportPeriodUi();
  scheduleExportPreviewRefresh();
});
els.exportMonth.addEventListener("change", () => scheduleExportPreviewRefresh());
els.exportYear.addEventListener("change", () => scheduleExportPreviewRefresh());

try {
  const savedGrouping = localStorage.getItem(EXPORT_GROUPING_STORAGE_KEY);
  if (savedGrouping) {
    els.exportGrouping.value = normalizeValidationDateGrouping(savedGrouping);
  }
} catch {
  // ignore
}
els.exportGrouping.addEventListener("change", () => {
  const v = normalizeValidationDateGrouping(els.exportGrouping.value);
  setExportGroupingValue(v);
  scheduleExportPreviewRefresh();
});

try {
  const saved = localStorage.getItem(EXPORT_MRN_FILTER_STORAGE_KEY) ?? "";
  if (!els.exportMrnFilter.value) els.exportMrnFilter.value = saved;
} catch {
  // ignore
}

let exportMrnDebounce: number | null = null;
els.exportMrnFilter.addEventListener("input", () => {
  const v = getExportMrnFilterValue();
  try {
    localStorage.setItem(EXPORT_MRN_FILTER_STORAGE_KEY, v);
  } catch {
    // ignore
  }
  if (exportMrnDebounce != null) window.clearTimeout(exportMrnDebounce);
  exportMrnDebounce = window.setTimeout(() => {
    exportMrnDebounce = null;
    scheduleExportPreviewRefresh();
  }, 280);
});

updateExportFiltersUi();
els.exportFilterImporter.addEventListener("input", () => {
  updateExportFiltersUi();
  scheduleExportPreviewRefresh();
});
els.exportAgentBtn.addEventListener("click", () => {
  if (els.exportAgentBtn.disabled) return;
  const open = els.exportAgentPopover.classList.contains("hidden");
  setExportAgentPopoverOpen(open);
});
els.exportAgentSearch.addEventListener("input", () => renderExportAgentList());
els.btnExportAgentClear.addEventListener("click", () => {
  exportAgentSelectedKeys.clear();
  updateExportAgentUi();
  renderExportAgentList();
  updateExportFiltersUi();
  scheduleExportPreviewRefresh();
});
els.exportAgentList.addEventListener("change", (e) => {
  const target = e.target as HTMLInputElement | null;
  if (!target || target.tagName !== "INPUT" || target.type !== "checkbox") return;
  const key = String(target.dataset.key ?? "").trim();
  if (!key) return;
  if (target.checked) exportAgentSelectedKeys.add(key);
  else exportAgentSelectedKeys.delete(key);
  updateExportAgentUi();
  updateExportFiltersUi();
  scheduleExportPreviewRefresh();
});
els.exportFilterDzial.addEventListener("input", () => {
  updateExportFiltersUi();
  scheduleExportPreviewRefresh();
});
els.btnExportFiltersClear.addEventListener("click", () => {
  clearExportFilters();
  scheduleExportPreviewRefresh();
});

if (!updateStatusUnsub) {
  updateStatusUnsub = window.api.onUpdateStatus((s) => {
    lastUpdateStatus = s;
    renderUpdateStatus(s);
  });
}

void refreshMeta();
void checkForUpdatesAndBlock();
void refreshAppVersion();
