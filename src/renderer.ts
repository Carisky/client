import './index.css';

type TabName = 'import' | 'preview' | 'settings';

const els = {
  tabImportBtn: document.getElementById('tab-btn-import') as HTMLButtonElement,
  tabPreviewBtn: document.getElementById('tab-btn-preview') as HTMLButtonElement,
  tabSettingsBtn: document.getElementById('tab-btn-settings') as HTMLButtonElement,

  tabImport: document.getElementById('tab-import') as HTMLElement,
  tabPreview: document.getElementById('tab-preview') as HTMLElement,
  tabSettings: document.getElementById('tab-settings') as HTMLElement,

  importBtn: document.getElementById('btn-import') as HTMLButtonElement,
  importProgress: document.getElementById('import-progress') as HTMLProgressElement,
  importProgressText: document.getElementById('import-progress-text') as HTMLElement,
  importStatus: document.getElementById('import-status') as HTMLElement,
  previewStatus: document.getElementById('preview-status') as HTMLElement,
  settingsStatus: document.getElementById('settings-status') as HTMLElement,

  meta: document.getElementById('meta') as HTMLElement,

  btnPrev: document.getElementById('btn-prev') as HTMLButtonElement,
  btnNext: document.getElementById('btn-next') as HTMLButtonElement,
  btnRefresh: document.getElementById('btn-refresh') as HTMLButtonElement,
  pageInfo: document.getElementById('page-info') as HTMLElement,
  pageSize: document.getElementById('page-size') as HTMLSelectElement,
  tableHead: document.querySelector('#data-table thead') as HTMLTableSectionElement,
  tableBody: document.querySelector('#data-table tbody') as HTMLTableSectionElement,

  dbPath: document.getElementById('db-path') as HTMLElement,
  btnShowDb: document.getElementById('btn-show-db') as HTMLButtonElement,
  btnClear: document.getElementById('btn-clear') as HTMLButtonElement,
};

const state = {
  tab: 'import' as TabName,
  page: 1,
  pageSize: 250,
  total: 0,
  columns: [] as Array<{ field: string; label: string }>,
};

function setStatus(el: HTMLElement, text: string) {
  el.textContent = text;
}

function errorMessage(e: unknown): string {
  if (e instanceof Error) return e.message;
  try {
    return String(e);
  } catch {
    return 'Nieznany błąd';
  }
}

function setTab(name: TabName) {
  state.tab = name;

  els.tabImportBtn.classList.toggle('active', name === 'import');
  els.tabPreviewBtn.classList.toggle('active', name === 'preview');
  els.tabSettingsBtn.classList.toggle('active', name === 'settings');

  els.tabImportBtn.setAttribute('aria-selected', String(name === 'import'));
  els.tabPreviewBtn.setAttribute('aria-selected', String(name === 'preview'));
  els.tabSettingsBtn.setAttribute('aria-selected', String(name === 'settings'));

  els.tabImport.classList.toggle('hidden', name !== 'import');
  els.tabPreview.classList.toggle('hidden', name !== 'preview');
  els.tabSettings.classList.toggle('hidden', name !== 'settings');

  if (name === 'preview') void refreshPreview();
  if (name === 'settings') void refreshSettings();
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function renderTable(page: { columns: Array<{ field: string; label: string }>; rows: Array<Record<string, string | null>> }) {
  const headers = page.columns;
  const thead = `<tr>${headers.map((c) => `<th title="${escapeHtml(c.label)}">${escapeHtml(c.label)}</th>`).join('')}</tr>`;
  els.tableHead.innerHTML = thead;

  const body = page.rows
    .map((row) => {
      const tds = headers
        .map((c) => {
          const v = row[c.field];
          const text = v == null ? '' : String(v);
          return `<td title="${escapeHtml(text)}">${escapeHtml(text)}</td>`;
        })
        .join('');
      return `<tr>${tds}</tr>`;
    })
    .join('');
  els.tableBody.innerHTML = body;
}

async function refreshMeta() {
  try {
    const meta = await window.api.getRaportMeta();
    if (!meta.importedAt || meta.rowCount === 0) {
      els.meta.textContent = 'Brak zaimportowanych danych';
      return;
    }
    const when = new Date(meta.importedAt).toLocaleString('pl-PL');
    const file = meta.sourceFile ? ` • Plik: ${meta.sourceFile}` : '';
    els.meta.textContent = `Zaimportowano: ${when} • Wiersze: ${meta.rowCount}${file}`;
  } catch {
    els.meta.textContent = 'Nie udało się odczytać informacji o imporcie';
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
  setStatus(els.previewStatus, 'Ładowanie danych…');
  await refreshMeta();

  try {
    const page = await window.api.getRaportPage(state.page, state.pageSize);
    state.total = page.total;
    state.columns = page.columns;
    renderTable(page);
    updatePagination();
    setStatus(els.previewStatus, page.total === 0 ? 'Brak danych do wyświetlenia.' : '');
  } catch (e: unknown) {
    setStatus(els.previewStatus, `Błąd: ${errorMessage(e)}`);
  }
}

async function refreshSettings() {
  setStatus(els.settingsStatus, '');
  try {
    const db = await window.api.getDbInfo();
    els.dbPath.textContent = db.filePath + (db.exists ? '' : ' (nie utworzono)');
  } catch (e: unknown) {
    els.dbPath.textContent = '—';
    setStatus(els.settingsStatus, `Błąd: ${errorMessage(e)}`);
  }
}

async function importRaport() {
  els.importBtn.disabled = true;
  setStatus(els.importStatus, 'Importowanie…');
  setStatus(els.previewStatus, '');
  els.importProgress.value = 0;
  setStatus(els.importProgressText, '');

  const unsubscribe = window.api.onImportProgress((p) => {
    const percent = p.total > 0 ? Math.max(0, Math.min(100, Math.round((p.current / p.total) * 100))) : 0;
    els.importProgress.value = p.stage === 'done' ? 100 : percent;
    const counter = p.total > 0 ? ` (${p.current}/${p.total})` : '';
    setStatus(els.importProgressText, `${p.message}${counter}`);
  });

  try {
    const res = await window.api.importRaport();
    if (!res.sourceFile) {
      setStatus(els.importStatus, 'Anulowano wybór pliku.');
      return;
    }
    setStatus(els.importStatus, `OK: zaimportowano ${res.rowCount} wierszy z pliku: ${res.sourceFile}`);
    state.page = 1;
    await refreshMeta();
  } catch (e: unknown) {
    setStatus(els.importStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    unsubscribe();
    els.importBtn.disabled = false;
  }
}

async function clearData() {
  const ok = window.confirm('Na pewno wyczyścić wszystkie dane z bazy? Tej operacji nie można cofnąć.');
  if (!ok) return;

  els.btnClear.disabled = true;
  setStatus(els.settingsStatus, 'Czyszczenie danych…');

  try {
    await window.api.clearRaport();
    setStatus(els.settingsStatus, 'OK: dane zostały wyczyszczone.');
    state.page = 1;
    state.total = 0;
    await refreshMeta();
    if (state.tab === 'preview') await refreshPreview();
  } catch (e: unknown) {
    setStatus(els.settingsStatus, `Błąd: ${errorMessage(e)}`);
  } finally {
    els.btnClear.disabled = false;
  }
}

els.tabImportBtn.addEventListener('click', () => setTab('import'));
els.tabPreviewBtn.addEventListener('click', () => setTab('preview'));
els.tabSettingsBtn.addEventListener('click', () => setTab('settings'));

els.importBtn.addEventListener('click', () => void importRaport());

els.btnPrev.addEventListener('click', () => {
  state.page = Math.max(1, state.page - 1);
  void refreshPreview();
});
els.btnNext.addEventListener('click', () => {
  state.page = state.page + 1;
  void refreshPreview();
});
els.btnRefresh.addEventListener('click', () => void refreshPreview());
els.pageSize.addEventListener('change', () => {
  const v = Number(els.pageSize.value);
  state.pageSize = Number.isFinite(v) && v > 0 ? v : 250;
  state.page = 1;
  void refreshPreview();
});

els.btnShowDb.addEventListener('click', () => void window.api.showDbInFolder().catch(() => {}));
els.btnClear.addEventListener('click', () => void clearData());

void refreshMeta();
