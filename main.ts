import { analyze, EDITABLE_FIELDS, groupColor, parseSheet, type Analysis, type CheckedLead, type Lead, type LeadField, type RowIdentity, type Snapshot } from './leads';

type Filter = 'all' | 'duplicate' | 'unique' | 'missing';
type Phase = 'idle' | 'scanning' | 'clear' | 'found' | 'error';

const app = document.querySelector('#app') as HTMLElement;
const checkbar = document.querySelector('#checkbar') as HTMLElement;
const checkLabel = document.querySelector('#check-label') as HTMLElement;
const scan = document.querySelector('#scan') as HTMLElement;
const statsEl = document.querySelector('#stats') as HTMLElement;
const rowsEl = document.querySelector('#rows') as HTMLElement;
const groupList = document.querySelector('#group-list') as HTMLElement;
const showing = document.querySelector('#showing') as HTMLElement;
const sourceName = document.querySelector('#source-name') as HTMLElement;
const sourceMeta = document.querySelector('#source-meta') as HTMLElement;
const liveDot = document.querySelector('#live-dot') as HTMLElement;
const search = document.querySelector('#search') as HTMLInputElement;
const sortToggle = document.querySelector('#sort-toggle') as HTMLButtonElement;
const refreshButton = document.querySelector('#refresh') as HTMLButtonElement;
const fileInput = document.querySelector('#file-input') as HTMLInputElement;
const useLive = document.querySelector('#use-live') as HTMLButtonElement;
const liveBanner = document.querySelector('#live-banner') as HTMLElement;
const loadLiveButton = document.querySelector('#load-live') as HTMLButtonElement;
const dropVeil = document.querySelector('#drop-veil') as HTMLElement;
const drawer = document.querySelector('#drawer') as HTMLDialogElement;
const drawerBody = document.querySelector('#drawer-body') as HTMLElement;
const drawerKicker = document.querySelector('#drawer-kicker') as HTMLElement;
const toastEl = document.querySelector('#toast') as HTMLElement;

const state = {
  mode: 'watch' as 'watch' | 'upload',
  fileName: 'Linkedin Leads updated.xlsx',
  filePath: '',
  modifiedAt: null as string | null,
  leads: [] as Lead[],
  analysis: null as Analysis | null,
  filter: 'all' as Filter,
  query: '',
  duplicatesFirst: true,
  focusGroup: null as number | null,
  checking: false,
  liveDirty: false,
  writing: false,
  drawerLead: null as CheckedLead | null,
};

let runToken = 0;
let toastTimer = 0;
let ignoreEventsUntil = 0;

const FIELD_LABELS: Record<LeadField, string> = {
  date: 'Date',
  contactName: 'Contact name',
  designation: 'Designation',
  school: 'School',
  location: 'Location',
  phone1: 'Phone 1',
  phone2: 'Phone 2',
  phone3: 'Phone 3',
  email1: 'Email 1',
  email2: 'Email 2',
  ownerName: 'Owner',
  linkedinUrl: 'LinkedIn URL',
};

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

function esc(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function safeUrl(value: string): string | null {
  try {
    const url = new URL(value, 'https://www.linkedin.com');
    if (url.protocol === 'http:' || url.protocol === 'https:') return url.href;
  } catch {
    return null;
  }
  return null;
}

function formatWhen(iso: string | null): string {
  if (!iso) return 'Not read yet';
  return new Date(iso).toLocaleString(undefined, {
    day: 'numeric',
    month: 'short',
    hour: 'numeric',
    minute: '2-digit',
  });
}

function setPhase(phase: Phase, message: string): void {
  checkbar.dataset.phase = phase;
  checkLabel.textContent = message;
  checkLabel.classList.remove('tick');
  void checkLabel.offsetWidth;
  checkLabel.classList.add('tick');
  const duplicates = state.analysis?.stats.duplicateGroups ?? 0;
  document.title = duplicates > 0 ? `LDC · ${duplicates} duplicate set${duplicates === 1 ? '' : 's'}` : 'LDC — LinkedIn Duplication Check';
}

function showToast(message: string): void {
  toastEl.textContent = message;
  toastEl.classList.add('show');
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toastEl.classList.remove('show'), 3200);
}

function replayScan(): void {
  scan.classList.remove('run');
  void scan.offsetWidth;
  scan.classList.add('run');
}

function statusLabel(lead: CheckedLead, checking: boolean): string {
  if (checking) return 'Checking';
  if (lead.status === 'duplicate') return `Duplicate · ${lead.duplicateCount}`;
  if (lead.status === 'unique') return 'Unique';
  if (lead.status === 'invalid') return 'Not a LinkedIn URL';
  return 'No URL';
}

function visibleLeads(): CheckedLead[] {
  const analysis = state.analysis;
  if (!analysis) return [];
  const query = state.query.trim().toLowerCase();
  let rows = analysis.leads;
  if (state.filter !== 'all') {
    rows = rows.filter((lead) => (state.filter === 'missing' ? lead.status === 'missing' || lead.status === 'invalid' : lead.status === state.filter));
  }
  if (query) {
    rows = rows.filter((lead) =>
      [lead.contactName, lead.designation, lead.school, lead.location, lead.ownerName, lead.linkedinUrl, lead.email1, lead.email2, lead.phone1, String(lead.row)]
        .join(' ')
        .toLowerCase()
        .includes(query),
    );
  }
  const rank = { duplicate: 0, unique: 1, invalid: 2, missing: 3 };
  return [...rows].sort((a, b) => {
    if (state.duplicatesFirst && rank[a.status] !== rank[b.status]) return rank[a.status] - rank[b.status];
    return a.row - b.row;
  });
}

function renderSource(): void {
  const live = state.mode === 'watch';
  sourceName.textContent = state.fileName || 'Spreadsheet';
  sourceMeta.textContent = live
    ? `Live file · updated ${formatWhen(state.modifiedAt)}`
    : `Dropped file · ${formatWhen(state.modifiedAt)}`;
  sourceMeta.title = state.filePath;
  liveDot.classList.toggle('is-upload', !live);
  liveDot.classList.toggle('is-error', checkbar.dataset.phase === 'error');
  useLive.hidden = live;
  liveBanner.hidden = !(state.liveDirty && !live);
}

function renderStats(): void {
  const stats = state.analysis?.stats;
  const uniqueRows = stats ? stats.withUrl - stats.duplicateRows : 0;
  const missingRows = (stats?.missing ?? 0) + (stats?.invalid ?? 0);
  const cards: Array<{ filter: Filter; value: number; label: string; note: string; className: string }> = [
    { filter: 'all', value: stats?.total ?? 0, label: 'Leads', note: 'Rows in the sheet', className: '' },
    { filter: 'unique', value: uniqueRows, label: 'Unique links', note: `${stats?.distinct ?? 0} distinct profiles`, className: 'good' },
    { filter: 'duplicate', value: stats?.duplicateGroups ?? 0, label: 'Duplicate sets', note: `${stats?.duplicateRows ?? 0} rows share a URL`, className: (stats?.duplicateGroups ?? 0) > 0 ? 'alert' : '' },
    { filter: 'missing', value: missingRows, label: 'Missing URLs', note: stats?.invalid ? `${stats.invalid} not a LinkedIn link` : 'Blank LinkedIn cells', className: '' },
  ];

  statsEl.innerHTML = cards
    .map((card) => {
      const active = state.filter === card.filter ? ' active' : '';
      const pop = !state.checking && card.filter === 'duplicate' && card.value > 0 ? ' pop' : '';
      return `<button type="button" class="stat ${card.className}${active}${pop}" data-filter="${card.filter}" aria-pressed="${state.filter === card.filter}">
        <b data-count="${card.value}">${state.checking ? '–' : card.value}</b>
        <span>${card.label}</span>
        <em>${esc(card.note)}</em>
      </button>`;
    })
    .join('');

  if (!state.checking) animateCounts();
}

function animateCounts(): void {
  statsEl.querySelectorAll<HTMLElement>('[data-count]').forEach((node) => {
    const target = Number(node.dataset.count ?? 0);
    const start = performance.now();
    const from = 0;
    const tick = (now: number) => {
      const progress = Math.min(1, (now - start) / 650);
      const eased = 1 - (1 - progress) ** 3;
      node.textContent = String(Math.round(from + (target - from) * eased));
      if (progress < 1) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  });
}

function linkHtml(url: string): string {
  if (!url) return '<span class="sub">—</span>';
  const href = safeUrl(url.startsWith('http') ? url : `https://${url}`);
  const label = esc(url.replace(/^https?:\/\/(www\.)?/i, ''));
  if (!href) return `<span class="url" title="${esc(url)}">${label}</span>`;
  return `<a class="url" href="${esc(href)}" target="_blank" rel="noopener noreferrer" title="${esc(url)}">${label}</a>`;
}

function renderRows(): void {
  const rows = visibleLeads();
  const total = state.analysis?.leads.length ?? 0;
  showing.textContent = total ? `Showing ${rows.length} of ${total}` : 'No rows yet';

  if (!rows.length) {
    rowsEl.innerHTML = `<tr><td class="empty" colspan="9">${state.analysis ? 'Nothing matches this view.' : 'Waiting for spreadsheet rows.'}</td></tr>`;
    return;
  }

  rowsEl.innerHTML = rows
    .map((lead, index) => {
      const color = lead.groupId ? groupColor(lead.groupId) : '';
      const focused = state.focusGroup != null && lead.groupId === state.focusGroup;
      const dimmed = state.focusGroup != null && !focused;
      const classes = [
        state.checking ? 'is-checking' : `is-${lead.status}`,
        focused ? 'is-focused' : '',
        dimmed ? 'is-dimmed' : '',
      ]
        .filter(Boolean)
        .join(' ');
      const badgeClass = state.checking ? 'badge-checking' : `badge-${lead.status}`;
      const times = !state.checking && lead.status === 'duplicate' ? `<span class="times">seen ${lead.duplicateCount}×</span>` : '';
      return `<tr class="${classes}" data-row="${lead.row}" style="--i:${index}; --group:${color}">
        <td><button type="button" class="badge ${badgeClass}" data-group="${lead.groupId ?? ''}">${esc(statusLabel(lead, state.checking))}</button></td>
        <td>${lead.row}</td>
        <td><span class="who">${esc(lead.contactName || 'Unnamed')}</span>${lead.designation ? `<span class="sub">${esc(lead.designation)}</span>` : ''}</td>
        <td class="col-role">${esc(lead.designation || '—')}</td>
        <td>${esc(lead.school || '—')}</td>
        <td class="col-place">${esc(lead.location || '—')}</td>
        <td class="col-owner">${esc(lead.ownerName || '—')}</td>
        <td>${linkHtml(lead.linkedinUrl)}${times}</td>
        <td class="actions">
          <button type="button" class="mini" data-edit="${lead.row}">Edit</button>
          <button type="button" class="mini remove" data-remove="${lead.row}">Remove</button>
        </td>
      </tr>`;
    })
    .join('');
}

function renderGroups(): void {
  const groups = state.analysis?.groups ?? [];
  if (state.checking) {
    groupList.innerHTML = `<div class="clear-card"><strong>Comparing URLs</strong><span>Matching profiles will land here.</span></div>`;
    return;
  }
  if (!groups.length) {
    groupList.innerHTML = `<div class="clear-card"><strong>All clear</strong><span>No LinkedIn URL appears more than once.</span></div>`;
    return;
  }

  groupList.innerHTML = groups
    .map((group, index) => {
      const names = group.names.map((name) => `<i>${esc(name)}</i>`).join('');
      return `<div role="button" tabindex="0" class="group-card${state.focusGroup === group.id ? ' active' : ''}" data-group="${group.id}" style="--group:${groupColor(group.id)}; --d:${index * 50}ms">
        <strong>Set ${group.id} · ${group.count} rows</strong>
        <span>${linkHtml(group.url)}</span>
        <small>Rows ${group.rowNumbers.join(', ')}</small>
        <div class="names">${names}</div>
      </div>`;
    })
    .join('');
}

function render(): void {
  renderSource();
  renderStats();
  renderRows();
  renderGroups();
}

function summary(analysis: Analysis): string {
  const { stats } = analysis;
  if (stats.duplicateGroups > 0) {
    const sets = `${stats.duplicateGroups} duplicate set${stats.duplicateGroups === 1 ? '' : 's'}`;
    return `${sets}. ${stats.duplicateRows} rows use a LinkedIn URL that also appears somewhere else.`;
  }
  if (stats.withUrl === 0) {
    return `Checked ${stats.total} leads. None of them have a LinkedIn URL yet.`;
  }
  return `Checked ${stats.withUrl} LinkedIn URL${stats.withUrl === 1 ? '' : 's'}. Each profile appears once.`;
}

async function present(leads: Lead[], announce = true, instant = false): Promise<void> {
  const token = ++runToken;
  state.leads = leads;
  state.analysis = analyze(leads);
  state.focusGroup = null;
  if (!instant) {
    state.checking = true;
    app.setAttribute('aria-busy', 'true');
    setPhase('scanning', 'Comparing LinkedIn URLs…');
    render();
    replayScan();
    await delay(880);
    if (token !== runToken) return;
  }
  state.checking = false;
  render();
  const found = (state.analysis?.stats.duplicateGroups ?? 0) > 0;
  setPhase(found ? 'found' : 'clear', state.analysis ? summary(state.analysis) : 'No leads to check.');
  app.removeAttribute('aria-busy');
  if (announce && state.mode === 'watch') {
    showToast(found ? 'Spreadsheet updated. Duplicates found.' : 'Spreadsheet updated. No duplicate LinkedIn URLs.');
  }
}

async function loadLive(announce = false): Promise<void> {
  refreshButton.disabled = true;
  try {
    const response = await fetch('/api/snapshot', { cache: 'no-store' });
    const snapshot = (await response.json()) as Snapshot;
    state.mode = 'watch';
    state.fileName = snapshot.fileName;
    state.filePath = snapshot.filePath;
    state.modifiedAt = snapshot.modifiedAt;
    state.liveDirty = false;
    if (!snapshot.ok) {
      setPhase('error', snapshot.error || 'Could not read the spreadsheet.');
      liveDot.classList.add('is-error');
      showToast(snapshot.error || 'Could not read the spreadsheet.');
      renderSource();
      return;
    }
    await present(snapshot.leads, announce);
  } catch {
    setPhase('error', 'The checker API is not running.');
    showToast('Could not reach the local checker. Start it with npm run dev.');
  } finally {
    refreshButton.disabled = false;
  }
}

async function openDroppedFile(file: File): Promise<void> {
  const extension = file.name.split('.').pop()?.toLowerCase() ?? '';
  if (!['xlsx', 'xlsm', 'xls', 'csv'].includes(extension)) {
    showToast('Use an Excel or CSV spreadsheet.');
    return;
  }
  const leads = parseSheet(await file.arrayBuffer());
  state.mode = 'upload';
  state.fileName = file.name;
  state.filePath = file.name;
  state.modifiedAt = new Date(file.lastModified).toISOString();
  await present(leads, false);
  showToast(`Checked ${file.name}.`);
}

function leadByRow(row: number): CheckedLead | undefined {
  return state.analysis?.leads.find((lead) => lead.row === row);
}

function fact(label: string, value: string): string {
  return `<dt>${label}</dt><dd>${value ? esc(value) : '—'}</dd>`;
}

function identityOf(lead: Lead): RowIdentity {
  return { contactName: lead.contactName, linkedinUrl: lead.linkedinUrl, school: lead.school };
}

function rememberFile(snapshot: Snapshot): void {
  state.mode = 'watch';
  state.fileName = snapshot.fileName;
  state.filePath = snapshot.filePath;
  state.modifiedAt = snapshot.modifiedAt;
  state.liveDirty = false;
}

async function commitChange(path: string, body: unknown, success: string, flashRow: number | null): Promise<void> {
  if (state.mode !== 'watch') {
    showToast('This is a dropped file. Watch the live spreadsheet to change Excel.');
    return;
  }
  ignoreEventsUntil = Date.now() + 5000;
  state.writing = true;
  try {
    const response = await fetch(path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const snapshot = (await response.json()) as Snapshot;
    rememberFile(snapshot);
    if (!snapshot.ok) {
      drawer.close();
      state.drawerLead = null;
      if (snapshot.leads.length) await present(snapshot.leads, false, true);
      else renderSource();
      showToast(snapshot.error || 'Could not update the spreadsheet.');
      return;
    }
    drawer.close();
    state.drawerLead = null;
    await present(snapshot.leads, false, true);
    if (flashRow != null) {
      rowsEl.querySelector(`tr[data-row="${flashRow}"]`)?.classList.add('is-saved');
    }
    showToast(success);
  } catch {
    rowsEl.querySelectorAll('.is-removing').forEach((row) => row.classList.remove('is-removing'));
    showToast('Could not update the spreadsheet. If it is open in Excel, close it and try again.');
  } finally {
    state.writing = false;
  }
}

function openDrawer(lead: CheckedLead): void {
  state.drawerLead = lead;
  const group = lead.groupId ? state.analysis?.groups.find((item) => item.id === lead.groupId) : null;
  drawerKicker.textContent = statusLabel(lead, false);
  const phones = [lead.phone1, lead.phone2, lead.phone3].filter(Boolean).join(' · ');
  const emails = [lead.email1, lead.email2].filter(Boolean).join(' · ');
  drawerBody.innerHTML = `
    <h3>${esc(lead.contactName || 'Unnamed')}</h3>
    <p>${esc([lead.designation, lead.school].filter(Boolean).join(' · ') || 'No designation')}</p>
    <dl class="facts">
      ${fact('Row', String(lead.row))}
      ${fact('Date', lead.date)}
      ${fact('Location', lead.location)}
      ${fact('Owner', lead.ownerName)}
      ${fact('Phone', phones)}
      ${fact('Email', emails)}
      <dt>LinkedIn</dt><dd>${lead.linkedinUrl ? linkHtml(lead.linkedinUrl) : '—'}</dd>
      ${fact('Matched as', lead.normalizedUrl ?? '')}
      ${group ? fact('Also on rows', group.rowNumbers.filter((row) => row !== lead.row).join(', ')) : ''}
    </dl>
    <div class="form-actions">
      <button type="button" data-action="edit">Edit</button>
      <button type="button" class="danger" data-action="remove">Remove</button>
    </div>`;
  if (!drawer.open) drawer.showModal();
}

function openEditor(lead: CheckedLead): void {
  state.drawerLead = lead;
  drawerKicker.textContent = `Edit row ${lead.row}`;
  const fields = EDITABLE_FIELDS.map((field) => {
    const wide = field === 'contactName' || field === 'school' || field === 'linkedinUrl' ? ' wide' : '';
    return `<label class="${wide.trim()}">${FIELD_LABELS[field]}<input name="${field}" value="${esc(lead[field])}" autocomplete="off" /></label>`;
  }).join('');
  drawerBody.innerHTML = `
    <h3>${esc(lead.contactName || 'Unnamed')}</h3>
    <p class="hint">Saving writes this row back into the Excel file.</p>
    <form id="edit-form" class="edit-form">
      ${fields}
      <div class="form-actions wide">
        <button type="submit">Save to spreadsheet</button>
        <button type="button" data-action="cancel">Cancel</button>
        <button type="button" class="danger" data-action="remove">Remove</button>
      </div>
    </form>`;
  if (!drawer.open) drawer.showModal();
}

function openRemove(lead: CheckedLead): void {
  state.drawerLead = lead;
  drawerKicker.textContent = 'Remove lead';
  drawerBody.innerHTML = `
    <h3>Remove ${esc(lead.contactName || 'this lead')}?</h3>
    <p class="hint">Row ${lead.row} will be deleted from the Excel file. Later rows move up, and the duplication check runs again.</p>
    <div class="form-actions">
      <button type="button" class="danger" data-action="confirm-remove">Remove from spreadsheet</button>
      <button type="button" data-action="cancel">Keep</button>
    </div>`;
  if (!drawer.open) drawer.showModal();
}

function focusGroup(groupId: number | null): void {
  state.focusGroup = state.focusGroup === groupId ? null : groupId;
  renderRows();
  renderGroups();
  if (state.focusGroup == null) return;
  const first = rowsEl.querySelector<HTMLElement>('.is-focused');
  first?.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

statsEl.addEventListener('click', (event) => {
  const button = (event.target as HTMLElement).closest<HTMLButtonElement>('[data-filter]');
  if (!button) return;
  const next = button.dataset.filter as Filter;
  state.filter = state.filter === next && next !== 'all' ? 'all' : next;
  renderStats();
  renderRows();
});

rowsEl.addEventListener('click', (event) => {
  const target = event.target as HTMLElement;
  if (target.closest('a')) return;
  const rowElement = target.closest<HTMLElement>('tr');
  const row = rowElement?.dataset.row;
  const lead = row ? leadByRow(Number(row)) : undefined;
  if (target.closest('[data-edit]')) {
    if (lead) openEditor(lead);
    return;
  }
  if (target.closest('[data-remove]')) {
    if (lead) openRemove(lead);
    return;
  }
  const badge = target.closest<HTMLButtonElement>('.badge');
  const groupId = badge?.dataset.group ? Number(badge.dataset.group) : 0;
  if (badge && groupId) {
    focusGroup(groupId);
    return;
  }
  if (lead) openDrawer(lead);
});

drawerBody.addEventListener('click', (event) => {
  const action = (event.target as HTMLElement).closest<HTMLElement>('[data-action]')?.dataset.action;
  const lead = state.drawerLead;
  if (!action || !lead) return;
  if (action === 'edit') openEditor(lead);
  if (action === 'remove') openRemove(lead);
  if (action === 'cancel') openDrawer(lead);
  if (action === 'confirm-remove') {
    if (state.writing) return;
    const row = rowsEl.querySelector(`tr[data-row="${lead.row}"]`);
    row?.classList.add('is-removing');
    const name = lead.contactName || 'the lead';
    void commitChange('/api/leads/delete', { row: lead.row, expect: identityOf(lead) }, `Removed ${name} from the spreadsheet.`, null);
  }
});

drawerBody.addEventListener('submit', (event) => {
  const form = event.target;
  if (!(form instanceof HTMLFormElement) || form.id !== 'edit-form') return;
  event.preventDefault();
  const lead = state.drawerLead;
  if (!lead || state.writing) return;
  const data = new FormData(form);
  const fields = Object.fromEntries(EDITABLE_FIELDS.map((field) => [field, String(data.get(field) ?? '')])) as Record<LeadField, string>;
  const name = fields.contactName || lead.contactName || 'the lead';
  void commitChange('/api/leads/update', { row: lead.row, expect: identityOf(lead), fields }, `Saved ${name} to the spreadsheet.`, lead.row);
});

function activateGroupCard(target: HTMLElement): void {
  const card = target.closest<HTMLElement>('[data-group]');
  if (!card?.dataset.group || target.closest('a')) return;
  state.filter = 'duplicate';
  focusGroup(Number(card.dataset.group));
  renderStats();
}

groupList.addEventListener('click', (event) => activateGroupCard(event.target as HTMLElement));
groupList.addEventListener('keydown', (event) => {
  if (event.key !== 'Enter' && event.key !== ' ') return;
  event.preventDefault();
  activateGroupCard(event.target as HTMLElement);
});

search.addEventListener('input', () => {
  state.query = search.value;
  renderRows();
});

sortToggle.addEventListener('click', () => {
  state.duplicatesFirst = !state.duplicatesFirst;
  sortToggle.setAttribute('aria-pressed', String(state.duplicatesFirst));
  sortToggle.textContent = state.duplicatesFirst ? 'Duplicates first' : 'Sheet order';
  renderRows();
});

refreshButton.addEventListener('click', () => {
  if (state.mode === 'upload') {
    showToast('This view is a dropped file. Watch the live file to re-read the spreadsheet.');
    return;
  }
  void loadLive(false);
});

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  fileInput.value = '';
  if (file) void openDroppedFile(file);
});

useLive.addEventListener('click', () => void loadLive(false));
loadLiveButton.addEventListener('click', () => void loadLive(false));

let dragDepth = 0;
window.addEventListener('dragenter', (event) => {
  if (!event.dataTransfer?.types.includes('Files')) return;
  event.preventDefault();
  dragDepth += 1;
  dropVeil.hidden = false;
});
window.addEventListener('dragover', (event) => {
  if (event.dataTransfer?.types.includes('Files')) event.preventDefault();
});
window.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) dropVeil.hidden = true;
});
window.addEventListener('drop', (event) => {
  if (!event.dataTransfer) return;
  event.preventDefault();
  dragDepth = 0;
  dropVeil.hidden = true;
  const file = event.dataTransfer.files?.[0];
  if (file) void openDroppedFile(file);
});

window.addEventListener('keydown', (event) => {
  if (event.key === '/' && document.activeElement !== search) {
    event.preventDefault();
    search.focus();
  }
});

const events = new EventSource('/api/events');
events.onmessage = () => {
  if (Date.now() < ignoreEventsUntil || state.writing) return;
  if (state.mode === 'upload') {
    state.liveDirty = true;
    renderSource();
    showToast('The live spreadsheet changed.');
    return;
  }
  void loadLive(true);
};

void loadLive(false);
