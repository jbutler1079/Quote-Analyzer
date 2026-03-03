/* ── State ─────────────────────────────────────────────────────────────────── */
const state = {
  caseId: null,
  files: [],      // { name, size, type, file: File }
  plans: [],      // parsed plan objects
  census: { ee: 0, es: 0, ec: 0, ef: 0 },
  recommendations: null,
  sortCol: null,
  sortDir: 'asc',
};

/* ── API helpers ──────────────────────────────────────────────────────────── */
function getApiBase() {
  const url = document.getElementById('apiUrl').value.trim();
  return url.replace(/\/$/, '');
}

function getToken() {
  return document.getElementById('apiToken').value.trim();
}

function authHeaders() {
  return { 'X-API-Token': getToken() };
}

async function apiPost(path, body) {
  const resp = await fetch(`${getApiBase()}${path}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders() },
    body: JSON.stringify(body),
  });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({ error: resp.statusText }));
    throw new Error(err.error || `HTTP ${resp.status}`);
  }
  return resp.json();
}

async function apiPostForm(path, formData) {
  const resp = await fetch(`${getApiBase()}${path}`, {
    method: 'POST',
    headers: authHeaders(),
    body: formData,
  });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({ error: resp.statusText }));
    throw new Error(err.error || `HTTP ${resp.status}`);
  }
  return resp.json();
}

async function apiPostBlob(path, body) {
  const resp = await fetch(`${getApiBase()}${path}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...authHeaders() },
    body: JSON.stringify(body),
  });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({ error: resp.statusText }));
    throw new Error(err.error || `HTTP ${resp.status}`);
  }
  const blob = await resp.blob();
  const cd = resp.headers.get('content-disposition') || '';
  const match = cd.match(/filename="?([^"]+)"?/);
  const filename = match ? match[1] : 'download';
  return { blob, filename };
}

/* ── Drag & Drop ──────────────────────────────────────────────────────────── */
(function initDragDrop() {
  const zone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');

  zone.addEventListener('click', () => fileInput.click());
  zone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });

  zone.addEventListener('dragenter', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', e => {
    if (!zone.contains(e.relatedTarget)) zone.classList.remove('drag-over');
  });
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    addFiles(Array.from(e.dataTransfer.files));
  });

  fileInput.addEventListener('change', () => {
    addFiles(Array.from(fileInput.files));
    fileInput.value = '';
  });
})();

function addFiles(newFiles) {
  const allowed = ['application/pdf',
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'text/csv'];
  const allowedExts = ['pdf', 'xlsx', 'xls', 'csv'];

  newFiles.forEach(f => {
    const ext = f.name.split('.').pop().toLowerCase();
    if (!allowed.includes(f.type) && !allowedExts.includes(ext)) {
      showToast(`Skipped "${f.name}": unsupported type`, 'warning');
      return;
    }
    if (f.size > 50 * 1024 * 1024) {
      showToast(`Skipped "${f.name}": exceeds 50 MB`, 'warning');
      return;
    }
    // Avoid duplicates
    if (state.files.some(sf => sf.name === f.name && sf.size === f.size)) return;
    state.files.push({ name: f.name, size: f.size, type: f.type, file: f });
  });

  renderFileList();
}

function removeFile(idx) {
  state.files.splice(idx, 1);
  renderFileList();
}

function renderFileList() {
  const list = document.getElementById('fileList');
  list.innerHTML = '';
  state.files.forEach((f, i) => {
    const ext = f.name.split('.').pop().toLowerCase();
    const li = document.createElement('li');
    li.className = 'file-item';
    li.innerHTML = `
      <span class="file-item-icon ${ext}">${ext.toUpperCase()}</span>
      <span class="file-item-info">
        <span class="file-item-name">${escHtml(f.name)}</span>
        <span class="file-item-size">${formatBytes(f.size)}</span>
      </span>
      <button class="file-remove" onclick="removeFile(${i})" aria-label="Remove ${escHtml(f.name)}">✕</button>
    `;
    list.appendChild(li);
  });

  document.getElementById('processBtn').disabled = state.files.length === 0;
}

/* ── Process Files ────────────────────────────────────────────────────────── */
async function processFiles() {
  if (state.files.length === 0) return;

  showLoading(true, 'Uploading files…');
  try {
    // 1. Upload
    const formData = new FormData();
    state.files.forEach(f => formData.append('files[]', f.file, f.name));
    const uploadData = await apiPostForm('/upload', formData);
    state.caseId = uploadData.caseId;

    // 2. Parse
    showLoading(true, 'Parsing quote documents…');
    const parseData = await apiPost('/parse', { caseId: state.caseId });
    state.plans = parseData.plans || [];

    if (parseData.warnings && parseData.warnings.length > 0) {
      console.warn('Parse warnings:', parseData.warnings);
    }

    renderPlansTable(state.plans);
    showSection('plansSection', true);
    document.getElementById('plansCount').textContent = state.plans.length;
    document.getElementById('recommendBtn').disabled = state.plans.length === 0;

    showLoading(false);

    if (state.plans.length === 0) {
      showToast('Files uploaded but no plans could be extracted. Check file format.', 'warning');
    } else {
      const warn = parseData.warnings && parseData.warnings.length > 0
        ? ` (${parseData.warnings.length} warning${parseData.warnings.length > 1 ? 's' : ''} — see console)`
        : '';
      showToast(`✓ Extracted ${state.plans.length} plan${state.plans.length !== 1 ? 's' : ''}${warn}`, 'success');
    }
  } catch (err) {
    showLoading(false);
    showToast(`Upload/parse failed: ${err.message}`, 'error');
    console.error(err);
  }
}

/* ── Plans Table ──────────────────────────────────────────────────────────── */
const EDITABLE_COLS = ['carrier', 'planName', 'networkType', 'metalLevel',
  'deductibleIndividual', 'oopMaxIndividual', 'copayPCP', 'premiumEE'];

function renderPlansTable(plans) {
  const tbody = document.getElementById('plansTableBody');
  tbody.innerHTML = '';

  if (plans.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--muted)">No plans extracted</td></tr>';
    return;
  }

  const fmtMoney = v => v != null ? `$${Number(v).toLocaleString()}` : '—';
  const fmtStr   = v => v || '—';

  plans.forEach((plan, rowIdx) => {
    const conf = plan.extractionConfidence || 0;
    const confPct = Math.round(conf * 100);
    const confClass = conf >= 0.7 ? '' : conf >= 0.4 ? 'medium' : 'low';

    const tr = document.createElement('tr');
    tr.setAttribute('data-idx', rowIdx);

    const cells = [
      { col: 'carrier',               val: fmtStr(plan.carrier),              editable: true },
      { col: 'planName',              val: fmtStr(plan.planName),             editable: true },
      { col: 'networkType',           val: fmtStr(plan.networkType),          editable: true },
      { col: 'metalLevel',            val: fmtStr(plan.metalLevel),           editable: true },
      { col: 'deductibleIndividual',  val: fmtMoney(plan.deductibleIndividual), editable: true },
      { col: 'oopMaxIndividual',      val: fmtMoney(plan.oopMaxIndividual),   editable: true },
      { col: 'copayPCP',              val: plan.copayPCP != null ? `$${plan.copayPCP}` : '—', editable: true },
      { col: 'premiumEE',             val: plan.premiumEE != null ? `$${plan.premiumEE.toFixed(2)}` : '—', editable: true },
      {
        col: 'extractionConfidence',
        val: `<div class="confidence-bar-wrap">
          <div class="confidence-bar"><div class="confidence-bar-fill ${confClass}" style="width:${confPct}%"></div></div>
          <span style="font-size:0.75rem;color:var(--muted);min-width:32px">${confPct}%</span>
        </div>`,
        editable: false,
      },
      { col: 'sourceFile', val: `<span style="font-size:0.75rem;color:var(--muted)" title="${escHtml(plan.sourceFile || '')}">${escHtml(truncate(plan.sourceFile || '—', 20))}</span>`, editable: false },
    ];

    cells.forEach(({ col, val, editable }) => {
      const td = document.createElement('td');
      if (editable) {
        td.className = 'editable';
        td.setAttribute('data-col', col);
        td.setAttribute('data-idx', rowIdx);
        td.innerHTML = val;
        td.addEventListener('click', startCellEdit);
      } else {
        td.innerHTML = val;
      }
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  // Attach sort handlers
  document.querySelectorAll('.plans-table thead th.sortable').forEach(th => {
    th.onclick = () => sortTable(th.getAttribute('data-col'));
  });
}

function startCellEdit(e) {
  const td = e.currentTarget;
  if (td.querySelector('input.cell-edit')) return; // already editing

  const col = td.getAttribute('data-col');
  const idx = parseInt(td.getAttribute('data-idx'), 10);
  const currentVal = state.plans[idx][col];
  const rawVal = currentVal != null ? String(currentVal) : '';

  const input = document.createElement('input');
  input.className = 'cell-edit';
  input.type = 'text';
  input.value = rawVal;
  td.innerHTML = '';
  td.appendChild(input);
  input.focus();
  input.select();

  const commit = () => {
    const newVal = input.value.trim();
    // Coerce to number for numeric fields
    const numericCols = ['deductibleIndividual', 'deductibleFamily', 'oopMaxIndividual', 'oopMaxFamily',
      'copayPCP', 'copaySpecialist', 'copayER', 'copayUrgentCare',
      'premiumEE', 'premiumES', 'premiumEC', 'premiumEF',
      'rxTier1', 'rxTier2', 'rxTier3'];
    if (numericCols.includes(col)) {
      const n = parseFloat(newVal.replace(/[$,]/g, ''));
      state.plans[idx][col] = isNaN(n) ? null : n;
    } else {
      state.plans[idx][col] = newVal || null;
    }
    renderPlansTable(state.plans);
    document.getElementById('plansCount').textContent = state.plans.length;
  };

  input.addEventListener('blur', commit);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
    if (e.key === 'Escape') { renderPlansTable(state.plans); }
  });
}

function sortTable(col) {
  if (state.sortCol === col) {
    state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    state.sortCol = col;
    state.sortDir = 'asc';
  }

  state.plans.sort((a, b) => {
    let va = a[col], vb = b[col];
    if (va == null && vb == null) return 0;
    if (va == null) return 1;
    if (vb == null) return -1;
    if (typeof va === 'number' && typeof vb === 'number') {
      return state.sortDir === 'asc' ? va - vb : vb - va;
    }
    va = String(va).toLowerCase();
    vb = String(vb).toLowerCase();
    if (va < vb) return state.sortDir === 'asc' ? -1 : 1;
    if (va > vb) return state.sortDir === 'asc' ? 1 : -1;
    return 0;
  });

  // Update header classes
  document.querySelectorAll('.plans-table thead th').forEach(th => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (th.getAttribute('data-col') === col) {
      th.classList.add(state.sortDir === 'asc' ? 'sort-asc' : 'sort-desc');
    }
  });

  renderPlansTable(state.plans);
}

/* ── Census ───────────────────────────────────────────────────────────────── */
function updateCensus() {
  state.census = {
    ee: parseInt(document.getElementById('eeCount').value, 10) || 0,
    es: parseInt(document.getElementById('esCount').value, 10) || 0,
    ec: parseInt(document.getElementById('ecCount').value, 10) || 0,
    ef: parseInt(document.getElementById('efCount').value, 10) || 0,
  };
  const total = state.census.ee + state.census.es + state.census.ec + state.census.ef;
  document.getElementById('censusTotal').textContent = total;
}

/* ── Recommendations ──────────────────────────────────────────────────────── */
async function getRecommendations() {
  if (!state.caseId) { showToast('Please process files first', 'warning'); return; }
  if (state.plans.length === 0) { showToast('No plans to score', 'warning'); return; }

  showLoading(true, 'Scoring plans…');
  try {
    const data = await apiPost('/recommend', { caseId: state.caseId, census: state.census });
    state.recommendations = data;
    renderRecommendations(data);
    showSection('recommendationsSection', true);
    showSection('outputsSection', true);
    showLoading(false);
    showToast(`✓ Recommendations generated for ${data.recommendations.length} top plans`, 'success');
  } catch (err) {
    showLoading(false);
    showToast(`Recommendations failed: ${err.message}`, 'error');
    console.error(err);
  }
}

function renderRecommendations(data) {
  const container = document.getElementById('recCards');
  container.innerHTML = '';
  const recs = data.recommendations || [];

  recs.forEach(plan => {
    const scoreVal = Math.round((plan.totalScore || 0) * 100);
    const card = document.createElement('div');
    card.className = `rec-card rank-${plan.rank}`;
    card.innerHTML = `
      <div class="rec-rank-badge">${plan.rank}</div>
      <div class="rec-carrier">${escHtml(plan.carrier || 'Unknown Carrier')}</div>
      <div class="rec-plan-name">${escHtml(plan.planName || 'Unknown Plan')}</div>

      <div class="score-section">
        <div class="score-label">
          <span>Overall Score</span>
          <span class="score-value">${scoreVal}<small style="font-size:0.7em;font-weight:400">/100</small></span>
        </div>
        <div class="score-bar">
          <div class="score-bar-fill" style="width:${scoreVal}%"></div>
        </div>
      </div>

      <div class="rec-metrics">
        <div class="rec-metric">
          <div class="rec-metric-label">Network</div>
          <div class="rec-metric-value">${escHtml(plan.networkType || '—')}</div>
        </div>
        <div class="rec-metric">
          <div class="rec-metric-label">Metal Level</div>
          <div class="rec-metric-value">${escHtml(plan.metalLevel || '—')}</div>
        </div>
        <div class="rec-metric">
          <div class="rec-metric-label">Deductible (Ind)</div>
          <div class="rec-metric-value">${plan.deductibleIndividual != null ? '$' + Number(plan.deductibleIndividual).toLocaleString() : '—'}</div>
        </div>
        <div class="rec-metric">
          <div class="rec-metric-label">OOP Max (Ind)</div>
          <div class="rec-metric-value">${plan.oopMaxIndividual != null ? '$' + Number(plan.oopMaxIndividual).toLocaleString() : '—'}</div>
        </div>
        <div class="rec-metric">
          <div class="rec-metric-label">PCP Copay</div>
          <div class="rec-metric-value">${plan.copayPCP != null ? '$' + plan.copayPCP : '—'}</div>
        </div>
        <div class="rec-metric">
          <div class="rec-metric-label">EE Premium/mo</div>
          <div class="rec-metric-value">${plan.premiumEE != null ? '$' + Number(plan.premiumEE).toFixed(2) : '—'}</div>
        </div>
      </div>

      <div class="rec-why">${escHtml(plan.whyRecommended || '')}</div>
    `;
    container.appendChild(card);
  });

  if (recs.length === 0) {
    container.innerHTML = '<p style="color:var(--muted);text-align:center;padding:24px">No recommendations available</p>';
  }
}

/* ── Downloads ────────────────────────────────────────────────────────────── */
async function downloadPPTX() {
  if (!state.caseId) { showToast('No case loaded', 'warning'); return; }
  showLoading(true, 'Building PowerPoint presentation…');
  try {
    const { blob, filename } = await apiPostBlob('/export/pptx', {
      caseId: state.caseId,
      clientName: document.getElementById('clientName').value || 'Client',
      effectiveDate: document.getElementById('effectiveDate').value || '',
    });
    triggerDownload(blob, filename || 'BenefitsAnalysis.pptx');
    showLoading(false);
    showToast('✓ PowerPoint downloaded successfully', 'success');
  } catch (err) {
    showLoading(false);
    showToast(`PPTX export failed: ${err.message}`, 'error');
    console.error(err);
  }
}

async function downloadXLSX() {
  if (!state.caseId) { showToast('No case loaded', 'warning'); return; }
  showLoading(true, 'Building Excel workbook…');
  try {
    const { blob, filename } = await apiPostBlob('/export/xlsx', {
      caseId: state.caseId,
      clientName: document.getElementById('clientName').value || 'Client',
      effectiveDate: document.getElementById('effectiveDate').value || '',
    });
    triggerDownload(blob, filename || 'BenefitsAnalysis.xlsx');
    showLoading(false);
    showToast('✓ Excel file downloaded successfully', 'success');
  } catch (err) {
    showLoading(false);
    showToast(`XLSX export failed: ${err.message}`, 'error');
    console.error(err);
  }
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

/* ── UI Helpers ───────────────────────────────────────────────────────────── */
let toastTimer = null;
function showToast(message, type = 'success') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast ${type}`;
  toast.classList.remove('hidden');
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.add('hidden'), 4500);
}

function showLoading(visible, message = 'Processing…') {
  const overlay = document.getElementById('loadingOverlay');
  const msg = document.getElementById('loadingMessage');
  if (visible) {
    msg.textContent = message;
    overlay.classList.remove('hidden');
  } else {
    overlay.classList.add('hidden');
  }
}

function showSection(id, visible) {
  const el = document.getElementById(id);
  if (el) el.classList.toggle('hidden', !visible);
}

/* ── Utilities ────────────────────────────────────────────────────────────── */
function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function truncate(str, len) {
  return str.length > len ? str.substring(0, len - 1) + '…' : str;
}

/* ── Init ─────────────────────────────────────────────────────────────────── */
(function init() {
  // Initialize census display
  updateCensus();
  // Sections that are hidden until data exists
  showSection('plansSection', false);
  showSection('recommendationsSection', false);
  showSection('outputsSection', false);
})();
