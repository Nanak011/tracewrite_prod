// Formatter page logic
let currentProjectId = null;
let formattedDocId = null;
let socket = null;

// Null-safe DOM readers. If an element is missing (e.g. the deployed
// formatter.html is out of sync with this script - which is exactly what
// happens when only one of the two files gets redeployed), these fall
// back to a sane default instead of throwing and aborting the whole
// formatting run.
function getChecked(id, fallback = true) {
  const el = document.getElementById(id);
  if (!el) {
    console.warn(`formatter.js: expected checkbox #${id} not found in HTML, using default (${fallback})`);
    return fallback;
  }
  return el.checked;
}

function getValue(id, fallback = '') {
  const el = document.getElementById(id);
  if (!el) {
    console.warn(`formatter.js: expected field #${id} not found in HTML, using default (${fallback})`);
    return fallback;
  }
  return el.value;
}

function setChecked(id, value) {
  const el = document.getElementById(id);
  if (el) el.checked = value;
}

function setValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value;
}

async function init() {
  const user = await getCurrentUser();
  if (!user) {
    window.location.href = '/login';
    return;
  }

  const urlParams = new URLSearchParams(window.location.search);
  currentProjectId = urlParams.get('projectId');

  if (!currentProjectId) {
    alert('No project specified');
    window.location.href = '/dashboard';
    return;
  }

  try {
    const response = await API.get(`/api/editor/${currentProjectId}/document`);
    document.getElementById('projectTitle').textContent =
      `Formatting: ${response.document.title || 'Untitled'}`;
  } catch (err) {
    console.error('Failed to load project', err);
  }

  socket = io();
  socket.emit('join_user_room', { userId: user.id });

  socket.on('format_progress', handleFormatProgress);
  socket.on('format_complete', handleFormatComplete);
  socket.on('format_error', handleFormatError);

  document.getElementById('startFormatBtn').addEventListener('click', startFormatting);
  document.getElementById('downloadBtn').addEventListener('click', downloadDocument);
  document.getElementById('backBtn').addEventListener('click', () => {
    window.location.href = `/editor?projectId=${currentProjectId}`;
  });
  document.getElementById('resetDefaultsBtn').addEventListener('click', resetToDefaults);
}

function resetToDefaults() {
  setValue('fontName', 'Times New Roman');
  setValue('bodyFontSize', '12');
  setValue('headingFontName', '');

  setValue('heading1Size', '16');
  setValue('heading2Size', '14');
  setValue('heading3Size', '13');

  setValue('heading1Color', '');
  setValue('heading2Color', '');

  setValue('lineSpacing', '1.5');
  setValue('paragraphSpacing', '6');

  setValue('bodyAlignment', 'justify');
  setChecked('optFirstLineIndent', true);

  setChecked('optTOC', true);
  setChecked('optNlpFallback', true);

  clearWarnings();
  addStatusLog('Settings reset to defaults', 'complete');
}

function addStatusLog(message, type = 'info') {
  const log = document.getElementById('statusLog');
  const item = document.createElement('div');
  item.className = `status-item ${type}`;
  item.innerHTML = `
    <small class="text-muted">${new Date().toLocaleTimeString()}</small><br>
    ${escapeHtml(message)}
  `;
  log.appendChild(item);
  log.scrollTop = log.scrollHeight;
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function clearWarnings() {
  const panel = document.getElementById('warningsPanel');
  panel.innerHTML = '';
  panel.style.display = 'none';
}

function showWarnings(warnings) {
  const panel = document.getElementById('warningsPanel');
  if (!warnings || warnings.length === 0) {
    clearWarnings();
    return;
  }
  panel.style.display = 'block';
  panel.innerHTML = `
    <h6 class="text-warning">⚠️ Needs your confirmation (${warnings.length})</h6>
    ${warnings.map(w => `<div class="status-item" style="background: rgba(255,193,7,0.2);">${escapeHtml(w)}</div>`).join('')}
  `;
}

function updateProgress(percent) {
  document.getElementById('progressBar').style.width = `${percent}%`;
}

async function startFormatting() {
  const startBtn = document.getElementById('startFormatBtn');
  startBtn.disabled = true;
  startBtn.textContent = 'Formatting...';
  clearWarnings();

  addStatusLog('Starting formatting process...', 'active');
  updateProgress(10);

  try {
    const options = {
      // Document features
      include_toc: getChecked('optTOC', true),
      enable_nlp_fallback: getChecked('optNlpFallback', true),

      // Font settings
      font_name: getValue('fontName', 'Times New Roman'),
      body_font_size_pt: parseInt(getValue('bodyFontSize', '12')),
      heading_font_name: getValue('headingFontName', '') || null,

      // Heading sizes
      heading_1_size_pt: parseInt(getValue('heading1Size', '16')),
      heading_2_size_pt: parseInt(getValue('heading2Size', '14')),
      heading_3_size_pt: parseInt(getValue('heading3Size', '13')),

      // Heading colors
      heading_1_color: getValue('heading1Color', '') || null,
      heading_2_color: getValue('heading2Color', '') || null,

      // Spacing
      line_spacing: parseFloat(getValue('lineSpacing', '1.5')),
      paragraph_space_after_pt: parseInt(getValue('paragraphSpacing', '6')),

      // Alignment
      body_alignment: getValue('bodyAlignment', 'justify'),
      enable_first_line_indent: getChecked('optFirstLineIndent', true),
    };

    addStatusLog('Sending document to formatter...', 'active');
    updateProgress(20);

    const response = await API.post(`/api/formatter/${currentProjectId}/format-live`, {
      options: options,
    });

    if (response.success) {
      addStatusLog('Formatting initiated', 'complete');
    } else {
      throw new Error(response.error || 'Formatting failed');
    }
  } catch (error) {
    console.error('Formatting error:', error);
    addStatusLog(`Error: ${error.message}`, 'error');
    startBtn.disabled = false;
    startBtn.textContent = '✨ Retry Formatting';
    updateProgress(0);
  }
}

function handleFormatProgress(data) {
  const { message, progress, preview } = data;

  addStatusLog(message, 'active');
  updateProgress(progress || 50);

  if (preview) {
    document.getElementById('previewContent').innerHTML = preview;
  }
}

function handleFormatComplete(data) {
  const { documentId, preview, message, warnings } = data;

  addStatusLog(message || 'Formatting complete!', 'complete');
  updateProgress(100);

  formattedDocId = documentId;

  if (preview) {
    document.getElementById('previewContent').innerHTML = preview;
  }

  showWarnings(warnings);

  document.getElementById('startFormatBtn').style.display = 'none';
  document.getElementById('downloadBtn').style.display = 'inline-block';
}

function handleFormatError(data) {
  const { error, details } = data;
  addStatusLog(`Error: ${error}`, 'error');
  if (details) {
    addStatusLog(details, 'error');
  }

  const startBtn = document.getElementById('startFormatBtn');
  startBtn.disabled = false;
  startBtn.textContent = '✨ Retry Formatting';
  updateProgress(0);
}

async function downloadDocument() {
  if (!formattedDocId) {
    alert('No formatted document available');
    return;
  }

  try {
    addStatusLog('Preparing download...', 'active');
    window.location.href = `/api/formatter/${currentProjectId}/download/${formattedDocId}`;
    addStatusLog('Download started', 'complete');
  } catch (error) {
    console.error('Download error:', error);
    addStatusLog(`Download failed: ${error.message}`, 'error');
  }
}

init();