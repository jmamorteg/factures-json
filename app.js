'use strict';

const STORAGE_KEY = 'factures-json:v1';
const state = {
  invoices: loadInvoices(),
  imageFile: null,
  deferredInstallPrompt: null,
};

const $ = (selector) => document.querySelector(selector);
const form = $('#invoiceForm');
const typeInputs = [...document.querySelectorAll('input[name="tipus"]')];

initialize();

function initialize() {
  $('#data').value = todayIso();
  typeInputs.forEach((input) => input.addEventListener('change', updateTypeFields));
  form.addEventListener('submit', saveInvoice);
  $('#resetFormButton').addEventListener('click', resetForm);
  $('#invoiceImage').addEventListener('change', handleImageSelection);
  $('#scanButton').addEventListener('click', runOcr);
  $('#clearImageButton').addEventListener('click', clearImage);
  $('#exportButton').addEventListener('click', exportJson);
  $('#importJson').addEventListener('change', importJson);
  $('#clearAllButton').addEventListener('click', clearAllInvoices);
  $('#invoiceList').addEventListener('click', handleInvoiceAction);
  window.addEventListener('beforeinstallprompt', handleInstallPrompt);
  $('#installButton').addEventListener('click', installApp);

  updateTypeFields();
  renderInvoices();
  registerServiceWorker();
}

function todayIso() {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  return new Date(now.getTime() - offset * 60000).toISOString().slice(0, 10);
}

function uuid() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function loadInvoices() {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function persistInvoices() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.invoices));
}

function selectedType() {
  return typeInputs.find((input) => input.checked)?.value || 'Emesa';
}

function updateTypeFields() {
  const isIssued = selectedType() === 'Emesa';
  $('#partyLabel').textContent = isIssued ? 'Client *' : 'Proveïdor *';
  $('#irpfField').hidden = !isIssued;
  $('#afectacioField').hidden = isIssued;
}

function saveInvoice(event) {
  event.preventDefault();
  hideFormError();

  if (!form.reportValidity()) return;

  const editingId = $('#editingId').value;
  const tipus = selectedType();
  const invoice = {
    id_extern: editingId || uuid(),
    tipus,
    data: $('#data').value,
    num_factura: $('#numFactura').value.trim(),
    descripcio: $('#descripcio').value.trim(),
    nom_alias: $('#nomAlias').value.trim(),
    nif: $('#nif').value.trim().toUpperCase(),
    bi: numberOrNull($('#bi').value),
    iva_pct: numberOrNull($('#ivaPct').value),
    irpf_pct: tipus === 'Emesa' ? numberOrNull($('#irpfPct').value) : null,
    afectacio_pct: tipus === 'Rebuda' ? numberOrNull($('#afectacioPct').value) : null,
    observacions: $('#observacions').value.trim(),
    updated_at: new Date().toISOString(),
  };

  if (!Number.isFinite(invoice.bi) || invoice.bi < 0) {
    showFormError('La base imposable no és vàlida.');
    return;
  }

  const duplicate = state.invoices.find((item) =>
    item.id_extern !== editingId &&
    item.tipus === invoice.tipus &&
    normalize(item.nif) === normalize(invoice.nif) &&
    normalize(item.num_factura) === normalize(invoice.num_factura) &&
    item.data === invoice.data &&
    Number(item.bi) === Number(invoice.bi)
  );

  if (duplicate && !window.confirm('Sembla que aquesta factura ja existeix. Vols guardar-la igualment?')) return;

  const existingIndex = state.invoices.findIndex((item) => item.id_extern === editingId);
  if (existingIndex >= 0) state.invoices[existingIndex] = invoice;
  else state.invoices.unshift(invoice);

  persistInvoices();
  renderInvoices();
  resetForm();
}

function numberOrNull(value) {
  if (value === '' || value == null) return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function normalize(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, '');
}

function resetForm() {
  form.reset();
  $('#editingId').value = '';
  $('#data').value = todayIso();
  typeInputs[0].checked = true;
  updateTypeFields();
  hideFormError();
  $('#form-title').textContent = 'Factura';
  form.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function showFormError(message) {
  const element = $('#formMessage');
  element.textContent = message;
  element.hidden = false;
}

function hideFormError() {
  $('#formMessage').hidden = true;
}

function renderInvoices() {
  const list = $('#invoiceList');
  list.replaceChildren();
  $('#invoiceCount').textContent = String(state.invoices.length);
  $('#emptyState').hidden = state.invoices.length > 0;
  $('#exportButton').disabled = state.invoices.length === 0;
  $('#clearAllButton').disabled = state.invoices.length === 0;

  for (const invoice of state.invoices) {
    const article = document.createElement('article');
    article.className = 'invoice-card';
    article.dataset.id = invoice.id_extern;

    const header = document.createElement('header');
    const text = document.createElement('div');
    const title = document.createElement('h3');
    title.textContent = `${invoice.tipus} · ${invoice.num_factura}`;
    const meta = document.createElement('p');
    meta.textContent = `${formatDate(invoice.data)} · ${invoice.nom_alias}${invoice.nif ? ` · ${invoice.nif}` : ''}`;
    text.append(title, meta);

    const amount = document.createElement('div');
    amount.className = 'invoice-amount';
    amount.textContent = formatMoney(invoice.bi);
    header.append(text, amount);

    const actions = document.createElement('div');
    actions.className = 'invoice-actions';
    actions.innerHTML = `
      <button class="button secondary" type="button" data-action="edit">Editar</button>
      <button class="button danger" type="button" data-action="delete">Esborrar</button>
    `;
    article.append(header, actions);
    list.append(article);
  }
}

function handleInvoiceAction(event) {
  const button = event.target.closest('button[data-action]');
  if (!button) return;
  const card = button.closest('.invoice-card');
  const invoice = state.invoices.find((item) => item.id_extern === card.dataset.id);
  if (!invoice) return;

  if (button.dataset.action === 'edit') editInvoice(invoice);
  if (button.dataset.action === 'delete') deleteInvoice(invoice);
}

function editInvoice(invoice) {
  $('#editingId').value = invoice.id_extern;
  typeInputs.find((input) => input.value === invoice.tipus).checked = true;
  updateTypeFields();
  $('#data').value = invoice.data;
  $('#numFactura').value = invoice.num_factura;
  $('#nomAlias').value = invoice.nom_alias;
  $('#nif').value = invoice.nif || '';
  $('#bi').value = invoice.bi ?? '';
  $('#ivaPct').value = String(invoice.iva_pct ?? 0.21);
  $('#irpfPct').value = String(invoice.irpf_pct ?? 0.15);
  $('#afectacioPct').value = String(invoice.afectacio_pct ?? 1);
  $('#descripcio').value = invoice.descripcio || '';
  $('#observacions').value = invoice.observacions || '';
  $('#form-title').textContent = 'Editar factura';
  form.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function deleteInvoice(invoice) {
  if (!window.confirm(`Esborrar la factura ${invoice.num_factura}?`)) return;
  state.invoices = state.invoices.filter((item) => item.id_extern !== invoice.id_extern);
  persistInvoices();
  renderInvoices();
}

function clearAllInvoices() {
  if (!window.confirm('Esborrar totes les factures guardades en aquest navegador?')) return;
  state.invoices = [];
  persistInvoices();
  renderInvoices();
}

function exportJson() {
  const payload = {
    schema_version: 1,
    export_id: uuid(),
    created_at: new Date().toISOString(),
    factures: state.invoices.map(({ updated_at, ...invoice }) => invoice),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `factures_${todayIso()}_${new Date().toTimeString().slice(0, 5).replace(':', '')}.json`;
  link.click();
  URL.revokeObjectURL(url);
}

async function importJson(event) {
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) return;
  try {
    const payload = JSON.parse(await file.text());
    if (payload.schema_version !== 1 || !Array.isArray(payload.factures)) throw new Error('Format no compatible');
    const existingIds = new Set(state.invoices.map((item) => item.id_extern));
    let imported = 0;
    for (const invoice of payload.factures) {
      if (!invoice.id_extern || existingIds.has(invoice.id_extern)) continue;
      state.invoices.push({ ...invoice, updated_at: new Date().toISOString() });
      existingIds.add(invoice.id_extern);
      imported += 1;
    }
    persistInvoices();
    renderInvoices();
    window.alert(`${imported} factures importades.`);
  } catch (error) {
    window.alert(`No s’ha pogut importar el JSON: ${error.message}`);
  }
}

function handleImageSelection(event) {
  clearImage(false);
  state.imageFile = event.target.files?.[0] || null;
  $('#scanButton').disabled = !state.imageFile;
  if (!state.imageFile) return;
  $('#imagePreview').src = URL.createObjectURL(state.imageFile);
  $('#imagePreview').hidden = false;
  $('#clearImageButton').hidden = false;
  setOcrStatus('Imatge preparada.');
}

function clearImage(resetInput = true) {
  if ($('#imagePreview').src) URL.revokeObjectURL($('#imagePreview').src);
  state.imageFile = null;
  $('#imagePreview').hidden = true;
  $('#imagePreview').removeAttribute('src');
  $('#scanButton').disabled = true;
  $('#clearImageButton').hidden = true;
  $('#ocrDetails').hidden = true;
  $('#ocrText').textContent = '';
  setOcrStatus('');
  if (resetInput) $('#invoiceImage').value = '';
}

async function runOcr() {
  if (!state.imageFile) return;
  if (!window.Tesseract) {
    setOcrStatus('No s’ha pogut carregar el motor OCR. Comprova la connexió.', true);
    return;
  }

  $('#scanButton').disabled = true;
  setOcrStatus('Carregant el reconeixement…');
  let worker;
  try {
    worker = await Tesseract.createWorker('spa+cat', 1, {
      logger: (message) => {
        if (message.status === 'recognizing text') {
          setOcrStatus(`Llegint la factura… ${Math.round((message.progress || 0) * 100)} %`);
        } else if (message.status) {
          setOcrStatus(humanizeOcrStatus(message.status));
        }
      },
    });
    const result = await worker.recognize(state.imageFile);
    const text = result.data.text || '';
    $('#ocrText').textContent = text;
    $('#ocrDetails').hidden = false;
    const suggestions = extractInvoiceFields(text);
    applyOcrSuggestions(suggestions);
    const fields = Object.keys(suggestions).length;
    setOcrStatus(fields ? `OCR completat. ${fields} camps suggerits; revisa’ls.` : 'OCR completat, però no he identificat camps amb prou confiança.');
  } catch (error) {
    console.error(error);
    setOcrStatus(`Error OCR: ${error.message}`, true);
  } finally {
    if (worker) await worker.terminate();
    $('#scanButton').disabled = false;
  }
}

function humanizeOcrStatus(status) {
  const messages = {
    'loading tesseract core': 'Carregant el motor OCR…',
    'initializing tesseract': 'Inicialitzant OCR…',
    'loading language traineddata': 'Carregant idiomes…',
    'initializing api': 'Preparant el reconeixement…',
  };
  return messages[status] || status;
}

function setOcrStatus(message, isError = false) {
  const element = $('#ocrStatus');
  element.textContent = message;
  element.classList.toggle('error', isError);
}

function extractInvoiceFields(rawText) {
  const text = rawText.replace(/\r/g, '');
  const lines = text.split('\n').map((line) => line.trim()).filter(Boolean);
  const result = {};

  const dateLine = findLine(lines, /\b(fecha|data|emisi[oó]|expedici[oó])\b/i);
  const dateMatch = (dateLine || text).match(/\b(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})\b|\b(20\d{2})[\/.\-](\d{1,2})[\/.\-](\d{1,2})\b/);
  if (dateMatch) result.data = normalizeDateMatch(dateMatch);

  const nifMatches = [...text.toUpperCase().matchAll(/\b(?:ES\s*)?([ABCDEFGHJNPQRSUVW]\d{7}[0-9A-J]|\d{8}[A-Z]|[XYZ]\d{7}[A-Z])\b/g)];
  if (nifMatches.length) result.nif = nifMatches[nifMatches.length - 1][1];

  const invoiceLine = findLine(lines, /\b(factura|invoice|n[úu]m(?:ero)?|n[ºo]\.?|ref(?:erencia)?)\b/i);
  if (invoiceLine) {
    const candidate = invoiceLine
      .replace(/.*?\b(factura|invoice|n[úu]m(?:ero)?|n[ºo]\.?|ref(?:erencia)?)\b\s*[:#.-]?\s*/i, '')
      .match(/[A-Z0-9][A-Z0-9_./-]{2,}/i);
    if (candidate) result.num_factura = candidate[0];
  }

  const baseLine = findLine(lines, /\b(base\s+imponible|base\s+imposable|subtotal|neto|net)\b/i);
  const base = baseLine ? lastMoney(baseLine) : null;
  if (base != null) result.bi = base;

  const ivaLine = findLine(lines, /\b(iva|vat)\b/i);
  const ivaMatch = ivaLine?.match(/\b(21|10|4|0)\s*%/);
  if (ivaMatch) result.iva_pct = Number(ivaMatch[1]) / 100;

  const irpfLine = findLine(lines, /\b(irpf|retenci[oó]n|retenci[oó])\b/i);
  const irpfMatch = irpfLine?.match(/\b(15|7|0)\s*%/);
  if (irpfMatch) result.irpf_pct = Number(irpfMatch[1]) / 100;

  const totalLine = [...lines].reverse().find((line) => /\b(total(?:\s+factura)?|a\s+pagar|importe\s+total|total\s+document)\b/i.test(line));
  const total = totalLine ? lastMoney(totalLine) : null;
  if (result.bi == null && total != null && result.iva_pct != null && result.irpf_pct == null) {
    const estimatedBase = total / (1 + result.iva_pct);
    if (estimatedBase > 0) result.bi = roundMoney(estimatedBase);
  }

  return result;
}

function findLine(lines, pattern) {
  return lines.find((line) => pattern.test(line));
}

function normalizeDateMatch(match) {
  let year;
  let month;
  let day;
  if (match[4]) {
    year = match[4]; month = match[5]; day = match[6];
  } else {
    day = match[1]; month = match[2]; year = match[3];
    if (year.length === 2) year = `20${year}`;
  }
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function lastMoney(line) {
  const matches = [...line.matchAll(/(?:€|EUR\s*)?(-?\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|-?\d+(?:[.,]\d{2}))(?:\s*(?:€|EUR))?/gi)];
  if (!matches.length) return null;
  return parseMoney(matches[matches.length - 1][1]);
}

function parseMoney(value) {
  let normalized = String(value).replace(/\s/g, '');
  const comma = normalized.lastIndexOf(',');
  const dot = normalized.lastIndexOf('.');
  if (comma > dot) normalized = normalized.replace(/\./g, '').replace(',', '.');
  else if (dot > comma && comma >= 0) normalized = normalized.replace(/,/g, '');
  else if (comma >= 0) normalized = normalized.replace(',', '.');
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function roundMoney(number) {
  return Math.round((number + Number.EPSILON) * 100) / 100;
}

function applyOcrSuggestions(suggestions) {
  if (suggestions.data) $('#data').value = suggestions.data;
  if (suggestions.num_factura) $('#numFactura').value = suggestions.num_factura;
  if (suggestions.nif) $('#nif').value = suggestions.nif;
  if (suggestions.bi != null) $('#bi').value = suggestions.bi;
  if (suggestions.iva_pct != null) setSelectValue($('#ivaPct'), suggestions.iva_pct);
  if (suggestions.irpf_pct != null) setSelectValue($('#irpfPct'), suggestions.irpf_pct);
}

function setSelectValue(select, value) {
  const target = String(value);
  if ([...select.options].some((option) => option.value === target)) select.value = target;
}

function formatMoney(value) {
  return new Intl.NumberFormat('ca-ES', { style: 'currency', currency: 'EUR' }).format(value || 0);
}

function formatDate(value) {
  if (!value) return '';
  return new Intl.DateTimeFormat('ca-ES').format(new Date(`${value}T12:00:00`));
}

function handleInstallPrompt(event) {
  event.preventDefault();
  state.deferredInstallPrompt = event;
  $('#installButton').hidden = false;
}

async function installApp() {
  if (!state.deferredInstallPrompt) return;
  state.deferredInstallPrompt.prompt();
  await state.deferredInstallPrompt.userChoice;
  state.deferredInstallPrompt = null;
  $('#installButton').hidden = true;
}

function registerServiceWorker() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(console.warn);
}
