// =====================================================================
//  IndexedDB Storage Layer
// =====================================================================
const IDB_NAME = 'police_files_idb';
const IDB_VERSION = 1;
const IDB_STORE = 'files';
const LEGACY_KEY = 'police_files_db'; // old localStorage key (for one-time migration)

let _dbPromise = null;

function getDB() {
    if (_dbPromise) return _dbPromise;
    _dbPromise = new Promise((resolve, reject) => {
        if (!('indexedDB' in window)) {
            reject(new Error('IndexedDB is not supported in this browser.'));
            return;
        }
        const req = indexedDB.open(IDB_NAME, IDB_VERSION);
        req.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(IDB_STORE)) {
                const store = db.createObjectStore(IDB_STORE, { keyPath: 'id' });
                store.createIndex('type', 'type', { unique: false });
                store.createIndex('date', 'date', { unique: false });
            }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
        req.onblocked = () => reject(new Error('IndexedDB blocked — close other tabs of this app.'));
    });
    return _dbPromise;
}

function idbGetAll() {
    return getDB().then(db => new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, 'readonly');
        const req = tx.objectStore(IDB_STORE).getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => reject(req.error);
    }));
}

function idbGet(id) {
    return getDB().then(db => new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, 'readonly');
        const req = tx.objectStore(IDB_STORE).get(id);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
    }));
}

function idbPut(record) {
    return getDB().then(db => new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, 'readwrite');
        tx.objectStore(IDB_STORE).put(record);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error);
    }));
}

function idbDelete(id) {
    return getDB().then(db => new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, 'readwrite');
        tx.objectStore(IDB_STORE).delete(id);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    }));
}

// --- One-time migration from localStorage to IndexedDB ---
async function migrateFromLocalStorage() {
    const raw = localStorage.getItem(LEGACY_KEY);
    if (!raw) return 0;
    try {
        const records = JSON.parse(raw);
        if (!Array.isArray(records) || !records.length) {
            localStorage.removeItem(LEGACY_KEY);
            return 0;
        }
        for (const rec of records) {
            if (rec && rec.id) await idbPut(rec);
        }
        localStorage.removeItem(LEGACY_KEY);
        console.log(`[migration] Moved ${records.length} records from localStorage → IndexedDB`);
        return records.length;
    } catch (e) {
        console.error('[migration] Failed:', e);
        return 0;
    }
}

// =====================================================================
//  Public Storage API (async)
// =====================================================================
async function getSavedFiles() {
    try {
        return await idbGetAll();
    } catch (e) {
        console.error('getSavedFiles failed:', e);
        return [];
    }
}

async function saveFileRecord(record, suppressAlert = false) {
    try {
        await idbPut(record);
        if (!suppressAlert) alert('File saved securely!');
        return true;
    } catch (e) {
        console.error('saveFileRecord failed:', e);
        if (!suppressAlert) {
            alert('Could not save file: ' + (e.message || e));
        }
        return false;
    }
}

async function deleteFileRecord(id) {
    if (!confirm('Are you sure you want to delete this file?')) return;
    try {
        await idbDelete(id);
        await renderSavedFiles();
    } catch (e) {
        console.error(e);
        alert('Delete failed: ' + (e.message || e));
    }
}

// =====================================================================
//  Tab Switching Logic
// =====================================================================
const btnMissing = document.getElementById('btn-missing');
const btnUd = document.getElementById('btn-ud');
const btnSaved = document.getElementById('btn-saved');
const formMissing = document.getElementById('missing-form');
const formUd = document.getElementById('ud-form');
const sectionSaved = document.getElementById('saved-files-section');

function switchTab(showMissing, showUd, showSaved) {
    btnMissing.classList.toggle('active', showMissing);
    btnUd.classList.toggle('active', showUd);
    if (btnSaved) btnSaved.classList.toggle('active', showSaved);

    formMissing.style.display = showMissing ? 'block' : 'none';
    formUd.style.display = showUd ? 'block' : 'none';
    if (sectionSaved) sectionSaved.style.display = showSaved ? 'block' : 'none';
}

btnMissing.addEventListener('click', () => switchTab(true, false, false));
btnUd.addEventListener('click', () => switchTab(false, true, false));

if (btnSaved) {
    btnSaved.addEventListener('click', async () => {
        switchTab(false, false, true);
        await renderSavedFiles();
    });
}

// =====================================================================
//  Form Serialization Helpers
// =====================================================================
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function serializeForm(formId) {
    const form = document.getElementById(formId);
    const data = {};
    const inputs = form.querySelectorAll('input, select, textarea');
    inputs.forEach(input => {
        if (input.type === 'radio') {
            if (input.checked) data[input.name] = input.value;
        } else if (input.type !== 'file') {
            data[input.id] = input.value;
        }
    });
    return data;
}

function deserializeForm(formId, data) {
    const form = document.getElementById(formId);
    Object.keys(data).forEach(key => {
        const input = document.getElementById(key);
        if (input) {
            input.value = data[key];
        } else {
            const radios = form.querySelectorAll(`input[name="${key}"]`);
            radios.forEach(r => { if (r.value === data[key]) r.checked = true; });
        }
    });
}

// =====================================================================
//  Search / Filter / Render
// =====================================================================
function escapeHtml(str) {
    return String(str == null ? '' : str).replace(/[&<>"']/g, c => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
}

function buildSearchIndex(file) {
    const parts = [
        file.title || '',
        file.type === 'missing' ? 'missing person mp' : 'ud unidentified dead body',
        file.date || ''
    ];
    const fd = file.formData || {};
    Object.keys(fd).forEach(k => {
        if (fd[k] !== undefined && fd[k] !== null && fd[k] !== '') {
            parts.push(String(fd[k]));
        }
    });
    return parts.join(' ').toLowerCase();
}

function getFilteredFiles(allFiles) {
    const searchEl = document.getElementById('saved-search');
    const typeEl = document.getElementById('saved-filter-type');
    const query = (searchEl ? searchEl.value : '').trim().toLowerCase();
    const type = typeEl ? typeEl.value : 'all';

    let files = allFiles.slice().reverse(); // newest first

    if (type !== 'all') files = files.filter(f => f.type === type);

    if (query) {
        const terms = query.split(/\s+/).filter(Boolean);
        files = files.filter(f => {
            const hay = buildSearchIndex(f);
            return terms.every(t => hay.includes(t));
        });
    }
    return files;
}

async function renderSavedFiles() {
    const list = document.getElementById('saved-files-list');
    if (!list) return;

    const allFiles = await getSavedFiles();
    const files = getFilteredFiles(allFiles);

    const countEl = document.getElementById('saved-count');
    if (countEl) {
        countEl.textContent = allFiles.length === 0
            ? ''
            : `Showing ${files.length} of ${allFiles.length} saved file${allFiles.length === 1 ? '' : 's'}`;
    }

    if (files.length === 0) {
        list.innerHTML = allFiles.length === 0
            ? '<p class="empty-msg">No saved files found. Your saved drafts will appear here.</p>'
            : '<p class="empty-msg">No files match your search.</p>';
        return;
    }

    list.innerHTML = files.map(file => {
        const isMissing = file.type === 'missing';
        const badgeClass = isMissing ? 'badge-missing' : 'badge-ud';
        const badgeText = isMissing ? 'MP' : 'UD';
        const typeName = isMissing ? 'Missing Person' : 'Unidentified Body';
        const rawTitle = (file.title || '').replace(/^(Missing:|UD Case:)\s*/i, '') || file.title || 'Untitled';

        return `
        <div class="saved-row">
            <div class="saved-row-main">
                <span class="saved-badge ${badgeClass}" title="${typeName}">${badgeText}</span>
                <span class="saved-row-title" title="${escapeHtml(file.title || '')}">${escapeHtml(rawTitle)}</span>
                <span class="saved-row-meta">${escapeHtml(typeName)} &middot; ${escapeHtml(file.date || '')}</span>
            </div>
            <div class="saved-row-actions">
                <button class="btn-action btn-edit" onclick="loadSavedFile('${file.id}')">Edit</button>
                <button class="btn-action btn-duplicate" style="background:#10b981;color:#fff;" onclick="downloadSavedDoc('${file.id}')">DOC</button>
                <button class="btn-action btn-delete" onclick="deleteFileRecord('${file.id}')">Delete</button>
            </div>
        </div>`;
    }).join('');
}

// --- Wire up the search box + type filter (debounced) ---
(function initSavedFileSearch() {
    const searchEl = document.getElementById('saved-search');
    const typeEl = document.getElementById('saved-filter-type');

    if (searchEl) {
        let timer;
        searchEl.addEventListener('input', () => {
            clearTimeout(timer);
            timer = setTimeout(() => renderSavedFiles(), 120);
        });
        searchEl.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                searchEl.value = '';
                renderSavedFiles();
            }
        });
    }

    if (typeEl) typeEl.addEventListener('change', () => renderSavedFiles());
})();

async function loadSavedFile(id) {
    const file = await idbGet(id);
    if (!file) {
        alert('File not found.');
        return;
    }

    const isMissing = file.type === 'missing';
    const formId = isMissing ? 'missing-form' : 'ud-form';
    const idInput = document.getElementById(isMissing ? 'missing-record-id' : 'ud-record-id');
    const preview = document.getElementById(isMissing ? 'image-preview' : 'ud-image-preview');

    switchTab(isMissing, !isMissing, false);

    if (idInput) idInput.value = file.id;
    deserializeForm(formId, file.formData);

    if (file.imageBase64 && preview) {
        preview.src = file.imageBase64;
        preview.style.display = 'block';
        preview.dataset.savedBase64 = file.imageBase64;
    } else if (preview) {
        preview.src = '';
        preview.style.display = 'none';
        delete preview.dataset.savedBase64;
    }
}

// =====================================================================
//  Data Mapping for Saved DOC Downloads
// =====================================================================
function mapMissingData(fd) {
    return {
        ps: fd['ps'] || '', pdType: fd['pd-type'] || '', pdName: fd['pd-name'] || '',
        gdeNo: fd['gde-no'] || '', gdeDate: formatDate(fd['gde-date']),
        caseNo: fd['case-no'] || '', caseDate: formatDate(fd['case-date']), caseSection: fd['case-section'] || '',
        name: fd['missing-name'] || '', alias: fd['alias'] || '', nickname: fd['nickname'] || '',
        guardianType: fd['guardian-type'] || '', guardianName: fd['guardian-name'] || '', address: fd['address'] || '',
        missingDate: formatDate(fd['missing-date']), missingTime: fd['missing-time'] || '',
        age: [fd['age-year'] ? fd['age-year'] + ' Years' : '', fd['age-month'] ? fd['age-month'] + ' Months' : ''].filter(Boolean).join(' '),
        sex: fd['sex'] || '',
        height: [fd['height-feet'] ? fd['height-feet'] + ' Feet' : '', fd['height-inch'] ? fd['height-inch'] + ' Inch' : ''].filter(Boolean).join(' '),
        complexion: fd['complexion'] || '', language: fd['language'] || '', build: fd['build'] || '',
        eyes: fd['eyes'] || '', hair: fd['hair'] || '', mental: fd['mental'] || '', apparel: fd['apparel'] || '',
        specialMarks: fd['special-marks'] || ''
    };
}

function mapUdData(fd) {
    return {
        ps: fd['ud-ps'] || '', pdType: fd['ud-pd-type'] || '', pdName: fd['ud-pd-name'] || '',
        caseNo: fd['ud-case-no'] || '', caseDate: formatDate(fd['ud-case-date']),
        tracingDate: formatDate(fd['ud-tracing-date']), tracingTime: fd['ud-tracing-time'] || '', tracingPlace: fd['ud-tracing-place'] || '',
        age: fd['ud-age-year'] ? fd['ud-age-year'] + ' Years' : '',
        sex: fd['ud-sex'] || '',
        height: [fd['ud-height-feet'] ? fd['ud-height-feet'] + ' Feet' : '', fd['ud-height-inch'] ? fd['ud-height-inch'] + ' inch' : ''].filter(Boolean).join(' '),
        complexion: fd['ud-complexion'] || '', build: fd['ud-build'] || '', eyes: fd['ud-eyes'] || '',
        hair: fd['ud-hair'] || '', apparel: fd['ud-apparel'] || '', specialMarks: fd['ud-special-marks'] || 'NIL',
        otherInfo: fd['ud-other-info'] || ''
    };
}

async function downloadSavedDoc(id) {
    const file = await idbGet(id);
    if (!file) return;

    const pageBreak = `<div style="page-break-before: always; clear: both;"></div>`;
    let combinedHtml = "";
    let filename = "";

    if (file.type === 'missing') {
        const data = mapMissingData(file.formData);
        combinedHtml = generateTemplate1(data, file.imageBase64) + pageBreak +
            generateTemplate2(data, file.imageBase64) + pageBreak +
            generateTemplate3(data, file.imageBase64) + pageBreak +
            generateTemplate4(data, file.imageBase64);
        filename = `${data.name || 'Missing'}_${data.age || ''}`;
    } else {
        const data = mapUdData(file.formData);
        combinedHtml = generateUdTemplate1(data, file.imageBase64) + pageBreak +
            generateUdTemplate2(data, file.imageBase64);
        filename = `${data.ps || 'UD'} UD Case ${data.caseNo ? data.caseNo.replace(/\//g, '-') : ''}`;
    }

    exportToWord(combinedHtml, filename);
}

// =====================================================================
//  Shared Helper Functions
// =====================================================================
function toBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);

        reader.onload = (event) => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');

                ctx.fillStyle = '#FFFFFF';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(img, 0, 0);

                resolve(canvas.toDataURL('image/jpeg', 0.9));
            };
            img.onerror = error => reject(error);
            img.src = event.target.result;
        };
        reader.onerror = error => reject(error);
    });
}

function formatDate(dateString) {
    if (!dateString) return '';
    const parts = dateString.split('-');
    if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
    return dateString;
}

function extractDistrict(addressStr) {
    if (!addressStr) return '';
    const regex = /dist(?:rict)?\s*[-:.]?\s*([^,]+)/i;
    const match = addressStr.match(regex);
    if (match && match[1]) {
        return match[1].replace(/\s*\bpin\b.*|\s+\d{6}.*/i, '').trim();
    }
    return '';
}

// =====================================================================
//  PDF EXPORT
// =====================================================================
function exportToPdf(htmlContent, filename) {
    if (!htmlContent || !htmlContent.trim()) {
        alert('Nothing to export.');
        return;
    }

    const pageBreakRegex = /<div style="page-break-before:\s*always;\s*clear:\s*both;"><\/div>/i;
    const pages = htmlContent.split(pageBreakRegex).map(p => p.trim()).filter(Boolean);
    const pageHtml = pages.map(p => `<div class="page">${p}</div>`).join('');

    const fullHtml = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>${filename}</title>
<style>
    @page { size: A4 portrait; margin: 15mm; }
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; background: #ffffff; color: #000000; }
    body { font-family: 'Times New Roman', Times, serif; font-size: 11pt; line-height: 1.15; }
    .page { page-break-after: always; break-after: page; padding: 0; margin: 0; }
    .page:last-child { page-break-after: auto; break-after: auto; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 15px; }
    table tr td { line-height: 1.15; padding: 3px 2px; margin: 0; vertical-align: top; }
    img { max-width: 100%; object-fit: contain; }
    table, tr, td { page-break-inside: avoid; break-inside: avoid; }
</style>
</head>
<body>${pageHtml}</body>
</html>`;

    const iframe = document.createElement('iframe');
    Object.assign(iframe.style, {
        position: 'fixed', right: '0', bottom: '0', width: '0', height: '0', border: '0', visibility: 'hidden'
    });
    document.body.appendChild(iframe);

    const blob = new Blob([fullHtml], { type: 'text/html' });
    const url = URL.createObjectURL(blob);

    iframe.onload = () => {
        const doc = iframe.contentWindow.document;
        const imgs = Array.from(doc.querySelectorAll('img'));
        Promise.all(imgs.map(img =>
            img.complete ? Promise.resolve() : new Promise(res => { img.onload = img.onerror = res; })
        )).then(() => {
            setTimeout(() => {
                try {
                    iframe.contentWindow.focus();
                    iframe.contentWindow.print();
                } catch (err) {
                    console.error('Print failed:', err);
                    alert('Could not open the print dialog. Please try again.');
                } finally {
                    setTimeout(() => {
                        URL.revokeObjectURL(url);
                        if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
                    }, 1500);
                }
            }, 300);
        });
    };
    iframe.src = url;
}

// =====================================================================
//  Export HTML string to Word Document
// =====================================================================
function exportToWord(htmlContent, filename) {
    htmlContent = htmlContent.replace(
        /<div style="page-break-before:\s*always;\s*clear:\s*both;"><\/div>/gi,
        '<br clear="all" style="page-break-before: always; mso-special-character: line-break;">'
    );

    let finalDocument = "";
    const base64Regex = /data:(image\/[^;]+);base64,([^"]+)/;
    const match = htmlContent.match(base64Regex);

    const preHtml = `<html xmlns:o='urn:schemas-microsoft-com:office:office'
        xmlns:w='urn:schemas-microsoft-com:office:word'
        xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
        <meta charset='utf-8'>
        <title>Export</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
            @page WordSection1 { size: 595.3pt 841.9pt; margin: 0.5in; }
            div.WordSection1 { page: WordSection1; }
            body { font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.15; }
            table, tr, td, p, div { page-break-inside: avoid; }
        </style>
    </head>
    <body><div class="WordSection1">`;
    const postHtml = "</div></body></html>";

    if (match) {
        const mimeType = match[1];
        const base64Data = match[2];
        const globalBase64Regex = /data:image\/[^;]+;base64,[^"]+/g;
        const modifiedHtml = htmlContent.replace(globalBase64Regex, "embedded_image");
        const fullHtmlDocument = preHtml + modifiedHtml + postHtml;

        finalDocument =
            `MIME-Version: 1.0\n` +
            `Content-Type: multipart/related; boundary="mht-boundary"\n\n` +
            `--mht-boundary\n` +
            `Content-Location: document.html\n` +
            `Content-Type: text/html; charset="utf-8"\n\n` +
            `${fullHtmlDocument}\n\n` +
            `--mht-boundary\n` +
            `Content-Location: embedded_image\n` +
            `Content-Transfer-Encoding: base64\n` +
            `Content-Type: ${mimeType}\n\n` +
            `${base64Data}\n` +
            `--mht-boundary--`;
    } else {
        finalDocument = preHtml + htmlContent + postHtml;
    }

    const blob = new Blob(['\ufeff', finalDocument], { type: 'application/msword' });
    const downloadLink = document.createElement("a");
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = filename.endsWith('.doc') ? filename : filename + '.doc';

    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}

// =====================================================================
//  Init: run one-time localStorage → IndexedDB migration
// =====================================================================
(async function initStorage() {
    try {
        const migrated = await migrateFromLocalStorage();
        if (migrated) console.log(`Migrated ${migrated} records into IndexedDB.`);
    } catch (e) {
        console.error('Storage init failed:', e);
    }
})();