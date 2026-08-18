// content.js — all_frames: true

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function waitFor(selectorFn, timeout = 12000, interval = 400) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const check = () => {
      const el = selectorFn();
      if (el) { resolve(el); return; }
      if (Date.now() - start > timeout) { reject(new Error('Timeout')); return; }
      setTimeout(check, interval);
    };
    check();
  });
}

function isMainFrame() {
  return !!document.querySelector('#containerextno') || !!document.querySelector('#m_search');
}

function isListFrame() {
  return !!(document.body && document.body.classList.contains('xc-list'));
}

// ── Main frame: fill #containerextno (排柜单号) + search ──────────────────────
async function doSearchInMainFrame(paiGuiNo) {
  const input = await waitFor(() => document.querySelector('#containerextno'), 10000);

  input.focus();
  input.value = '';
  input.dispatchEvent(new Event('input', { bubbles: true }));
  input.dispatchEvent(new Event('change', { bubbles: true }));
  await sleep(300);

  input.value = paiGuiNo;
  input.dispatchEvent(new Event('input', { bubbles: true }));
  input.dispatchEvent(new Event('change', { bubbles: true }));
  await sleep(400);

  const btn = await waitFor(() => document.querySelector('#m_search'), 5000);
  btn.click();
  await sleep(500);
  return { success: true };
}

// A row/cell counts as a downloadable target if its text mentions
// "packing list" (old file naming) or "clearance_doc" (new file naming,
// e.g. "EU_Clearance_Doc..xlsx"). Add more keywords here if TMS renames
// the file again.
function isTargetName(text) {
  const t = text.toLowerCase();
  return t.includes('packing list')
    || t.includes('packing_list')
    || t.includes('clearance_doc')
    || t.includes('clearance doc');
}

// Fallback file, e.g. "惠州市-英国英国063721=清关.zip" — a ZIP whose name
// contains "清关". Only used when no packing list / clearance_doc XLSX
// exists for this container.
function isQingGuanZipName(text) {
  return text.includes('清关');
}

// Last-resort fallback: none of the named patterns above matched anything
// (e.g. "宁波-英国英国064774.zip" — just the container number, no
// "packing list" / "clearance_doc" / "清关" wording at all). At this point
// we no longer try to match by name — any row that is a ZIP is accepted,
// first one found wins.
function isAnyName() {
  return true;
}

function clickDownloadLink(row) {
  const link = row.querySelector('a[href*="/file/"]');
  if (!link) return false;
  const dl = document.createElement('a');
  dl.href = link.href;
  dl.download = '';
  dl.style.display = 'none';
  document.body.appendChild(dl);
  dl.click();
  document.body.removeChild(dl);
  return true;
}

// True if some cell either IS the format code exactly (e.g. a dedicated
// "XLSX"/"ZIP" type column) OR its text ends with ".xlsx"/".zip" (the
// extension embedded in the filename itself). Some rows only expose one
// of these, so both are checked.
function rowHasExtension(cells, ext) {
  const upperExt = ext.toUpperCase();
  const dotExt = '.' + ext.toLowerCase();
  return cells.some(td => {
    const raw = td.textContent.trim();
    return raw.toUpperCase() === upperExt || raw.toLowerCase().endsWith(dotExt);
  });
}

// Scans rows for one whose text matches nameMatch and whose format column
// / filename extension matches ext (e.g. "xlsx", "zip"); clicks its
// download link/name.
async function tryDownloadMatch(rows, nameMatch, ext) {
  for (const row of rows) {
    const cells = [...row.querySelectorAll('td')];
    if (!nameMatch(row.textContent)) continue;

    if (!rowHasExtension(cells, ext)) continue;

    if (clickDownloadLink(row)) {
      await sleep(1000);
      return { success: true };
    }

    const nameCell = cells.find(td => nameMatch(td.textContent));
    if (nameCell) {
      const a = nameCell.querySelector('a');
      if (a) { a.click(); } else { nameCell.click(); }
      await sleep(1000);
      return { success: true };
    }
  }
  return null;
}

// ── List frame: wait for rows + download ──────────────────────────────────────
async function doDownloadInListFrame() {
  // Wait for old rows to clear
  const clearStart = Date.now();
  while (Date.now() - clearStart < 5000) {
    const rows = document.querySelectorAll('tbody tr');
    if (rows.length === 0) break;
    await sleep(300);
  }

  // Wait for new rows to appear (up to 20 seconds)
  await waitFor(() => {
    const rows = document.querySelectorAll('tbody tr');
    return rows.length > 0 ? true : null;
  }, 20000, 500);

  await sleep(600);
  window.scrollTo(0, document.body.scrollHeight);
  await sleep(600);

  const rows = [...document.querySelectorAll('tbody tr')];
  if (rows.length === 0) return { success: false, status: 'No results' };

  // 1st choice: packing list / clearance_doc XLSX
  const xlsxHit = await tryDownloadMatch(rows, isTargetName, 'xlsx');
  if (xlsxHit) return xlsxHit;

  // 2nd choice: "=清关" ZIP, only when no XLSX target was found above
  const zipHit = await tryDownloadMatch(rows, isQingGuanZipName, 'zip');
  if (zipHit) return { success: true, status: 'Downloaded ZIP (清关)' };

  // Last resort: none of the named patterns matched anything at all —
  // just grab the first ZIP in the list, whatever its name is
  // (e.g. a plain "宁波-英国英国064774.zip").
  const anyZipHit = await tryDownloadMatch(rows, isAnyName, 'zip');
  if (anyZipHit) return { success: true, status: 'Downloaded ZIP (fallback, no name match)' };

  return { success: false, status: 'No XLSX' };
}

// ── Message handler ───────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'doSearch' && isMainFrame()) {
    doSearchInMainFrame(msg.container)
      .then(r => sendResponse(r))
      .catch(e => sendResponse({ success: false, status: 'Error: ' + e.message }));
    return true;
  }
  if (msg.action === 'doDownload' && isListFrame()) {
    doDownloadInListFrame()
      .then(r => sendResponse(r))
      .catch(e => sendResponse({ success: false, status: 'Error: ' + e.message }));
    return true;
  }
  if (msg.action === 'hasRows' && isListFrame()) {
    const rows = document.querySelectorAll('tbody tr');
    sendResponse({ hasRows: rows.length > 0 });
    return false;
  }
  if (msg.action === 'ping') {
    sendResponse({
      alive: true,
      isMain: isMainFrame(),
      isList: isListFrame(),
      url: window.location.href,
    });
    return false;
  }
});
