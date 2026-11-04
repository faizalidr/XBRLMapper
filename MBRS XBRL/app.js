/**
 * MBRS Excel Mapper
 *
 * Reads a filled MBRS source .xlsm and copies all cell values into a
 * destination MBRS .xlsm template, matching by sheet name + cell reference.
 *
 * Both files are OOXML (ZIP) archives.  We use JSZip to read/write them
 * and plain DOMParser to manipulate the worksheet XML.
 */

// ─── DOM refs ────────────────────────────────────────────────────────
const sourceInput   = document.getElementById('source-file');
const destInput     = document.getElementById('dest-file');
const sourceCard    = document.getElementById('source-card');
const destCard      = document.getElementById('dest-card');
const sourceName    = document.getElementById('source-name');
const destName      = document.getElementById('dest-name');
const sourceInfo    = document.getElementById('source-info');
const destInfo      = document.getElementById('dest-info');
const runBtn        = document.getElementById('run-btn');
const progressSec   = document.getElementById('progress-section');
const progressBar   = document.getElementById('progress-bar');
const progressText  = document.getElementById('progress-text');
const logSection    = document.getElementById('log-section');
const logContainer  = document.getElementById('log-container');
const resultSection = document.getElementById('result-section');
const resultSummary = document.getElementById('result-summary');
const downloadBtn   = document.getElementById('download-btn');

let sourceFile = null;
let destFile   = null;
let resultBlob = null;

// ─── File input handlers ─────────────────────────────────────────────
sourceInput.addEventListener('change', e => {
    sourceFile = e.target.files[0];
    if (sourceFile) {
        sourceName.textContent = sourceFile.name;
        sourceName.classList.add('active');
        sourceCard.classList.add('loaded');
        sourceInfo.textContent = `Size: ${formatSize(sourceFile.size)}`;
    }
    checkReady();
});

destInput.addEventListener('change', e => {
    destFile = e.target.files[0];
    if (destFile) {
        destName.textContent = destFile.name;
        destName.classList.add('active');
        destCard.classList.add('loaded');
        destInfo.textContent = `Size: ${formatSize(destFile.size)}`;
    }
    checkReady();
});

function checkReady() {
    runBtn.disabled = !(sourceFile && destFile);
}

// ─── Run button ──────────────────────────────────────────────────────
runBtn.addEventListener('click', async () => {
    runBtn.disabled = true;
    resultSection.style.display = 'none';
    progressSec.style.display = 'block';
    logSection.style.display = 'block';
    logContainer.innerHTML = '';

    try {
        await runMapper();
    } catch (err) {
        log(`Fatal error: ${err.message}`, 'error');
        console.error(err);
    }

    runBtn.disabled = false;
});

// ─── Download button ─────────────────────────────────────────────────
downloadBtn.addEventListener('click', () => {
    if (!resultBlob) return;
    const url = URL.createObjectURL(resultBlob);
    const a = document.createElement('a');
    a.href = url;
    // Keep the original destination filename but prefix with "filled_"
    a.download = 'filled_' + (destFile ? destFile.name : 'output.xlsm');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
});

// ─── Main mapping logic ──────────────────────────────────────────────
async function runMapper() {
    // 1. Load both ZIP archives
    setProgress(5, 'Loading source file…');
    log('Loading source file…', 'info');
    const srcZip = await JSZip.loadAsync(sourceFile);

    setProgress(10, 'Loading destination file…');
    log('Loading destination file…', 'info');
    const dstZip = await JSZip.loadAsync(destFile);

    // 2. Parse workbook.xml to get sheet-name → rId mapping
    setProgress(15, 'Parsing workbook structure…');
    const srcSheetMap = await getSheetMap(srcZip);
    const dstSheetMap = await getSheetMap(dstZip);

    log(`Source sheets: ${srcSheetMap.size}`, 'info');
    log(`Destination sheets: ${dstSheetMap.size}`, 'info');

    // 3. Parse workbook.xml.rels to get rId → sheetN.xml mapping
    const srcRels = await getRels(srcZip);
    const dstRels = await getRels(dstZip);

    // 4. Read source shared strings
    setProgress(20, 'Reading shared strings…');
    const srcStrings = await getSharedStrings(srcZip);
    log(`Source shared strings: ${srcStrings.length}`, 'dim');

    // 5. Read destination shared strings (we'll need to add new ones)
    const dstStrings = await getSharedStrings(dstZip);
    const dstStringIndex = new Map();
    dstStrings.forEach((s, i) => dstStringIndex.set(s, i));
    let dstStringsDirty = false;
    log(`Destination shared strings: ${dstStrings.length}`, 'dim');

    // 6. For each matching sheet, copy values
    const sheetNames = [...srcSheetMap.keys()];
    let totalCopied = 0;
    let totalSheets = 0;

    for (let i = 0; i < sheetNames.length; i++) {
        const name = sheetNames[i];
        const pct = 25 + Math.round((i / sheetNames.length) * 65);
        setProgress(pct, `Processing sheet: ${name}`);

        if (!dstSheetMap.has(name)) {
            log(`⚠ Sheet "${name}" not found in destination – skipped`, 'warn');
            continue;
        }

        const srcRId = srcSheetMap.get(name);
        const dstRId = dstSheetMap.get(name);
        const srcSheetPath = srcRels.get(srcRId);
        const dstSheetPath = dstRels.get(dstRId);

        if (!srcSheetPath || !dstSheetPath) continue;

        const srcFullPath = 'xl/' + srcSheetPath;
        const dstFullPath = 'xl/' + dstSheetPath;

        if (!srcZip.file(srcFullPath) || !dstZip.file(dstFullPath)) continue;

        // Parse source sheet to extract cell values
        const srcXml = await srcZip.file(srcFullPath).async('string');
        const srcCells = extractCellValues(srcXml, srcStrings);

        if (srcCells.size === 0) continue;

        // Parse destination sheet and inject values
        const dstXml = await dstZip.file(dstFullPath).async('string');
        const { xml: newXml, copied, newStrings } = injectCellValues(
            dstXml, srcCells, dstStrings, dstStringIndex
        );

        if (newStrings > 0) dstStringsDirty = true;

        if (copied > 0) {
            dstZip.file(dstFullPath, newXml);
            totalCopied += copied;
            totalSheets++;
            log(`✓ ${name}: ${copied} values copied`, 'success');
        } else {
            log(`– ${name}: no new values to copy`, 'dim');
        }
    }

    // 7. If we added shared strings, rewrite sharedStrings.xml
    if (dstStringsDirty) {
        setProgress(92, 'Updating shared strings…');
        const newSSXml = buildSharedStringsXml(dstStrings);
        dstZip.file('xl/sharedStrings.xml', newSSXml);
        log(`Updated shared strings (now ${dstStrings.length})`, 'info');
    }

    // 8. Generate output
    setProgress(95, 'Generating output file…');
    log('Generating output file…', 'info');

    // Determine MIME type based on file extension
    const isXlsm = destFile.name.toLowerCase().endsWith('.xlsm');
    const mimeType = isXlsm
        ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    resultBlob = await dstZip.generateAsync({
        type: 'blob',
        mimeType: mimeType,
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
    });

    setProgress(100, 'Done!');
    log(`\n✅ Mapping complete! ${totalCopied} values copied across ${totalSheets} sheets.`, 'success');

    resultSummary.textContent =
        `${totalCopied} cell values were copied across ${totalSheets} sheets.`;
    resultSection.style.display = 'block';
}

// ─── XML Helpers ─────────────────────────────────────────────────────

/**
 * Parse workbook.xml → Map<sheetName, rId>
 */
async function getSheetMap(zip) {
    const xml = await zip.file('xl/workbook.xml').async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const ns  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const rns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const map = new Map();
    const sheets = doc.getElementsByTagNameNS(ns, 'sheet');
    for (const s of sheets) {
        map.set(s.getAttribute('name'), s.getAttributeNS(rns, 'id'));
    }
    return map;
}

/**
 * Parse workbook.xml.rels → Map<rId, relative path>
 */
async function getRels(zip) {
    const xml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const map = new Map();
    const rels = doc.getElementsByTagName('Relationship');
    for (const r of rels) {
        map.set(r.getAttribute('Id'), r.getAttribute('Target'));
    }
    return map;
}

/**
 * Parse sharedStrings.xml → string[]
 */
async function getSharedStrings(zip) {
    const file = zip.file('xl/sharedStrings.xml');
    if (!file) return [];
    const xml  = await file.async('string');
    const doc  = new DOMParser().parseFromString(xml, 'application/xml');
    const ns   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const arr  = [];
    const items = doc.getElementsByTagNameNS(ns, 'si');
    for (const si of items) {
        // Collect all <t> text nodes (handles rich text with <r><t>…)
        let text = '';
        const tNodes = si.getElementsByTagNameNS(ns, 't');
        for (const t of tNodes) text += t.textContent;
        arr.push(text);
    }
    return arr;
}

/**
 * Extract all cell values from a worksheet XML string.
 * Returns Map<cellRef, { value, type }>
 *   type: 'n' (number/inline), 's' (already resolved string), 'b' (bool), 'str' (formula string)
 */
function extractCellValues(xml, sharedStrings) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const ns  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const map = new Map();
    const cells = doc.getElementsByTagNameNS(ns, 'c');

    for (const c of cells) {
        const ref  = c.getAttribute('r');
        const type = c.getAttribute('t') || '';
        const vEl  = c.getElementsByTagNameNS(ns, 'v')[0];
        if (!vEl) continue;

        const raw = vEl.textContent;
        if (raw === '') continue;

        let resolvedValue = raw;
        let resolvedType  = type || 'n';

        if (type === 's') {
            // Shared string – resolve to actual text
            const idx = parseInt(raw, 10);
            if (idx < sharedStrings.length) {
                resolvedValue = sharedStrings[idx];
            }
            resolvedType = 's';
        }

        map.set(ref, { value: resolvedValue, type: resolvedType });
    }
    return map;
}

/**
 * Inject source cell values into destination worksheet XML.
 * - If a destination cell already has a <v> with the same value → skip.
 * - If the destination cell exists but has no <v> or different value → set it.
 * - If the destination cell doesn't exist → create it in the correct row.
 *
 * For string values, we look up or add to the destination shared-string table.
 *
 * Returns { xml, copied, newStrings }
 */
function injectCellValues(dstXml, srcCells, dstStrings, dstStringIndex) {
    const doc = new DOMParser().parseFromString(dstXml, 'application/xml');
    const ns  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    let copied = 0;
    let newStrings = 0;

    // Build a map of existing destination cells for quick lookup
    const dstCellElements = new Map();
    const cells = doc.getElementsByTagNameNS(ns, 'c');
    for (const c of cells) {
        dstCellElements.set(c.getAttribute('r'), c);
    }

    // Build a map of existing rows
    const dstRowElements = new Map();
    const rows = doc.getElementsByTagNameNS(ns, 'row');
    for (const r of rows) {
        dstRowElements.set(r.getAttribute('r'), r);
    }

    for (const [ref, { value, type }] of srcCells) {
        // Determine what value/type to write into destination
        let writeType = type;
        let writeValue = value;

        if (type === 's') {
            // Need to find or add in destination shared strings
            let idx = dstStringIndex.get(value);
            if (idx === undefined) {
                idx = dstStrings.length;
                dstStrings.push(value);
                dstStringIndex.set(value, idx);
                newStrings++;
            }
            writeValue = String(idx);
            writeType  = 's';
        }

        const existing = dstCellElements.get(ref);
        if (existing) {
            // Cell element exists in destination
            const vEl = existing.getElementsByTagNameNS(ns, 'v')[0];
            if (vEl && vEl.textContent === writeValue) {
                // Same value already there – skip
                continue;
            }
            // Set type
            if (writeType && writeType !== 'n') {
                existing.setAttribute('t', writeType);
            }
            // Set or create <v>
            if (vEl) {
                vEl.textContent = writeValue;
            } else {
                const newV = doc.createElementNS(ns, 'v');
                newV.textContent = writeValue;
                existing.appendChild(newV);
            }
            copied++;
        } else {
            // Cell doesn't exist – we need to create it in the right row
            const rowNum = ref.replace(/[A-Z]+/i, '');
            let rowEl = dstRowElements.get(rowNum);

            if (!rowEl) {
                // Row doesn't exist either – create it
                rowEl = doc.createElementNS(ns, 'row');
                rowEl.setAttribute('r', rowNum);
                // Find sheetData and append
                const sheetData = doc.getElementsByTagNameNS(ns, 'sheetData')[0];
                if (sheetData) {
                    // Insert row in order
                    const rn = parseInt(rowNum, 10);
                    let inserted = false;
                    for (const existingRow of sheetData.getElementsByTagNameNS(ns, 'row')) {
                        if (parseInt(existingRow.getAttribute('r'), 10) > rn) {
                            sheetData.insertBefore(rowEl, existingRow);
                            inserted = true;
                            break;
                        }
                    }
                    if (!inserted) sheetData.appendChild(rowEl);
                    dstRowElements.set(rowNum, rowEl);
                }
            }

            // Create cell element
            const newCell = doc.createElementNS(ns, 'c');
            newCell.setAttribute('r', ref);
            if (writeType && writeType !== 'n') {
                newCell.setAttribute('t', writeType);
            }
            const newV = doc.createElementNS(ns, 'v');
            newV.textContent = writeValue;
            newCell.appendChild(newV);

            // Insert cell in column order within the row
            const colName = ref.replace(/[0-9]+/, '');
            let cellInserted = false;
            for (const existingCell of rowEl.getElementsByTagNameNS(ns, 'c')) {
                const existCol = existingCell.getAttribute('r').replace(/[0-9]+/, '');
                if (compareColumns(existCol, colName) > 0) {
                    rowEl.insertBefore(newCell, existingCell);
                    cellInserted = true;
                    break;
                }
            }
            if (!cellInserted) rowEl.appendChild(newCell);

            dstCellElements.set(ref, newCell);
            copied++;
        }
    }

    const serializer = new XMLSerializer();
    return {
        xml: serializer.serializeToString(doc),
        copied,
        newStrings
    };
}

/**
 * Build a new sharedStrings.xml from the string array
 */
function buildSharedStringsXml(strings) {
    const count = strings.length;
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
    xml += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${count}" uniqueCount="${count}">`;
    for (const s of strings) {
        xml += `<si><t xml:space="preserve">${escapeXml(s)}</t></si>`;
    }
    xml += '</sst>';
    return xml;
}

// ─── Utility helpers ─────────────────────────────────────────────────

function escapeXml(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/** Compare Excel column letters: A < B < … < Z < AA < AB … */
function compareColumns(a, b) {
    if (a.length !== b.length) return a.length - b.length;
    return a < b ? -1 : a > b ? 1 : 0;
}

function setProgress(pct, text) {
    progressBar.style.width = pct + '%';
    progressText.textContent = text;
}

function log(msg, cls = 'info') {
    const div = document.createElement('div');
    div.className = 'log-entry ' + cls;
    div.textContent = msg;
    logContainer.appendChild(div);
    logContainer.scrollTop = logContainer.scrollHeight;
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
}

