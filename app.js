let templateFile = null;
let targetFiles = [];

const templateZone = document.getElementById('templateZone');
const targetZone = document.getElementById('targetZone');
const updateBtn = document.getElementById('updateBtn');
const status = document.getElementById('status');

setupDropZone(templateZone, 'template');
setupDropZone(targetZone, 'target');

function setupDropZone(zone, type) {
    ['dragenter', 'dragover'].forEach(evt => {
        zone.addEventListener(evt, (e) => {
            e.preventDefault();
            zone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(evt => {
        zone.addEventListener(evt, (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
        });
    });

    zone.addEventListener('drop', (e) => {
        const files = Array.from(e.dataTransfer.files).filter(f =>
            f.name.toLowerCase().endsWith('.docx')
        );

        if (files.length === 0) {
            showStatus('Please drop .docx files only', 'error');
            return;
        }

        if (type === 'template') {
            templateFile = files[0];
            renderFileList('templateFiles', [templateFile]);
            zone.classList.add('has-files');
        } else {
            const existingNames = new Set(targetFiles.map(f => f.name));
            const newFiles = files.filter(f => !existingNames.has(f.name));
            targetFiles = targetFiles.concat(newFiles);
            renderTargetFileList();
            if (targetFiles.length > 0) zone.classList.add('has-files');
        }

        checkReady();
    });

    zone.addEventListener('click', (e) => {
        if (e.target.closest('.remove-btn') || e.target.closest('.clear-btn')) return;

        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.docx';
        input.multiple = (type === 'target');
        input.onchange = () => {
            const files = Array.from(input.files);
            if (type === 'template') {
                templateFile = files[0];
                renderFileList('templateFiles', [templateFile]);
                zone.classList.add('has-files');
            } else {
                const existingNames = new Set(targetFiles.map(f => f.name));
                const newFiles = files.filter(f => !existingNames.has(f.name));
                targetFiles = targetFiles.concat(newFiles);
                renderTargetFileList();
                if (targetFiles.length > 0) zone.classList.add('has-files');
            }
            checkReady();
        };
        input.click();
    });
}

function renderFileList(containerId, files) {
    const container = document.getElementById(containerId);
    container.innerHTML = files.map(f =>
        `<div class="file-item">${f.name}</div>`
    ).join('');
}

function renderTargetFileList() {
    const container = document.getElementById('targetFiles');
    if (targetFiles.length === 0) {
        container.innerHTML = '';
        targetZone.classList.remove('has-files');
        return;
    }
    container.innerHTML =
        `<div class="file-list-header">${targetFiles.length} file(s) <button class="clear-btn" onclick="clearTargets(event)">Clear all</button></div>` +
        targetFiles.map((f, i) =>
            `<div class="file-item">${f.name} <button class="remove-btn" onclick="removeTarget(event, ${i})">&#10005;</button></div>`
        ).join('');
}

function removeTarget(e, index) {
    e.stopPropagation();
    targetFiles.splice(index, 1);
    renderTargetFileList();
    checkReady();
}

function clearTargets(e) {
    e.stopPropagation();
    targetFiles = [];
    renderTargetFileList();
    checkReady();
}

window.removeTarget = removeTarget;
window.clearTargets = clearTargets;

function checkReady() {
    updateBtn.disabled = !(templateFile && targetFiles.length > 0);
}

function showStatus(message, type) {
    status.textContent = message;
    status.className = 'status ' + type;
}

// ===== XML PARSING HELPERS =====

const PARSER = new DOMParser();
const SERIALIZER = new XMLSerializer();

function parseXml(xmlString) {
    return PARSER.parseFromString(xmlString, 'application/xml');
}

function serializeXml(doc) {
    return SERIALIZER.serializeToString(doc);
}

// ===== DOCX UPDATE LOGIC =====

async function updateDocument(templateZip, targetBuffer) {
    const targetZip = await JSZip.loadAsync(targetBuffer);

    // 1. Strip comments FIRST
    await stripComments(targetZip);

    // 2. Strip tracked changes
    await stripRevisions(targetZip);

    // 3. Transfer styles, theme, numbering (simple file copy)
    await transferFile(templateZip, targetZip, 'word/styles.xml');
    await transferFile(templateZip, targetZip, 'word/theme/theme1.xml');
    await transferFile(templateZip, targetZip, 'word/numbering.xml');

    // 4. Transfer headers/footers using proper XML parsing
    await transferHeadersFooters(templateZip, targetZip);

    return await targetZip.generateAsync({ type: 'blob' });
}

async function transferFile(templateZip, targetZip, path) {
    const file = templateZip.file(path);
    if (file) {
        const content = await file.async('uint8array');
        targetZip.file(path, content);
    }
}

// ===== COMMENT STRIPPING =====

async function stripComments(targetZip) {
    // Remove all comment-related files
    const commentFiles = [
        'word/comments.xml', 'word/commentsExtended.xml',
        'word/commentsIds.xml', 'word/commentsExtensible.xml',
        'word/people.xml',
        'word/_rels/comments.xml.rels',
        'word/_rels/commentsExtended.xml.rels',
        'word/_rels/commentsIds.xml.rels',
        'word/_rels/commentsExtensible.xml.rels'
    ];
    commentFiles.forEach(path => {
        if (targetZip.file(path)) targetZip.remove(path);
    });

    // Collect all XML files to clean (document + headers + footers)
    const xmlFiles = ['word/document.xml'];
    targetZip.folder('word').forEach((path) => {
        if (/^(header|footer)\d*\.xml$/.test(path)) {
            xmlFiles.push('word/' + path);
        }
    });

    for (const filePath of xmlFiles) {
        const file = targetZip.file(filePath);
        if (!file) continue;
        let xml = await file.async('string');

        // Use DOMParser for robust comment removal
        const doc = parseXml(xml);
        const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

        // Remove comment range markers
        const tagNames = ['commentRangeStart', 'commentRangeEnd', 'commentReference', 'annotationRef'];
        for (const tagName of tagNames) {
            const elements = doc.getElementsByTagNameNS(W_NS, tagName);
            while (elements.length > 0) {
                elements[0].parentNode.removeChild(elements[0]);
            }
        }

        // Remove w15:commentEx elements
        const w15Elements = doc.getElementsByTagName('w15:commentEx');
        while (w15Elements.length > 0) {
            w15Elements[0].parentNode.removeChild(w15Elements[0]);
        }

        // Remove w16cid:commentId elements
        const w16Elements = doc.getElementsByTagName('w16cid:commentId');
        while (w16Elements.length > 0) {
            w16Elements[0].parentNode.removeChild(w16Elements[0]);
        }

        // Remove <w:r> runs that only contain a commentReference (no <w:t>)
        // After removing commentReference above, check for empty runs
        const runs = doc.getElementsByTagNameNS(W_NS, 'r');
        const runsToRemove = [];
        for (let i = 0; i < runs.length; i++) {
            const run = runs[i];
            const hasText = run.getElementsByTagNameNS(W_NS, 't').length > 0;
            const hasDrawing = run.getElementsByTagNameNS(W_NS, 'drawing').length > 0;
            const hasPict = run.getElementsByTagNameNS(W_NS, 'pict').length > 0;
            const hasBreak = run.getElementsByTagNameNS(W_NS, 'br').length > 0;
            const hasTab = run.getElementsByTagNameNS(W_NS, 'tab').length > 0;
            const hasSym = run.getElementsByTagNameNS(W_NS, 'sym').length > 0;
            const hasFldChar = run.getElementsByTagNameNS(W_NS, 'fldChar').length > 0;
            const hasInstrText = run.getElementsByTagNameNS(W_NS, 'instrText').length > 0;

            // If run has no meaningful content, mark for removal
            if (!hasText && !hasDrawing && !hasPict && !hasBreak && !hasTab && !hasSym && !hasFldChar && !hasInstrText) {
                // Check if it only has rPr (formatting) and nothing else useful
                const children = run.childNodes;
                let hasContent = false;
                for (let j = 0; j < children.length; j++) {
                    const child = children[j];
                    if (child.nodeType === 1) { // Element node
                        const localName = child.localName;
                        if (localName !== 'rPr') {
                            hasContent = true;
                            break;
                        }
                    }
                }
                if (!hasContent) {
                    runsToRemove.push(run);
                }
            }
        }
        for (const run of runsToRemove) {
            run.parentNode.removeChild(run);
        }

        targetZip.file(filePath, serializeXml(doc));
    }

    // Clean relationships using DOMParser
    const relsFile = targetZip.file('word/_rels/document.xml.rels');
    if (relsFile) {
        const relsXml = await relsFile.async('string');
        const doc = parseXml(relsXml);
        const rels = doc.getElementsByTagName('Relationship');
        const toRemove = [];
        for (let i = 0; i < rels.length; i++) {
            const type = rels[i].getAttribute('Type') || '';
            if (/\/(comments|commentsExtended|commentsIds|commentsExtensible|people)$/.test(type)) {
                toRemove.push(rels[i]);
            }
        }
        for (const el of toRemove) {
            el.parentNode.removeChild(el);
        }
        targetZip.file('word/_rels/document.xml.rels', serializeXml(doc));
    }

    // Clean content types using DOMParser
    const ctFile = targetZip.file('[Content_Types].xml');
    if (ctFile) {
        const ctXml = await ctFile.async('string');
        const doc = parseXml(ctXml);
        const overrides = doc.getElementsByTagName('Override');
        const toRemove = [];
        for (let i = 0; i < overrides.length; i++) {
            const partName = overrides[i].getAttribute('PartName') || '';
            if (/\/word\/(comments|commentsExtended|commentsIds|commentsExtensible|people)\.xml$/.test(partName)) {
                toRemove.push(overrides[i]);
            }
        }
        for (const el of toRemove) {
            el.parentNode.removeChild(el);
        }
        targetZip.file('[Content_Types].xml', serializeXml(doc));
    }
}

// ===== REVISION STRIPPING =====

async function stripRevisions(targetZip) {
    const xmlFiles = ['word/document.xml'];
    targetZip.folder('word').forEach((path) => {
        if (/^(header|footer)\d*\.xml$/.test(path)) {
            xmlFiles.push('word/' + path);
        }
    });

    for (const filePath of xmlFiles) {
        const file = targetZip.file(filePath);
        if (!file) continue;
        let xml = await file.async('string');

        const doc = parseXml(xml);
        const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

        // Remove <w:del> elements entirely (deleted text goes away)
        let dels = doc.getElementsByTagNameNS(W_NS, 'del');
        while (dels.length > 0) {
            dels[0].parentNode.removeChild(dels[0]);
        }

        // Unwrap <w:ins> (keep children, remove wrapper)
        let ins = doc.getElementsByTagNameNS(W_NS, 'ins');
        while (ins.length > 0) {
            const insEl = ins[0];
            const parent = insEl.parentNode;
            while (insEl.firstChild) {
                parent.insertBefore(insEl.firstChild, insEl);
            }
            parent.removeChild(insEl);
        }

        // Remove <w:moveFrom> entirely
        let moveFroms = doc.getElementsByTagNameNS(W_NS, 'moveFrom');
        while (moveFroms.length > 0) {
            moveFroms[0].parentNode.removeChild(moveFroms[0]);
        }

        // Unwrap <w:moveTo>
        let moveTos = doc.getElementsByTagNameNS(W_NS, 'moveTo');
        while (moveTos.length > 0) {
            const el = moveTos[0];
            const parent = el.parentNode;
            while (el.firstChild) {
                parent.insertBefore(el.firstChild, el);
            }
            parent.removeChild(el);
        }

        // Remove property change elements
        const changeTypes = ['rPrChange', 'pPrChange', 'sectPrChange',
            'tblPrChange', 'tcPrChange', 'trPrChange', 'tblGridChange'];
        for (const changeType of changeTypes) {
            let elements = doc.getElementsByTagNameNS(W_NS, changeType);
            while (elements.length > 0) {
                elements[0].parentNode.removeChild(elements[0]);
            }
        }

        // Remove move range markers
        const moveMarkers = ['moveFromRangeStart', 'moveFromRangeEnd',
            'moveToRangeStart', 'moveToRangeEnd'];
        for (const marker of moveMarkers) {
            let elements = doc.getElementsByTagNameNS(W_NS, marker);
            while (elements.length > 0) {
                elements[0].parentNode.removeChild(elements[0]);
            }
        }

        targetZip.file(filePath, serializeXml(doc));
    }
}

// ===== HEADER/FOOTER TRANSFER =====

async function transferHeadersFooters(templateZip, targetZip) {
    // --- Step 1: Parse template's document.xml.rels to find header/footer entries ---
    const templateRelsFile = templateZip.file('word/_rels/document.xml.rels');
    if (!templateRelsFile) return;
    const templateRelsXml = await templateRelsFile.async('string');
    const templateRelsDoc = parseXml(templateRelsXml);

    const templateHFEntries = []; // {originalId, target, typeUrl, isHeader}
    const templateRelEls = templateRelsDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < templateRelEls.length; i++) {
        const el = templateRelEls[i];
        const typeUrl = el.getAttribute('Type') || '';
        if (/\/(header|footer)$/.test(typeUrl)) {
            templateHFEntries.push({
                originalId: el.getAttribute('Id'),
                target: el.getAttribute('Target'),
                typeUrl: typeUrl,
                isHeader: /\/header$/.test(typeUrl)
            });
        }
    }

    if (templateHFEntries.length === 0) return;

    // --- Step 2: Parse target's document.xml.rels ---
    const targetRelsFile = targetZip.file('word/_rels/document.xml.rels');
    if (!targetRelsFile) return;
    const targetRelsXml = await targetRelsFile.async('string');
    const targetRelsDoc = parseXml(targetRelsXml);
    const targetRelsRoot = targetRelsDoc.documentElement;

    // Find max rId in target
    let maxId = 0;
    const targetRelEls = targetRelsDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < targetRelEls.length; i++) {
        const id = targetRelEls[i].getAttribute('Id') || '';
        const num = parseInt(id.replace(/\D/g, ''));
        if (!isNaN(num) && num > maxId) maxId = num;
    }

    // Remove existing header/footer relationships from target
    const toRemoveRels = [];
    for (let i = 0; i < targetRelEls.length; i++) {
        const typeUrl = targetRelEls[i].getAttribute('Type') || '';
        if (/\/(header|footer)$/.test(typeUrl)) {
            toRemoveRels.push(targetRelEls[i]);
        }
    }
    for (const el of toRemoveRels) {
        el.parentNode.removeChild(el);
    }

    // --- Step 3: Remove existing header/footer files from target ---
    const existingHF = [];
    targetZip.folder('word').forEach((path) => {
        if (/^(header|footer)\d*\.xml$/.test(path)) {
            existingHF.push('word/' + path);
        }
    });
    for (const path of existingHF) {
        targetZip.remove(path);
        const relsPath = 'word/_rels/' + path.replace('word/', '') + '.rels';
        if (targetZip.file(relsPath)) targetZip.remove(relsPath);
    }

    // --- Step 4: Copy template header/footer files and assign fresh rIds ---
    const idMapping = {}; // template originalId -> new rId

    for (const entry of templateHFEntries) {
        maxId++;
        const newId = 'rId' + maxId;
        idMapping[entry.originalId] = newId;

        // Add relationship to target rels
        const newRel = targetRelsDoc.createElementNS(
            targetRelsRoot.namespaceURI, 'Relationship'
        );
        newRel.setAttribute('Id', newId);
        newRel.setAttribute('Type', entry.typeUrl);
        newRel.setAttribute('Target', entry.target);
        targetRelsRoot.appendChild(newRel);

        // Copy header/footer XML file
        const srcPath = 'word/' + entry.target;
        const srcFile = templateZip.file(srcPath);
        if (srcFile) {
            const content = await srcFile.async('uint8array');
            targetZip.file(srcPath, content);
        }

        // Copy its .rels file
        const srcRelsPath = 'word/_rels/' + entry.target + '.rels';
        const srcRelsFile = templateZip.file(srcRelsPath);
        if (srcRelsFile) {
            const content = await srcRelsFile.async('uint8array');
            targetZip.file(srcRelsPath, content);
        }
    }

    // Save updated target rels
    targetZip.file('word/_rels/document.xml.rels', serializeXml(targetRelsDoc));

    // --- Step 5: Copy media files referenced by headers/footers ---
    for (const entry of templateHFEntries) {
        const relsPath = 'word/_rels/' + entry.target + '.rels';
        const relsFile = templateZip.file(relsPath);
        if (!relsFile) continue;
        const relsContent = await relsFile.async('string');
        const relsDoc = parseXml(relsContent);
        const rels = relsDoc.getElementsByTagName('Relationship');
        for (let i = 0; i < rels.length; i++) {
            const mediaTarget = rels[i].getAttribute('Target') || '';
            if (mediaTarget.startsWith('http://') || mediaTarget.startsWith('https://')) continue;
            const mediaPath = 'word/' + mediaTarget.replace(/^\.\//, '');
            const mediaFile = templateZip.file(mediaPath);
            if (mediaFile) {
                const content = await mediaFile.async('uint8array');
                targetZip.file(mediaPath, content);
            }
        }
    }

    // --- Step 6: Update [Content_Types].xml ---
    const ctFile = targetZip.file('[Content_Types].xml');
    if (ctFile) {
        const ctXml = await ctFile.async('string');
        const ctDoc = parseXml(ctXml);
        const ctRoot = ctDoc.documentElement;

        // Remove existing header/footer overrides
        const overrides = ctDoc.getElementsByTagName('Override');
        const toRemoveCT = [];
        for (let i = 0; i < overrides.length; i++) {
            const partName = overrides[i].getAttribute('PartName') || '';
            if (/\/word\/(header|footer)\d*\.xml$/.test(partName)) {
                toRemoveCT.push(overrides[i]);
            }
        }
        for (const el of toRemoveCT) {
            el.parentNode.removeChild(el);
        }

        // Add new overrides
        const addedParts = new Set();
        for (const entry of templateHFEntries) {
            const partName = '/word/' + entry.target;
            if (addedParts.has(partName)) continue;
            addedParts.add(partName);

            const ct = entry.isHeader
                ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
                : 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

            const override = ctDoc.createElementNS(ctRoot.namespaceURI, 'Override');
            override.setAttribute('PartName', partName);
            override.setAttribute('ContentType', ct);
            ctRoot.appendChild(override);
        }

        targetZip.file('[Content_Types].xml', serializeXml(ctDoc));
    }

    // --- Step 7: Update header/footer references in ALL sectPr elements ---
    const templateDocFile = templateZip.file('word/document.xml');
    if (!templateDocFile) return;
    const templateDocXml = await templateDocFile.async('string');
    const templateDocDoc = parseXml(templateDocXml);

    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    // Get ALL header/footer references from template's last sectPr
    const templateSectPrs = templateDocDoc.getElementsByTagNameNS(W_NS, 'sectPr');
    if (templateSectPrs.length === 0) return;
    const templateSectPr = templateSectPrs[templateSectPrs.length - 1];

    // Collect template's headerReference and footerReference elements
    const templateRefs = [];
    for (let i = 0; i < templateSectPr.childNodes.length; i++) {
        const child = templateSectPr.childNodes[i];
        if (child.nodeType !== 1) continue;
        if (child.localName === 'headerReference' || child.localName === 'footerReference') {
            templateRefs.push({
                tagName: child.localName,
                type: child.getAttributeNS(W_NS, 'type') || child.getAttribute('w:type') || '',
                originalRid: child.getAttributeNS(R_NS, 'id') || child.getAttribute('r:id') || ''
            });
        }
    }

    // Also collect page layout elements from template sectPr
    const layoutTagNames = ['pgSz', 'pgMar', 'cols', 'docGrid', 'pgBorders', 'lnNumType', 'pgNumType'];
    const templateLayoutXmls = {};
    for (const tagName of layoutTagNames) {
        const els = templateSectPr.getElementsByTagNameNS(W_NS, tagName);
        if (els.length > 0) {
            templateLayoutXmls[tagName] = serializeXml(els[0]);
        }
    }

    // Now update target document.xml
    const targetDocFile = targetZip.file('word/document.xml');
    if (!targetDocFile) return;
    const targetDocXml = await targetDocFile.async('string');
    const targetDocDoc = parseXml(targetDocXml);

    const targetSectPrs = targetDocDoc.getElementsByTagNameNS(W_NS, 'sectPr');

    for (let s = 0; s < targetSectPrs.length; s++) {
        const sectPr = targetSectPrs[s];

        // Remove existing headerReference and footerReference elements
        const toRemoveHF = [];
        for (let i = 0; i < sectPr.childNodes.length; i++) {
            const child = sectPr.childNodes[i];
            if (child.nodeType !== 1) continue;
            if (child.localName === 'headerReference' || child.localName === 'footerReference') {
                toRemoveHF.push(child);
            }
        }
        for (const el of toRemoveHF) {
            sectPr.removeChild(el);
        }

        // Add template's header/footer references with remapped rIds
        for (const ref of templateRefs) {
            const newEl = targetDocDoc.createElementNS(W_NS, 'w:' + ref.tagName);
            newEl.setAttributeNS(W_NS, 'w:type', ref.type);
            const newRid = idMapping[ref.originalRid] || ref.originalRid;
            newEl.setAttributeNS(R_NS, 'r:id', newRid);
            // Insert before first child to keep proper order
            if (sectPr.firstChild) {
                sectPr.insertBefore(newEl, sectPr.firstChild);
            } else {
                sectPr.appendChild(newEl);
            }
        }

        // Replace page layout properties
        for (const tagName of layoutTagNames) {
            // Remove existing
            const existing = sectPr.getElementsByTagNameNS(W_NS, tagName);
            while (existing.length > 0) {
                sectPr.removeChild(existing[0]);
            }

            // Add from template
            if (templateLayoutXmls[tagName]) {
                const tempDoc = parseXml(templateLayoutXmls[tagName]);
                const imported = targetDocDoc.importNode(tempDoc.documentElement, true);
                sectPr.appendChild(imported);
            }
        }
    }

    targetZip.file('word/document.xml', serializeXml(targetDocDoc));
}

// ===== MAIN UPDATE HANDLER =====

updateBtn.addEventListener('click', async () => {
    if (!templateFile || targetFiles.length === 0) return;

    updateBtn.disabled = true;
    showStatus('Processing documents...', 'info');

    try {
        const templateBuffer = await templateFile.arrayBuffer();
        const templateZip = await JSZip.loadAsync(templateBuffer);

        if (targetFiles.length === 1) {
            const targetBuffer = await targetFiles[0].arrayBuffer();
            const updatedBlob = await updateDocument(templateZip, targetBuffer);
            const newName = targetFiles[0].name.replace('.docx', '_updated.docx');
            saveAs(updatedBlob, newName);
        } else {
            const outputZip = new JSZip();
            for (let i = 0; i < targetFiles.length; i++) {
                showStatus(`Processing ${i + 1} of ${targetFiles.length}...`, 'info');
                const targetBuffer = await targetFiles[i].arrayBuffer();
                const updatedBlob = await updateDocument(templateZip, targetBuffer);
                const newName = targetFiles[i].name.replace('.docx', '_updated.docx');
                outputZip.file(newName, updatedBlob);
            }
            const zipBlob = await outputZip.generateAsync({ type: 'blob' });
            saveAs(zipBlob, 'updated_documents.zip');
        }

        showStatus(`Successfully updated ${targetFiles.length} document(s)!`, 'success');
    } catch (err) {
        showStatus('Error: ' + err.message, 'error');
        console.error(err);
    } finally {
        updateBtn.disabled = false;
    }
});
