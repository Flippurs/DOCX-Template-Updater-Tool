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

// ===== DOCX UPDATE LOGIC =====

async function updateDocument(templateZip, targetBuffer) {
    const targetZip = await JSZip.loadAsync(targetBuffer);

    // 1. Strip comments FIRST
    await stripComments(targetZip);

    // 2. Strip tracked changes
    await stripRevisions(targetZip);

    // 3. Transfer styles
    await transferFile(templateZip, targetZip, 'word/styles.xml');

    // 4. Transfer theme
    await transferFile(templateZip, targetZip, 'word/theme/theme1.xml');

    // 5. Transfer numbering
    await transferFile(templateZip, targetZip, 'word/numbering.xml');

    // 6. Transfer headers/footers with fresh relationship IDs
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
    // Remove comment files
    const commentFiles = [
        'word/comments.xml', 'word/commentsExtended.xml',
        'word/commentsIds.xml', 'word/commentsExtensible.xml',
        'word/people.xml',
        'word/_rels/comments.xml.rels',
        'word/_rels/commentsExtended.xml.rels'
    ];
    commentFiles.forEach(path => {
        if (targetZip.file(path)) targetZip.remove(path);
    });

    // Clean all XML files that can contain comment markers
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

        // Remove comment range markers (self-closing)
        xml = xml.replace(/<w:commentRangeStart[^\/]*\/>/g, '');
        xml = xml.replace(/<w:commentRangeEnd[^\/]*\/>/g, '');

        // Remove comment references (self-closing)
        xml = xml.replace(/<w:commentReference[^\/]*\/>/g, '');
        xml = xml.replace(/<w:annotationRef[^\/]*\/>/g, '');

        // Remove w15/w16 comment extensions
        xml = xml.replace(/<w15:commentEx[^\/]*\/>/g, '');
        xml = xml.replace(/<w16cid:commentId[^\/]*\/>/g, '');

        // Remove runs that only contain comment references
        // Pattern: <w:r>...<w:commentReference .../>...</w:r> where there's no <w:t>
        xml = xml.replace(/<w:r\b[^>]*>(?:(?!<w:t[ >]|<w:t\/>)[\s\S])*?<w:commentReference[^\/]*\/>(?:(?!<w:t[ >]|<w:t\/>)[\s\S])*?<\/w:r>/g, '');

        targetZip.file(filePath, xml);
    }

    // Clean relationships
    const relsFile = targetZip.file('word/_rels/document.xml.rels');
    if (relsFile) {
        let xml = await relsFile.async('string');
        xml = xml.replace(/<Relationship[^>]*Type="[^"]*\/(comments|commentsExtended|commentsIds|commentsExtensible|people)"[^>]*\/>/g, '');
        targetZip.file('word/_rels/document.xml.rels', xml);
    }

    // Clean content types
    const ctFile = targetZip.file('[Content_Types].xml');
    if (ctFile) {
        let xml = await ctFile.async('string');
        xml = xml.replace(/<Override[^>]*PartName="\/word\/(comments|commentsExtended|commentsIds|commentsExtensible|people)\.xml"[^>]*\/>/g, '');
        targetZip.file('[Content_Types].xml', xml);
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

        // Remove deleted content
        xml = xml.replace(/<w:del\b[\s\S]*?<\/w:del>/g, '');

        // Unwrap insertions
        xml = xml.replace(/<w:ins\b[^>]*>/g, '');
        xml = xml.replace(/<\/w:ins>/g, '');

        // Remove move-from content
        xml = xml.replace(/<w:moveFrom\b[\s\S]*?<\/w:moveFrom>/g, '');

        // Unwrap move-to
        xml = xml.replace(/<w:moveTo\b[^>]*>/g, '');
        xml = xml.replace(/<\/w:moveTo>/g, '');

        // Remove property changes
        xml = xml.replace(/<w:rPrChange\b[\s\S]*?<\/w:rPrChange>/g, '');
        xml = xml.replace(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g, '');
        xml = xml.replace(/<w:sectPrChange\b[\s\S]*?<\/w:sectPrChange>/g, '');
        xml = xml.replace(/<w:tblPrChange\b[\s\S]*?<\/w:tblPrChange>/g, '');
        xml = xml.replace(/<w:tcPrChange\b[\s\S]*?<\/w:tcPrChange>/g, '');
        xml = xml.replace(/<w:trPrChange\b[\s\S]*?<\/w:trPrChange>/g, '');
        xml = xml.replace(/<w:tblGridChange\b[\s\S]*?<\/w:tblGridChange>/g, '');

        // Remove move range markers
        xml = xml.replace(/<w:moveFromRangeStart[^\/]*\/>/g, '');
        xml = xml.replace(/<w:moveFromRangeEnd[^\/]*\/>/g, '');
        xml = xml.replace(/<w:moveToRangeStart[^\/]*\/>/g, '');
        xml = xml.replace(/<w:moveToRangeEnd[^\/]*\/>/g, '');

        targetZip.file(filePath, xml);
    }
}

// ===== HEADER/FOOTER TRANSFER =====

/**
 * Transfers headers/footers from template using fresh rId numbers
 * that won't conflict with existing target relationships.
 */
async function transferHeadersFooters(templateZip, targetZip) {
    // Get template's document.xml.rels to find header/footer relationships
    const templateRelsFile = templateZip.file('word/_rels/document.xml.rels');
    if (!templateRelsFile) return;
    const templateRelsXml = await templateRelsFile.async('string');

    // Extract header/footer relationships from template
    const hfRelRegex = /<Relationship\s+[^>]*Type="[^"]*\/(header|footer)"[^>]*\/>/gi;
    const templateHFRels = [];
    let m;
    while ((m = hfRelRegex.exec(templateRelsXml)) !== null) {
        const rel = m[0];
        const idMatch = rel.match(/Id="([^"]+)"/);
        const targetMatch = rel.match(/Target="([^"]+)"/);
        const typeMatch = rel.match(/Type="([^"]+)"/);
        if (idMatch && targetMatch && typeMatch) {
            templateHFRels.push({
                originalId: idMatch[1],
                target: targetMatch[1],
                type: typeMatch[1],
                isHeader: m[1].toLowerCase() === 'header'
            });
        }
    }

    if (templateHFRels.length === 0) return;

    // Get target's existing rels to find max rId
    const targetRelsFile = targetZip.file('word/_rels/document.xml.rels');
    if (!targetRelsFile) return;
    let targetRelsXml = await targetRelsFile.async('string');

    // Find highest existing rId number in target
    let maxId = 0;
    const allIdRegex = /Id="rId(\d+)"/g;
    let idM;
    while ((idM = allIdRegex.exec(targetRelsXml)) !== null) {
        const num = parseInt(idM[1]);
        if (num > maxId) maxId = num;
    }

    // Remove existing header/footer relationships from target
    targetRelsXml = targetRelsXml.replace(/<Relationship[^>]*Type="[^"]*\/(header|footer)"[^>]*\/>/gi, '');

    // Remove existing header/footer files from target
    const toRemove = [];
    targetZip.folder('word').forEach((path) => {
        if (/^(header|footer)\d*\.xml$/.test(path)) {
            toRemove.push('word/' + path);
        }
    });
    toRemove.forEach(path => {
        targetZip.remove(path);
        const relsPath = 'word/_rels/' + path.replace('word/', '') + '.rels';
        if (targetZip.file(relsPath)) targetZip.remove(relsPath);
    });

    // Assign fresh rId numbers and copy files
    const idMapping = {}; // oldId -> newId
    const newRels = [];

    for (const rel of templateHFRels) {
        maxId++;
        const newId = 'rId' + maxId;
        idMapping[rel.originalId] = newId;

        // Add new relationship entry
        newRels.push(`<Relationship Id="${newId}" Type="${rel.type}" Target="${rel.target}"/>`);

        // Copy the header/footer file
        const filePath = 'word/' + rel.target;
        const file = templateZip.file(filePath);
        if (file) {
            const content = await file.async('uint8array');
            targetZip.file(filePath, content);
        }

        // Copy its .rels file if exists
        const relsPath = 'word/_rels/' + rel.target + '.rels';
        const relsFile = templateZip.file(relsPath);
        if (relsFile) {
            const content = await relsFile.async('uint8array');
            targetZip.file(relsPath, content);
        }
    }

    // Insert new relationships before </Relationships>
    const insertPoint = targetRelsXml.lastIndexOf('</Relationships>');
    if (insertPoint !== -1) {
        targetRelsXml = targetRelsXml.substring(0, insertPoint) +
            newRels.join('\n') + '\n' +
            targetRelsXml.substring(insertPoint);
    }
    targetZip.file('word/_rels/document.xml.rels', targetRelsXml);

    // Copy media referenced by headers/footers
    for (const rel of templateHFRels) {
        const relsPath = 'word/_rels/' + rel.target + '.rels';
        const relsFile = templateZip.file(relsPath);
        if (!relsFile) continue;
        const relsContent = await relsFile.async('string');
        const mediaRegex = /Target="([^"]+)"/g;
        let mediaMatch;
        while ((mediaMatch = mediaRegex.exec(relsContent)) !== null) {
            const mediaTarget = mediaMatch[1];
            if (mediaTarget.startsWith('http')) continue;
            const mediaPath = 'word/' + mediaTarget.replace(/^\.\//, '');
            const mediaFile = templateZip.file(mediaPath);
            if (mediaFile) {
                const content = await mediaFile.async('uint8array');
                targetZip.file(mediaPath, content);
            }
        }
    }

    // Update content types
    await updateContentTypes(targetZip, templateHFRels);

    // Now get template's sectPr to find header/footer reference pattern
    const templateDocFile = templateZip.file('word/document.xml');
    if (!templateDocFile) return;
    const templateDoc = await templateDocFile.async('string');

    // Get the last sectPr from template
    const sectPrMatches = templateDoc.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g);
    if (!sectPrMatches) return;
    const templateSectPr = sectPrMatches[sectPrMatches.length - 1];

    // Extract headerReference/footerReference tags and remap their IDs
    const hfRefRegex = /<w:(headerReference|footerReference)\s+[^>]*\/>/g;
    const templateRefs = [];
    let refMatch;
    while ((refMatch = hfRefRegex.exec(templateSectPr)) !== null) {
        let ref = refMatch[0];
        // Remap the r:id
        const ridMatch = ref.match(/r:id="([^"]+)"/);
        if (ridMatch && idMapping[ridMatch[1]]) {
            ref = ref.replace(ridMatch[1], idMapping[ridMatch[1]]);
        }
        templateRefs.push(ref);
    }

    if (templateRefs.length === 0) return;

    // Also extract page layout properties from template sectPr
    const layoutProps = [];
    const layoutRegex = /<w:(pgSz|pgMar|cols|docGrid)\b[^\/]*\/>/g;
    let layoutMatch;
    while ((layoutMatch = layoutRegex.exec(templateSectPr)) !== null) {
        layoutProps.push(layoutMatch[0]);
    }

    // Update ALL sectPr elements in target document
    const targetDocFile = targetZip.file('word/document.xml');
    if (!targetDocFile) return;
    let targetDoc = await targetDocFile.async('string');

    targetDoc = targetDoc.replace(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g, (sectPr) => {
        // Remove existing header/footer references
        let updated = sectPr.replace(/<w:(headerReference|footerReference)\s+[^>]*\/>/g, '');

        // Remove existing page layout props that we're replacing
        if (layoutProps.length > 0) {
            updated = updated.replace(/<w:(pgSz|pgMar|cols|docGrid)\b[^\/]*\/>/g, '');
        }

        // Insert new refs and layout after the opening tag
        const closingTag = '</w:sectPr>';
        const insertPos = updated.lastIndexOf(closingTag);
        if (insertPos !== -1) {
            const newContent = templateRefs.join('') + layoutProps.join('');
            updated = updated.substring(0, insertPos) + newContent + closingTag;
        }

        return updated;
    });

    targetZip.file('word/document.xml', targetDoc);
}

async function updateContentTypes(targetZip, templateHFRels) {
    const ctFile = targetZip.file('[Content_Types].xml');
    if (!ctFile) return;
    let ctXml = await ctFile.async('string');

    // Remove existing header/footer overrides
    ctXml = ctXml.replace(/<Override[^>]*PartName="\/word\/(header|footer)\d*\.xml"[^>]*\/>/gi, '');

    // Add new ones
    const added = new Set();
    for (const rel of templateHFRels) {
        const partName = '/word/' + rel.target;
        if (added.has(partName)) continue;
        added.add(partName);

        const ct = rel.isHeader
            ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
            : 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

        const override = `<Override PartName="${partName}" ContentType="${ct}"/>`;
        const insertPoint = ctXml.lastIndexOf('</Types>');
        if (insertPoint !== -1) {
            ctXml = ctXml.substring(0, insertPoint) + override + ctXml.substring(insertPoint);
        }
    }

    targetZip.file('[Content_Types].xml', ctXml);
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
