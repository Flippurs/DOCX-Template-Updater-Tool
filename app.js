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

    // 1. Strip comments FIRST (before any other processing)
    await stripComments(targetZip);

    // 2. Strip tracked changes / revisions
    await stripRevisions(targetZip);

    // 3. Transfer styles
    await transferFile(templateZip, targetZip, 'word/styles.xml');

    // 4. Transfer theme (fonts, colors)
    await transferFile(templateZip, targetZip, 'word/theme/theme1.xml');

    // 5. Transfer numbering/list definitions
    await transferFile(templateZip, targetZip, 'word/numbering.xml');

    // 6. Transfer headers and footers + update ALL section references
    await transferHeadersFootersFull(templateZip, targetZip);

    // 7. Transfer section properties (margins, page size) to ALL sections
    await transferSectionProperties(templateZip, targetZip);

    return await targetZip.generateAsync({ type: 'blob' });
}

async function transferFile(templateZip, targetZip, path) {
    const file = templateZip.file(path);
    if (file) {
        const content = await file.async('uint8array');
        targetZip.file(path, content);
    }
}

/**
 * Thoroughly strips ALL comments from the document.
 * Handles all known comment XML variants and markup patterns.
 */
async function stripComments(targetZip) {
    // Remove all comment-related XML files
    const commentFiles = [
        'word/comments.xml',
        'word/commentsExtended.xml',
        'word/commentsIds.xml',
        'word/commentsExtensible.xml',
        'word/people.xml',
        'word/_rels/comments.xml.rels',
        'word/_rels/commentsExtended.xml.rels',
        'word/_rels/commentsIds.xml.rels',
        'word/_rels/commentsExtensible.xml.rels'
    ];
    commentFiles.forEach(path => {
        if (targetZip.file(path)) targetZip.remove(path);
    });

    // Clean document.xml - remove all comment markers
    const docFile = targetZip.file('word/document.xml');
    if (docFile) {
        let xml = await docFile.async('string');

        // Self-closing comment tags
        xml = xml.replace(/<w:commentRangeStart[^>]*\/>/g, '');
        xml = xml.replace(/<w:commentRangeEnd[^>]*\/>/g, '');
        xml = xml.replace(/<w:commentReference[^>]*\/>/g, '');
        xml = xml.replace(/<w:annotationRef[^>]*\/>/g, '');

        // Non-self-closing comment tags (some DOCX files use these)
        xml = xml.replace(/<w:commentRangeStart[^>]*>[^<]*<\/w:commentRangeStart>/g, '');
        xml = xml.replace(/<w:commentRangeEnd[^>]*>[^<]*<\/w:commentRangeEnd>/g, '');

        // Word 2013+ extended comment tags (w15 namespace)
        xml = xml.replace(/<w15:commentEx[^>]*\/>/g, '');
        xml = xml.replace(/<w15:commentEx[^>]*>[\s\S]*?<\/w15:commentEx>/g, '');

        // Word 16 comment ID tags
        xml = xml.replace(/<w16cid:commentId[^>]*\/>/g, '');
        xml = xml.replace(/<w16cid:commentId[^>]*>[\s\S]*?<\/w16cid:commentId>/g, '');

        // Remove <w:r> runs that ONLY contain a commentReference (with or without rPr)
        xml = xml.replace(/<w:r\b[^>]*>\s*(?:<w:rPr>[\s\S]*?<\/w:rPr>\s*)?<w:commentReference[^>]*\/>\s*<\/w:r>/g, '');
        xml = xml.replace(/<w:r\b[^>]*>\s*(?:<w:rPr[^>]*\/>\s*)?<w:commentReference[^>]*\/>\s*<\/w:r>/g, '');

        // Remove <w:r> runs that ONLY contain an annotationRef
        xml = xml.replace(/<w:r\b[^>]*>\s*(?:<w:rPr>[\s\S]*?<\/w:rPr>\s*)?<w:annotationRef[^>]*\/>\s*<\/w:r>/g, '');
        xml = xml.replace(/<w:r\b[^>]*>\s*(?:<w:rPr[^>]*\/>\s*)?<w:annotationRef[^>]*\/>\s*<\/w:r>/g, '');

        targetZip.file('word/document.xml', xml);
    }

    // Clean header and footer files too (comments can appear there)
    const allFiles = Object.keys(targetZip.files);
    for (const filePath of allFiles) {
        if (/^word\/(header|footer)\d*\.xml$/.test(filePath)) {
            const file = targetZip.file(filePath);
            if (file) {
                let xml = await file.async('string');
                xml = xml.replace(/<w:commentRangeStart[^>]*\/>/g, '');
                xml = xml.replace(/<w:commentRangeEnd[^>]*\/>/g, '');
                xml = xml.replace(/<w:commentReference[^>]*\/>/g, '');
                xml = xml.replace(/<w:annotationRef[^>]*\/>/g, '');
                xml = xml.replace(/<w:commentRangeStart[^>]*>[^<]*<\/w:commentRangeStart>/g, '');
                xml = xml.replace(/<w:commentRangeEnd[^>]*>[^<]*<\/w:commentRangeEnd>/g, '');
                xml = xml.replace(/<w15:commentEx[^>]*\/>/g, '');
                xml = xml.replace(/<w16cid:commentId[^>]*\/>/g, '');
                xml = xml.replace(/<w:r\b[^>]*>\s*(?:<w:rPr>[\s\S]*?<\/w:rPr>\s*)?<w:commentReference[^>]*\/>\s*<\/w:r>/g, '');
                xml = xml.replace(/<w:r\b[^>]*>\s*(?:<w:rPr[^>]*\/>\s*)?<w:commentReference[^>]*\/>\s*<\/w:r>/g, '');
                targetZip.file(filePath, xml);
            }
        }
    }

    // Remove comment relationships from document.xml.rels
    const relsFile = targetZip.file('word/_rels/document.xml.rels');
    if (relsFile) {
        let relsXml = await relsFile.async('string');
        relsXml = relsXml.replace(/<Relationship[^>]*Type="[^"]*\/(comments|commentsExtended|commentsIds|commentsExtensible|people)[^"]*"[^>]*\/>/gi, '');
        targetZip.file('word/_rels/document.xml.rels', relsXml);
    }

    // Remove comment content types
    const ctFile = targetZip.file('[Content_Types].xml');
    if (ctFile) {
        let ctXml = await ctFile.async('string');
        ctXml = ctXml.replace(/<Override[^>]*PartName="[^"]*(comments|commentsExtended|commentsIds|commentsExtensible|people)[^"]*"[^>]*\/>/gi, '');
        targetZip.file('[Content_Types].xml', ctXml);
    }
}

/**
 * Strips tracked changes (revisions) - accepts insertions, removes deletions.
 * Also cleans headers/footers.
 */
async function stripRevisions(targetZip) {
    // Process document.xml and all headers/footers
    const filesToClean = ['word/document.xml'];
    const allFiles = Object.keys(targetZip.files);
    for (const filePath of allFiles) {
        if (/^word\/(header|footer)\d*\.xml$/.test(filePath)) {
            filesToClean.push(filePath);
        }
    }

    for (const filePath of filesToClean) {
        const file = targetZip.file(filePath);
        if (!file) continue;

        let xml = await file.async('string');

        // Remove deleted content entirely
        xml = xml.replace(/<w:del\b[\s\S]*?<\/w:del>/g, '');

        // Unwrap inserted content (keep content, remove wrapper)
        xml = xml.replace(/<w:ins\b[^>]*>/g, '');
        xml = xml.replace(/<\/w:ins>/g, '');

        // Remove move-from (content that was moved away)
        xml = xml.replace(/<w:moveFrom\b[\s\S]*?<\/w:moveFrom>/g, '');

        // Unwrap move-to (keep the content that was moved here)
        xml = xml.replace(/<w:moveTo\b[^>]*>/g, '');
        xml = xml.replace(/<\/w:moveTo>/g, '');

        // Remove revision property changes
        xml = xml.replace(/<w:rPrChange[\s\S]*?<\/w:rPrChange>/g, '');
        xml = xml.replace(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/g, '');
        xml = xml.replace(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/g, '');
        xml = xml.replace(/<w:tblPrChange[\s\S]*?<\/w:tblPrChange>/g, '');
        xml = xml.replace(/<w:tcPrChange[\s\S]*?<\/w:tcPrChange>/g, '');
        xml = xml.replace(/<w:trPrChange[\s\S]*?<\/w:trPrChange>/g, '');
        xml = xml.replace(/<w:tblGridChange[\s\S]*?<\/w:tblGridChange>/g, '');

        // Remove bookmark-style move markers
        xml = xml.replace(/<w:moveFromRangeStart[^>]*\/>/g, '');
        xml = xml.replace(/<w:moveFromRangeEnd[^>]*\/>/g, '');
        xml = xml.replace(/<w:moveToRangeStart[^>]*\/>/g, '');
        xml = xml.replace(/<w:moveToRangeEnd[^>]*\/>/g, '');

        // Remove rsid attributes (revision save IDs) - cosmetic cleanup
        xml = xml.replace(/\s+w:rsid\w*="[^"]*"/g, '');

        targetZip.file(filePath, xml);
    }
}

/**
 * Transfers headers/footers from template and updates ALL sectPr references
 * in the target document (not just the last one).
 */
async function transferHeadersFootersFull(templateZip, targetZip) {
    // Get template header/footer files
    const templateHF = [];
    templateZip.folder('word').forEach((relativePath) => {
        if (/^(header|footer)\d*\.xml$/.test(relativePath)) {
            templateHF.push(relativePath);
        }
    });

    if (templateHF.length === 0) return;

    // Remove existing headers/footers from target
    const toRemove = [];
    targetZip.folder('word').forEach((relativePath) => {
        if (/^(header|footer)\d*\.xml$/.test(relativePath)) {
            toRemove.push('word/' + relativePath);
        }
    });
    toRemove.forEach(path => targetZip.remove(path));

    // Also remove their .rels files
    toRemove.forEach(path => {
        const relsPath = path.replace('word/', 'word/_rels/') + '.rels';
        if (targetZip.file(relsPath)) targetZip.remove(relsPath);
    });

    // Copy header/footer XML files from template
    for (const relativePath of templateHF) {
        const content = await templateZip.file('word/' + relativePath).async('uint8array');
        targetZip.file('word/' + relativePath, content);
    }

    // Copy header/footer .rels files from template
    for (const relativePath of templateHF) {
        const relsPath = 'word/_rels/' + relativePath + '.rels';
        const relsFile = templateZip.file(relsPath);
        if (relsFile) {
            const content = await relsFile.async('uint8array');
            targetZip.file(relsPath, content);
        }
    }

    // Copy any media referenced by template headers/footers
    await copyHeaderFooterMedia(templateZip, targetZip, templateHF);

    // Merge header/footer relationships into target's document.xml.rels
    await mergeHeaderFooterRels(templateZip, targetZip);

    // Update content types
    await mergeContentTypes(templateZip, targetZip, templateHF);

    // Extract template's header/footer references from its sectPr
    const templateDoc = await templateZip.file('word/document.xml').async('string');
    const templateSectPrMatch = templateDoc.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g);
    if (!templateSectPrMatch) return;

    // Get the last (document-level) sectPr from template
    const templateSectPr = templateSectPrMatch[templateSectPrMatch.length - 1];

    // Extract header/footer reference tags from template sectPr
    const hfRefRegex = /<w:(headerReference|footerReference)[^>]*\/>/g;
    const templateHFRefs = [];
    let m;
    while ((m = hfRefRegex.exec(templateSectPr)) !== null) {
        templateHFRefs.push(m[0]);
    }

    if (templateHFRefs.length === 0) return;

    // Now update ALL sectPr elements in target document
    const targetDocFile = targetZip.file('word/document.xml');
    if (!targetDocFile) return;
    let targetDoc = await targetDocFile.async('string');

    // Replace header/footer references in every sectPr
    targetDoc = targetDoc.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, (sectPr) => {
        // Remove existing header/footer references
        let cleaned = sectPr.replace(/<w:(headerReference|footerReference)[^>]*\/>/g, '');

        // Insert template's header/footer references after the opening <w:sectPr...> tag
        const openTagEnd = cleaned.indexOf('>') + 1;
        cleaned = cleaned.substring(0, openTagEnd) +
            '\n' + templateHFRefs.join('\n') + '\n' +
            cleaned.substring(openTagEnd);

        return cleaned;
    });

    targetZip.file('word/document.xml', targetDoc);
}

/**
 * Copy media files (images, etc.) referenced by template headers/footers
 */
async function copyHeaderFooterMedia(templateZip, targetZip, templateHF) {
    for (const relativePath of templateHF) {
        const relsPath = 'word/_rels/' + relativePath + '.rels';
        const relsFile = templateZip.file(relsPath);
        if (!relsFile) continue;

        const relsXml = await relsFile.async('string');
        const targetRegex = /Target="([^"]+)"/g;
        let match;
        while ((match = targetRegex.exec(relsXml)) !== null) {
            const target = match[1];
            // Skip external URLs
            if (target.startsWith('http://') || target.startsWith('https://')) continue;

            // Resolve relative path
            const mediaPath = 'word/' + target.replace(/^\.\//, '');
            const mediaFile = templateZip.file(mediaPath);
            if (mediaFile) {
                const content = await mediaFile.async('uint8array');
                targetZip.file(mediaPath, content);
            }
        }
    }
}

async function mergeHeaderFooterRels(templateZip, targetZip) {
    const templateRelsFile = templateZip.file('word/_rels/document.xml.rels');
    const targetRelsFile = targetZip.file('word/_rels/document.xml.rels');

    if (!templateRelsFile || !targetRelsFile) return;

    let templateRels = await templateRelsFile.async('string');
    let targetRels = await targetRelsFile.async('string');

    // Remove existing header/footer relationships from target
    targetRels = targetRels.replace(
        /<Relationship[^>]*Type="[^"]*\/(header|footer)"[^>]*\/>/gi, ''
    );

    // Extract header/footer relationships from template
    const hfRelRegex = /<Relationship[^>]*Type="[^"]*\/(header|footer)"[^>]*\/>/gi;
    const hfRels = [];
    let match;
    while ((match = hfRelRegex.exec(templateRels)) !== null) {
        hfRels.push(match[0]);
    }

    // Check for ID conflicts and remap if needed
    if (hfRels.length > 0) {
        // Get all existing IDs in target
        const existingIds = new Set();
        const idRegex = /Id="(rId\d+)"/g;
        let idMatch;
        while ((idMatch = idRegex.exec(targetRels)) !== null) {
            existingIds.add(idMatch[1]);
        }

        // Find max rId number
        let maxId = 0;
        existingIds.forEach(id => {
            const num = parseInt(id.replace('rId', ''));
            if (num > maxId) maxId = num;
        });

        // Remap conflicting IDs
        const idMap = {};
        const remappedRels = hfRels.map(rel => {
            const relIdMatch = rel.match(/Id="(rId\d+)"/);
            if (relIdMatch && existingIds.has(relIdMatch[1])) {
                maxId++;
                const newId = 'rId' + maxId;
                idMap[relIdMatch[1]] = newId;
                return rel.replace(relIdMatch[1], newId);
            }
            return rel;
        });

        const insertPoint = targetRels.lastIndexOf('</Relationships>');
        if (insertPoint !== -1) {
            targetRels = targetRels.substring(0, insertPoint) +
                remappedRels.join('\n') + '\n' +
                targetRels.substring(insertPoint);
        }

        // If we remapped IDs, update the references in document.xml
        if (Object.keys(idMap).length > 0) {
            const docFile = targetZip.file('word/document.xml');
            if (docFile) {
                let docXml = await docFile.async('string');
                for (const [oldId, newId] of Object.entries(idMap)) {
                    docXml = docXml.replace(new RegExp(`r:id="${oldId}"`, 'g'), `r:id="${newId}"`);
                }
                targetZip.file('word/document.xml', docXml);
            }
        }
    }

    targetZip.file('word/_rels/document.xml.rels', targetRels);
}

async function mergeContentTypes(templateZip, targetZip, templateHF) {
    const targetCTFile = targetZip.file('[Content_Types].xml');
    if (!targetCTFile) return;

    let targetCT = await targetCTFile.async('string');

    // Remove existing header/footer overrides
    targetCT = targetCT.replace(/<Override[^>]*PartName="\/word\/(header|footer)\d*\.xml"[^>]*\/>/gi, '');

    // Add overrides for template headers/footers
    for (const relativePath of templateHF) {
        const partName = '/word/' + relativePath;
        const type = relativePath.startsWith('header') ? 'header' : 'footer';
        const override = `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
        const insertPoint = targetCT.lastIndexOf('</Types>');
        if (insertPoint !== -1) {
            targetCT = targetCT.substring(0, insertPoint) +
                override + '\n' +
                targetCT.substring(insertPoint);
        }
    }

    targetZip.file('[Content_Types].xml', targetCT);
}

/**
 * Transfers page layout (margins, page size, orientation) from template.
 * Updates ALL section properties, not just the last one.
 * Preserves header/footer references that were already set.
 */
async function transferSectionProperties(templateZip, targetZip) {
    const templateDocFile = templateZip.file('word/document.xml');
    const targetDocFile = targetZip.file('word/document.xml');

    if (!templateDocFile || !targetDocFile) return;

    const templateDoc = await templateDocFile.async('string');
    let targetDoc = await targetDocFile.async('string');

    // Get template's document-level sectPr (last one)
    const sectPrRegex = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/g;
    const templateMatches = templateDoc.match(sectPrRegex);
    if (!templateMatches) return;
    const templateSectPr = templateMatches[templateMatches.length - 1];

    // Extract layout properties from template (pgSz, pgMar, cols, docGrid)
    const layoutTags = ['pgSz', 'pgMar', 'cols', 'docGrid', 'pgBorders', 'lnNumType', 'pgNumType'];
    const templateLayouts = {};
    for (const tag of layoutTags) {
        const selfClose = templateSectPr.match(new RegExp(`<w:${tag}[^>]*\\/>`));
        const withContent = templateSectPr.match(new RegExp(`<w:${tag}[^>]*>[\\s\\S]*?<\\/w:${tag}>`));
        if (selfClose) templateLayouts[tag] = selfClose[0];
        else if (withContent) templateLayouts[tag] = withContent[0];
    }

    // Update each sectPr in target - replace only layout tags, keep references
    targetDoc = targetDoc.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, (sectPr) => {
        let updated = sectPr;

        for (const tag of layoutTags) {
            if (templateLayouts[tag]) {
                // Remove existing tag
                updated = updated.replace(new RegExp(`<w:${tag}[^>]*\\/>`), '');
                updated = updated.replace(new RegExp(`<w:${tag}[^>]*>[\\s\\S]*?<\\/w:${tag}>`), '');

                // Insert template's tag before closing </w:sectPr>
                updated = updated.replace('</w:sectPr>', templateLayouts[tag] + '</w:sectPr>');
            }
        }

        return updated;
    });

    targetZip.file('word/document.xml', targetDoc);
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
