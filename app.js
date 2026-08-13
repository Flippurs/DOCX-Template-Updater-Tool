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
            targetFiles = files;
            renderFileList('targetFiles', targetFiles);
            zone.classList.add('has-files');
        }

        checkReady();
    });

    zone.addEventListener('click', () => {
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
                targetFiles = files;
                renderFileList('targetFiles', targetFiles);
                zone.classList.add('has-files');
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

    // 1. Transfer styles (paragraph, character, table styles)
    await transferFile(templateZip, targetZip, 'word/styles.xml');

    // 2. Transfer theme (fonts, colors)
    await transferFile(templateZip, targetZip, 'word/theme/theme1.xml');

    // 3. Transfer numbering/list definitions
    await transferFile(templateZip, targetZip, 'word/numbering.xml');

    // 4. Transfer section properties (margins, page size, orientation)
    await transferSectionProperties(templateZip, targetZip);

    // 5. Transfer headers and footers + merge relationships safely
    await transferHeadersFootersSafe(templateZip, targetZip);

    // 6. Strip comments from the document
    await stripComments(targetZip);

    // 7. Accept/remove tracked changes (revisions) for a clean document
    await stripRevisions(targetZip);

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
 * Removes all comments from the document.
 * - Deletes word/comments.xml and word/commentsExtended.xml
 * - Removes comment markers from document body
 * - Removes comment relationships from .rels
 * - Cleans up content types
 */
async function stripComments(targetZip) {
    // Remove comment XML files
    ['word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml',
     'word/commentsExtensible.xml'].forEach(path => {
        if (targetZip.file(path)) {
            targetZip.remove(path);
        }
    });

    // Remove comment rels files
    ['word/_rels/comments.xml.rels'].forEach(path => {
        if (targetZip.file(path)) {
            targetZip.remove(path);
        }
    });

    // Remove comment markers from document.xml
    const docFile = targetZip.file('word/document.xml');
    if (docFile) {
        let docXml = await docFile.async('string');
        // Remove <w:commentRangeStart>, <w:commentRangeEnd>, and <w:commentReference> tags
        docXml = docXml.replace(/<w:commentRangeStart[^>]*\/>/g, '');
        docXml = docXml.replace(/<w:commentRangeEnd[^>]*\/>/g, '');
        docXml = docXml.replace(/<w:commentReference[^>]*\/>/g, '');
        // Remove entire <w:r> runs that only contain a commentReference
        docXml = docXml.replace(/<w:r[^>]*>\s*<w:rPr[^>]*\/>\s*<w:commentReference[^>]*\/>\s*<\/w:r>/g, '');
        docXml = docXml.replace(/<w:r>\s*<w:commentReference[^>]*\/>\s*<\/w:r>/g, '');
        targetZip.file('word/document.xml', docXml);
    }

    // Remove comment relationships from document.xml.rels
    const relsFile = targetZip.file('word/_rels/document.xml.rels');
    if (relsFile) {
        let relsXml = await relsFile.async('string');
        relsXml = relsXml.replace(
            /<Relationship[^>]*Type="[^"]*\/comments[^"]*"[^>]*\/>/gi, ''
        );
        targetZip.file('word/_rels/document.xml.rels', relsXml);
    }

    // Remove comment content types
    const ctFile = targetZip.file('[Content_Types].xml');
    if (ctFile) {
        let ctXml = await ctFile.async('string');
        ctXml = ctXml.replace(
            /<Override[^>]*PartName="[^"]*comments[^"]*"[^>]*\/>/gi, ''
        );
        targetZip.file('[Content_Types].xml', ctXml);
    }
}

/**
 * Removes tracked changes (revisions) markup for a clean document.
 * Accepts insertions and removes deletions.
 */
async function stripRevisions(targetZip) {
    const docFile = targetZip.file('word/document.xml');
    if (!docFile) return;

    let docXml = await docFile.async('string');

    // Remove <w:del>...</w:del> entirely (deleted text goes away)
    docXml = docXml.replace(/<w:del\b[\s\S]*?<\/w:del>/g, '');

    // Unwrap <w:ins>...</w:ins> (keep the inserted content, remove the ins wrapper)
    docXml = docXml.replace(/<w:ins\b[^>]*>/g, '');
    docXml = docXml.replace(/<\/w:ins>/g, '');

    // Remove rPr revision markers
    docXml = docXml.replace(/<w:rPrChange[\s\S]*?<\/w:rPrChange>/g, '');
    docXml = docXml.replace(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/g, '');
    docXml = docXml.replace(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/g, '');
    docXml = docXml.replace(/<w:tblPrChange[\s\S]*?<\/w:tblPrChange>/g, '');
    docXml = docXml.replace(/<w:tcPrChange[\s\S]*?<\/w:tcPrChange>/g, '');
    docXml = docXml.replace(/<w:trPrChange[\s\S]*?<\/w:trPrChange>/g, '');

    targetZip.file('word/document.xml', docXml);
}

async function transferHeadersFootersSafe(templateZip, targetZip) {
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

    // Copy header/footer XML files from template
    for (const relativePath of templateHF) {
        const content = await templateZip.file('word/' + relativePath).async('uint8array');
        targetZip.file('word/' + relativePath, content);
    }

    // Copy header/footer relationship files
    for (const relativePath of templateHF) {
        const relsPath = 'word/_rels/' + relativePath + '.rels';
        const relsFile = templateZip.file(relsPath);
        if (relsFile) {
            const content = await relsFile.async('uint8array');
            targetZip.file(relsPath, content);
        }
    }

    // Merge header/footer rels into target's document.xml.rels
    await mergeHeaderFooterRels(templateZip, targetZip);

    // Merge content types for headers/footers
    await mergeContentTypes(templateZip, targetZip, templateHF);
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

    if (hfRels.length > 0) {
        const insertPoint = targetRels.lastIndexOf('</Relationships>');
        if (insertPoint !== -1) {
            targetRels = targetRels.substring(0, insertPoint) +
                hfRels.join('\n') + '\n' +
                targetRels.substring(insertPoint);
        }
    }

    targetZip.file('word/_rels/document.xml.rels', targetRels);
}

async function mergeContentTypes(templateZip, targetZip, templateHF) {
    const targetCTFile = targetZip.file('[Content_Types].xml');
    if (!targetCTFile) return;

    let targetCT = await targetCTFile.async('string');

    for (const relativePath of templateHF) {
        const partName = '/word/' + relativePath;
        if (!targetCT.includes(partName)) {
            const type = relativePath.startsWith('header') ? 'header' : 'footer';
            const override = `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
            const insertPoint = targetCT.lastIndexOf('</Types>');
            if (insertPoint !== -1) {
                targetCT = targetCT.substring(0, insertPoint) +
                    override + '\n' +
                    targetCT.substring(insertPoint);
            }
        }
    }

    targetZip.file('[Content_Types].xml', targetCT);
}

async function transferSectionProperties(templateZip, targetZip) {
    const templateDocFile = templateZip.file('word/document.xml');
    const targetDocFile = targetZip.file('word/document.xml');

    if (!templateDocFile || !targetDocFile) return;

    const templateDoc = await templateDocFile.async('string');
    const targetDoc = await targetDocFile.async('string');

    const sectPrRegex = /<w:sectPr[\s\S]*?<\/w:sectPr>/g;
    let templateSectPr = null;
    let match;
    while ((match = sectPrRegex.exec(templateDoc)) !== null) {
        templateSectPr = match[0];
    }
    if (!templateSectPr) return;

    const targetSectPrRegex = /<w:sectPr[\s\S]*?<\/w:sectPr>/g;
    let lastMatch = null;
    while ((match = targetSectPrRegex.exec(targetDoc)) !== null) {
        lastMatch = match;
    }

    let updatedTarget;
    if (lastMatch) {
        updatedTarget = targetDoc.substring(0, lastMatch.index) +
            templateSectPr +
            targetDoc.substring(lastMatch.index + lastMatch[0].length);
    } else {
        updatedTarget = targetDoc.replace('</w:body>', templateSectPr + '</w:body>');
    }

    targetZip.file('word/document.xml', updatedTarget);
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
