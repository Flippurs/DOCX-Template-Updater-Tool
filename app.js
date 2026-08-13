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

/**
 * Transfers document-level formatting from template to target.
 * Text content in the target is NOT modified.
 */
async function updateDocument(templateZip, targetBuffer) {
    const targetZip = await JSZip.loadAsync(targetBuffer);

    // Transfer styles
    await transferFile(templateZip, targetZip, 'word/styles.xml');

    // Transfer theme (fonts, colors)
    await transferFile(templateZip, targetZip, 'word/theme/theme1.xml');

    // Transfer numbering/list definitions
    await transferFile(templateZip, targetZip, 'word/numbering.xml');

    // Transfer document settings
    await transferFile(templateZip, targetZip, 'word/settings.xml');

    // Transfer document properties/metadata
    await transferFile(templateZip, targetZip, 'docProps/core.xml');
    await transferFile(templateZip, targetZip, 'docProps/app.xml');

    // Transfer headers and footers
    await transferHeadersFooters(templateZip, targetZip);

    // Transfer section properties (margins, page size, orientation)
    await transferSectionProperties(templateZip, targetZip);

    // Update content types
    await transferFile(templateZip, targetZip, '[Content_Types].xml');

    // Update relationships for headers/footers
    await transferFile(templateZip, targetZip, 'word/_rels/document.xml.rels');

    return await targetZip.generateAsync({ type: 'blob' });
}

async function transferFile(templateZip, targetZip, path) {
    const file = templateZip.file(path);
    if (file) {
        const content = await file.async('uint8array');
        targetZip.file(path, content);
    }
}

async function transferHeadersFooters(templateZip, targetZip) {
    // Remove existing headers/footers from target
    const toRemove = [];
    targetZip.folder('word').forEach((relativePath) => {
        if (/^(header|footer)\d*\.xml$/.test(relativePath)) {
            toRemove.push('word/' + relativePath);
        }
    });
    toRemove.forEach(path => targetZip.remove(path));

    // Copy headers/footers from template
    const toCopy = [];
    templateZip.folder('word').forEach((relativePath) => {
        if (/^(header|footer)\d*\.xml$/.test(relativePath)) {
            toCopy.push(relativePath);
        }
    });

    for (const relativePath of toCopy) {
        const content = await templateZip.file('word/' + relativePath).async('uint8array');
        targetZip.file('word/' + relativePath, content);
    }
}

async function transferSectionProperties(templateZip, targetZip) {
    const templateDocFile = templateZip.file('word/document.xml');
    const targetDocFile = targetZip.file('word/document.xml');

    if (!templateDocFile || !targetDocFile) return;

    const templateDoc = await templateDocFile.async('string');
    const targetDoc = await targetDocFile.async('string');

    // Extract sectPr from template
    const sectPrMatch = templateDoc.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
    if (!sectPrMatch) return;

    const templateSectPr = sectPrMatch[0];

    // Replace last sectPr in target
    const targetSectPrRegex = /<w:sectPr[\s\S]*?<\/w:sectPr>/g;
    let lastMatch = null;
    let match;
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
            // Single file - download directly
            const targetBuffer = await targetFiles[0].arrayBuffer();
            const updatedBlob = await updateDocument(templateZip, targetBuffer);
            const newName = targetFiles[0].name.replace('.docx', '_updated.docx');
            saveAs(updatedBlob, newName);
        } else {
            // Multiple files - zip them
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
