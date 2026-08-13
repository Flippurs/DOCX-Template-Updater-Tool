const JSZip = require('jszip');

/**
 * Transfers document-level formatting from template to target.
 * Text content in the target is NOT modified.
 *
 * What gets transferred:
 * - word/styles.xml (paragraph, character, table styles)
 * - word/theme/theme1.xml (fonts, colors, theme)
 * - word/numbering.xml (list/numbering definitions)
 * - word/settings.xml (document settings)
 * - docProps/core.xml (metadata)
 * - docProps/app.xml (app properties)
 * - word/header*.xml and word/footer*.xml (headers/footers)
 * - Section properties (margins, page size, orientation) from word/document.xml
 *
 * What stays unchanged:
 * - All text content in the body
 * - Images and media
 * - Body structure (paragraphs, tables, content)
 */
async function updateDocuments(templateBuffer, targetBuffer) {
    const templateZip = await JSZip.loadAsync(templateBuffer);
    const targetZip = await JSZip.loadAsync(targetBuffer);

    // 1. Transfer styles.xml
    await transferFile(templateZip, targetZip, 'word/styles.xml');

    // 2. Transfer theme
    await transferFile(templateZip, targetZip, 'word/theme/theme1.xml');

    // 3. Transfer numbering definitions
    await transferFile(templateZip, targetZip, 'word/numbering.xml');

    // 4. Transfer document settings
    await transferFile(templateZip, targetZip, 'word/settings.xml');

    // 5. Transfer document properties
    await transferFile(templateZip, targetZip, 'docProps/core.xml');
    await transferFile(templateZip, targetZip, 'docProps/app.xml');

    // 6. Transfer headers and footers
    await transferHeadersFooters(templateZip, targetZip);

    // 7. Transfer section properties (margins, page size, orientation)
    await transferSectionProperties(templateZip, targetZip);

    // 8. Update content types if needed
    await transferFile(templateZip, targetZip, '[Content_Types].xml');

    // 9. Update relationships for headers/footers
    await transferFile(templateZip, targetZip, 'word/_rels/document.xml.rels');

    // Generate updated file
    return await targetZip.generateAsync({ type: 'nodebuffer' });
}

/**
 * Transfer a file from template to target zip if it exists in template
 */
async function transferFile(templateZip, targetZip, path) {
    const file = templateZip.file(path);
    if (file) {
        const content = await file.async('nodebuffer');
        targetZip.file(path, content);
    }
}

/**
 * Transfer all header and footer XML files from template to target
 */
async function transferHeadersFooters(templateZip, targetZip) {
    // Remove existing headers/footers from target
    targetZip.folder('word').forEach((relativePath, file) => {
        if (/^(header|footer)\d*\.xml$/.test(relativePath)) {
            targetZip.remove('word/' + relativePath);
        }
    });

    // Copy headers/footers from template
    templateZip.folder('word').forEach(async (relativePath, file) => {
        if (/^(header|footer)\d*\.xml$/.test(relativePath)) {
            const content = await file.async('nodebuffer');
            targetZip.file('word/' + relativePath, content);
        }
    });
}

/**
 * Transfer section properties (margins, page size, orientation) from template
 * to target without modifying body text content.
 */
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

    // Replace sectPr in target (last occurrence, which is the document-level sectPr)
    let updatedTarget;
    const targetSectPrMatch = targetDoc.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);

    if (targetSectPrMatch) {
        // Replace the last sectPr
        const lastIndex = targetDoc.lastIndexOf(targetSectPrMatch[0]);
        updatedTarget = targetDoc.substring(0, lastIndex) +
            templateSectPr +
            targetDoc.substring(lastIndex + targetSectPrMatch[0].length);
    } else {
        // Insert before closing body tag
        updatedTarget = targetDoc.replace('</w:body>', templateSectPr + '</w:body>');
    }

    targetZip.file('word/document.xml', updatedTarget);
}

module.exports = { updateDocuments };
