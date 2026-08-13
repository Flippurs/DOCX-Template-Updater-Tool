const functions = require('firebase-functions');
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const { updateDocuments } = require('./docxUpdater');
const JSZip = require('jszip');

const app = express();
app.use(cors({ origin: true }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.post('/update', upload.fields([
    { name: 'template', maxCount: 1 },
    { name: 'targets', maxCount: 50 }
]), async (req, res) => {
    try {
        if (!req.files['template'] || !req.files['targets']) {
            return res.status(400).json({ error: 'Template and at least one target file required' });
        }

        const templateBuffer = req.files['template'][0].buffer;
        const targetFiles = req.files['targets'];

        const results = [];
        for (const target of targetFiles) {
            const updated = await updateDocuments(templateBuffer, target.buffer);
            results.push({
                name: target.originalname.replace('.docx', '_updated.docx'),
                buffer: updated
            });
        }

        // If single file, return directly; if multiple, zip them
        if (results.length === 1) {
            res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.set('Content-Disposition', `attachment; filename="${results[0].name}"`);
            return res.send(results[0].buffer);
        }

        const zip = new JSZip();
        results.forEach(r => zip.file(r.name, r.buffer));
        const zipBuffer = await zip.generateAsync({ type: 'nodebuffer' });

        res.set('Content-Type', 'application/zip');
        res.set('Content-Disposition', 'attachment; filename="updated_documents.zip"');
        res.send(zipBuffer);
    } catch (err) {
        console.error('Error processing documents:', err);
        res.status(500).json({ error: 'Failed to process documents: ' + err.message });
    }
});

exports.api = functions.https.onRequest(app);
