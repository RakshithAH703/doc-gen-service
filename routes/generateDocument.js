const express = require('express');
const router = express.Router();
const { generateDocument, generatePresentation } = require('../server/documentGenerator');
const path = require('path');

router.post('/', async (req, res) => {
  try {
    const data = req.body;
    const type = req.query.type || 'docx'; // Defaults to word document if not specified
    const responseFormat = req.query.responseFormat || 'buffer'; // buffer or base64
    
    // Validate that we received an object
    if (!data || typeof data !== 'object') {
      return res.status(400).json({ error: 'Invalid JSON payload provided' });
    }

    // Determine template path and headers based on the requested type
    let templatePath;
    let contentType;
    let fileName;
    let generatedDoc;

    if (type === 'kickoff') {
      templatePath = path.join(__dirname, '../templates/KickOffdecktemplate.pptx');
      contentType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
      fileName = 'kickoff-presentation.pptx';
      // Use the programmatic builder instead of the text-replacement engine
      generatedDoc = await generatePresentation(templatePath, data);
    } else {
      if (type === 'pptx_full') {
        templatePath = path.join(__dirname, '../templates/n8n-full-template.pptx');
        contentType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
        fileName = 'generated-n8n-presentation.pptx';
      } else if (type === 'pptx') {
        templatePath = path.join(__dirname, '../templates/presentation-template.pptx');
        contentType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
        fileName = 'generated-presentation.pptx';
      } else {
        // Default to docx
        templatePath = path.join(__dirname, '../templates/report-template.docx');
        contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        fileName = 'generated-report.docx';
      }
      // Use standard docxtemplater text-replacement 
      generatedDoc = await generateDocument(templatePath, data);
    }

    // Determine return strategy based on 'responseFormat' query param
    if (responseFormat === 'base64') {
        // Return JSON containing the raw base64 string and an optional viewer helper script
        // so a frontend can render the document directly in the browser.
        const base64 = generatedDoc.toString('base64');

        // Viewer script uses docx-preview (CDN) to render into a container div
        // and also provides a download button for the generated Word file.
        const viewerScript = `// This script is intended to be used in a browser environment.
// Example usage:
// 1) Add <div id="docx-viewer"></div> and <button id="docx-download">Download</button> to your HTML.
// 2) Insert this script after the div/button.
// 3) Ensure you load docx-preview from a CDN (shown below).

// Load docx-preview (unpkg CDN) if not already loaded.
if (typeof docx === 'undefined') {
  const s = document.createElement('script');
  s.src = 'https://unpkg.com/docx-preview@3.2.0/dist/docx-preview.min.js';
  document.head.appendChild(s);
}

function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function downloadDocx(base64, fileName) {
  const bytes = base64ToUint8Array(base64);
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName || 'generated.docx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

(function renderDocx(base64, fileName) {
  const target = document.getElementById('docx-viewer');
  const downloadBtn = document.getElementById('docx-download');

  if (!target) {
    console.warn('No element with id "docx-viewer" found. Add <div id="docx-viewer"></div> to your page.');
  }

  if (!downloadBtn) {
    console.warn('No element with id "docx-download" found. Add <button id="docx-download">Download</button> to your page.');
  }

  if (downloadBtn) {
    downloadBtn.addEventListener('click', () => downloadDocx(base64, fileName));
  }

  const tryRender = () => {
    if (typeof docx !== 'undefined' && docx.renderAsync && target) {
      const bytes = base64ToUint8Array(base64);
      docx.renderAsync(bytes.buffer, target).catch(console.error);
    } else {
      setTimeout(tryRender, 50);
    }
  };

  tryRender();
})("${base64}", "${fileName}");`;
        return res.json({
          fileName,
          contentType,
          base64,
          viewerScript,
        });
    }

    // Default response: send the generated file buffer as a download
    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);

    return res.send(generatedDoc);
  } catch (error) {
    console.error('Error generating document:', error);
    res.status(500).json({ error: 'Failed to generate document', details: error.message });
  }
});

module.exports = router;
