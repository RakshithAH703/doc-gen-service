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
        // Return JSON containing the raw base64 string
        const base64String = generatedDoc.toString('base64');
        return res.json({
            fileName: fileName,
            contentType: contentType,
            base64: base64String
        });
    } else {
        // Set headers to trigger a direct file download stream
        res.setHeader('Content-Type', contentType);
        res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);

        // Send the generated file buffer
        return res.send(generatedDoc);
    }
  } catch (error) {
    console.error('Error generating document:', error);
    res.status(500).json({ error: 'Failed to generate document', details: error.message });
  }
});

module.exports = router;
