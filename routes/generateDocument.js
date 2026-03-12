const express = require('express');
const router = express.Router();
const { generateDocument } = require('../server/documentGenerator');
const path = require('path');

router.post('/', async (req, res) => {
  try {
    const data = req.body;
    
    // Validate that we received an object
    if (!data || typeof data !== 'object') {
      return res.status(400).json({ error: 'Invalid JSON payload provided' });
    }

    // Path to the template
    const templatePath = path.join(__dirname, '../templates/report-template.docx');

    // Generate the document
    const generatedDoc = await generateDocument(templatePath, data);

    // Set headers to trigger a file download in the browser
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=generated-report.docx');

    // Send the generated file buffer
    res.send(generatedDoc);
  } catch (error) {
    console.error('Error generating document:', error);
    res.status(500).json({ error: 'Failed to generate document', details: error.message });
  }
});

module.exports = router;
