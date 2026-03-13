const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = require('docx');

/**
 * Generates a Word document dynamically from JSON data without a template.
 * Creates a professional document with headings, paragraphs, and bullet points.
 * @param {object} data - The JSON data to convert into a document.
 * @returns {Buffer} - The binary buffer of the generated .docx file.
 */
async function generateDocumentFromJson(data) {
  const doc = new Document({
    sections: [{
      properties: {},
      children: parseJsonToDocxElements(data),
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  return buffer;
}

/**
 * Generates an Excel workbook from JSON data.
 * Expected JSON shape:
 * {
 *   phases: [
 *     { phase: string, items: [ { responsibility, vendor, client, consultant } ] }
 *   ]
 * }
 */
async function generateExcelFromJson(data) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet 1');

  // Define headers
  sheet.addRow(['Phase', 'Responsibility', 'Vendor', 'Client', 'Consultant']);

  if (data && Array.isArray(data.phases)) {
    data.phases.forEach(phase => {
      const phaseName = phase.phase || '';
      const items = Array.isArray(phase.items) ? phase.items : [];
      items.forEach(item => {
        sheet.addRow([
          phaseName,
          item.responsibility || '',
          item.vendor || '',
          item.client || '',
          item.consultant || '',
        ]);
      });
    });
  }

  // Style header
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true };
  sheet.columns.forEach(column => {
    column.width = 25;
  });

  return workbook.xlsx.writeBuffer();
}

/**
 * Recursively parses JSON data and converts it to DOCX elements.
 * @param {any} data - The JSON data to parse.
 * @param {number} level - The current nesting level for headings.
 * @returns {Array} - Array of DOCX elements.
 */
function parseJsonToDocxElements(data, level = 0) {
  const elements = [];

  if (typeof data === 'string') {
    // Check if the string is valid JSON, if so, parse it recursively
    try {
      const parsed = JSON.parse(data);
      elements.push(...parseJsonToDocxElements(parsed, level));
    } catch (e) {
      // Not JSON, treat as regular string
      elements.push(
        new Paragraph({
          children: [new TextRun({ text: data, size: 24 })],
          spacing: { after: 200 },
        })
      );
    }
  } else if (Array.isArray(data)) {
    // Arrays become bullet lists
    data.forEach(item => {
      if (typeof item === 'string') {
        elements.push(
          new Paragraph({
            children: [new TextRun({ text: `• ${item}`, size: 24 })],
            indent: { left: 720 }, // Indent for bullets
            spacing: { after: 120 },
          })
        );
      } else {
        // Nested objects in arrays
        elements.push(...parseJsonToDocxElements(item, level + 1));
      }
    });
  } else if (typeof data === 'object' && data !== null) {
    // Objects: keys become headings, values become content
    Object.entries(data).forEach(([key, value]) => {
      // Add heading for the key
      const headingLevel = Math.min(level + 1, 6); // Max heading level 6
      elements.push(
        new Paragraph({
          children: [new TextRun({ text: key.replace(/_/g, ' ').toUpperCase(), bold: true, size: 28 - (headingLevel * 2) })],
          heading: HeadingLevel[`HEADING_${headingLevel}`],
          spacing: { after: 200, before: 400 },
        })
      );

      // Add content for the value
      elements.push(...parseJsonToDocxElements(value, level + 1));
    });
  } else {
    // Other types (numbers, booleans) convert to string
    elements.push(
      new Paragraph({
        children: [new TextRun({ text: String(data), size: 24 })],
        spacing: { after: 200 },
      })
    );
  }

  return elements;
}

/**
 * Generates a Word or PowerPoint document from a template using the provided JSON data.
 * @param {string} templatePath - The absolute path to the .docx or .pptx template file.
 * @param {object} data - The structured JSON data containing values to inject.
 * @returns {Buffer} - The binary buffer of the generated file.
 */
async function generateDocument(templatePath, data) {
  // Check if template exists, if not, generate dynamically
  if (!fs.existsSync(templatePath)) {
    return generateDocumentFromJson(data);
  }

  // Load the template file as a binary stream
  const content = fs.readFileSync(templatePath, 'binary');

  // Load the binary into a PizZip instance
  const zip = new PizZip(content);

  let doc;
  try {
    // Create the docxtemplater instance
    // paragraphLoop: true ensures arrays maintain proper layout spacing
    // linebreaks: true allows \n in strings to create actual linebreaks in word
    doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
  } catch (error) {
    // Catch compilation errors (e.g., malformed template tags)
    throw error;
  }

  // Set the template variables from our JSON payload
  doc.render(data);

  // Get the rendered doc as a zip buffer
  const buf = doc.getZip().generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  });

  return buf;
}

/**
 * Dynamically builds a new PowerPoint presentation inheriting its 
 * themes and master slide definitions from an existing .pptx file.
 * @param {string} masterTemplatePath - The absolute path to the master .pptx template.
 * @param {object} data - The structured JSON n8n data containing values to inject.
 * @returns {Buffer} - The binary buffer of the generated file.
 */
async function generatePresentation(masterTemplatePath, data) {
  const PptxGenJS = require('pptxgenjs');
  let pres = new PptxGenJS();
  
  // NOTE: Currently PptxGenJS does not natively support cloning an existing 
  // PPTX file and injecting custom slides into it seamlessly from a Buffer/Path.
  // It only supports outputting. 
  // For the sake of this implementation plan where we are reading `project_brochure`,
  // `internal_members`, and `external_members`, we define the styling programmatically
  // assuming these mimic the company's "Kickoff Deck" theme colors and layouts.

  // Define Master Slide settings targeting a modern "kickoff" theme
  pres.defineSlideMaster({
      title: "KICKOFF_MASTER",
      background: { color: "FFFFFF" },
      objects: [
          { rect: { x: 0, y: 0, w: "100%", h: 0.8, fill: { color: "1E3A8A" } } },
          { text: { text: "PROJECT KICKOFF", options: { x: 0.5, y: 0.1, w: 9, fontSize: 24, color: "FFFFFF", bold: true } } },
          { rect: { x: 0, y: 5.2, w: "100%", h: 0.4, fill: { color: "1E3A8A" } } },
          { text: { text: "Confidential - Internal Use Only", options: { x: 0.5, y: 5.3, w: 9, fontSize: 10, color: "FFFFFF" } } }
      ]
  });

  // -------------------------------------------------------------
  // SLIDE 1: Title & Executive Summary (Uses Project Brochure)
  // -------------------------------------------------------------
  let slide1 = pres.addSlide({ masterName: "KICKOFF_MASTER" });
  slide1.addText("Executive Summary", { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: "1E3A8A" });
  
  if (data.project_brochure) {
      slide1.addText(data.project_brochure, { 
          x: 0.5, y: 1.8, w: 9, h: 3, 
          fontSize: 16, 
          color: "333333",
          align: "left",
          valign: "top" 
      });
  }

  // -------------------------------------------------------------
  // SLIDE 2: Project Teams (Internal & External Members)
  // -------------------------------------------------------------
  if (data.internal_members || data.external_members) {
      let slide2 = pres.addSlide({ masterName: "KICKOFF_MASTER" });
      slide2.addText("Team Structure", { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: "1E3A8A" });

      if (data.internal_members && Array.isArray(data.internal_members)) {
          slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.7, w: 4.25, h: 3, fill: { color: "F3F4F6" }, line: { color: "D1D5DB", width: 1 } });
          slide2.addText("Internal Members", { x: 0.7, y: 1.9, w: 4, fontSize: 18, bold: true, color: "1F2937" });
          
          let internalText = data.internal_members.map(m => `• ${m.name} | ${m.designation}`).join('\n');
          slide2.addText(internalText, { x: 0.7, y: 2.3, w: 4, h: 2, fontSize: 14, color: "4B5563", valign: "top" });
      }

      if (data.external_members && Array.isArray(data.external_members)) {
          slide2.addShape(pres.ShapeType.rect, { x: 5.25, y: 1.7, w: 4.25, h: 3, fill: { color: "FEF3C7" }, line: { color: "FBBF24", width: 1 } });
          slide2.addText("External Advisors", { x: 5.45, y: 1.9, w: 4, fontSize: 18, bold: true, color: "92400E" });
          
          let externalText = data.external_members.map(m => `• ${m.name} | ${m.designation}`).join('\n');
          slide2.addText(externalText, { x: 5.45, y: 2.3, w: 4, h: 2, fontSize: 14, color: "92400E", valign: "top" });
      }
  }

  const buffer = await pres.write({ outputType: 'nodebuffer' });
  return buffer;
}


module.exports = {
  generateDocument,
  generatePresentation,
  generateExcelFromJson,
};
