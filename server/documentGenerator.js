const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

/**
 * Generates a Word document from a template using the provided JSON data.
 * @param {string} templatePath - The absolute path to the .docx template file.
 * @param {object} data - The structured JSON data containing values to inject.
 * @returns {Buffer} - The binary buffer of the generated .docx file.
 */
async function generateDocument(templatePath, data) {
  // Load the docx file as a binary stream
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

module.exports = {
  generateDocument
};
