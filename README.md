# Document Generation Microservice

This service provides an HTTP REST API to generate `.docx` (Word documents) from a structured JSON payload. It uses `docxtemplater` to inject data into placeholder variables located within a template Word document.

## Running the Service

1. Ensure you have [Node.js](https://nodejs.org/) installed (v18+ recommended).
2. Install dependencies:
   ```bash
   npm install
   ```
3. Start the internal server:
   ```bash
   npm start
   ```
   Or explicitly using node:
   ```bash
   node server/server.js
   ```
   *The server runs on http://localhost:3000 by default.*

## Creating the Template (`/templates/report-template.docx`)

You must place a Word document named `report-template.docx` in the `/templates` folder.

### Example Template Placeholders

In your Word document, type out the tags like this. They will be formatted dynamically by docxtemplater based on your JSON values:

**Simple string mapping:**
Project Background: `{textbox_project_background}`

**Looping an array of strings:**
Project Scope:
`{#textbox_project_scope}`
• `{.}`
`{/textbox_project_scope}`

**Looping an array of objects:**
Team Members:
`{#internal_members}`
Name: `{name}` - `{designation}`
`{/internal_members}`

**Accessing nested objects:**
Week 1 Vendor Responsibilities:
`{#key_next_steps_dependencies.week_01.vendor_responsibilities}`
- `{.}`
`{/key_next_steps_dependencies.week_01.vendor_responsibilities}`

*Note: The array loop syntax `{array}` and `{/array}` controls what text inside is repeated.*

## API Request Example

You can test the generation using an HTTP client like curl or Postman:

```bash
curl -X POST http://localhost:3000/generate-document \
-H "Content-Type: application/json" \
-d '{
  "textbox_project_background": "This project aims to revitalize the core infrastructure.",
  "textbox_project_scope": ["Backend API", "Frontend UX", "Database Migration"],
  "internal_members": [{"name": "Alice", "designation": "Manager"}],
  "key_next_steps_dependencies": {
    "week_01": {
      "vendor_responsibilities": ["Setup servers", "Configure domains"]
    }
  }
}' --output generated-report.docx
```

## Frontend Integration Guide

A frontend application can call the API and automatically download the generated document using the `fetch` API:

```javascript
async function downloadGeneratedDocument(jsonData) {
  try {
    const response = await fetch('http://localhost:3000/generate-document', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(jsonData)
    });

    if (!response.ok) {
      throw new Error(`Error: ${response.statusText}`);
    }

    // Convert response to an invisible blob / file
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    
    // Create an invisible anchor tag to trigger the browser's download prompt
    const a = document.createElement('a');
    a.href = url;
    a.download = 'custom-report.docx';
    document.body.appendChild(a);
    a.click();
    
    // Cleanup
    a.remove();
    window.URL.revokeObjectURL(url);
  } catch (error) {
    console.error("Failed to generate document:", error);
  }
}
```

## Workflow (n8n) Integration

This service shines when placed inside an n8n automation pipeline.

1. **Trigger**: An HTTP Trigger or Webhook fires in n8n when the frontend uploads a document.
2. **AI Processing**: An n8n AI node (like an OpenAI LLM Chain) processes the raw text from the document and outputs a perfectly structured JSON object conforming to your template schema.
3. **HTTP Request Node**: The AI output is passed into an n8n HTTP Request node:
   - **Method**: `POST`
   - **URL**: `http://<your-service-ip>:3000/generate-document`
   - **Authentication**: (None by default, add headers if you implement auth)
   - **Body Type**: `json`
   - **Body Parameters**: Map the JSON output of the AI node into the body.
   - **Response Type**: `File`
4. **Conclusion**: n8n receives the binary file back from this API and can either save it to Google Drive, email it, or pass the binary response back to the frontend webhook origin.

## Internal Rendering Logic

This microservice uses `pizzip` and `docxtemplater` to manipulate the document file internally. Here is how it works under the hood:

1. A DOCX file is secretly just a disguised `.zip` archive holding an intricate web of XML files.
2. `pizzip` unzips the `report-template.docx` entirely in memory (RAM).
3. `docxtemplater` parses the raw XML string of the document, searching for string literals surrounded by `{}`. 
4. It recursively compiles those literals, injecting the matching JSON payload variables. For arrays, it creates copies of the XML nodes to loop the content properly.
5. The modified XML tree is then re-architected and zipped back up into a `Buffer` mimicking a DOCX executable file, and streamed back to the requester via Express.
