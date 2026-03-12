const data = {
  "textbox_project_background": "This project aims to automate our entire contract lifecycle using AI.",
  "textbox_project_scope": [
    "Identify current bottlenecks",
    "Develop initial proof of concept",
    "Deploy microservice to production"
  ],
  "textbox_project_success_factors": [
    "Over 90% document generation accuracy",
    "Under 2 seconds generation latency"
  ],
  "internal_members": [
    { "name": "Jane Doe", "designation": "Project Manager" },
    { "name": "John Smith", "designation": "Lead Developer" }
  ],
  "external_members": [
    { "name": "Alice Wonderland", "designation": "Vendor Consultant" }
  ],
  "key_next_steps_dependencies": {
    "week_01": {
      "vendor_responsibilities": [
        "Provide server access",
        "Sign NDA"
      ],
      "client_responsibilities": [
        "Provide sample templates",
        "Define JSON schema"
      ]
    },
    "week_02": {
      "vendor_responsibilities": [
        "Deploy beta version"
      ],
      "client_responsibilities": [
        "Perform UAT testing"
      ]
    }
  }
};

async function testGeneration() {
  console.log("Sending POST request to localhost:3000/generate-document...");
  try {
    const response = await fetch('http://localhost:3000/generate-document', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(data)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Server returned ${response.status}: ${errorText}`);
    }

    const buffer = await response.arrayBuffer();
    const fs = require('fs');
    fs.writeFileSync('test-output.docx', Buffer.from(buffer));
    console.log("Success! File saved as test-output.docx");
  } catch (err) {
    console.error("Test failed:", err.message);
  }
}

testGeneration();
