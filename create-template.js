const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun } = require('docx');

async function createTemplate() {
    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Project Report Template"),
                    ],
                    heading: "Heading1",
                }),
                new Paragraph("Background: {textbox_project_background}"),
                new Paragraph("Project Scope:"),
                new Paragraph("{#textbox_project_scope} - {.} {/textbox_project_scope}"),
                new Paragraph("Success Factors:"),
                new Paragraph("{#textbox_project_success_factors} - {.} {/textbox_project_success_factors}"),
                new Paragraph("Internal Members:"),
                new Paragraph("{#internal_members} {name} - {designation} {/internal_members}"),
                new Paragraph("External Members:"),
                new Paragraph("{#external_members} {name} - {designation} {/external_members}"),
                new Paragraph("Next Steps - Week 1 (Vendor):"),
                new Paragraph("{#key_next_steps_dependencies.week_01.vendor_responsibilities} * {.} {/key_next_steps_dependencies.week_01.vendor_responsibilities}"),
                new Paragraph("Next Steps - Week 1 (Client):"),
                new Paragraph("{#key_next_steps_dependencies.week_01.client_responsibilities} * {.} {/key_next_steps_dependencies.week_01.client_responsibilities}"),
                new Paragraph("Next Steps - Week 2 (Vendor):"),
                new Paragraph("{#key_next_steps_dependencies.week_02.vendor_responsibilities} * {.} {/key_next_steps_dependencies.week_02.vendor_responsibilities}"),
                new Paragraph("Next Steps - Week 2 (Client):"),
                new Paragraph("{#key_next_steps_dependencies.week_02.client_responsibilities} * {.} {/key_next_steps_dependencies.week_02.client_responsibilities}"),
            ],
        }],
    });

    const buffer = await Packer.toBuffer(doc);
    const templatesDir = path.join(__dirname, 'templates');
    if (!fs.existsSync(templatesDir)) {
        fs.mkdirSync(templatesDir);
    }
    fs.writeFileSync(path.join(templatesDir, 'report-template.docx'), buffer);
    console.log('Template created at templates/report-template.docx');
}

createTemplate().catch(console.error);
