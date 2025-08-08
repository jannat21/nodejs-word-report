const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

// Load the docx file as binary content
const content = fs.readFileSync(
  path.resolve(__dirname, "template.docx"),
  "binary"
);

// Create a new PizZip instance with the template content
const zip = new PizZip(content);

// Create a new Docxtemplater instance
const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
});

// Data to fill in the template
const data = {
  name: "علی احمدی",
  company: "شرکت فناوری نوین",
  items: [
    { item: "محصول A" },
    { item: "محصول B" },
    { item: "محصول C" },
  ],
};

// Render the document with the provided data
doc.render(data);

// Generate the output document
const buf = doc.getZip().generate({
  type: "nodebuffer",
  compression: "DEFLATE",
});

// Save the output document
fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
console.log("Report generated successfully as output.docx");