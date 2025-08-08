const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const fs = require("fs");
const path = require("path");

// Check if image file exists
const imagePath = path.resolve(__dirname, "logo.png");
if (!fs.existsSync(imagePath)) {
  console.error("Error: Image file 'logo.png' not found at", imagePath);
  process.exit(1);
}

// Load the docx file as binary content
const content = fs.readFileSync(
  path.resolve(__dirname, "template.docx"),
  "binary"
);

// Create a new PizZip instance with the template content
const zip = new PizZip(content);

// Image module configuration
const imageOpts = {
  centered: false, // Image will not be centered
  getImage: (tagValue) => {
    return fs.readFileSync(tagValue); // Read image file from path
  },
  getSize: () => {
    return [150, 150]; // Image size in pixels (width, height)
  },
};

// Create a new Docxtemplater instance with image module
const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
  modules: [new ImageModule(imageOpts)],
});

// Data to fill in the template
const data = {
  name: "علی احمدی",
  company: "شرکت فناوری نوین",
  logo: imagePath, // Path to the image file
  items: [
    { item: "محصول A", price: 100000, quantity: 5 },
    { item: "محصول B", price: 200000, quantity: 3 },
    { item: "محصول C", price: 150000, quantity: 10 },
  ],
};

try {
  // Render the document with the provided data
  doc.render(data);

  // Generate the output document
  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  // Save the output document
  fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
  console.log("Report with table and image generated successfully as output.docx");
} catch (error) {
  console.error("Error rendering document:", error);
}