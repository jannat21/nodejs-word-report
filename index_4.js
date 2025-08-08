const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { createCanvas } = require("canvas");
const Chart = require("chart.js/auto");
const fs = require("fs");
const path = require("path");

// Function to generate chart as an image
async function generateChartImage() {
  const canvas = createCanvas(600, 400);
  const ctx = canvas.getContext("2d");

  new Chart(ctx, {
    type: "bar",
    data: {
      labels: ["محصول A", "محصول B", "محصول C"],
      datasets: [
        {
          label: "تعداد فروش",
          data: [5, 3, 10], // Quantities from data
          backgroundColor: ["#36A2EB", "#FF6384", "#FFCE56"],
          borderColor: ["#36A2EB", "#FF6384", "#FFCE56"],
          borderWidth: 1,
        },
      ],
    },
    options: {
      scales: {
        y: {
          beginAtZero: true,
          title: {
            display: true,
            text: "تعداد",
          },
        },
        x: {
          title: {
            display: true,
            text: "محصولات",
          },
        },
      },
      plugins: {
        legend: {
          display: true,
          position: "top",
        },
      },
    },
  });

  // Save chart as PNG
  const chartPath = path.resolve(__dirname, "chart.png");
  const out = fs.createWriteStream(chartPath);
  const stream = canvas.createPNGStream();
  stream.pipe(out);
  return new Promise((resolve) => {
    out.on("finish", () => resolve(chartPath));
  });
}

// Check if image files exist
const logoPath = path.resolve(__dirname, "logo.png");
if (!fs.existsSync(logoPath)) {
  console.error("Error: Image file 'logo.png' not found at", logoPath);
  process.exit(1);
}

// Main function to generate report
async function generateReport() {
  // Generate chart image
  const chartPath = await generateChartImage();

  // Load the docx file as binary content
  const content = fs.readFileSync(
    path.resolve(__dirname, "template.docx"),
    "binary"
  );

  // Create a new PizZip instance with the template content
  const zip = new PizZip(content);

  // Image module configuration
  const imageOpts = {
    centered: false,
    getImage: (tagValue) => {
      return fs.readFileSync(tagValue);
    },
    getSize: (tagValue) => {
      return tagValue.includes("chart") ? [300, 200] : [150, 150]; // Different sizes for chart and logo
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
    logo: logoPath,
    chart: chartPath,
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
    console.log("Report with table, logo, and chart generated successfully as output.docx");
  } catch (error) {
    console.error("Error rendering document:", error);
  }
}

// Run the report generation
generateReport();