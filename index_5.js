const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { createCanvas, registerFont } = require("canvas");
const Chart = require("chart.js/auto");
const fs = require("fs");
const path = require("path");

// Register Persian font (Vazir) for chart
const fontPath = path.resolve(__dirname, "Vazir.ttf");
if (fs.existsSync(fontPath)) {
  registerFont(fontPath, { family: "Vazir" });
} else {
  console.warn("Warning: Vazir font file not found. Chart labels may not display in Persian correctly.");
}

// Function to generate chart as an image with larger size
async function generateChartImage(data) {
  const canvas = createCanvas(1400, 840); // Increased resolution for larger, high-quality chart
  const ctx = canvas.getContext("2d");

  new Chart(ctx, {
    type: "bar",
    data: {
      labels: data.items.map(item => item.item), // Dynamic product names
      datasets: [
        {
          label: "تعداد فروش",
          data: data.items.map(item => item.quantity), // Dynamic quantities
          backgroundColor: ["#36A2EB", "#FF6384", "#FFCE56"],
          borderColor: ["#36A2EB", "#FF6384", "#FFCE56"],
          borderWidth: 2,
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
            font: { family: "Vazir", size: 28 }, // Larger font for y-axis title
          },
          ticks: { font: { family: "Vazir", size: 20 } }, // Larger font for y-axis labels
        },
        x: {
          title: {
            display: true,
            text: "محصولات",
            font: { family: "Vazir", size: 28 }, // Larger font for x-axis title
          },
          ticks: { font: { family: "Vazir", size: 20 } }, // Larger font for x-axis labels
        },
      },
      plugins: {
        legend: {
          display: true,
          position: "top",
          labels: {
            font: { family: "Vazir", size: 22 }, // Larger font for legend
          },
        },
      },
    },
  });

  // Save chart as PNG
  const chartPath = path.resolve(__dirname, "chart.png");
  const out = fs.createWriteStream(chartPath);
  const stream = canvas.createPNGStream({ compressionLevel: 2 }); // Optimized compression
  stream.pipe(out);
  return new Promise((resolve) => {
    out.on("finish", () => resolve(chartPath));
  });
}

// Check if image and font files exist
const logoPath = path.resolve(__dirname, "logo.png");
if (!fs.existsSync(logoPath)) {
  console.error("Error: Image file 'logo.png' not found at", logoPath);
  process.exit(1);
}

const templatePath = path.resolve(__dirname, "template.docx");
if (!fs.existsSync(templatePath)) {
  console.error("Error: Template file 'template.docx' not found at", templatePath);
  process.exit(1);
}

// Main function to generate report
async function generateReport() {
  // Data to fill in the template
  const data = {
    name: "علی احمدی",
    company: "شرکت فناوری نوین",
    logo: logoPath,
    items: [
      { item: "محصول A", price: 100000, quantity: 5 },
      { item: "محصول B", price: 200000, quantity: 3 },
      { item: "محصول C", price: 150000, quantity: 10 },
    ],
  };

  // Generate chart image
  const chartPath = await generateChartImage(data);
  data.chart = chartPath;

  // Load the docx file as binary content
  const content = fs.readFileSync(templatePath, "binary");

  // Create a new PizZip instance with the template content
  const zip = new PizZip(content);

  // Image module configuration with larger chart size
  const imageOpts = {
    centered: true, // Center images
    getImage: (tagValue) => {
      if (!fs.existsSync(tagValue)) {
        console.error(`Error: Image file not found at ${tagValue}`);
        throw new Error(`Image file not found: ${tagValue}`);
      }
      return fs.readFileSync(tagValue);
    },
    getSize: (tagValue) => {
      if (tagValue.includes("chart")) {
        return [600, 360]; // Larger size for chart in document
      }
      return [250, 250]; // Size for logo
    },
  };

  // Create a new Docxtemplater instance with image module
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [new ImageModule(imageOpts)],
  });

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
    console.log("Report with RTL, table, logo, and larger chart generated successfully as output.docx");
  } catch (error) {
    console.error("Error rendering document:", error.message);
    if (error.properties && error.properties.errors) {
      console.error("Template errors:", error.properties.errors);
    }
  }
}

// Run the report generation
generateReport();