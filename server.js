const express = require("express");
const cors = require("cors");
const { v4: uuidv4 } = require("uuid");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  TableRow,
  TableCell,
  Table,
  WidthType,
} = require("docx");
const path = require("path");
const fs = require("fs");

const app = express();
app.use(cors());
app.use(express.json({ limit: "5mb" }));
app.use(express.text({ limit: "5mb", type: "text/plain" }));
app.use(express.urlencoded({ extended: true, limit: "5mb" }));

// Serve generated files
const OUTPUT_DIR = path.join(__dirname, "generated");
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}
app.use("/files", express.static(OUTPUT_DIR));

// Clean up files older than 1 hour
function cleanupOldFiles() {
  const files = fs.readdirSync(OUTPUT_DIR);
  const now = Date.now();
  files.forEach((file) => {
    const filePath = path.join(OUTPUT_DIR, file);
    const stats = fs.statSync(filePath);
    if (now - stats.mtimeMs > 3600000) {
      fs.unlinkSync(filePath);
    }
  });
}

function parseReportToDocx(reportText) {
  const lines = reportText.split("\n");
  const children = [];

  // Title
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: "FEEDBACK ANALYSIS REPORT",
          bold: true,
          size: 36,
          font: "Calibri",
          color: "1B3A5C",
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
    })
  );

  // Separator line
  children.push(
    new Paragraph({
      border: {
        bottom: { color: "1B3A5C", space: 1, style: BorderStyle.SINGLE, size: 6 },
      },
      spacing: { after: 200 },
    })
  );

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Skip the title line and "Prepared by" / "Date" lines at the top (already handled)
    if (line === "FEEDBACK ANALYSIS REPORT") continue;

    // Prepared by and Date lines
    if (line.startsWith("Prepared by:") || line.startsWith("Date:")) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line,
              italics: true,
              size: 20,
              font: "Calibri",
              color: "666666",
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
        })
      );
      continue;
    }

    // Section headers (lines starting with a number followed by a period and space)
    const sectionMatch = line.match(/^(\d+)\.\s+(.+)/);
    if (sectionMatch && line === line.toUpperCase().replace(/[^A-Z0-9.\s]/g, "").trim() === false) {
      // Check if it's a main section header (all uppercase or mostly uppercase)
    }

    if (sectionMatch && isMainSectionHeader(line)) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line,
              bold: true,
              size: 26,
              font: "Calibri",
              color: "1B3A5C",
            }),
          ],
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 300, after: 150 },
          border: {
            bottom: { color: "E0E0E0", space: 1, style: BorderStyle.SINGLE, size: 2 },
          },
        })
      );
      continue;
    }

    // Sub-section headers (like "2.1 Experience Ratings" or numbered items under sections)
    const subSectionMatch = line.match(/^(\d+\.\d+)\s+(.+)/);
    if (subSectionMatch) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line,
              bold: true,
              size: 22,
              font: "Calibri",
              color: "2C5F8A",
            }),
          ],
          spacing: { before: 200, after: 100 },
        })
      );
      continue;
    }

    // Table-like rows (Rating/Respondents/Percentage headers and data)
    if (isTableRow(line)) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line,
              size: 20,
              font: "Consolas",
              color: "333333",
            }),
          ],
          spacing: { after: 40 },
        })
      );
      continue;
    }

    // Priority lines
    if (line.startsWith("Priority:")) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line,
              bold: true,
              size: 20,
              font: "Calibri",
              color: line.includes("High") ? "CC3333" : line.includes("Medium") ? "CC8833" : "339933",
            }),
          ],
          spacing: { before: 100, after: 40 },
          indent: { left: 400 },
        })
      );
      continue;
    }

    // Severity lines
    if (line.startsWith("Severity:")) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line,
              bold: true,
              size: 20,
              font: "Calibri",
              color: line.includes("High") ? "CC3333" : line.includes("Medium") ? "CC8833" : "339933",
            }),
          ],
          spacing: { after: 40 },
          indent: { left: 400 },
        })
      );
      continue;
    }

    // Label lines (Recommendation:, Rationale:, Expected Impact:, etc.)
    const labelMatch = line.match(/^(Recommendation|Rationale|Expected Impact|Detail|Evidence|Representative quote|Comment|Experience|Comment Status|Respondents reflected):\s*(.*)/);
    if (labelMatch) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: labelMatch[1] + ": ",
              bold: true,
              size: 20,
              font: "Calibri",
              color: "444444",
            }),
            new TextRun({
              text: labelMatch[2],
              size: 20,
              font: "Calibri",
              color: "333333",
            }),
          ],
          spacing: { after: 60 },
          indent: { left: 400 },
        })
      );
      continue;
    }

    // Numbered sub-items (like "1. Strong Overall Experience Ratings")
    const numberedMatch = line.match(/^(\d+)\.\s+(.+)/);
    if (numberedMatch && !isMainSectionHeader(line)) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `${numberedMatch[1]}. `,
              bold: true,
              size: 20,
              font: "Calibri",
              color: "1B3A5C",
            }),
            new TextRun({
              text: numberedMatch[2],
              bold: true,
              size: 20,
              font: "Calibri",
              color: "333333",
            }),
          ],
          spacing: { before: 150, after: 60 },
          indent: { left: 200 },
        })
      );
      continue;
    }

    // Empty lines
    if (line === "") {
      children.push(
        new Paragraph({
          spacing: { after: 80 },
        })
      );
      continue;
    }

    // Regular text
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: line,
            size: 20,
            font: "Calibri",
            color: "333333",
          }),
        ],
        spacing: { after: 60 },
        indent: { left: line.startsWith("   ") ? 400 : 0 },
      })
    );
  }

  return children;
}

function isMainSectionHeader(line) {
  const mainHeaders = [
    "EXECUTIVE SUMMARY",
    "RESPONSE BREAKDOWN",
    "POSITIVE FINDINGS",
    "AREAS OF CONCERN",
    "CUSTOMER COMMENTS REVIEW",
    "KEY INSIGHTS",
    "RECOMMENDATIONS",
  ];
  return mainHeaders.some((h) => line.toUpperCase().includes(h));
}

function isTableRow(line) {
  // Lines with multiple spaces between words (tabular data)
  const parts = line.split(/\s{3,}/);
  if (parts.length >= 3) return true;
  // Header-like lines
  if (line.match(/^(Rating|Comment Status|Total)\s+/)) return true;
  return false;
}

// API endpoint
app.post("/api/generate-report", async (req, res) => {
  try {
    let report;

    // Support both JSON body {"report": "..."} and plain text body
    if (typeof req.body === "string") {
      report = req.body;
    } else if (req.body && req.body.report) {
      report = req.body.report;
    } else {
      // Try to extract from any body format
      report = JSON.stringify(req.body);
    }

    if (!report || report.length < 10) {
      return res.status(400).json({ error: "Missing report content in request body. Send as plain text or JSON with 'report' field." });
    }

    cleanupOldFiles();

    const docChildren = parseReportToDocx(report);

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: {
              font: "Calibri",
              size: 20,
            },
          },
        },
      },
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 1440,
                right: 1440,
                bottom: 1440,
                left: 1440,
              },
            },
          },
          children: docChildren,
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);
    const fileId = uuidv4();
    const fileName = `Feedback_Analysis_Report_${new Date().toISOString().split("T")[0]}_${fileId.slice(0, 8)}.docx`;
    const filePath = path.join(OUTPUT_DIR, fileName);

    fs.writeFileSync(filePath, buffer);

    const protocol = req.headers['x-forwarded-proto'] || req.protocol || 'https';
    const host = req.headers['host'];
    const baseUrl = `${protocol}://${host}`;
    const downloadUrl = `${baseUrl}/files/${fileName}`;

    res.json({
      success: true,
      downloadUrl: downloadUrl,
      fileName: fileName,
    });
  } catch (error) {
    console.error("Error generating report:", error);
    res.status(500).json({ error: "Failed to generate report" });
  }
});

// Health check
app.get("/", (req, res) => {
  res.json({ status: "ok", service: "DOCX Report Generator" });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`DOCX Generator running on port ${PORT}`);
});
