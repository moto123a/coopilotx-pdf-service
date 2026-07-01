const express = require("express");
const puppeteer = require("puppeteer");
const cors = require("cors");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, SpacingType,
} = require("docx");

const app = express();
app.use(cors());
app.use(express.json({ limit: "10mb" }));

/* ══════════════════════════════════════════════════════
   HEALTH CHECK
══════════════════════════════════════════════════════ */
app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

/* ══════════════════════════════════════════════════════
   PDF GENERATION
══════════════════════════════════════════════════════ */
app.post("/generate-pdf", async (req, res) => {
  const { html, paperSize } = req.body;

  if (!html) {
    return res.status(400).json({ error: "html is required" });
  }

  let browser;
  try {
    browser = await puppeteer.launch({
      headless: "new",
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
      ],
    });

    const page = await browser.newPage();

    // ── SSRF protection ──────────────────────────────────────────────
    // The HTML comes from user resume data, so a crafted <img>/<link>/@import
    // could make headless Chromium fetch internal resources (cloud metadata
    // at 169.254.169.254, file://, localhost, private LAN, docker services).
    // Intercept every sub-resource request and abort anything that isn't a
    // data: URI or a public http(s) host. Normal resumes only use inline CSS,
    // data-URI images, and public font/image CDNs, so this never blocks a real
    // resume while cutting off the SSRF vector entirely.
    await page.setRequestInterception(true);
    page.on("request", (reqI) => {
      const url = reqI.url();
      if (url.startsWith("data:") || url === "about:blank") return reqI.continue();
      let parsed;
      try { parsed = new URL(url); } catch { return reqI.abort(); }
      if (parsed.protocol !== "http:" && parsed.protocol !== "https:") return reqI.abort();
      const host = parsed.hostname.toLowerCase();
      const isPrivate =
        host === "localhost" ||
        host === "0.0.0.0" ||
        host === "[::1]" || host === "::1" ||
        host.endsWith(".internal") || host.endsWith(".local") ||
        /^127\./.test(host) ||
        /^10\./.test(host) ||
        /^192\.168\./.test(host) ||
        /^169\.254\./.test(host) ||                       // link-local + cloud metadata
        /^172\.(1[6-9]|2\d|3[01])\./.test(host) ||        // 172.16.0.0/12
        /^fe80:/i.test(host) || /^fc/i.test(host) || /^fd/i.test(host); // IPv6 link-local/ULA
      if (isPrivate) return reqI.abort();
      return reqI.continue();
    });

    await page.setContent(html, { waitUntil: "networkidle0" });

    const format = paperSize === "letter" ? "Letter" : "A4";

    const pdf = await page.pdf({
      format,
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
      displayHeaderFooter: false, // ← NO browser headers/footers
    });

    await browser.close();

    res.set({
      "Content-Type": "application/pdf",
      "Content-Disposition": "attachment; filename=resume.pdf",
      "Content-Length": pdf.length,
    });
    res.send(pdf);

  } catch (err) {
    if (browser) await browser.close();
    console.error("PDF generation error:", err);
    res.status(500).json({ error: "PDF generation failed: " + err.message });
  }
});

/* ══════════════════════════════════════════════════════
   WORD GENERATION
══════════════════════════════════════════════════════ */
app.post("/generate-word", async (req, res) => {
  const { data, styles } = req.body;

  if (!data) {
    return res.status(400).json({ error: "data is required" });
  }

  try {
    const pi = data.personalInfo || {};
    const ac = styles?.ac || "#2563eb";
    const fs = styles?.fs || 11;

    const hexToRgb = (hex) => {
      const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16),
      } : { r: 45, g: 91, b: 227 };
    };

    const acRgb = hexToRgb(ac);
    const accentColor = `${acRgb.r.toString(16).padStart(2,"0")}${acRgb.g.toString(16).padStart(2,"0")}${acRgb.b.toString(16).padStart(2,"0")}`;

    const children = [];

    // ── HEADER ──
    children.push(
      new Paragraph({
        children: [new TextRun({ text: pi.name || "Your Name", bold: true, size: 36, color: accentColor })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
      }),
      new Paragraph({
        children: [new TextRun({ text: pi.headline || "", size: 20, color: "666666" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
      }),
      new Paragraph({
        children: [new TextRun({
          text: [pi.email, pi.phone, pi.location].filter(Boolean).join("  ·  "),
          size: 18, color: "555555",
        })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
      })
    );

    if (pi.linkedin || pi.github || pi.portfolio) {
      children.push(new Paragraph({
        children: [new TextRun({
          text: [pi.linkedin, pi.github, pi.portfolio].filter(Boolean).join("  ·  "),
          size: 16, color: "888888",
        })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
      }));
    }

    // ── SECTION HEADING ──
    const sectionHeading = (title) => new Paragraph({
      children: [new TextRun({ text: title.toUpperCase(), bold: true, size: 18, color: accentColor })],
      spacing: { before: 240, after: 80 },
      border: {
        bottom: { style: BorderStyle.SINGLE, size: 6, color: accentColor },
      },
    });

    // ── SUMMARY ──
    if (data.summary) {
      children.push(sectionHeading("Professional Summary"));
      children.push(new Paragraph({
        children: [new TextRun({ text: data.summary, size: fs * 2 })],
        spacing: { after: 120 },
      }));
    }

    // ── SKILLS ──
    if ((data.skillCategories || []).length > 0) {
      children.push(sectionHeading("Technical Skills"));
      data.skillCategories.forEach(cat => {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: `${cat.name}: `, bold: true, size: fs * 2 }),
            new TextRun({ text: cat.skills, size: fs * 2 }),
          ],
          spacing: { after: 60 },
        }));
      });
    }

    // ── EXPERIENCE ──
    if ((data.experience || []).length > 0) {
      children.push(sectionHeading("Work Experience"));
      data.experience.forEach((exp, i) => {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: `${exp.company}${exp.role ? ` — ${exp.role}` : ""}`, bold: true, size: fs * 2 }),
            new TextRun({ text: `\t${exp.period || ""}`, size: (fs - 1) * 2, color: "666666" }),
          ],
          spacing: { before: i === 0 ? 0 : 160, after: 40 },
        }));
        if (exp.location) {
          children.push(new Paragraph({
            children: [new TextRun({ text: exp.location, size: (fs - 1) * 2, color: "888888", italics: true })],
            spacing: { after: 60 },
          }));
        }
        (exp.bullets || []).forEach(bullet => {
          children.push(new Paragraph({
            children: [new TextRun({ text: bullet, size: fs * 2 })],
            bullet: { level: 0 },
            spacing: { after: 40 },
          }));
        });
      });
    }

    // ── PROJECTS ──
    if ((data.projects || []).length > 0) {
      children.push(sectionHeading("Projects"));
      data.projects.forEach((proj, i) => {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: proj.title || proj.name || "", bold: true, size: fs * 2, color: accentColor }),
            new TextRun({ text: `\t${proj.period || ""}`, size: (fs - 1) * 2, color: "666666" }),
          ],
          spacing: { before: i === 0 ? 0 : 140, after: 40 },
        }));
        if (proj.tech) {
          children.push(new Paragraph({
            children: [new TextRun({ text: proj.tech, size: (fs - 1) * 2, color: "666666", italics: true })],
            spacing: { after: 60 },
          }));
        }
        (proj.bullets || []).forEach(bullet => {
          children.push(new Paragraph({
            children: [new TextRun({ text: bullet, size: fs * 2 })],
            bullet: { level: 0 },
            spacing: { after: 40 },
          }));
        });
      });
    }

    // ── EDUCATION ──
    if ((data.education || []).length > 0) {
      children.push(sectionHeading("Education"));
      data.education.forEach((edu, i) => {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: edu.school || edu.institution || "", bold: true, size: fs * 2 }),
            new TextRun({ text: `\t${edu.period || edu.year || ""}`, size: (fs - 1) * 2, color: "666666" }),
          ],
          spacing: { before: i === 0 ? 0 : 140, after: 40 },
        }));
        children.push(new Paragraph({
          children: [new TextRun({
            text: `${edu.degree || ""}${edu.gpa ? ` — GPA: ${edu.gpa}` : ""}`,
            size: (fs - 0.5) * 2, color: "444444",
          })],
          spacing: { after: 60 },
        }));
      });
    }

    // ── CERTIFICATIONS ──
    const validCerts = (data.certifications || []).filter(c =>
      typeof c === "string" ? c.trim() : c?.name?.trim()
    );
    if (validCerts.length > 0) {
      children.push(sectionHeading("Certifications"));
      validCerts.forEach(cert => {
        const text = typeof cert === "string"
          ? cert
          : `${cert.name || ""}${cert.issuer ? ` — ${cert.issuer}` : ""}${cert.year ? ` (${cert.year})` : ""}`;
        children.push(new Paragraph({
          children: [new TextRun({ text, size: fs * 2 })],
          bullet: { level: 0 },
          spacing: { after: 40 },
        }));
      });
    }

    // ── BUILD DOCUMENT ──
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 720, bottom: 720, left: 864, right: 864 },
          },
        },
        children,
      }],
    });

    const buffer = await Packer.toBuffer(doc);

    const name = pi.name ? pi.name.replace(/\s+/g, "_") : "Resume";
    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename=${name}_Resume.docx`,
      "Content-Length": buffer.length,
    });
    res.send(buffer);

  } catch (err) {
    console.error("Word generation error:", err);
    res.status(500).json({ error: "Word generation failed: " + err.message });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`✅ PDF/Word service running on port ${PORT}`);
});