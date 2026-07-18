const express = require("express");
const puppeteer = require("puppeteer");
const crypto = require("crypto");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, SpacingType,
} = require("docx");

const app = express();
app.use(express.json({ limit: "10mb" }));

/* ══════════════════════════════════════════════════════
   AUTH — this service is internal-only (called by the
   Spring backend over the Docker network). Every request
   must carry the shared X-Service-Token; without the env
   var configured the generate endpoints refuse to run
   rather than default to open. No CORS: browsers are not
   a legitimate caller of this service.
══════════════════════════════════════════════════════ */
const SERVICE_TOKEN = process.env.PDF_SERVICE_TOKEN || "";
if (!SERVICE_TOKEN) {
  console.error("⚠ PDF_SERVICE_TOKEN is not set — generate endpoints will return 503. " +
                "Set the same token on this service and the backend (openssl rand -hex 32).");
}

function requireServiceToken(req, res, next) {
  if (!SERVICE_TOKEN) {
    return res.status(503).json({ error: "Service not configured" });
  }
  const provided = req.get("X-Service-Token") || "";
  const a = Buffer.from(provided);
  const b = Buffer.from(SERVICE_TOKEN);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

/* ══════════════════════════════════════════════════════
   CONCURRENCY CAP — each render costs a Chromium page;
   without a cap a request burst exhausts the host.
══════════════════════════════════════════════════════ */
const MAX_CONCURRENT_RENDERS = 3;
let activeRenders = 0;

/* ══════════════════════════════════════════════════════
   SHARED BROWSER — one Chromium for the process lifetime
   (a fresh page per request) instead of a full launch per
   request. Relaunches automatically if it crashes.
   --no-sandbox is required because the container runs as
   root; SSRF is instead blocked per-request below.
══════════════════════════════════════════════════════ */
let browserPromise = null;

async function getBrowser() {
  if (browserPromise) {
    try {
      const b = await browserPromise;
      if (b.connected) return b;
    } catch { /* fall through to relaunch */ }
  }
  browserPromise = puppeteer.launch({
    headless: true,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
      "--disable-gpu",
    ],
  });
  return browserPromise;
}

// Only these external hosts may be fetched while rendering (webfonts).
// Everything else — cloud metadata endpoints, localhost, internal Docker
// hostnames, arbitrary internet URLs — is aborted, which closes the SSRF
// hole where attacker HTML exfiltrates internal responses into the PDF.
const ALLOWED_REMOTE_HOSTS = new Set([
  "fonts.googleapis.com",
  "fonts.gstatic.com",
]);

function isRequestAllowed(url) {
  if (url.startsWith("data:") || url === "about:blank") return true;
  try {
    const parsed = new URL(url);
    return (parsed.protocol === "https:" || parsed.protocol === "http:")
        && ALLOWED_REMOTE_HOSTS.has(parsed.hostname);
  } catch {
    return false;
  }
}

/* ══════════════════════════════════════════════════════
   HEALTH CHECK (unauthenticated — used by Docker/monitor)
══════════════════════════════════════════════════════ */
app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

/* ══════════════════════════════════════════════════════
   PDF GENERATION
══════════════════════════════════════════════════════ */
app.post("/generate-pdf", requireServiceToken, async (req, res) => {
  const { html, paperSize } = req.body;

  if (!html || !html.trim()) {
    return res.status(400).json({ error: "html is required and cannot be empty" });
  }

  if (activeRenders >= MAX_CONCURRENT_RENDERS) {
    return res.status(429).json({ error: "Renderer busy, retry shortly" });
  }
  activeRenders++;

  let page;
  try {
    const browser = await getBrowser();
    page = await browser.newPage();

    // Block all network access from the rendered document except data: URIs
    // and the webfont allowlist (see isRequestAllowed).
    await page.setRequestInterception(true);
    page.on("request", (request) => {
      if (isRequestAllowed(request.url())) request.continue();
      else request.abort();
    });

    await page.setContent(html, { waitUntil: "networkidle0", timeout: 30000 });

    const format = paperSize === "letter" ? "Letter" : "A4";

    // puppeteer ≥22 returns a Uint8Array; Express JSON-serializes those, so
    // wrap in a Buffer to send real binary.
    const pdf = Buffer.from(await page.pdf({
      format,
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
      displayHeaderFooter: false, // ← NO browser headers/footers
    }));

    res.set({
      "Content-Type": "application/pdf",
      "Content-Disposition": "attachment; filename=resume.pdf",
      "Content-Length": pdf.length,
    });
    res.send(pdf);

  } catch (err) {
    console.error("PDF generation error:", err);
    res.status(500).json({ error: "PDF generation failed: " + err.message });
  } finally {
    if (page) { try { await page.close(); } catch {} }
    activeRenders--;
  }
});

/* ══════════════════════════════════════════════════════
   WORD GENERATION
══════════════════════════════════════════════════════ */
app.post("/generate-word", requireServiceToken, async (req, res) => {
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
      if (!result) {
        console.warn(`hexToRgb: invalid hex color "${hex}", falling back to default`);
        return { r: 45, g: 91, b: 227 };
      }
      return {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16),
      };
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

    const name = pi.name ? pi.name.replace(/\s+/g, "_") : `Resume_${Date.now()}`;
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
