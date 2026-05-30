const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType,
  VerticalAlign, TabStopType, ExternalHyperlink, UnderlineType
} = require('docx');
const fs = require('fs');

// ─────────────────────────────────────────
//  STEP 1 — Customize your color palette
// ─────────────────────────────────────────
const COLOR_ACCENT = "1A56A0"; // Blue accent (headings, bullets, links)
const COLOR_DARK   = "1A1A2E"; // Near-black  (name, company names)
const COLOR_MID    = "444444"; // Body text
const COLOR_LIGHT  = "666666"; // Dates, subtitles
const CONTENT_WIDTH = 9360;    // US Letter, 1-inch margins (DXA units)

// ─────────────────────────────────────────
//  Shared border helpers
// ─────────────────────────────────────────
const noBorder  = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ─────────────────────────────────────────
//  Layout helpers
// ─────────────────────────────────────────
function rule(color = COLOR_ACCENT, size = 8) {
  return new Paragraph({
    spacing: { before: 0, after: 0 },
    border: { bottom: { style: BorderStyle.SINGLE, size, color, space: 1 } },
    children: [new TextRun("")]
  });
}

function spacer(before = 80, after = 0) {
  return new Paragraph({ spacing: { before, after }, children: [new TextRun("")] });
}

// ─────────────────────────────────────────
//  Section heading  (ATS-parseable: plain
//  uppercase text + bottom border)
// ─────────────────────────────────────────
function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLOR_ACCENT, space: 2 } },
    children: [
      new TextRun({
        text: text.toUpperCase(),
        bold: true, size: 22,
        color: COLOR_ACCENT,
        font: "Arial",
        characterSpacing: 40
      })
    ]
  });
}

// ─────────────────────────────────────────
//  Bullet paragraph
// ─────────────────────────────────────────
function bullet(runs) {
  return new Paragraph({
    numbering: { reference: "resume-bullets", level: 0 },
    spacing: { before: 30, after: 30 },
    indent: { left: 360, hanging: 220 },
    children: runs
  });
}

// ─────────────────────────────────────────
//  Text-run helpers (mix inside a bullet or
//  paragraph to apply different styles)
// ─────────────────────────────────────────
function bold(text, color = COLOR_DARK, size = 19) {
  return new TextRun({ text, bold: true, color, size, font: "Arial" });
}

function normal(text, color = COLOR_MID, size = 19) {
  return new TextRun({ text, color, size, font: "Arial" });
}

function accent(text, size = 19) {
  return new TextRun({ text, bold: true, color: COLOR_ACCENT, size, font: "Arial" });
}

// ─────────────────────────────────────────
//  Hyperlink helper
// ─────────────────────────────────────────
function link(displayText, url, size = 18) {
  return new ExternalHyperlink({
    link: url,
    children: [
      new TextRun({
        text: displayText,
        size, color: COLOR_ACCENT, font: "Arial",
        underline: { type: UnderlineType.SINGLE, color: COLOR_ACCENT }
      })
    ]
  });
}

// Dot separator for the contact bar
function dot(size = 18) {
  return new TextRun({ text: "   •   ", size, color: COLOR_ACCENT, font: "Arial" });
}

// ─────────────────────────────────────────
//  Job entry header
//  company | role          [certificate]  date
// ─────────────────────────────────────────
function jobHeader(company, role, period, certUrl) {
  const children = [
    new TextRun({ text: company, bold: true, size: 21, color: COLOR_DARK,  font: "Arial" }),
    new TextRun({ text: "  |  ",             size: 20, color: COLOR_LIGHT, font: "Arial" }),
    new TextRun({ text: role,   bold: true,  size: 20, color: COLOR_ACCENT, font: "Arial" }),
  ];
  if (certUrl) {
    children.push(new TextRun({ text: "  ", size: 18, font: "Arial" }));
    children.push(
      new ExternalHyperlink({
        link: certUrl,
        children: [new TextRun({
          text: "[Certificate]", size: 17,
          color: COLOR_ACCENT, font: "Arial",
          underline: { type: UnderlineType.SINGLE, color: COLOR_ACCENT },
          italics: true
        })]
      })
    );
  }
  // Tab + right-aligned date
  children.push(new TextRun({ text: "\t", font: "Arial" }));
  children.push(new TextRun({ text: period, size: 18, color: COLOR_LIGHT, font: "Arial", italics: true }));

  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH }],
    spacing: { before: 120, after: 20 },
    children
  });
}

// ─────────────────────────────────────────
//  Project entry header
//  title  [GitHub]                         date
// ─────────────────────────────────────────
function projHeader(title, period, githubUrl) {
  const children = [
    new TextRun({ text: title, bold: true, size: 21, color: COLOR_DARK, font: "Arial" }),
  ];
  if (githubUrl) {
    children.push(new TextRun({ text: "  ", size: 18, font: "Arial" }));
    children.push(
      new ExternalHyperlink({
        link: githubUrl,
        children: [new TextRun({
          text: "[GitHub]", size: 17,
          color: COLOR_ACCENT, font: "Arial",
          underline: { type: UnderlineType.SINGLE, color: COLOR_ACCENT },
          italics: true
        })]
      })
    );
  }
  children.push(new TextRun({ text: "\t", font: "Arial" }));
  children.push(new TextRun({ text: period, size: 18, color: COLOR_LIGHT, font: "Arial", italics: true }));

  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH }],
    spacing: { before: 120, after: 20 },
    children
  });
}

// ─────────────────────────────────────────
//  Skills table row  (Label | Value)
// ─────────────────────────────────────────
function skillRow(label, value) {
  return new TableRow({
    children: [
      new TableCell({
        borders: noBorders,
        width: { size: 1600, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 0, right: 80 },
        children: [new Paragraph({
          children: [new TextRun({ text: label, bold: true, size: 19, color: COLOR_DARK, font: "Arial" })]
        })]
      }),
      new TableCell({
        borders: noBorders,
        width: { size: 7760, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 80, right: 0 },
        children: [new Paragraph({
          children: [new TextRun({ text: value, size: 19, color: COLOR_MID, font: "Arial" })]
        })]
      })
    ]
  });
}

// ═════════════════════════════════════════
//  STEP 2 — Fill in YOUR information below
// ═════════════════════════════════════════

const doc = new Document({
  // Bullet style definition (▸ symbol, accent color)
  numbering: {
    config: [{
      reference: "resume-bullets",
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: "▸",
        alignment: AlignmentType.LEFT,
        style: {
          run: { color: COLOR_ACCENT, size: 18, font: "Arial" },
          paragraph: { indent: { left: 360, hanging: 220 } }
        }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20, color: COLOR_MID } } }
  },

  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 }, // US Letter
        margin: { top: 1008, right: 1080, bottom: 1008, left: 1080 }
      }
    },

    children: [

      // ──────────────────────────────────
      //  NAME
      // ──────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [new TextRun({
          text: "ALEX MORGAN",          // ← Your full name (ALL CAPS)
          bold: true, size: 52,
          color: COLOR_DARK, font: "Arial", characterSpacing: 60
        })]
      }),

      // ── TAGLINE ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({
          text: "Software Engineer  ·  Full Stack Developer  ·  Open Source Contributor",
          size: 20, color: COLOR_LIGHT, font: "Arial", italics: true
        })]
      }),

      // ── CONTACT BAR ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [
          new TextRun({ text: "San Francisco, CA", size: 18, color: COLOR_LIGHT, font: "Arial" }),
          dot(),
          new TextRun({ text: "+1 415-555-0192", size: 18, color: COLOR_LIGHT, font: "Arial" }),
          dot(),
          link("alex@example.com", "mailto:alex@example.com"),
          dot(),
          link("linkedin.com/in/alexmorgan", "https://www.linkedin.com/in/alexmorgan"),
          dot(),
          link("github.com/alexmorgan", "https://github.com/alexmorgan"),
          dot(),
          link("alexmorgan.dev", "https://alexmorgan.dev"),
        ]
      }),

      spacer(80, 0),
      rule(COLOR_ACCENT, 12),
      spacer(60, 0),

      // ══════════════════════════════════
      //  TECHNICAL SKILLS
      // ══════════════════════════════════
      sectionHeading("Technical Skills"),
      spacer(60, 0),

      new Table({
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        columnWidths: [1600, 7760],
        rows: [
          skillRow("Languages",      "Python, JavaScript, TypeScript, Go, SQL, Java"),
          skillRow("Backend",        "Node.js, Express.js, FastAPI, Spring Boot, GraphQL, REST APIs"),
          skillRow("Frontend",       "React.js, Next.js, Tailwind CSS, Redux, HTML5, CSS3"),
          skillRow("Databases",      "PostgreSQL, MongoDB, MySQL, Redis, Elasticsearch, DynamoDB"),
          skillRow("Cloud & DevOps", "AWS (EC2, S3, Lambda, RDS, ECS), Docker, Kubernetes, Terraform, GitHub Actions CI/CD"),
          skillRow("Tools",          "Git, Postman, JIRA, Figma, Linux, Nginx, Webpack"),
          skillRow("Coursework",     "Data Structures & Algorithms, System Design, DBMS, Operating Systems"),
        ]
      }),

      spacer(120, 0),

      // ══════════════════════════════════
      //  PROFESSIONAL EXPERIENCE
      // ══════════════════════════════════
      sectionHeading("Professional Experience"),
      spacer(60, 0),

      // ── Job 1 ──
      jobHeader("Stripe", "Software Engineer Intern", "Jan 2025 – Present", null),
      bullet([normal("Built a "), bold("real-time payment reconciliation dashboard"), normal(" using "), bold("React.js + Node.js"), normal(", reducing manual auditing time by "), accent("65%"), normal(".")]),
      bullet([normal("Optimized PostgreSQL query performance with indexing and query rewrites, cutting average response time from "), accent("3.2s to 0.4s"), normal(" (87% improvement).")]),
      bullet([normal("Collaborated with a cross-functional team of "), bold("8 engineers"), normal(" to ship a fraud-detection microservice serving "), accent("500K+ transactions/day"), normal(".")]),

      spacer(80, 0),

      // ── Job 2 — with certificate link ──
      jobHeader("Acme Corp", "Backend Developer", "Jun 2024 – Dec 2024", "https://example.com/certificate/acme"),
      bullet([normal("Designed and deployed a "), bold("RESTful API"), normal(" for an inventory management system serving "), accent("200+ enterprise clients"), normal(" and processing "), accent("15,000+ daily requests"), normal(".")]),
      bullet([normal("Integrated "), bold("WebSockets"), normal(" for live order status updates, improving client-reported satisfaction scores by "), accent("30%"), normal(".")]),
      bullet([normal("Implemented "), bold("JWT authentication"), normal(" and role-based access control, reducing unauthorized access incidents by "), accent("100%"), normal(" post-deployment.")]),

      spacer(80, 0),

      // ── Job 3 — with certificate link ──
      jobHeader("StartupXYZ", "Full Stack Developer", "Jan 2024 – May 2024", "https://example.com/certificate/startupxyz"),
      bullet([normal("Developed "), accent("12+ reusable React components"), normal(" for a SaaS analytics platform, cutting UI development time by "), accent("40%"), normal(".")]),
      bullet([normal("Migrated a legacy monolith to "), bold("microservices architecture"), normal(" on AWS ECS, improving deployment frequency from monthly to "), accent("weekly releases"), normal(".")]),
      bullet([normal("Wrote "), accent("90+ unit and integration tests"), normal(" with Jest and Supertest, raising code coverage from 32% to "), accent("87%"), normal(".")]),

      spacer(120, 0),

      // ══════════════════════════════════
      //  PROJECTS
      // ══════════════════════════════════
      sectionHeading("Projects"),
      spacer(60, 0),

      // ── Project 1 ──
      projHeader("DevTracker — Open Source Issue Tracker", "2025", "https://github.com/alexmorgan/devtracker"),
      new Paragraph({
        spacing: { before: 20, after: 30 },
        children: [
          new TextRun({ text: "Stack: ", bold: true, size: 18, color: COLOR_DARK, font: "Arial" }),
          new TextRun({ text: "TypeScript · React.js · Node.js · PostgreSQL · Docker · GitHub Actions", size: 18, color: COLOR_LIGHT, font: "Arial", italics: true })
        ]
      }),
      bullet([normal("Built a full-featured project management tool with real-time collaboration using "), bold("WebSockets"), normal(", supporting "), accent("500+ concurrent users"), normal(" in load tests.")]),
      bullet([normal("Implemented "), bold("CI/CD pipeline"), normal(" with GitHub Actions — automated linting, testing, and Docker image publishing on every PR merge.")]),
      bullet([normal("Achieved "), accent("2,000+ GitHub stars"), normal(" and "), accent("80+ contributors"), normal(" within 3 months of launch.")]),

      spacer(80, 0),

      // ── Project 2 ──
      projHeader("SmartBudget — Personal Finance API", "2024", "https://github.com/alexmorgan/smartbudget"),
      new Paragraph({
        spacing: { before: 20, after: 30 },
        children: [
          new TextRun({ text: "Stack: ", bold: true, size: 18, color: COLOR_DARK, font: "Arial" }),
          new TextRun({ text: "Python · FastAPI · PostgreSQL · Redis · AWS Lambda · Terraform", size: 18, color: COLOR_LIGHT, font: "Arial", italics: true })
        ]
      }),
      bullet([normal("Designed a "), bold("serverless financial analytics API"), normal(" on AWS Lambda + API Gateway with sub-"), accent("200ms p99 latency"), normal(" under production load.")]),
      bullet([normal("Used "), bold("Redis caching"), normal(" for frequently queried budget summaries, reducing database load by "), accent("70%"), normal(".")]),
      bullet([normal("Provisioned all infrastructure with "), bold("Terraform"), normal(", enabling one-command reproducible deployments across dev/staging/prod environments.")]),

      spacer(120, 0),

      // ══════════════════════════════════
      //  EDUCATION
      // ══════════════════════════════════
      sectionHeading("Education"),
      spacer(60, 0),

      new Table({
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        columnWidths: [7200, 2160],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: noBorders,
                width: { size: 7200, type: WidthType.DXA },
                margins: { top: 40, bottom: 40, left: 0, right: 80 },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "B.S. in Computer Science", bold: true, size: 21, color: COLOR_DARK, font: "Arial" })] }),
                  new Paragraph({ spacing: { before: 20 }, children: [new TextRun({ text: "University of California, Berkeley", size: 19, color: COLOR_MID, font: "Arial" })] })
                ]
              }),
              new TableCell({
                borders: noBorders,
                width: { size: 2160, type: WidthType.DXA },
                verticalAlign: VerticalAlign.CENTER,
                margins: { top: 40, bottom: 40, left: 80, right: 0 },
                children: [
                  new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "GPA: 3.87 / 4.0", bold: true, size: 20, color: COLOR_ACCENT, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "2021 – 2025 (Expected)", size: 18, color: COLOR_LIGHT, font: "Arial", italics: true })] })
                ]
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                borders: noBorders,
                width: { size: 7200, type: WidthType.DXA },
                margins: { top: 40, bottom: 0, left: 0, right: 80 },
                children: [new Paragraph({ children: [new TextRun({ text: "Relevant Coursework:", bold: true, size: 19, color: COLOR_DARK, font: "Arial" })] }),
                           new Paragraph({ spacing: { before: 10 }, children: [new TextRun({ text: "Algorithms, Distributed Systems, Machine Learning, Database Systems, Computer Networks", size: 18, color: COLOR_MID, font: "Arial" })] })]
              }),
              new TableCell({
                borders: noBorders,
                width: { size: 2160, type: WidthType.DXA },
                margins: { top: 40, bottom: 0, left: 80, right: 0 },
                children: [new Paragraph({ children: [new TextRun("")] })]
              })
            ]
          })
        ]
      }),

      spacer(120, 0),

      // ══════════════════════════════════
      //  ACHIEVEMENTS & CONTRIBUTIONS
      // ══════════════════════════════════
      sectionHeading("Achievements & Contributions"),
      spacer(60, 0),

      bullet([
        bold("npm Package — cli-scaffold: "),
        normal("Published a CLI tool that scaffolds project structures from config files, reaching "),
        accent("3,200+ downloads"),
        normal(". Available at "),
        new ExternalHyperlink({
          link: "https://www.npmjs.com/package/cli-scaffold",
          children: [new TextRun({ text: "npmjs.com/package/cli-scaffold", size: 19, color: COLOR_ACCENT, font: "Arial", underline: { type: UnderlineType.SINGLE, color: COLOR_ACCENT } })]
        }),
        normal(".")
      ]),

      bullet([
        bold("Published Research — Efficient Graph Indexing: "),
        normal("Co-authored a paper on "),
        bold("approximate nearest-neighbor search"),
        normal(" for large-scale knowledge graphs, accepted at "),
        bold("IEEE ICDE 2025"),
        normal(". "),
        new ExternalHyperlink({
          link: "https://example.com/paper",
          children: [new TextRun({ text: "[View Paper]", size: 19, color: COLOR_ACCENT, font: "Arial", underline: { type: UnderlineType.SINGLE, color: COLOR_ACCENT } })]
        }),
      ]),

      bullet([bold("HackMIT 2024 — 1st Place: "), normal("Won Best Developer Tool award among "), accent("400+ participants"), normal(" for an AI-powered code review bot.")]),
      bullet([bold("Google Code Jam 2024: "), normal("Ranked in the "), accent("top 2%"), normal(" globally (Round 2).")]),
      bullet([bold("Tech Blog — alexmorgan.dev: "), normal("Write about system design and backend engineering; articles have received "), accent("50K+ total views"), normal(".")]),

    ]
  }]
});

// ─────────────────────────────────────────
//  Write the .docx file
// ─────────────────────────────────────────
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("resume.docx", buffer);
  console.log("✅  resume.docx generated successfully!");
});
