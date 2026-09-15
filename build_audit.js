const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, ImageRun
} = require("docx");
const fs = require("fs");
const path = require("path");

const WORDMARK_PNG = fs.readFileSync(path.join(__dirname, "logo-assets", "nudge-wordmark-ink.png"));

const PAGE_WIDTH = 12240; // US Letter
const PAGE_HEIGHT = 15840;
const MARGIN = 1080;

const ACCENT = "1F4E5F";
const LIGHT = "EAF1F3";
const GRAY = "595959";
const WARN = "B5651D";

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    border: { bottom: { color: ACCENT, space: 4, style: BorderStyle.SINGLE, size: 6 } },
    children: [new TextRun({ text, bold: true, color: ACCENT, size: 28 })],
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, size: 21, ...opts })],
  });
}

function bracket(text) {
  return new TextRun({ text, size: 21, color: "8C8C8C", italics: true });
}

function bullet(children, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { after: 60 },
    children,
  });
}

function cell(children, opts = {}) {
  const { width, shade = null, verticalAlign } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade } : undefined,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: Array.isArray(children) ? children : [children],
  });
}

function headerCell(text, width) {
  return cell(
    [new Paragraph({ children: [new TextRun({ text, bold: true, size: 19, color: "FFFFFF" })] })],
    { width, shade: ACCENT }
  );
}

function bodyCell(text, width, shade = null, opts = {}) {
  return cell(
    [new Paragraph({ children: [new TextRun({ text, size: 19, ...opts })] })],
    { width, shade }
  );
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 260 } } } },
      ],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
        margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
      },
    },
    children: [
      // Brand header
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new ImageRun({
            type: "png",
            data: WORDMARK_PNG,
            transformation: { width: 118, height: 50 },
          }),
        ],
      }),
      new Paragraph({
        spacing: { after: 260 },
        children: [new TextRun({ text: "Give your business its hours back.", size: 20, color: GRAY, italics: true })],
      }),

      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: "WORKFLOW AUDIT", bold: true, size: 32, color: "000000" })],
      }),

      // Meta table (fill-in fields)
      new Table({
        width: { size: 10080, type: WidthType.DXA },
        columnWidths: [2500, 5580, 2000],
        rows: [
          new TableRow({ children: [
            bodyCell("Prepared for:", 2500, LIGHT, { bold: true }),
            cell([new Paragraph({ children: [bracket("[Client business name]")] })], { width: 5580 }),
            bodyCell("Date:", 2000, LIGHT, { bold: true }),
          ]}),
          new TableRow({ children: [
            bodyCell("Prepared by:", 2500, LIGHT, { bold: true }),
            cell([new Paragraph({ children: [bracket("[Your name] · Nudge")] })], { width: 5580 }),
            cell([new Paragraph({ children: [bracket("[Date]")] })], { width: 2000 }),
          ]}),
        ],
      }),

      p(""),

      // Section 1
      h1("1. Business Snapshot"),
      bullet([new TextRun({ text: "Business: ", bold: true, size: 21 }), bracket("[name, industry, team size]")]),
      bullet([new TextRun({ text: "Tools currently in use: ", bold: true, size: 21 }), bracket("[CRM, scheduling tool, email, spreadsheets, etc.]")]),
      bullet([new TextRun({ text: "Goal for this audit: ", bold: true, size: 21 }), bracket("[what they said they wanted looked at]")]),

      // Section 2
      h1("2. Workflows Reviewed"),
      p("The processes below were walked through during the audit call/observation."),
      new Table({
        width: { size: 10080, type: WidthType.DXA },
        columnWidths: [2500, 3300, 1700, 2580],
        rows: [
          new TableRow({ children: [
            headerCell("Workflow", 2500),
            headerCell("Current process", 3300),
            headerCell("Time/week", 1700),
            headerCell("Pain points", 2580),
          ]}),
          new TableRow({ children: [
            bodyCell("[e.g. Customer inquiries]", 2500, LIGHT),
            cell([new Paragraph({ children: [bracket("[how it's done today]")] })], { width: 3300, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[hrs]")] })], { width: 1700, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[delays, errors, etc.]")] })], { width: 2580, shade: LIGHT }),
          ]}),
          new TableRow({ children: [
            bodyCell("[Workflow 2]", 2500),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 3300 }),
            cell([new Paragraph({ children: [bracket("[hrs]")] })], { width: 1700 }),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 2580 }),
          ]}),
          new TableRow({ children: [
            bodyCell("[Workflow 3]", 2500, LIGHT),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 3300, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[hrs]")] })], { width: 1700, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 2580, shade: LIGHT }),
          ]}),
        ],
      }),

      // Section 3
      h1("3. Where Time & Money Are Leaking"),
      p("Plain-language summary of the biggest findings — written for a business owner, not a technical audience."),
      bullet([new TextRun({ text: "Finding 1: ", bold: true, size: 21 }), bracket("[e.g. “Customer questions are answered manually, one at a time, even though 70% are the same 5 questions.”]")]),
      bullet([new TextRun({ text: "Finding 2: ", bold: true, size: 21 }), bracket("[...]")]),
      bullet([new TextRun({ text: "Finding 3: ", bold: true, size: 21 }), bracket("[...]")]),

      // Section 4
      h1("4. Recommended Fixes"),
      p("Ranked by impact vs. effort — highest-impact, lowest-effort items first."),
      new Table({
        width: { size: 10080, type: WidthType.DXA },
        columnWidths: [1200, 3200, 2600, 1600, 1480],
        rows: [
          new TableRow({ children: [
            headerCell("Priority", 1200),
            headerCell("Recommended fix", 3200),
            headerCell("Approach / tool", 2600),
            headerCell("Setup time", 1600),
            headerCell("Est. time saved/wk", 1480),
          ]}),
          new TableRow({ children: [
            cell([new Paragraph({ children: [new TextRun({ text: "1", bold: true, size: 19 })] })], { width: 1200, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[fix]")] })], { width: 3200, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[tool/approach]")] })], { width: 2600, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[days]")] })], { width: 1600, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[hrs]")] })], { width: 1480, shade: LIGHT }),
          ]}),
          new TableRow({ children: [
            cell([new Paragraph({ children: [new TextRun({ text: "2", bold: true, size: 19 })] })], { width: 1200 }),
            cell([new Paragraph({ children: [bracket("[fix]")] })], { width: 3200 }),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 2600 }),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 1600 }),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 1480 }),
          ]}),
          new TableRow({ children: [
            cell([new Paragraph({ children: [new TextRun({ text: "3", bold: true, size: 19 })] })], { width: 1200, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[fix]")] })], { width: 3200, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 2600, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 1600, shade: LIGHT }),
            cell([new Paragraph({ children: [bracket("[...]")] })], { width: 1480, shade: LIGHT }),
          ]}),
        ],
      }),

      // Section 5
      h1("5. Estimated Return"),
      new Table({
        width: { size: 10080, type: WidthType.DXA },
        columnWidths: [5040, 5040],
        rows: [
          new TableRow({ children: [
            bodyCell("Current cost of the problem", 5040, LIGHT, { bold: true }),
            cell([new Paragraph({ children: [bracket("[hrs/week × hourly value = $X/month]")] })], { width: 5040, shade: LIGHT }),
          ]}),
          new TableRow({ children: [
            bodyCell("Projected time saved after fix", 5040, null, { bold: true }),
            cell([new Paragraph({ children: [bracket("[hrs/week]")] })], { width: 5040 }),
          ]}),
          new TableRow({ children: [
            bodyCell("Estimated monthly value", 5040, LIGHT, { bold: true }),
            cell([new Paragraph({ children: [bracket("[$X/month]")] })], { width: 5040, shade: LIGHT }),
          ]}),
          new TableRow({ children: [
            bodyCell("Payback period on implementation cost", 5040, null, { bold: true }),
            cell([new Paragraph({ children: [bracket("[X weeks/months]")] })], { width: 5040 }),
          ]}),
        ],
      }),
      new Paragraph({ spacing: { before: 140 }, children: [new TextRun({ text: "Use conservative estimates here — under-promising and over-delivering builds the case study you need far more than an inflated number does.", italics: true, size: 19, color: WARN })] }),

      // Section 6
      h1("6. Recommended Next Step"),
      p("Based on this audit, the recommended next step is:"),
      bullet([new TextRun({ text: "Implementation package — ", bold: true, size: 21 }), bracket("[fix #1 above], estimated timeline [X weeks], investment [$X].")]),
      p(""),
      p("Questions or want to move forward? Reach out anytime.", { italics: true, color: GRAY }),
      new Paragraph({ spacing: { before: 40 }, children: [new TextRun({ text: "[Your name]  ·  [email]  ·  [phone]  ·  Nudge", size: 19, color: GRAY })] }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync("Nudge_Workflow_Audit_Template.docx", buf);
  console.log("written");
});
