const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, convertInchesToTwip, ImageRun
} = require("docx");
const fs = require("fs");
const path = require("path");

const WORDMARK_PNG = fs.readFileSync(path.join(__dirname, "logo-assets", "nudge-wordmark-ink.png"));

const PAGE_WIDTH = 12240; // US Letter
const PAGE_HEIGHT = 15840;
const MARGIN = 1080; // 0.75"

const ACCENT = "1F4E5F";
const LIGHT = "EAF1F3";
const GRAY = "595959";

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    border: { bottom: { color: ACCENT, space: 4, style: BorderStyle.SINGLE, size: 6 } },
    children: [new TextRun({ text, bold: true, color: ACCENT, size: 30 })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, bold: true, color: ACCENT, size: 24 })],
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, size: 21, ...opts })],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { after: 60 },
    children: [new TextRun({ text, size: 21 })],
  });
}

function boldLead(lead, rest) {
  return new Paragraph({
    spacing: { after: 60 },
    numbering: { reference: "bullets", level: 0 },
    children: [
      new TextRun({ text: lead, bold: true, size: 21 }),
      new TextRun({ text: rest, size: 21 }),
    ],
  });
}

function cell(text, opts = {}) {
  const { width, bold = false, shade = null, align = AlignmentType.LEFT, color = "000000" } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade } : undefined,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold, size: 19, color })],
    })],
  });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 260 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 260 } } } },
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
      new Paragraph({
        spacing: { after: 100 },
        children: [
          new ImageRun({
            type: "png",
            data: WORDMARK_PNG,
            transformation: { width: 118, height: 50 },
          }),
        ],
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: "AI CONSULTING LAUNCH PLAN", bold: true, size: 40, color: ACCENT })],
      }),
      new Paragraph({
        spacing: { after: 300 },
        children: [new TextRun({ text: "Positioning, pricing, and a 90-day path to first clients", size: 21, color: GRAY, italics: true })],
      }),

      // ---- Section: Why niche ----
      h1("Why Niche Down"),
      p("The market for AI consulting is real and growing fast (~26% CAGR, $10–12B today heading toward $70–90B by the early 2030s), but “I do AI consulting” is no longer a differentiated pitch — too many people are saying it since ChatGPT. The data points to a specific, sellable problem instead:"),
      boldLead("The gap, not the hype: ", "55% of small businesses already use AI, but 80% report no measurable business impact. That gap — tools adopted, no results — is a stronger opening pitch than “you need AI.”"),
      boldLead("Premium goes to specificity: ", "Industry-specialist consultants command 30–40% higher fees than generalists. AI-specific expertise commands a 40–60% premium over generalist IT consulting."),
      boldLead("Generalist AI strategy is crowded: ", "McKinsey/BCG/Accenture own the top; a flood of new “AI consultants” crowd the bottom. The open lane is a specific outcome, for a specific type of client, in a specific industry."),
      p("You said you have some background across tech, a specific industry, and marketing/sales/ops, and that you're open to serving businesses of any size. That's a genuine asset — but it's an asset for choosing among strong options, not for pitching all of them at once. Below are three concrete directions built from that mix. Pick one to lead with; the others can become add-on services once you have traction."),

      // ---- Section: Three directions ----
      h1("Three Positioning Directions"),

      h2("Option A — AI Operations Consultant for Small Business"),
      p("Best if: your tech + ops/marketing background is stronger than any single industry tie."),
      bullet("Pitch: “I find the 2–3 workflows costing you the most time — customer service, scheduling, or reporting — and implement AI + automation to fix them, with numbers to prove it.”"),
      bullet("Why it works: SMBs want workflow redesign and implementation, not tool tours. Most consultants skip straight to “use this tool” and never touch the underlying process — which is exactly why 80% see no impact."),
      bullet("Client size: small business / solopreneur. Fastest sales cycle, lowest budgets, easiest to land first paying clients and case studies."),
      bullet("Ceiling: lower per-deal size, but higher deal volume and repeatable playbooks."),

      h2("Option B — Industry-Specialist AI Implementation Partner"),
      p("Best if: you have real depth (worked in, sold into, or studied) one specific non-tech industry — e.g. healthcare, legal, real estate, manufacturing, professional services."),
      bullet("Pitch: “I'm the AI consultant who's actually worked in [industry] — I know your compliance constraints, your systems, and where the real time sinks are.”"),
      bullet("Why it works: domain credibility shortens the sales cycle and justifies a 30–40% fee premium. Buyers in regulated or high-trust industries want someone who won't need a crash course in how their business works."),
      bullet("Client size: mid-size companies (50–500 employees) are the sweet spot — enough budget to fund real projects, small enough that you can reach a decision-maker directly."),
      bullet("Ceiling: highest — this is the path to premium day rates and retainers, and eventually a fractional/part-time “Head of AI” role ($100–200K/yr for 1–3 days/week is an emerging category)."),

      h2("Option C — AI Marketing & Sales Systems Consultant"),
      p("Best if: your marketing/sales/ops background is your strongest card."),
      bullet("Pitch: “I build AI-powered lead gen, content, and follow-up systems that get your sales team more qualified conversations without more headcount.”"),
      bullet("Why it works: 77% of small businesses see marketing as the highest-impact area for AI, and marketing/sales ROI is easy to measure — which makes it easy to sell and easy to prove."),
      bullet("Client size: works across small business and mid-size; marketing budgets exist even when “AI transformation” budgets don't."),
      bullet("Ceiling: medium — clear ROI story, but more competition from marketing agencies also bolting on “AI”."),

      h2("Recommendation"),
      p("Given a mixed background and no committed niche yet: start with Option A or C to get 2–3 paying clients and real case studies fast (short sales cycle, low risk for the buyer), while you decide whether a specific industry (Option B) is worth specializing into once you see which clients respond best. Don't market all three at once — pick the language for your first outreach and let the case studies decide the rest."),

      // ---- Section: Pricing ----
      h1("Pricing Packages"),
      p("Rates below reflect current market benchmarks for independent / boutique AI consultants (2026)."),

      new Table({
        width: { size: 10080, type: WidthType.DXA },
        columnWidths: [2800, 3800, 3480],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              cell("Package", { width: 2800, bold: true, shade: ACCENT, color: "FFFFFF" }),
              cell("What's included", { width: 3800, bold: true, shade: ACCENT, color: "FFFFFF" }),
              cell("Price", { width: 3480, bold: true, shade: ACCENT, color: "FFFFFF" }),
            ],
          }),
          new TableRow({
            children: [
              cell("Discovery / Audit", { width: 2800, shade: LIGHT, bold: true }),
              cell("Workflow assessment + prioritized implementation roadmap. Low-risk entry point for new clients.", { width: 3800, shade: LIGHT }),
              cell("$2,000 – $5,000\nflat fee", { width: 3480, shade: LIGHT }),
            ],
          }),
          new TableRow({
            children: [
              cell("Implementation", { width: 2800, bold: true }),
              cell("Build and deploy one workflow (e.g. customer service automation, lead follow-up, reporting). Includes team training.", { width: 3800 }),
              cell("$5,000 – $25,000\nper project", { width: 3480 }),
            ],
          }),
          new TableRow({
            children: [
              cell("AI Voice Agent", { width: 2800, shade: LIGHT, bold: true }),
              cell("Custom-trained voice agent for lead intake, appointment booking, or FAQ handling. SMB single-purpose tier.", { width: 3800, shade: LIGHT }),
              cell("$3,000 – $12,000\nper project", { width: 3480, shade: LIGHT }),
            ],
          }),
          new TableRow({
            children: [
              cell("Predictive Analytics Dashboard", { width: 2800, bold: true }),
              cell("Turn existing data into a forecasting / decision-support dashboard. Single-dashboard scope.", { width: 3800 }),
              cell("$5,000 – $15,000\nper project", { width: 3480 }),
            ],
          }),
          new TableRow({
            children: [
              cell("Ongoing Retainer", { width: 2800, shade: LIGHT, bold: true }),
              cell("Monitoring, maintenance, API/tool updates, quarterly strategy check-ins.", { width: 3800, shade: LIGHT }),
              cell("$500 – $3,000\nper month", { width: 3480, shade: LIGHT }),
            ],
          }),
          new TableRow({
            children: [
              cell("Hourly / Ad hoc", { width: 2800, bold: true }),
              cell("For scoping calls, one-off troubleshooting, or clients not ready for a project.", { width: 3800 }),
              cell("$75 – $250/hr\n(higher with industry specialization)", { width: 3480 }),
            ],
          }),
        ],
      }),
      new Paragraph({ spacing: { before: 160, after: 100 }, children: [new TextRun({ text: "Early on, price toward the low end of Discovery and Implementation to build case studies quickly — raise rates as you accumulate proof, not before.", italics: true, size: 19, color: GRAY })] }),
      p("AI Voice Agent and Predictive Analytics Dashboard are upsells, not the lead offer. Pitch them to clients who've already been through a Discovery or Implementation engagement and trust you — not cold prospects. Leading with these before you have workflow-automation case studies undercuts the fast, low-friction sales motion the rest of this plan is built around."),

      // ---- Section: Phase 2 recurring revenue ----
      h1("Phase 2 — Recurring Revenue Ladder (Later)"),
      p("Do not lead with this. A monthly subscription is a bigger ask than a bounded project fee, and asking a cold stranger to commit to one before you've proven anything is a harder sell than it needs to be — it works against the low-friction motion this whole plan is built around. Introduce this ladder only after you've completed 2–3 project engagements and have real, named results to point to. At that point you're not selling a subscription to a stranger — you're offering an existing client a way to keep the value going, which is a much easier yes."),
      p("This also solves a real operational risk: promising ongoing weekly service before you've delivered a single project risks overcommitting capacity you haven't tested yet. Build the muscle on bounded projects first."),

      new Table({
        width: { size: 10080, type: WidthType.DXA },
        columnWidths: [2400, 4680, 3000],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              cell("Tier", { width: 2400, bold: true, shade: ACCENT, color: "FFFFFF" }),
              cell("What's included", { width: 4680, bold: true, shade: ACCENT, color: "FFFFFF" }),
              cell("Price", { width: 3000, bold: true, shade: ACCENT, color: "FFFFFF" }),
            ],
          }),
          new TableRow({
            children: [
              cell("Core Care", { width: 2400, shade: LIGHT, bold: true }),
              cell("Monitoring & maintenance of what's already been built, monthly performance report, minor tweaks, priority email support.", { width: 4680, shade: LIGHT }),
              cell("$997/mo\nno long-term contract", { width: 3000, shade: LIGHT }),
            ],
          }),
          new TableRow({
            children: [
              cell("Growth Partner", { width: 2400, bold: true }),
              cell("Everything in Core Care, plus AI voice agent monitoring & optimization, a quarterly strategy session, and one new minor workflow per quarter.", { width: 4680 }),
              cell("$2,497/mo\nno long-term contract", { width: 3000 }),
            ],
          }),
          new TableRow({
            children: [
              cell("Full Systems Partner", { width: 2400, shade: LIGHT, bold: true }),
              cell("Everything in Growth Partner, plus predictive analytics dashboard maintenance & reporting, a dedicated monthly strategist call, and new workflow additions included.", { width: 4680, shade: LIGHT }),
              cell("$4,997/mo\nno long-term contract", { width: 3000, shade: LIGHT }),
            ],
          }),
        ],
      }),
      new Paragraph({ spacing: { before: 160, after: 0 }, children: [new TextRun({ text: "Unlike the project packages, these tiers assume the underlying system is already built — this is a management/optimization subscription for existing clients, not a bundled build-plus-subscription offer to new ones. “No long-term contract” lowers the perceived risk of the ask, the same way it does in the project pricing.", italics: true, size: 19, color: GRAY })] }),

      // ---- Section: 90 day plan ----
      h1("First 90 Days"),

      h2("Days 1–15 — Commit and Position"),
      bullet("Pick one option (A, B, or C) as your lead positioning. Write a one-line pitch and a one-paragraph “who I help / what changes” statement."),
      bullet("List 20 people or businesses you already know (past colleagues, local businesses, your network) who fit the target client profile."),
      bullet("Set up a simple landing page or LinkedIn profile rewrite reflecting the new positioning — not a full website yet."),

      h2("Days 16–45 — Land a Pilot Client"),
      bullet("Offer 1–2 people from your list a discounted or free Discovery engagement in exchange for a case study and testimonial. Real, named results are your most valuable asset right now."),
      bullet("Scope the engagement tightly: one workflow, one measurable outcome (hours saved, response time, conversion rate)."),
      bullet("Deliver, measure the before/after, and write it up as a one-page case study with real numbers."),

      h2("Days 46–75 — Convert Proof into Pipeline"),
      bullet("Use the case study in outreach: cold email/LinkedIn to 5–10 prospects per week in your target niche, referencing the specific, measurable result."),
      bullet("Ask your pilot client for 2–3 warm referrals — referred leads convert far faster than cold ones."),
      bullet("Start charging full Discovery/Implementation rates for new clients."),

      h2("Days 76–90 — Decide and Double Down"),
      bullet("Review which pitch (A, B, or C) got the fastest yeses and best-fit clients — let real response data, not preference, decide your niche."),
      bullet("If an industry pattern emerged (e.g. most interest came from one vertical), consider narrowing into Option B for the premium and defensibility it offers."),
      bullet("Set a recurring lead-generation habit (weekly outreach quota, content, or referral ask) so pipeline doesn't stall between projects."),

      // ---- Footer note ----
      new Paragraph({ spacing: { before: 300 }, border: { top: { color: "CCCCCC", space: 4, style: BorderStyle.SINGLE, size: 4 } }, children: [
        new TextRun({ text: "Prepared for Ahmed — based on 2026 AI consulting market research (Fortune Business Insights, ColorWhistle, Layer3 Labs, Dan Cumberland Labs).", size: 17, color: GRAY, italics: true }),
      ]}),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync("AI_Consulting_Launch_Plan.docx", buf);
  console.log("written");
});
