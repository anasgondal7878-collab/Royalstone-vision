const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, TableOfContents,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
  VerticalAlign, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 1080;
const CONTENT_W = PAGE_W - MARGIN * 2; // 10080

const NAVY = "1F2D3D";
const GOLD = "8C6A2E";
const LIGHTGOLD = "F3ECDD";
const LIGHTGREY = "F2F2F2";
const BORDERCLR = "CCCCCC";

const border = { style: BorderStyle.SINGLE, size: 1, color: BORDERCLR };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, opts = {}) {
  const { width, bold = false, fill = null, align = AlignmentType.LEFT, color = "000000", size = 21 } = opts;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold, color, size })]
    })]
  });
}

function sectionHeading(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function subHeading(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function subsubHeading(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160 },
    children: [new TextRun({ text, italics: opts.italics || false, bold: opts.bold || false, size: opts.size || 22 })]
  });
}
function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: opts.size || 22, bold: opts.bold || false })]
  });
}
function note(text) {
  return new Paragraph({
    spacing: { after: 200, before: 80 },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: GOLD, space: 8 } },
    children: [new TextRun({ text, italics: true, size: 20, color: "555555" })]
  });
}
function divider() {
  return new Paragraph({
    spacing: { before: 100, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 1 } },
    children: [new TextRun("")]
  });
}

function makeTable(headers, rows, widths) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => cell(h, { width: widths[i], bold: true, fill: NAVY, color: "FFFFFF", align: AlignmentType.LEFT }))
  });
  const dataRows = rows.map((r, idx) => new TableRow({
    children: r.map((val, i) => cell(String(val), { width: widths[i], fill: idx % 2 === 1 ? LIGHTGREY : null }))
  }));
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, ...dataRows]
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Calibri", size: 22, color: "222222" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Calibri", color: NAVY },
        paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 4 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Calibri", color: NAVY },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 23, bold: true, font: "Calibri", color: GOLD },
        paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            new TextRun({ text: "ROYAL STONE INTELLIGENCE", size: 16, color: "888888", bold: true }),
            new TextRun({ text: "\tStrategic Plan & Financial Model", size: 16, color: "888888" })
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888" }),
            new TextRun({ text: " of ", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "888888" }),
          ]
        })]
      })
    },
    children: [

      // ===== COVER =====
      new Paragraph({ spacing: { before: 1600 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "ROYAL STONE", size: 64, bold: true, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
        children: [new TextRun({ text: "INTELLIGENCE", size: 44, bold: true, color: GOLD })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 600 },
        children: [new TextRun({ text: "A PropTech Intelligence Platform for the Islamabad-Rawalpindi Corridor", size: 24, italics: true, color: "555555" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 800 },
        children: [new TextRun({ text: "Family Strategy & Vision Document", size: 22, bold: true })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100 },
        children: [new TextRun({ text: "Prepared by Anas  |  June 2026", size: 20, color: "666666" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1600 },
        children: [new TextRun({ text: "All figures in this document are planning projections based on stated assumptions, not guaranteed outcomes.", size: 18, italics: true, color: "999999" })] }),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== TOC =====
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Table of Contents")] }),
      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 1. VISION =====
      sectionHeading("1. The Vision"),
      body("Royal Stone has operated as a real estate business in Islamabad for years, built on relationships, local knowledge, and trust across the housing societies of the Islamabad-Rawalpindi corridor. This document lays out a plan to build on that foundation with something new: a data and intelligence layer that turns Royal Stone from a single real estate office into a platform serving the entire region."),
      body("The core idea is simple. Right now, anyone buying, selling, or investing in property across these societies is making decisions with incomplete information — scattered listings, word-of-mouth pricing, no organized record of trends. Royal Stone Intelligence closes that gap, and charges a subscription for it."),
      body("This is not a replacement for the existing real estate business. It is a new, additional revenue stream that uses the same network, the same relationships, and the same on-the-ground knowledge the business already has — packaged into products people pay for every month."),
      note("A note on how to read this document: every number from this point forward is a projection built on assumptions stated clearly in each section. They are planning tools, not promises. The purpose is to show a credible, honest path — not to oversell."),

      // ===== 2. THE FOUR PRODUCTS =====
      sectionHeading("2. The Four Products"),
      body("The platform is built around four products, each serving a different type of customer at a different price point. They are designed to work together — the free products bring people in, the paid products are where the revenue comes from."),

      subHeading("2.1 Market Intelligence Report — PKR 2,500/month"),
      body("A monthly report covering price trends, new listings, and market movement across the covered societies. This is the entry-level product: affordable, useful to almost anyone with an interest in the local property market, and the easiest product to convert a free user into a paying one."),

      subHeading("2.2 Portfolio Tracker — PKR 8,000/quarter"),
      body("For people who already own property and want to track its value and the surrounding market over time. Priced quarterly because the value is best demonstrated over a longer window, not a single month. The Market Intelligence Report is included free with this subscription, since the data overlaps."),

      subHeading("2.3 First-Time Investor Onboarding — Free"),
      body("An AI-guided onboarding flow that asks a new investor a series of questions about budget, goals, and risk appetite, then matches them to relevant listings. This product is intentionally free. Its job is not to generate direct revenue but to bring new people into the platform, build trust with first-time buyers, and create the pipeline that later upgrades into paid subscriptions or direct property transactions (where Royal Stone's core brokerage business earns its usual commission)."),

      subHeading("2.4 Inner Circle — Free for existing customers / PKR 5,000/month for associate members"),
      body("Early access (48 hours before public release) to new property listings. Past Royal Stone customers get this free, as a loyalty and referral tool. New \u201Cassociate members\u201D who have no prior history with Royal Stone pay PKR 5,000/month for the same early access. The Market Intelligence Report is included free with this tier as well."),

      subsubHeading("How the bundling works"),
      bullet("Market Intelligence Report: standalone entry product at PKR 2,500/month"),
      bullet("Portfolio Tracker: PKR 8,000/quarter, includes the Market Intelligence Report at no extra cost"),
      bullet("Inner Circle (associate): PKR 5,000/month, includes the Market Intelligence Report at no extra cost"),
      bullet("Inner Circle (past customers): free, a retention and referral tool rather than a revenue line"),
      bullet("First-Time Investor Onboarding: always free, the top-of-funnel lead generator"),

      makeTable(
        ["Product", "Price", "Includes Market Intel?", "Primary Role"],
        [
          ["Market Intelligence Report", "PKR 2,500/mo", "—", "Entry-level revenue"],
          ["Portfolio Tracker", "PKR 8,000/quarter", "Yes, free", "Higher-value revenue"],
          ["Investor Onboarding (AI)", "Free", "No", "Lead generation"],
          ["Inner Circle (past customer)", "Free", "Yes, free", "Retention / referral"],
          ["Inner Circle (associate)", "PKR 5,000/mo", "Yes, free", "Premium revenue"],
        ],
        [3200, 2200, 2300, 2380]
      ),
      new Paragraph({ spacing: { after: 240 } }),

      // ===== 3. WHERE TO BEGIN =====
      sectionHeading("3. Where to Begin"),
      body("This is being built from scratch — there is no existing subscriber base or live product today. The honest starting point is the network: an estimated 50-200 active agents and dealers already connected to Royal Stone across the covered societies. That network is the seed for everything that follows."),

      subHeading("3.1 What's needed before launch"),
      bullet("A working MVP: web app with the AI-guided onboarding flow and the basic structure for the subscription products."),
      bullet("Organized listings and pricing data across the societies — this currently exists informally through the family's network but needs to be collected and structured properly. This is the single most important piece of groundwork."),
      bullet("Payment infrastructure (Stripe or a Pakistan-compatible alternative) and basic subscriber management."),
      bullet("An initial group of 20-30 first-time investors run through the AI onboarding flow manually, to validate the matching logic before opening it up publicly."),

      subHeading("3.2 Who builds it"),
      body("In the first phase, Anas builds the MVP, the offer/landing pages, and the AI integration directly. Royal Stone's existing network and market credibility (through Anas's father) provide the data, trust, and initial customer access. No paid hires are planned until the platform shows real traction — the first hire will be a part-time support or operations person once there's a subscriber base big enough to need one."),

      // ===== 4. PHASED ROLLOUT =====
      sectionHeading("4. Phased Rollout (12-Month Build, Multi-Year Expansion)"),
      body("The plan moves in four phases. The first year is mostly building and proving the model in a small number of societies before expanding."),

      subHeading("Phase 0 — Months 1-4: Build"),
      bullet("Build the MVP: web app, AI onboarding flow, basic admin tools."),
      bullet("Begin structured data collection from the existing agent/dealer network."),
      bullet("Run the first 20-30 investor leads through the AI onboarding manually, no paid products live yet."),

      subHeading("Phase 1 — Months 5-8: Pilot Launch"),
      bullet("Launch the Market Intelligence Report and Inner Circle in 2-3 pilot societies where data quality is strongest."),
      bullet("Existing Royal Stone customers get Inner Circle access free, to build goodwill and referrals."),
      bullet("Begin onboarding the first associate (paid) Inner Circle members."),

      subHeading("Phase 2 — Months 9-12: First Cohort"),
      bullet("Launch the Portfolio Tracker."),
      bullet("Expand coverage from 2-3 societies to 6-8."),
      bullet("Target: 150-300 paying subscribers by the end of month 12."),

      subHeading("Phase 3 — Year 2: Scale Locally"),
      bullet("Expand to the full 14-16 societies across the corridor."),
      bullet("Make the first part-time hire (operations or customer support)."),
      bullet("Target: 1,000-1,500 paying subscribers."),

      subHeading("Phase 4 — Year 3: Regional Expansion"),
      bullet("If the model is proven in Islamabad-Rawalpindi, expand to a second corridor (for example, Lahore)."),
      bullet("Target: 3,000-5,000 paying subscribers."),
      bullet("This is the stage where serious conversations about acquisition interest could realistically begin."),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 5. COSTS =====
      sectionHeading("5. Running Costs"),
      body("One of the strongest features of this business model is how little it costs to run, especially in the first year. Below is the Year 1 cost breakdown, based on a tight working budget."),

      makeTable(
        ["Cost Item", "Year 1 Estimate (PKR)", "Notes"],
        [
          ["Domain & hosting", "25,000", "Vercel + Supabase, low usage tier"],
          ["AI API costs (Claude)", "150,000", "Scales with onboarding & report volume"],
          ["Payment gateway / messaging", "80,000", "Stripe setup, WhatsApp Business API, email"],
          ["Marketing & ads", "100,000", "Organic-first, small paid test budget"],
          ["Total Year 1 running cost", "355,000", "~US$1,270 — fits a tight budget"],
        ],
        [3600, 2800, 3680]
      ),
      note("Anas's own time is not included as a cost in Year 1 — this is sweat equity, not a salary line. A salary or founder draw becomes realistic once subscription revenue is consistent (see Section 7)."),

      subHeading("5.1 Costs as the business scales"),
      makeTable(
        ["Year", "Estimated Annual Cost (PKR)", "What Changes"],
        [
          ["Year 1", "355,000", "Solo-built, minimal infrastructure"],
          ["Year 2", "~2,600,000", "First part-time hire, higher AI/infra usage, bigger marketing budget"],
          ["Year 3", "~6,000,000 (estimate)", "Small team, broader marketing for regional expansion"],
        ],
        [2200, 3600, 4280]
      ),

      // ===== 6. REVENUE PROJECTIONS =====
      sectionHeading("6. Revenue Projections"),
      body("These projections use a blended average revenue per paying subscriber of approximately PKR 3,500-4,200 per month, reflecting a mix of Market Intelligence-only subscribers (larger in number, lower price) and Portfolio Tracker / associate Inner Circle subscribers (fewer in number, higher price). This is a model assumption, not a guarantee of actual subscriber mix."),

      makeTable(
        ["Year", "Paying Subscribers (est.)", "Blended ARPU (PKR/mo)", "Monthly Run-Rate (PKR)", "Annualized Run-Rate (PKR)"],
        [
          ["Year 1 (end)", "~200", "3,500", "700,000", "8,400,000"],
          ["Year 2 (end)", "~1,200", "3,800", "4,560,000", "54,720,000"],
          ["Year 3 (end)", "~4,000", "4,200", "16,800,000", "201,600,000"],
        ],
        [1800, 2600, 2400, 1900, 1380]
      ),
      note("\u201CRun-rate\u201D means the revenue pace at that point in time, annualized. It is not the actual cash collected across the year, which will be lower in Years 1-2 because subscriber numbers build up gradually rather than starting at the year-end figure."),

      subHeading("6.1 Year 1 net position"),
      makeTable(
        ["Item", "Estimate (PKR)"],
        [
          ["Actual Year 1 revenue collected (ramping, not flat)", "2,000,000 - 2,800,000"],
          ["Year 1 running costs", "355,000"],
          ["Estimated Year 1 surplus (before founder salary)", "1,600,000 - 2,400,000"],
        ],
        [6500, 3580]
      ),
      body("From Year 2 onward, margins are expected to be strong — typically 60-70% or higher — because the core cost base (hosting, AI, a small team) grows much more slowly than subscription revenue does. This is the normal pattern for subscription data businesses, not something unique to this plan, but it only holds if subscriber growth tracks the projections above."),
      new Paragraph({ children: [new PageBreak()] }),

      // ===== 7. MARKETING & DATA COLLECTION =====
      sectionHeading("7. Marketing Strategy"),
      body("Marketing leans on what Royal Stone already has: real relationships with agents and dealers across the corridor, and Anas's experience building audiences for Reset90. The approach in Year 1 is deliberately low-cost and relationship-driven rather than paid-ad-heavy."),
      bullet("Direct outreach through the existing 50-200 agent/dealer network — the fastest, lowest-cost path to the first paying subscribers."),
      bullet("Free Inner Circle access for past Royal Stone customers, used explicitly as a referral and word-of-mouth driver."),
      bullet("The free AI Investor Onboarding tool itself doubles as a marketing asset — something genuinely useful to share, that naturally introduces people to the paid products."),
      bullet("Local content (market updates, society-specific price trends) shared through WhatsApp groups and social channels already used by agents in the corridor."),
      bullet("A small paid ad test budget in Phase 1-2 to validate which channels convert, before committing larger marketing spend in Year 2."),

      sectionHeading("8. Data Collection Methods"),
      body("Reliable data is the foundation the entire platform is built on — without accurate listings and pricing, none of the four products hold value. This is the area requiring the most hands-on work in Phase 0."),
      bullet("Direct collection from the existing agent/dealer network — manually at first, structured into a consistent format (price, location, size, society, date)."),
      bullet("Society-by-society verification, starting with the 2-3 pilot societies with the strongest existing relationships, before expanding."),
      bullet("The AI Investor Onboarding flow itself generates a feed of real demand data (what buyers are actually looking for) that sharpens the Market Intelligence Reports over time."),
      bullet("As Portfolio Tracker subscribers come on board, their tracked properties add an ongoing, self-updating layer of verified data."),
      note("This is the part of the plan with the least margin for shortcuts. Bad data undermines every product at once, so Phase 0 deliberately allocates the most time to getting this right before any paid product goes live."),

      // ===== 9. EXIT / IPO =====
      sectionHeading("9. Long-Term Potential: Acquisition or IPO"),
      body("It's important to be honest about this section, since it's the one most prone to overpromising. The realistic long-term outcome for a business like this is acquisition by a larger player, not a stock market listing."),

      subHeading("9.1 Why acquisition, not IPO"),
      body("An IPO on the Pakistan Stock Exchange requires a scale of revenue, profit history, and regulatory readiness far beyond what this business will reach even at the optimistic Year 3 numbers above. IPO is mentioned here only as a distant, low-probability, Year 5-plus scenario contingent on successful regional expansion — not as a near-term goal."),
      body("A far more realistic path is acquisition by a larger real estate portal, a regional fintech, or a property group looking to add a data and intelligence layer to their existing platform. The value being built here is the data relationship across 14-16 societies and the trust Royal Stone already carries locally — both of which are genuinely hard for a larger company to replicate quickly, which is exactly what makes the business attractive to acquire rather than copy."),

      subHeading("9.2 Illustrative acquisition scenario"),
      makeTable(
        ["Metric", "Year 3 Estimate"],
        [
          ["Annualized run-rate revenue (ARR)", "PKR 201,600,000 (~US$720,000)"],
          ["Illustrative valuation multiple (subscription data businesses)", "3x - 6x ARR"],
          ["Illustrative valuation range", "PKR ~605M - 1.21B (~US$2.2M - 4.3M)"],
        ],
        [5800, 4280]
      ),
      note("This valuation range is an illustration of how subscription businesses are sometimes valued elsewhere, applied to a projected number — not a forecast, an offer, or anything resembling a guarantee. Actual outcomes depend on real subscriber growth, data quality, and market conditions that can't be predicted today."),

      // ===== 10. CLOSING =====
      sectionHeading("10. Closing Note"),
      body("This plan is built to be honest rather than impressive. The numbers above are a model, not a promise — they exist to show that there is a credible, low-cost, low-risk way to start, with a clear path to something significant if the early phases prove the model out."),
      body("The starting point is small and deliberately so: a working MVP, the network the family already has, and a tight first-year budget. Every stage after that depends on real subscribers, real data, and real results, not on assumptions holding up automatically."),
      body("This is shared as a vision for where Royal Stone can go next, built on everything the business already is."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/royalstone/RoyalStone_Intelligence_Strategy.docx", buffer);
  console.log("done");
});
