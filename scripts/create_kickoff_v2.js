const pptxgen = require("pptxgenjs");

// Colors
const BL = "2D5F8A", DBL = "1B3A5C", GR = "00A443", DGR = "007A33";
const CH = "333333", DG = "555555", MG = "888888", LG = "F5F5F5", WH = "FFFFFF";
const RD = "DC3545", OR = "FF8C00", LBL = "E8F0FE", LGR = "E8F5E9";

// W=10, H=5.625
function hdr(s, title, sub) {
  s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.85, fill: { color: BL } });
  s.addText(title, { x: 0.4, y: 0.08, w: 8, h: 0.42, fontSize: 22, fontFace: "Arial", bold: true, color: WH, margin: 0 });
  if (sub) s.addText(sub, { x: 0.4, y: 0.47, w: 8, h: 0.28, fontSize: 10, fontFace: "Arial", color: WH, margin: 0 });
}
function ftr(s) {
  s.addShape("rect", { x: 0, y: 5.28, w: 10, h: 0.35, fill: { color: LG } });
  s.addText("Avangrid  |  Digital Products Transformation", { x: 0.4, y: 5.3, w: 5, h: 0.28, fontSize: 8, color: MG });
  s.addText("NTT DATA", { x: 8, y: 5.3, w: 1.6, h: 0.28, fontSize: 9, color: DBL, bold: true, align: "right" });
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "NTT DATA";
  pres.title = "Digital Products - Project Kickoff";

  // ═══ SLIDE 1: Cover ═══
  let s = pres.addSlide();
  s.background = { color: DBL };
  s.addText("Digital Products", { x: 0.6, y: 1.2, w: 8, h: 0.9, fontSize: 40, bold: true, color: WH, fontFace: "Arial" });
  s.addText("Project Kickoff — New Ways of Working", { x: 0.6, y: 2.1, w: 8, h: 0.5, fontSize: 18, color: "AACCEE", fontFace: "Arial" });
  s.addText("Wave 1  •  14 Weeks  •  BRS Pilot", { x: 0.6, y: 2.7, w: 8, h: 0.4, fontSize: 13, color: "88AACC", fontFace: "Arial" });
  s.addShape("rect", { x: 0, y: 4.85, w: 10, h: 0.78, fill: { color: GR } });
  s.addText("NTT DATA × Avangrid", { x: 0.6, y: 4.92, w: 4, h: 0.35, fontSize: 14, bold: true, color: WH });
  s.addText("April 2026", { x: 7, y: 4.92, w: 2.5, h: 0.35, fontSize: 12, color: WH, align: "right" });

  // ═══ SLIDE 2: Why We're Here ═══
  s = pres.addSlide();
  hdr(s, "Why We're Here", "The challenge and the value we'll deliver");
  ftr(s);

  s.addText("Digital Products is growing 30-50% — but without a unified operating model, growth creates chaos.", {
    x: 0.4, y: 0.95, w: 9.2, h: 0.32, fontSize: 10, color: DG
  });

  s.addText("THE CHALLENGE", { x: 0.4, y: 1.35, w: 3, h: 0.28, fontSize: 12, bold: true, color: RD });
  s.addText("THE VALUE", { x: 5.3, y: 1.35, w: 3, h: 0.28, fontSize: 12, bold: true, color: GR });

  const pairs = [
    ["Fragmented Intake", "Email, Teams, ServiceNow — no front door", "Demand Clarity", "One structured channel — teams know what to do"],
    ["Reactive Prioritization", "Urgency wins — no value criteria", "Strategic Decisions", "Value-based scoring replaces opinions"],
    ["No Traceability", "Can't connect request to outcome", "E2E Visibility", "Request → outcome in one view"],
    ["Zero Transparency", "Leadership flies blind", "Real-Time Dashboards", "Portfolio health visible in real time"],
    ["Tribal Knowledge", "Processes in people's heads", "Self-Sustaining Model", "Documented WoW — runs without us"],
  ];

  pairs.forEach(([ct, cd, vt, vd], i) => {
    const y = 1.72 + i * 0.58;
    s.addShape("rect", { x: 0.4, y, w: 4.1, h: 0.5, fill: { color: i % 2 === 0 ? "FFF0F0" : WH } });
    s.addText([
      { text: ct, options: { bold: true, fontSize: 10, color: CH, breakLine: true } },
      { text: cd, options: { fontSize: 8, color: DG } }
    ], { x: 0.55, y: y + 0.04, w: 3.8, h: 0.42, valign: "middle" });

    s.addText("→", { x: 4.6, y: y + 0.05, w: 0.3, h: 0.4, fontSize: 14, bold: true, color: GR, align: "center", valign: "middle" });

    s.addShape("rect", { x: 5.1, y, w: 4.5, h: 0.5, fill: { color: i % 2 === 0 ? LGR : WH } });
    s.addText([
      { text: vt, options: { bold: true, fontSize: 10, color: "006630", breakLine: true } },
      { text: vd, options: { fontSize: 8, color: DG } }
    ], { x: 5.25, y: y + 0.04, w: 4.2, h: 0.42, valign: "middle" });
  });

  s.addShape("rect", { x: 0.4, y: 4.7, w: 9.2, h: 0.4, fill: { color: DBL } });
  s.addText("14 weeks  •  1 pilot (BRS)  •  Built on proven assets from 4 projects  •  Scalable to all OPCOs", {
    x: 0.5, y: 4.72, w: 9, h: 0.36, fontSize: 10, color: WH, bold: true, align: "center"
  });

  // ═══ SLIDE 3: Approach — 4 Columns ═══
  s = pres.addSlide();
  hdr(s, "Our Approach", "Four dimensions driving the transformation");
  ftr(s);

  const dims = [
    { t: "Lifecycle &\nStrategy", sub: "Standardize how work flows", col: BL, items: ["Single intake channel", "Value-based prioritization", "E2E traceability"] },
    { t: "Tools &\nTransparency", sub: "Connected execution & visibility", col: "2E7D32", items: ["Jira workflows", "PPMNow portfolio", "Real-time reporting"] },
    { t: "People &\nGovernance", sub: "Clear ownership & rhythm", col: "E65100", items: ["RACI per phase", "Champions per area", "Weekly cadence"] },
    { t: "Pilot &\nAdoption", sub: "Validate before scaling", col: "6A1B9A", items: ["BRS E2E pilot", "WoW Playbook", "KPI framework"] },
  ];

  dims.forEach((d, i) => {
    const x = 0.3 + i * 2.4;
    const y = 1.05;

    // Card background
    s.addShape("rect", { x, y, w: 2.2, h: 4.0, fill: { color: WH },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.08 } });
    // Color accent top bar
    s.addShape("rect", { x, y, w: 2.2, h: 0.06, fill: { color: d.col } });
    // Color circle with initial letter
    s.addShape("oval", { x: x + 0.75, y: y + 0.25, w: 0.7, h: 0.7, fill: { color: d.col } });
    s.addText(d.t.charAt(0), { x: x + 0.75, y: y + 0.25, w: 0.7, h: 0.7, fontSize: 24, bold: true, color: WH, align: "center", valign: "middle" });
    // Title
    s.addText(d.t, { x: x + 0.1, y: y + 1.1, w: 2, h: 0.6, fontSize: 13, bold: true, color: CH, align: "center", valign: "middle" });
    // Subtitle
    s.addText(d.sub, { x: x + 0.15, y: y + 1.7, w: 1.9, h: 0.3, fontSize: 8, color: MG, align: "center" });
    // Divider line
    s.addShape("line", { x: x + 0.3, y: y + 2.1, w: 1.6, h: 0, line: { color: "E0E0E0", width: 1 } });
    // Items with green check marks
    d.items.forEach((item, j) => {
      s.addText("✓", { x: x + 0.2, y: y + 2.25 + j * 0.45, w: 0.25, h: 0.25, fontSize: 12, bold: true, color: GR });
      s.addText(item, { x: x + 0.45, y: y + 2.25 + j * 0.45, w: 1.6, h: 0.25, fontSize: 9, color: DG, valign: "middle" });
    });
  });

  // ═══ SLIDE 4: Scope & Success ═══
  s = pres.addSlide();
  hdr(s, "Scope & Success Criteria", "What's in Wave 1 and how we measure success");
  ftr(s);

  s.addText("WAVE 1 SCOPE", { x: 0.4, y: 0.98, w: 3, h: 0.28, fontSize: 12, bold: true, color: GR });
  const scope = [
    "Review of Prior Materials (4 projects)", "E2E Product Lifecycle Definition",
    "Value-Driven Prioritization Framework", "Live Pilot Execution (BRS)",
    "Tool Ecosystem Design (Jira + PPMNow)", "WoW Playbook + RACI Definition",
  ];
  scope.forEach((item, i) => {
    const y = 1.35 + i * 0.45;
    s.addShape("rect", { x: 0.4, y, w: 4.3, h: 0.38, fill: { color: i % 2 === 0 ? LGR : WH } });
    s.addText("✓", { x: 0.5, y, w: 0.25, h: 0.38, fontSize: 12, bold: true, color: GR, valign: "middle" });
    s.addText(item, { x: 0.8, y, w: 3.8, h: 0.38, fontSize: 10, color: CH, valign: "middle" });
  });
  s.addText("Wave 2: Capacity Mgmt, Governance, Org Redesign, Broad Change Mgmt", {
    x: 0.4, y: 4.15, w: 4.3, h: 0.25, fontSize: 7, color: MG, italic: true
  });

  s.addText("SUCCESS CRITERIA", { x: 5.2, y: 0.98, w: 3, h: 0.28, fontSize: 12, bold: true, color: BL });
  const crit = [
    ["Intake Clarity", "100% BRS demand via one structured channel"],
    ["Prioritization", "All demands scored by Value Framework V1.0"],
    ["E2E Lifecycle", "Pilot completes full lifecycle: Intake → Report"],
    ["Tool Integration", "Jira as single env + live dashboards"],
    ["Playbook", "WoW v1.0 validated, ready for Wave 2"],
  ];
  crit.forEach(([t, d], i) => {
    const y = 1.35 + i * 0.55;
    s.addShape("rect", { x: 5.2, y, w: 4.4, h: 0.48, fill: { color: i % 2 === 0 ? LBL : WH } });
    s.addShape("oval", { x: 5.35, y: y + 0.08, w: 0.3, h: 0.3, fill: { color: BL } });
    s.addText(String(i + 1), { x: 5.35, y: y + 0.08, w: 0.3, h: 0.3, fontSize: 10, color: WH, bold: true, align: "center", valign: "middle" });
    s.addText([
      { text: t, options: { bold: true, fontSize: 10, color: CH, breakLine: true } },
      { text: d, options: { fontSize: 8, color: DG } }
    ], { x: 5.8, y: y + 0.03, w: 3.7, h: 0.42, valign: "middle" });
  });

  // ═══ SLIDE 5: Timeline ═══
  s = pres.addSlide();
  hdr(s, "Wave 1 Timeline (14 Weeks)", "Key checkpoints ensure validation and readiness for scaling");
  ftr(s);

  s.addShape("rect", { x: 0.3, y: 1.1, w: 9.4, h: 0.28, fill: { color: LBL } });
  for (let w = 1; w <= 14; w++) {
    s.addText("W" + w, { x: 0.3 + (w - 1) * (9.4 / 14), y: 1.1, w: 9.4 / 14, h: 0.28, fontSize: 7, color: BL, align: "center", valign: "middle", bold: true });
  }

  const steps = [
    { name: "Step 1 — Review & Baseline", start: 1, end: 2, y: 1.6 },
    { name: "Step 2 — Define Lifecycle", start: 2, end: 5, y: 2.2 },
    { name: "Step 3 — Design Tools", start: 4, end: 7, y: 2.8 },
    { name: "Step 4 — Pilot Execution (BRS)", start: 7, end: 13, y: 3.4 },
    { name: "Step 5 — Learn & Adjust", start: 13, end: 14, y: 4.0 },
  ];
  steps.forEach(st => {
    const sx = 0.3 + (st.start - 1) * (9.4 / 14);
    const sw = (st.end - st.start + 1) * (9.4 / 14);
    s.addShape("rect", { x: sx, y: st.y - 0.03, w: sw, h: 0.03, fill: { color: BL } });
    s.addShape("rect", { x: sx, y: st.y, w: Math.min(sw, 2.8), h: 0.45, fill: { color: WH },
      shadow: { type: "outer", color: "000000", blur: 3, offset: 1, angle: 135, opacity: 0.06 } });
    s.addText(st.name, { x: sx + 0.1, y: st.y + 0.02, w: 2.6, h: 0.4, fontSize: 9, color: CH, bold: true, valign: "middle" });
  });

  const gates = [
    { w: 5, label: "Pilot Ready", col: RD }, { w: 10, label: "Mid-Pilot", col: OR },
    { w: 13, label: "Adjusted", col: GR }, { w: 14, label: "Decision", col: RD },
  ];
  gates.forEach(g => {
    const gx = 0.3 + (g.w - 1) * (9.4 / 14) + (9.4 / 14) / 2 - 0.1;
    s.addShape("diamond", { x: gx, y: 1.42, w: 0.22, h: 0.22, fill: { color: g.col } });
    s.addText(g.label, { x: gx - 0.3, y: 1.02, w: 0.8, h: 0.35, fontSize: 6, color: g.col, align: "center", bold: true });
  });

  // ═══ SLIDE 6: Team & Governance ═══
  s = pres.addSlide();
  hdr(s, "Team & Governance Model", "Joint team, engagement cadence, and required inputs");
  ftr(s);

  // Left: Team Structure
  s.addShape("rect", { x: 0.3, y: 0.95, w: 4.3, h: 3.5, fill: { color: WH },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.06 } });
  s.addText("Joint Team Structure", { x: 0.45, y: 1.0, w: 3, h: 0.3, fontSize: 13, bold: true, color: CH });

  // NTT DATA
  s.addShape("rect", { x: 0.5, y: 1.45, w: 1.8, h: 0.04, fill: { color: GR } });
  s.addText("NTT DATA", { x: 0.5, y: 1.52, w: 1.8, h: 0.22, fontSize: 10, bold: true, color: GR, align: "center" });
  ["Engagement Lead — David Zorrilla", "Delivery Lead — Gustavo Castellano", "Sr Consultant — Alfredo Soto", "Sr Consultant — Ana Sancha"].forEach((r, i) => {
    s.addShape("rect", { x: 0.5, y: 1.8 + i * 0.35, w: 1.8, h: 0.28, fill: { color: WH }, line: { color: "E0E0E0", width: 0.5 } });
    s.addText(r, { x: 0.55, y: 1.8 + i * 0.35, w: 1.7, h: 0.28, fontSize: 7, color: DG, valign: "middle" });
  });

  // Digital Products
  s.addShape("rect", { x: 2.6, y: 1.45, w: 1.8, h: 0.04, fill: { color: GR } });
  s.addText("Digital Products", { x: 2.6, y: 1.52, w: 1.8, h: 0.22, fontSize: 10, bold: true, color: GR, align: "center" });
  ["Exec Sponsor — Vitor Cruz", "PO/PMs — Heather L./Kumar J.", "Pilot Team — [BRS TBC]", "PMO Liaison — [TBC]"].forEach((r, i) => {
    s.addShape("rect", { x: 2.6, y: 1.8 + i * 0.35, w: 1.8, h: 0.28, fill: { color: WH }, line: { color: "E0E0E0", width: 0.5 } });
    s.addText(r, { x: 2.65, y: 1.8 + i * 0.35, w: 1.7, h: 0.28, fontSize: 7, color: DG, valign: "middle" });
  });

  s.addText("Dedication:  ○ Partial   ● Full", { x: 0.5, y: 3.35, w: 3, h: 0.2, fontSize: 7, color: MG });

  // Right top: Cadence
  s.addShape("rect", { x: 4.85, y: 0.95, w: 4.85, h: 1.7, fill: { color: WH },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.06 } });
  s.addText("Engagement Cadence", { x: 5.0, y: 1.0, w: 3, h: 0.3, fontSize: 13, bold: true, color: CH });
  [
    ["🔄  Weekly Core Team Stand-up", "30-min sync on progress, blockers, next steps"],
    ["🔧  Workshops & Sessions", "Process definition, tool config, pilot prep (Wk 1-6)"],
    ["🚀  Pilot Ceremonies", "Stand-ups, reviews, retros during BRS pilot (Wk 7-13)"],
  ].forEach(([t, d], i) => {
    s.addText([
      { text: t, options: { bold: true, fontSize: 9, color: CH, breakLine: true } },
      { text: d, options: { fontSize: 7, color: MG } }
    ], { x: 5.1, y: 1.38 + i * 0.38, w: 4.4, h: 0.38, valign: "middle" });
  });

  // Right bottom: Client Inputs
  s.addShape("rect", { x: 4.85, y: 2.85, w: 4.85, h: 1.6, fill: { color: WH },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.06 } });
  s.addText("Required Client Inputs", { x: 5.0, y: 2.9, w: 3, h: 0.3, fontSize: 13, bold: true, color: CH });
  [
    { t: "Org Data", d: "Roles & Org Chart", c: GR, bg: LGR },
    { t: "Access", d: "Jira Admin + pilot env", c: BL, bg: LBL },
    { t: "Projects", d: "Landscape & docs", c: GR, bg: LGR },
    { t: "Alignment", d: "PPMNow + PMO intro", c: BL, bg: LBL },
  ].forEach((inp, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const ix = 5.0 + col * 2.35, iy = 3.28 + row * 0.55;
    s.addShape("rect", { x: ix, y: iy, w: 2.2, h: 0.48, fill: { color: inp.bg } });
    s.addText([
      { text: inp.t, options: { bold: true, fontSize: 9, color: inp.c, breakLine: true } },
      { text: inp.d, options: { fontSize: 7, color: DG } }
    ], { x: ix + 0.1, y: iy + 0.04, w: 2, h: 0.4, valign: "middle" });
  });

  // ═══ SLIDE 7: Next Steps ═══
  s = pres.addSlide();
  hdr(s, "Immediate Next Steps", "Actions for the next 48 hours");
  ftr(s);

  // NTT DATA
  s.addShape("rect", { x: 0.3, y: 0.95, w: 4.5, h: 4.0, fill: { color: WH },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.06 } });
  s.addShape("rect", { x: 0.3, y: 0.95, w: 4.5, h: 0.06, fill: { color: GR } });
  s.addText("NTT DATA", { x: 0.5, y: 1.1, w: 2, h: 0.3, fontSize: 14, bold: true, color: GR });

  [
    ["Schedule Weekly Stand-up + Steering", "Day 1"],
    ["Schedule Week 1 Discovery sessions", "Day 1"],
    ["Share Discovery Sessions Workbook", "Day 2"],
    ["Submit Jira Admin access request", "Day 1"],
    ["Send RFI Tracker with templates", "Day 1"],
  ].forEach(([a, d], i) => {
    const y = 1.55 + i * 0.55;
    s.addShape("oval", { x: 0.5, y: y + 0.02, w: 0.24, h: 0.24, fill: { color: GR } });
    s.addText(String(i + 1), { x: 0.5, y: y + 0.02, w: 0.24, h: 0.24, fontSize: 9, color: WH, bold: true, align: "center", valign: "middle" });
    s.addText(a, { x: 0.85, y: y, w: 2.9, h: 0.28, fontSize: 9, color: CH, valign: "middle" });
    s.addShape("rect", { x: 3.9, y: y + 0.22, w: 0.65, h: 0.2, fill: { color: LGR } });
    s.addText(d, { x: 3.9, y: y + 0.22, w: 0.65, h: 0.2, fontSize: 7, color: GR, bold: true, align: "center", valign: "middle" });
  });

  // Avangrid
  s.addShape("rect", { x: 5.2, y: 0.95, w: 4.5, h: 4.0, fill: { color: WH },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.06 } });
  s.addShape("rect", { x: 5.2, y: 0.95, w: 4.5, h: 0.06, fill: { color: BL } });
  s.addText("AVANGRID", { x: 5.4, y: 1.1, w: 2, h: 0.3, fontSize: 14, bold: true, color: BL });

  [
    ["Grant system access (Jira, PPMNow, SP)", "Day 2"],
    ["Gather RFI items: Org Chart, Roles", "Wk 1"],
    ["Accept Discovery session invites", "Day 2"],
    ["Champion proposes interview slots", "Day 3"],
    ["Nominate 3 BRS pilot candidates", "Wk 1"],
  ].forEach(([a, d], i) => {
    const y = 1.55 + i * 0.55;
    s.addShape("oval", { x: 5.4, y: y + 0.02, w: 0.24, h: 0.24, fill: { color: BL } });
    s.addText(String(i + 1), { x: 5.4, y: y + 0.02, w: 0.24, h: 0.24, fontSize: 9, color: WH, bold: true, align: "center", valign: "middle" });
    s.addText(a, { x: 5.75, y: y, w: 3, h: 0.28, fontSize: 9, color: CH, valign: "middle" });
    s.addShape("rect", { x: 8.8, y: y + 0.22, w: 0.65, h: 0.2, fill: { color: LBL } });
    s.addText(d, { x: 8.8, y: y + 0.22, w: 0.65, h: 0.2, fontSize: 7, color: BL, bold: true, align: "center", valign: "middle" });
  });

  // ═══ SLIDE 8: Thank You ═══
  s = pres.addSlide();
  s.background = { color: DBL };
  s.addText("Thank you!", { x: 0, y: 1.5, w: 10, h: 1.2, fontSize: 44, bold: true, color: WH, align: "center" });
  s.addText("Digital Products  •  New Ways of Working  •  NTT DATA × Avangrid", {
    x: 0, y: 2.8, w: 10, h: 0.5, fontSize: 14, color: "88AACC", align: "center"
  });
  s.addShape("rect", { x: 0, y: 4.85, w: 10, h: 0.78, fill: { color: GR } });
  s.addText("Wave 1 — April 2026", { x: 0, y: 4.95, w: 10, h: 0.4, fontSize: 12, color: WH, align: "center" });

  const out = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Digital Products - Kickoff V2 (Modern).pptx";
  await pres.writeFile({ fileName: out });
  console.log("Saved: " + out);
}

main().catch(console.error);
