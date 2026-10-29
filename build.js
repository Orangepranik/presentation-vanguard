// Vanguard — investor presentation (Norway). Run: node build.js
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/9142bc29-2fd9-4804-9843-442ee676b21d_c706edd8-eb4c-48fb-82d2-52b724836e47/pptx/scripts/apply_theme.js");

const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "Vanguard_Investor_Presentation_Norway.pptx");

// ---------- theme ----------
const INK = "0B1220", PANEL = "111B2E", LINE = "22334F", STEEL = "2F4468";
const IVORY = "EDE8DC", MUTED = "98A4B8", BRASS = "C2A25D", OLIVE = "6B7A55";
const HEAD = "Cambria", BODY = "Calibri";
const THEME = {
  name: "Vanguard Military",
  headFontFace: HEAD,
  bodyFontFace: BODY,
  colors: { dk1: INK, lt1: IVORY, dk2: PANEL, lt2: MUTED, accent1: BRASS, accent2: STEEL, accent3: OLIVE, accent4: "8C7A4B", accent5: LINE, accent6: "B7BFCC", hlink: BRASS, folHlink: MUTED },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pres.title = "Vanguard — Technologies that Protect";
pres.author = "URS Vanguard";
pres.company = "URS Vanguard";

const W = 13.333, M = 0.7;

// ---------- helpers ----------
const dims = {};
async function loadDims() {
  for (const f of fs.readdirSync(path.join(__dirname, "assets"))) {
    const m = await sharp(A(f)).metadata();
    dims[f] = m.width / m.height;
  }
}
// place image inside a box, keeping aspect ratio; ax/ay = 0..1 alignment
function fit(s, file, x, y, bw, bh, alt, ax = 0.5, ay = 0.5, extra = {}) {
  const ar = dims[file];
  let w = bw, h = bw / ar;
  if (h > bh) { h = bh; w = bh * ar; }
  s.addImage({ path: A(file), x: x + (bw - w) * ax, y: y + (bh - h) * ay, w, h, altText: alt, ...extra });
}
async function icon(Comp, color = "#" + BRASS, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
const eyebrow = (s, text, x = M) =>
  s.addText(text.toUpperCase(), { x, y: 0.85, w: 9, h: 0.3, fontFace: BODY, fontSize: 12, bold: true, color: BRASS, charSpacing: 4, margin: 0, isTextBox: true, objectName: "Eyebrow" });
const title = (s, text, opt = {}) =>
  s.addText(text, { x: M, y: 1.15, w: 11.5, h: 0.9, fontFace: HEAD, fontSize: 36, bold: true, color: IVORY, margin: 0, valign: "top", isTextBox: true, objectName: "Title", ...opt });
const panel = (s, x, y, w, h, name) =>
  s.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: PANEL }, line: { color: LINE, width: 0.75 }, objectName: name || "Panel" });
const body = (s, text, x, y, w, h, opt = {}) =>
  s.addText(text, { x, y, w, h, fontFace: BODY, fontSize: 16, color: IVORY, margin: 0, valign: "top", isTextBox: true, ...opt });
const label = (s, text, x, y, w) =>
  s.addText(text.toUpperCase(), { x, y, w, h: 0.3, fontFace: BODY, fontSize: 12, bold: true, color: BRASS, charSpacing: 3, margin: 0, isTextBox: true });
const iconCircle = async (s, Comp, x, y, d = 0.62) => {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: INK }, line: { color: BRASS, width: 1 }, objectName: "IconRing" });
  const k = d * 0.46;
  s.addImage({ data: await icon(Comp), x: x + (d - k) / 2, y: y + (d - k) / 2, w: k, h: k, altText: "icon" });
};
function stat(s, x, y, w, h, big, small, bigSize = 34) {
  panel(s, x, y, w, h, "Stat card");
  s.addText(big, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: h * 0.52, fontFace: HEAD, fontSize: bigSize, bold: true, color: BRASS, margin: 0, valign: "middle", isTextBox: true });
  body(s, small, x + 0.2, y + h * 0.62, w - 0.4, h * 0.34, { fontSize: 12, color: MUTED });
}
function bullets(s, items, x, y, w, h, opt = {}) {
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 6 } })),
    { x, y, w, h, fontFace: BODY, fontSize: 14, color: IVORY, margin: 0, valign: "top", isTextBox: true, ...opt });
}
const th = (t) => ({ text: t, options: { bold: true, color: BRASS, fill: { color: INK }, fontFace: BODY, fontSize: 13 } });
const tc = (t, c = IVORY, o = {}) => ({ text: t, options: { color: c, fill: { color: PANEL }, fontFace: BODY, fontSize: 13, ...o } });
const tableOpts = { border: { type: "solid", color: LINE, pt: 0.75 }, margin: [0.04, 0.12, 0.04, 0.12], valign: "middle" };

// ---------- layouts ----------
pres.defineSlideMaster({ title: "COVER", background: { color: INK }, objects: [] });
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: INK },
  objects: [
    { image: { x: 9.6, y: 3.6, w: 4.4, h: 4.3, path: A("emblem_brass.png"), transparency: 94 } },
    { rect: { x: M, y: 0.62, w: W - 2 * M, h: 0.01, fill: { color: LINE } } },
    { text: { text: "VANGUARD   ·   TECHNOLOGIES THAT PROTECT", options: { x: M, y: 0.22, w: 7, h: 0.3, fontFace: BODY, fontSize: 10, color: MUTED, charSpacing: 3, margin: 0 } } },
  ],
  slideNumber: { x: W - M - 0.8, y: 0.22, w: 0.8, h: 0.3, fontFace: BODY, fontSize: 10, color: BRASS, align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: INK },
  objects: [{ image: { x: 8.2, y: -0.6, w: 6.4, h: 6.2, path: A("emblem_brass.png"), transparency: 92 } }],
});

(async () => {
  await loadDims();
  const content = (sec) => pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });

  // ======= 1. COVER =======
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "COVER", sectionTitle: "Introduction" });
  s.addImage({ path: A("emblem_brass.png"), x: -1.2, y: 2.4, w: 6.6, h: 6.4, transparency: 91, altText: "Vanguard emblem" });
  s.addImage({ path: A("logo_ivory.png"), x: M, y: 0.7, w: 3.9, h: 3.9 / dims["logo_ivory.png"], altText: "Vanguard — Ukrainian Radioelectronic Systems" });
  label(s, "Investor presentation   ·   Norway   ·   2026", M, 2.95, 8);
  s.addText("Technologies that Protect", { x: M, y: 3.35, w: 7.4, h: 1.7, fontFace: HEAD, fontSize: 50, bold: true, color: IVORY, margin: 0, valign: "top", isTextBox: true, objectName: "Title" });
  body(s, "Integrated protection ecosystem: EW, ER and next-generation counter-UAS. Combat-proven, built on proprietary SDR technologies for NATO-market expansion.", M, 5.4, 6.6, 1.0, { fontSize: 18, color: MUTED });
  body(s, "URS “Vanguard”   ·   Kyiv, Ukraine", M, 6.65, 6, 0.3, { fontSize: 12, color: MUTED, charSpacing: 2 });
  fit(s, "kozhan_detail.png", 8.1, 0.55, 4.7, 6.4, "KOZHAN 4.0 drone detector");
  s.addNotes("Opening. Introduce Vanguard and the purpose of the meeting: present the 2026 product line and discuss partnership with Norwegian investors.");

  // ======= MISSION =======
  pres.addSection({ title: "Company" });
  s = content("Company");
  eyebrow(s, "About URS Vanguard");
  title(s, "Building Europe’s DefenseTech ecosystem");
  body(s, "Kyiv-based developer of electronic warfare (EW), radioelectronic reconnaissance (ER) and counter-UAS systems on proprietary SDR technologies.", M, 2.2, 11.5, 0.7, { fontSize: 18, color: MUTED });
  const mv = [
    [fa.FaBullseye, "Our mission", "Develop real-time technologies that significantly enhance the protection and survival of military personnel and the success of their missions."],
    [fa.FaChartLine, "Vision, 3–5 years", "Become one of Europe’s key manufacturers of EW, ER and personal protection systems, unified on a single interoperable information-network platform."],
    [fa.FaGlobeEurope, "Global expansion", "Turn battle-tested front-line experience into certified, standardized solutions for NATO countries and international partners."],
  ];
  for (let i = 0; i < 3; i++) {
    const x = M + i * 4.0;
    panel(s, x, 3.15, 3.93, 3.65, "Mission card");
    await iconCircle(s, mv[i][0], x + 0.3, 3.45, 0.7);
    s.addText(mv[i][1], { x: x + 0.3, y: 4.35, w: 3.4, h: 0.45, fontFace: HEAD, fontSize: 22, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, mv[i][2], x + 0.3, 4.9, 3.35, 1.8, { fontSize: 14, color: MUTED });
  }
  s.addNotes("Company, mission and vision (from the earlier Vanguard presentation).");

  // ======= TRACTION =======
  s = content("Company");
  eyebrow(s, "Traction");
  title(s, "Proven demand and rapid growth");
  const tr = [
    ["3x", "Year-over-year growth", "Consistent tripling of demand each year."],
    ["300+", "Completed deliveries", "Production capacity booked more than a month ahead."],
    ["15", "Core team", "In-house specialists covering the full development and production cycle."],
    ["6–12", "Months of pipeline", "Long-term direct contracts with military units under negotiation."],
  ];
  const sw = 2.82, sg = 0.21;
  tr.forEach((c, i) => {
    const x = M + i * (sw + sg);
    panel(s, x, 2.3, sw, 4.4, "Stat card");
    s.addText(c[0], { x: x + 0.3, y: 2.65, w: sw - 0.6, h: 1.3, fontFace: HEAD, fontSize: 66, bold: true, color: BRASS, margin: 0, valign: "middle", isTextBox: true });
    s.addText(c[1], { x: x + 0.3, y: 4.2, w: sw - 0.6, h: 0.7, fontFace: HEAD, fontSize: 19, bold: true, color: IVORY, margin: 0, valign: "top", isTextBox: true });
    body(s, c[2], x + 0.3, 5.05, sw - 0.6, 1.5, { fontSize: 14, color: MUTED });
  });
  s.addNotes("Traction figures are taken from the earlier Vanguard presentation; confirm they are current before the meeting.");

  // ======= ADVANTAGES =======
  s = content("Company");
  eyebrow(s, "Our advantages");
  title(s, "Control, speed and proven results");
  const adv = [
    [fa.FaMicrochip, "100% proprietary software", "AI-based signal processing, SDR architecture, in-house antenna topology and circuit design."],
    [fa.FaSatelliteDish, "Widest scanning range", "Maximum frequency coverage for detecting masked and non-standard UAV control channels."],
    [fa.FaCheckDouble, "Minimized false alarms", "Proprietary digital filtering delivers accurate situational awareness."],
    [fa.FaProjectDiagram, "Unified ecosystem", "Detection (ER) and suppression (EW) integrated into a single automated network."],
    [fa.FaIndustry, "Full control of the chain", "R&D, software and production in one company: a high barrier to entry."],
    [fa.FaBolt, "Faster adaptation", "Rapid update cycle and direct feedback from front-line operators."],
  ];
  for (let i = 0; i < 6; i++) {
    const x = M + (i % 3) * 4.0, y = 2.3 + Math.floor(i / 3) * 2.4;
    panel(s, x, y, 3.93, 2.2, "Benefit card");
    await iconCircle(s, adv[i][0], x + 0.3, y + 0.25, 0.62);
    s.addText(adv[i][1], { x: x + 0.3, y: y + 1.0, w: 3.4, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, adv[i][2], x + 0.3, y + 1.45, 3.4, 0.7, { fontSize: 12.5, color: MUTED });
  }
  s.addNotes("Advantages taken from the earlier Vanguard presentation.");

  // ======= KOZHAN HERO =======
  pres.addSection({ title: "KOZHAN 4.0" });
  s = content("KOZHAN 4.0");
  eyebrow(s, "Flagship product  ·  RF reconnaissance device (drone detector)");
  s.addText("KOZHAN 4.0", { x: M, y: 1.12, w: 7.5, h: 1.05, fontFace: HEAD, fontSize: 60, bold: true, color: IVORY, margin: 0, isTextBox: true, objectName: "Title" });
  body(s, "A multi-band detector that scans the set frequency ranges and picks up the analog video signal of UAVs from 3 km. Three automatic scan modes, a log of detected signals and fast software updates over Wi-Fi. Interface languages: Ukrainian, English, German and Polish.", M, 2.35, 7.1, 1.5, { fontSize: 15 });
  s.addShape(pres.ShapeType.rect, { x: M, y: 3.78, w: 6.2, h: 0.42, fill: { color: INK }, line: { color: BRASS, width: 1 }, objectName: "NATO badge" });
  s.addText("FINAL STAGE OF NATO CODIFICATION  ·  NATO STOCK NUMBER", { x: M, y: 3.78, w: 6.2, h: 0.42, fontFace: BODY, fontSize: 11, bold: true, color: BRASS, charSpacing: 2, align: "center", valign: "middle", margin: 0, isTextBox: true });
  const steps = [[fa.FaBroadcastTower, "Scan", "Automatic monitoring of 4 radio bands in real time"], [fa.FaCrosshairs, "Detect", "Locks onto the analog UAV video signal"], [fa.FaBell, "Alert", "Instant operator notification and log entry"]];
  for (let i = 0; i < 3; i++) {
    const x = M + i * 2.5;
    panel(s, x, 4.5, 2.35, 2.2, "Step card");
    await iconCircle(s, steps[i][0], x + 0.25, 4.7, 0.56);
    s.addText(`0${i + 1}  ${steps[i][1].toUpperCase()}`, { x: x + 0.25, y: 5.38, w: 2.0, h: 0.3, fontFace: BODY, fontSize: 13, bold: true, color: BRASS, charSpacing: 2, margin: 0, isTextBox: true });
    body(s, steps[i][2], x + 0.25, 5.75, 1.95, 0.9, { fontSize: 13, color: MUTED });
  }
  fit(s, "kozhan_front.png", 8.3, 1.2, 4.35, 5.6, "KOZHAN 4.0 detector with four antennas");
  s.addNotes("KOZHAN 4.0 is the focus of this presentation. How it works: scan, detect, alert.");

  // ======= KOZHAN SPECS =======
  s = content("KOZHAN 4.0");
  eyebrow(s, "KOZHAN 4.0");
  title(s, "Technical characteristics");
  const stats = [["4", "radio bands\n1G · 3G · 5G · 6G"], ["3 km+", "detection range,\nfrom 3 km"], ["12 h", "autonomous\noperation"], ["IP54", "protected\ncase"]];
  stats.forEach((st, i) => {
    const x = M + (i % 2) * 2.95, y = 2.3 + Math.floor(i / 2) * 2.1;
    panel(s, x, y, 2.75, 1.9, "Stat card");
    s.addText(st[0], { x: x + 0.25, y: y + 0.15, w: 2.3, h: 0.85, fontFace: HEAD, fontSize: 44, bold: true, color: BRASS, margin: 0, isTextBox: true });
    body(s, st[1], x + 0.25, y + 1.1, 2.3, 0.7, { fontSize: 14, color: MUTED });
  });
  body(s, "Bands, MHz:  900–1800  ·  2700–4500  ·  4850–6150  ·  6000–8200", M, 6.55, 5.8, 0.35, { fontSize: 12, color: MUTED });
  const row = (a, b) => [tc(a, MUTED, { fontSize: 13 }), tc(b, IVORY, { fontSize: 13 })];
  s.addTable([
    [th("Parameter"), th("Value")],
    row("Detection range", "from 3 km"),
    row("Scan time", "up to 7 seconds"),
    row("Scan modes", "3 automatic, signal log"),
    row("Display", "IPS 5″ (800 × 480); external via AV"),
    row("Power", "Built-in battery 75.6 Wh, USB-C charging"),
    row("Weight", "780 g"),
    row("Dimensions", "280×110×45 mm; 155×110×45 mm"),
    row("Ingress protection", "IP54"),
    row("Software", "Wi-Fi updates; UA, EN, DE, PL"),
    row("EW auto-activation", "MT-30 connector"),
  ], { x: 6.85, y: 2.3, w: 5.78, colW: [2.2, 3.58], rowH: 0.41, ...tableOpts, objectName: "Specs table" });
  s.addNotes("Specifications are taken from the product page on the company website (updated 23.10.2026). The earlier PDF catalog gave 2 km detection range and an upper band edge of 8260 MHz; the website values are used here.");

  // ======= KOZHAN KIT =======
  s = content("KOZHAN 4.0");
  eyebrow(s, "KOZHAN 4.0");
  title(s, "Complete kit, ready for the field");
  fit(s, "kozhan_detail.png", M, 2.2, 2.6, 4.6, "KOZHAN 4.0 with antennas");
  panel(s, 3.55, 2.3, 2.95, 4.5, "Kit panel");
  label(s, "Delivery set", 3.8, 2.5, 2.5);
  bullets(s, ["4-band receiver 1G+3G+5G+6G in a solid case", "Antennas 1G, 3G, 5G, 6G", "USB-C charging cable", "Device mount", "Velcro strap set"], 3.8, 2.95, 2.5, 3.7, { fontSize: 15, paraSpaceAfter: 12 });
  panel(s, 6.7, 2.3, 2.95, 4.5, "Interfaces panel");
  label(s, "Interfaces", 6.95, 2.5, 2.5);
  const ifs = [[fa.FaVideo, "AV", "External display"], [fa.FaBolt, "USB-C", "Charging"], [fa.FaPlug, "MT-30", "Auto-start of EW system"]];
  for (let i = 0; i < 3; i++) {
    const y = 3.05 + i * 1.2;
    await iconCircle(s, ifs[i][0], 6.95, y, 0.55);
    s.addText(ifs[i][1], { x: 7.7, y: y - 0.02, w: 1.8, h: 0.32, fontFace: HEAD, fontSize: 17, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, ifs[i][2], 7.7, y + 0.32, 1.85, 0.6, { fontSize: 12, color: MUTED });
  }
  panel(s, 9.85, 2.3, 2.78, 4.5, "Accessory panel");
  label(s, "Accessory", 10.1, 2.5, 2.4);
  fit(s, "antenna.png", 10.0, 2.95, 2.5, 1.0, "Active remote antenna");
  s.addText("Active remote antenna", { x: 10.1, y: 4.1, w: 2.4, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: IVORY, margin: 0, isTextBox: true });
  body(s, "Magnetic mount, four antenna channels with built-in amplifiers, 3 m RG316 cables. For KOZHAN 4.0 and 3.0.", 10.1, 4.55, 2.4, 2.1, { fontSize: 12.5, color: MUTED });
  s.addNotes("Delivery set, connectors and the optional active remote antenna (powered via MT-30; amplifier gain to be confirmed in the catalog).");

  // ======= 4. PORTFOLIO =======
  pres.addSection({ title: "Portfolio" });
  s = content("Portfolio");
  eyebrow(s, "Product portfolio");
  title(s, "One ecosystem: detect, observe, suppress");
  body(s, "Delfin and Kraken are in serial production and used in real operations; EW platforms can be started automatically by KOZHAN.", M, 1.78, 11.5, 0.35, { fontSize: 13, color: MUTED });
  label(s, "Detection and surveillance", M, 2.15, 6);
  const top = [
    ["kozhan_front.png", "KOZHAN 4.0", "Multi-band drone detector, 4 bands"],
    ["antenna.png", "Remote antenna", "Magnetic mount, built-in amplifiers"],
    ["thermal.png", "NICH", "Pan-controlled thermal camera"],
  ];
  const tw = 3.85, tg = 0.19;
  top.forEach((c, i) => {
    const x = M + i * (tw + tg);
    panel(s, x, 2.5, tw, 1.95, "Product card");
    fit(s, c[0], x + 0.2, 2.62, 1.7, 1.6, c[1]);
    s.addText(c[1], { x: x + 2.0, y: 2.85, w: tw - 2.1, h: 0.4, fontFace: HEAD, fontSize: 15, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, c[2], x + 2.0, 3.3, tw - 2.15, 0.9, { fontSize: 12, color: MUTED });
    // small image left already; caption below image area
  });
  label(s, "Electronic warfare, 300–6000 MHz", M, 4.62, 8);
  const bot = [
    ["delfin_box.png", "Delfin", "Disguised as a civilian car roof box"],
    ["hydra_kit.png", "Hydra 2U", "19″ rack for heavy vehicles, 8–20 bands"],
    ["dokatka.png", "Dokatka", "Hidden in truck-wheel form factor"],
    ["kraken_full.png", "Kraken", "Directional sector suppression"],
  ];
  const bw = 2.86, bg = 0.15;
  bot.forEach((c, i) => {
    const x = M + i * (bw + bg);
    panel(s, x, 4.95, bw, 1.95, "Product card");
    fit(s, c[0], x + 0.15, 5.02, bw - 0.3, 0.95, c[1]);
    s.addText(c[1], { x: x + 0.2, y: 6.02, w: bw - 0.4, h: 0.3, fontFace: HEAD, fontSize: 16, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, c[2], x + 0.2, 6.33, bw - 0.4, 0.5, { fontSize: 12, color: MUTED });
  });
  s.addNotes("Seven products in three groups. All EW platforms can be switched on automatically by the KOZHAN detector.");

  // ======= SCALE =======
  pres.addSection({ title: "Readiness" });
  s = content("Readiness");
  eyebrow(s, "Scale and readiness");
  title(s, "Field-validated, ready to scale");
  const cap = [["80%", "vertical integration"], ["500", "detectors per month, up to"], ["50", "EW systems per month, up to"], ["Days", "for firmware and hardware updates"]];
  cap.forEach((c, i) => stat(s, M + i * 3.0, 2.2, 2.85, 1.55, c[0], c[1], 38));
  const rd = [
    [fa.FaCertificate, "TRL 9 readiness", "Products passed full verification in real combat conditions, with hundreds of positive feedback reports from military units."],
    [fa.FaFlagUsa, "International forums", "Participation in, and official invitations to, defense events in the USA, Lithuania, Norway, Italy and others."],
    [fa.FaHandshake, "Partnerships and certification", "Preliminary cooperation agreements reached; codification, patenting and international certification under way."],
  ];
  for (let i = 0; i < 3; i++) {
    const x = M + i * 4.0;
    panel(s, x, 4.05, 3.93, 2.75, "Readiness card");
    await iconCircle(s, rd[i][0], x + 0.3, 4.3, 0.62);
    s.addText(rd[i][1], { x: x + 1.1, y: 4.3, w: 2.7, h: 0.62, fontFace: HEAD, fontSize: 17, bold: true, color: IVORY, margin: 0, valign: "middle", isTextBox: true });
    body(s, rd[i][2], x + 0.3, 5.15, 3.35, 1.6, { fontSize: 13, color: MUTED });
  }
  s.addNotes("Capacity, TRL and international engagement figures are taken from the earlier Vanguard presentation; confirm they are current.");

  // ======= ROADMAP =======
  s = content("Readiness");
  eyebrow(s, "Roadmap");
  title(s, "Our path to the NATO market");
  const rm = [
    ["Export and certification", "Complete codification, obtain export permits and international certificates."],
    ["European hub", "Establish an official presence and legal entity in Poland or Italy."],
    ["Partner network", "Build distribution and joint ventures with Western defense contractors."],
    ["Serial supply", "Scale production to meet direct defense contracts with partners."],
  ];
  s.addShape(pres.ShapeType.line, { x: M + 0.35, y: 2.75, w: 3 * 3.0 + 0.0, h: 0, line: { color: STEEL, width: 1.5 } });
  rm.forEach((r, i) => {
    const x = M + i * 3.0;
    s.addShape(pres.ShapeType.ellipse, { x, y: 2.4, w: 0.7, h: 0.7, fill: { color: INK }, line: { color: BRASS, width: 1.5 }, objectName: "Step marker" });
    s.addText(String(i + 1), { x, y: 2.4, w: 0.7, h: 0.7, fontFace: HEAD, fontSize: 22, bold: true, color: BRASS, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(r[0], { x, y: 3.35, w: 2.75, h: 0.75, fontFace: HEAD, fontSize: 19, bold: true, color: IVORY, margin: 0, valign: "top", isTextBox: true });
    body(s, r[1], x, 4.15, 2.65, 1.1, { fontSize: 13, color: MUTED });
  });
  panel(s, M, 5.45, W - 2 * M, 1.4, "Market panel");
  label(s, "Why now", M + 0.3, 5.58, 3);
  body(s, "Demand for counter-drone systems is rising sharply as European armed forces re-arm. Western buyers want solutions validated in real conditions against a technologically advanced adversary, with better price and speed than bulky alternatives.", M + 0.3, 5.93, 11.3, 0.85, { fontSize: 14 });
  s.addNotes("Roadmap and market drivers are taken from the earlier Vanguard presentation.");

  // ======= 11. CONTACT =======
  pres.addSection({ title: "Contact" });
  s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Contact" });
  s.addImage({ path: A("logo_ivory.png"), x: M, y: 0.7, w: 3.2, h: 3.2 / dims["logo_ivory.png"], altText: "Vanguard logo" });
  s.addText("Let’s build the future of defense technology together", { x: M, y: 2.5, w: 7.2, h: 1.6, fontFace: HEAD, fontSize: 40, bold: true, color: IVORY, margin: 0, valign: "top", isTextBox: true, objectName: "Title" });
  const ct = [[fa.FaBuilding, "URS “Vanguard” LLC, Kyiv, Ukraine"], [fa.FaPhone, "+38 (093) 968 59 73"], [fa.FaEnvelope, "vanguardltd25@gmail.com"], [fa.FaGlobe, "ursvanguard.com"]];
  for (let i = 0; i < 4; i++) {
    const y = 4.3 + i * 0.66;
    await iconCircle(s, ct[i][0], M, y, 0.46);
    s.addText(ct[i][1], { x: M + 0.7, y, w: 6, h: 0.46, fontFace: BODY, fontSize: 18, color: IVORY, margin: 0, valign: "middle", isTextBox: true, hyperlink: i === 3 ? { url: "https://ursvanguard.com" } : undefined });
  }
  s.addShape(pres.ShapeType.rect, { x: 8.55, y: 2.1, w: 4.1, h: 4.6, fill: { color: IVORY }, line: { color: BRASS, width: 1.5 }, objectName: "QR card" });
  s.addImage({ path: A("qr.png"), x: 8.85, y: 2.4, w: 3.5, h: 3.5, altText: "QR code linking to ursvanguard.com", hyperlink: { url: "https://ursvanguard.com" } });
  s.addText("SCAN TO VISIT US", { x: 8.55, y: 6.0, w: 4.1, h: 0.3, fontFace: BODY, fontSize: 12, bold: true, color: INK, align: "center", charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("ursvanguard.com", { x: 8.55, y: 6.3, w: 4.1, h: 0.3, fontFace: BODY, fontSize: 12, color: STEEL, align: "center", margin: 0, isTextBox: true });
  s.addNotes("Closing slide with contacts and the QR code to the company website.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
