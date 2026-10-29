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
  body(s, "Ukrainian radioelectronic systems: drone detection, thermal surveillance and electronic warfare, developed in Kyiv.", M, 5.4, 6.6, 1.0, { fontSize: 18, color: MUTED });
  body(s, "URS “Vanguard”   ·   Kyiv, Ukraine", M, 6.65, 6, 0.3, { fontSize: 12, color: MUTED, charSpacing: 2 });
  fit(s, "kozhan_detail.png", 8.1, 0.55, 4.7, 6.4, "KOZHAN 4.0 drone detector");
  s.addNotes("Opening. Introduce Vanguard and the purpose of the meeting: present the 2026 product line and discuss partnership with Norwegian investors.");

  // ======= 2. ABOUT =======
  pres.addSection({ title: "Company" });
  s = content("Company");
  eyebrow(s, "About Vanguard");
  title(s, "Ukrainian engineering, built for the front line");
  body(s, "URS “Vanguard” is a Kyiv-based developer of radioelectronic systems. Our motto, Technologies that Protect, is a design brief: give operators early warning, clear situational awareness and effective countermeasures where conditions are hardest.", M, 2.3, 5.5, 2.4, { fontSize: 18 });
  body(s, "Seven products in the 2026 catalog: a drone detector with an active antenna, a thermal camera and four mobile electronic warfare systems.", M, 4.85, 5.5, 1.2, { fontSize: 16, color: MUTED });
  const pill = [
    [fa.FaSatelliteDish, "Radioelectronic reconnaissance", "KOZHAN 4.0 multi-band drone detector."],
    [fa.FaEye, "Video surveillance", "NICH thermal camera for observation in darkness."],
    [fa.FaShieldAlt, "Electronic warfare", "Delfin, Hydra 2U, Dokatka and Kraken, 300–6000 MHz."],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 2.3 + i * 1.62;
    panel(s, 6.9, y, 5.73, 1.42, "Card");
    await iconCircle(s, pill[i][0], 7.2, y + 0.4, 0.62);
    s.addText(pill[i][1], { x: 8.1, y: y + 0.22, w: 4.4, h: 0.4, fontFace: HEAD, fontSize: 20, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, pill[i][2], 8.1, y + 0.72, 4.4, 0.6, { fontSize: 14, color: MUTED });
  }
  s.addNotes("Company overview and the three capability areas. Detailed product slides follow.");

  // ======= ADVANTAGES =======
  s = content("Company");
  eyebrow(s, "Our advantages");
  title(s, "Why partner with Vanguard");
  const adv = [
    [fa.FaShieldAlt, "Born in the field", "Developed in Ukraine, where drone threats are a daily reality."],
    [fa.FaLayerGroup, "Full cycle from one vendor", "Detection, thermal surveillance and electronic warfare."],
    [fa.FaFeatherAlt, "Compact and rugged", "KOZHAN 4.0: 780 g, IP54 case, up to 12 hours on battery."],
    [fa.FaPlug, "Built to integrate", "Auto-start of EW systems, external display, remote antenna."],
    [fa.FaWifi, "Always up to date", "Software updates over Wi-Fi, three automatic scan modes."],
    [fa.FaLanguage, "Ready for Europe", "Interface in Ukrainian, English, German and Polish."],
  ];
  for (let i = 0; i < 6; i++) {
    const x = M + (i % 3) * 4.0, y = 2.3 + Math.floor(i / 3) * 2.4;
    panel(s, x, y, 3.93, 2.2, "Benefit card");
    await iconCircle(s, adv[i][0], x + 0.3, y + 0.3, 0.62);
    s.addText(adv[i][1], { x: x + 0.3, y: y + 1.1, w: 3.4, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, adv[i][2], x + 0.3, y + 1.55, 3.4, 0.6, { fontSize: 13, color: MUTED });
  }
  s.addNotes("Our advantages. The emphasis of this presentation is the KOZHAN 4.0 detector, covered in the next slides.");
  // ======= 3. CHALLENGE =======
  s = content("Company");
  eyebrow(s, "The challenge");
  title(s, "Detect first, then neutralise");
  const pts = [
    ["01", "Low-cost drones are everywhere", "Small UAVs threaten people, vehicles and positions, often with little warning."],
    ["02", "Their signal gives them away", "Analog video links leave a radio signature that can be scanned for in real time."],
    ["03", "Detection triggers the response", "An instant alert, or automatic activation of an EW system, closes the gap."],
  ];
  pts.forEach((p, i) => {
    const y = 2.35 + i * 1.5;
    s.addText(p[0], { x: M, y, w: 0.9, h: 0.7, fontFace: HEAD, fontSize: 36, bold: true, color: BRASS, margin: 0, isTextBox: true });
    s.addText(p[1], { x: 1.75, y: y + 0.02, w: 5.3, h: 0.4, fontFace: HEAD, fontSize: 20, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, p[2], 1.75, y + 0.5, 5.3, 0.8, { fontSize: 14, color: MUTED });
  });
  const cx = 10.0, cy = 4.5;
  [2.35, 1.6, 0.85].forEach((r, i) =>
    s.addShape(pres.ShapeType.ellipse, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color: PANEL, transparency: i === 0 ? 40 : 100 }, line: { color: STEEL, width: 1 }, objectName: "Radar ring" }));
  s.addShape(pres.ShapeType.line, { x: cx - 2.35, y: cy, w: 4.7, h: 0, line: { color: LINE, width: 0.75 } });
  s.addShape(pres.ShapeType.line, { x: cx, y: cy - 2.35, w: 0, h: 4.7, line: { color: LINE, width: 0.75 } });
  s.addShape(pres.ShapeType.ellipse, { x: cx - 0.09, y: cy - 0.09, w: 0.18, h: 0.18, fill: { color: BRASS }, line: { color: BRASS, width: 0.5 }, objectName: "Operator" });
  [[1.3, -1.4], [-1.55, -0.6], [0.7, 1.55]].forEach(([dx, dy]) =>
    s.addShape(pres.ShapeType.diamond, { x: cx + dx - 0.12, y: cy + dy - 0.12, w: 0.24, h: 0.24, fill: { color: "B5483C" }, line: { color: IVORY, width: 0.75 }, objectName: "Contact" }));
  s.addText("2 km+", { x: cx + 1.62, y: cy + 0.05, w: 0.72, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: BRASS, margin: 0, isTextBox: true });
  s.addText("KOZHAN 4.0 detection range (illustrative)", { x: cx - 2.5, y: cy + 2.5, w: 5.0, h: 0.3, fontFace: BODY, fontSize: 11, color: MUTED, align: "center", margin: 0, isTextBox: true });
  s.addNotes("The problem we solve. The radar graphic is illustrative; KOZHAN 4.0 detects from 2 km.");

  // ======= KOZHAN HERO =======
  pres.addSection({ title: "KOZHAN 4.0" });
  s = content("KOZHAN 4.0");
  eyebrow(s, "Flagship product  ·  Drone detector");
  s.addText("KOZHAN 4.0", { x: M, y: 1.12, w: 7.5, h: 1.05, fontFace: HEAD, fontSize: 60, bold: true, color: IVORY, margin: 0, isTextBox: true, objectName: "Title" });
  body(s, "A multi-band detector that scans the set frequency ranges and picks up the analog video signal of UAVs. Three automatic scan modes, a log of detected signals and fast software updates over Wi-Fi. Interface languages: Ukrainian, English, German and Polish.", M, 2.35, 7.1, 1.8, { fontSize: 16 });
  const steps = [[fa.FaBroadcastTower, "Scan", "Automatic monitoring of 4 radio bands in real time"], [fa.FaCrosshairs, "Detect", "Locks onto the analog UAV video signal"], [fa.FaBell, "Alert", "Instant operator notification and log entry"]];
  for (let i = 0; i < 3; i++) {
    const x = M + i * 2.5;
    panel(s, x, 4.4, 2.35, 2.2, "Step card");
    await iconCircle(s, steps[i][0], x + 0.25, 4.6, 0.56);
    s.addText(`0${i + 1}  ${steps[i][1].toUpperCase()}`, { x: x + 0.25, y: 5.28, w: 2.0, h: 0.3, fontFace: BODY, fontSize: 13, bold: true, color: BRASS, charSpacing: 2, margin: 0, isTextBox: true });
    body(s, steps[i][2], x + 0.25, 5.65, 1.95, 1.0, { fontSize: 13, color: MUTED });
  }
  fit(s, "kozhan_front.png", 8.3, 1.2, 4.35, 5.6, "KOZHAN 4.0 detector with four antennas");
  s.addNotes("KOZHAN 4.0 is the focus of this presentation. How it works: scan, detect, alert.");

  // ======= KOZHAN SPECS =======
  s = content("KOZHAN 4.0");
  eyebrow(s, "KOZHAN 4.0");
  title(s, "Technical characteristics");
  const stats = [["4", "radio bands\n1G · 3G · 5G · 6G"], ["2 km+", "detection\nrange"], ["12 h", "autonomous\noperation"], ["7 s", "full-scan\ntime"]];
  stats.forEach((st, i) => {
    const x = M + (i % 2) * 2.95, y = 2.3 + Math.floor(i / 2) * 2.1;
    panel(s, x, y, 2.75, 1.9, "Stat card");
    s.addText(st[0], { x: x + 0.25, y: y + 0.15, w: 2.3, h: 0.85, fontFace: HEAD, fontSize: 44, bold: true, color: BRASS, margin: 0, isTextBox: true });
    body(s, st[1], x + 0.25, y + 1.1, 2.3, 0.7, { fontSize: 14, color: MUTED });
  });
  body(s, "Bands, MHz:  900–1800  ·  2700–4500  ·  4850–6150  ·  6000–8260", M, 6.55, 5.8, 0.35, { fontSize: 12, color: MUTED });
  const row = (a, b) => [tc(a, MUTED, { fontSize: 14 }), tc(b, IVORY, { fontSize: 14 })];
  s.addTable([
    [th("Parameter"), th("Value")],
    row("Weight", "780 g"),
    row("Dimensions", "280×110×45 mm; 155×110×45 mm"),
    row("Display", "IPS 5″ (800 × 480)"),
    row("Built-in battery", "75.6 Wh"),
    row("Ingress protection", "IP-54"),
    row("External display", "AV output"),
    row("EW auto-activation", "MT-30 connector"),
  ], { x: 6.85, y: 2.3, w: 5.78, colW: [2.4, 3.38], rowH: 0.52, ...tableOpts, objectName: "Specs table" });
  s.addNotes("Specifications from the 2026 catalog. The catalog gives the upper band edge as 8260 MHz on the overview page and 8200 MHz in the specification table; confirm before sending.");

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

  // ======= SYSTEM =======
  s = content("KOZHAN 4.0");
  eyebrow(s, "KOZHAN 4.0 in a system");
  title(s, "From detection to response");
  const flow = [
    [fa.FaBroadcastTower, "DETECT", "KOZHAN 4.0 scans the 1G, 3G, 5G and 6G bands for the UAV video signal."],
    [fa.FaBell, "ALERT", "The operator is notified instantly and the event is logged."],
    [fa.FaShieldAlt, "RESPOND", "Optional automatic start of an EW system via the MT-30 connector."],
    [fa.FaEye, "OBSERVE", "NICH thermal camera extends awareness in darkness."],
  ];
  const cw = 2.78, gap = 0.27;
  for (let i = 0; i < 4; i++) {
    const x = M + i * (cw + gap);
    panel(s, x, 2.3, cw, 2.85, "Flow card");
    await iconCircle(s, flow[i][0], x + 0.3, 2.55, 0.62);
    s.addText(`0${i + 1}`, { x: x + cw - 0.9, y: 2.55, w: 0.6, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true, color: STEEL, align: "right", margin: 0, isTextBox: true });
    s.addText(flow[i][1], { x: x + 0.3, y: 3.4, w: cw - 0.6, h: 0.4, fontFace: HEAD, fontSize: 21, bold: true, color: IVORY, margin: 0, isTextBox: true });
    body(s, flow[i][2], x + 0.3, 3.9, cw - 0.6, 1.2, { fontSize: 13, color: MUTED });
    if (i < 3) s.addImage({ data: await icon(fa.FaChevronRight, "#" + BRASS), x: x + cw + 0.04, y: 3.65, w: 0.19, h: 0.19, altText: "next" });
  }
  panel(s, M, 5.4, W - 2 * M, 1.45, "Platforms panel");
  label(s, "EW platforms compatible with KOZHAN", M + 0.3, 5.52, 6);
  [["delfin_box.png", "Delfin"], ["hydra_kit.png", "Hydra 2U"], ["dokatka.png", "Dokatka"], ["kraken_full.png", "Kraken"]].forEach((p, i) => {
    const x = M + 0.3 + i * 2.95;
    fit(s, p[0], x, 5.88, 1.45, 0.85, p[1]);
    s.addText(p[1], { x: x + 1.55, y: 6.1, w: 1.3, h: 0.4, fontFace: HEAD, fontSize: 15, bold: true, color: IVORY, margin: 0, valign: "middle", isTextBox: true });
  });
  s.addNotes("Each layer works on its own and is stronger together. Delfin, Hydra 2U, Dokatka and Kraken all support connection to the KOZHAN detector.");
  // ======= 4. PORTFOLIO =======
  pres.addSection({ title: "Products" });
  s = content("Products");
  eyebrow(s, "Product portfolio");
  title(s, "One catalog: detect, observe, suppress");
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

  // ======= EW FAMILY =======
  pres.addSection({ title: "Portfolio" });
  s = content("Portfolio");
  eyebrow(s, "Complementary systems");
  title(s, "Electronic warfare family and thermal vision");
  s.addTable([
    [th("Platform"), th("Built for"), th("Bands / range"), th("Power"), th("Weight")],
    [tc("Delfin", IVORY, { bold: true }), tc("Concealed on a civilian car"), tc("up to 13 bands, 300–6000 MHz"), tc("50–100 W per module"), tc("60–80 kg")],
    [tc("Hydra 2U", IVORY, { bold: true }), tc("Heavy vehicles, APCs, command vehicles"), tc("8–20 bands, 300–6000 MHz"), tc("50 W per band"), tc("60–80 kg")],
    [tc("Dokatka", IVORY, { bold: true }), tc("Concealed on trucks and mobile positions"), tc("up to 12 bands, 300–6000 MHz"), tc("50–100 W per band"), tc("30–40 kg")],
    [tc("Kraken", IVORY, { bold: true }), tc("Directional suppression in a sector"), tc("300–4100 MHz"), tc("25–50 W per 100 MHz"), tc("25–30 kg")],
  ], { x: M, y: 2.2, w: W - 2 * M, colW: [1.7, 3.5, 3.0, 2.4, 1.33], rowH: 0.5, ...tableOpts, objectName: "EW comparison" });
  panel(s, M, 4.95, W - 2 * M, 1.9, "NICH panel");
  fit(s, "thermal.png", M + 0.25, 5.1, 3.1, 1.6, "NICH thermal camera with display unit");
  label(s, "NICH  ·  Thermal camera", 4.7, 5.12, 5);
  body(s, "Pan-controlled thermal camera with an uncooled VOx sensor (256 × 192 px, 50 Hz, 50° × 37.2° field of view) and a 5″ display. Cable 5–50 m, supply 9–24 V.", 4.7, 5.5, 4.4, 1.3, { fontSize: 13, color: MUTED });
  s.addImage({ path: A("thermal_view.png"), x: 9.35, y: 5.1, w: 1.6 * dims["thermal_view.png"] * 1.0, h: 1.6, altText: "Thermal image: a person detected in darkness" });
  s.addNotes("The EW platforms can be switched on automatically by the KOZHAN detector; NICH provides thermal observation at night. Partnership topics for discussion: distribution, localisation and joint development with Norwegian partners.");

  // ======= 11. CONTACT =======
  pres.addSection({ title: "Contact" });
  s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Contact" });
  s.addImage({ path: A("logo_ivory.png"), x: M, y: 0.7, w: 3.2, h: 3.2 / dims["logo_ivory.png"], altText: "Vanguard logo" });
  s.addText("Let’s build protection together", { x: M, y: 2.5, w: 7.2, h: 1.6, fontFace: HEAD, fontSize: 44, bold: true, color: IVORY, margin: 0, valign: "top", isTextBox: true, objectName: "Title" });
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
