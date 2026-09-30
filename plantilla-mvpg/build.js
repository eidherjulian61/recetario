const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const path = require("path");

const DIR = path.join(__dirname, "assets");
const OUT = process.argv[2] || path.join(__dirname, "Plantilla_MVPG.pptx");

// ---------- Paleta (solo azul oscuro, mostaza, crema y sus variaciones) ----------
const C = {
  NAVY_900: "0F1A30", NAVY: "1B2A4A", NAVY_600: "2D4270", NAVY_70: "5C667A", NAVY_40: "9EA2AA",
  MUST_TEXT: "8A6A1F", MUST: "C9A23F", MUST_LIGHT: "E0C06E", MUST_PALE: "F0E4C2",
  CREAM_LIGHT: "FBFAF6", CREAM: "F5F2EA", CREAM_DARK: "EAE4D6", CREAM_DEEP: "DDD4C0",
};
const HEAD = "Arial";
const BODY = "Calibri";

// ---------- Retícula ----------
const W = 13.333, H = 7.5, M = 0.6, CW = W - 2 * M;
const TOP = 1.9; // inicio del área de contenido
const BOTTOM = 6.5; // fin del área de contenido
const LOGO_W = 0.95, LOGO_H = LOGO_W / 2.134;

const LOGO_LIGHT = path.join(DIR, "logo_light.png");
const LOGO_DARK = path.join(DIR, "logo_dark.png");

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Imagen de relleno: el docente la reemplaza con clic derecho › Cambiar imagen
async function placeholderImg(w, h) {
  const pw = Math.round(w * 150), ph = Math.round(h * 150);
  const s = Math.min(pw, ph) * 0.22;
  const cx = pw / 2, cy = ph / 2 - s * 0.25;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${pw}" height="${ph}">
    <rect width="${pw}" height="${ph}" fill="#${C.CREAM_DARK}"/>
    <rect x="${cx - s}" y="${cy - s * 0.7}" width="${2 * s}" height="${1.4 * s}" rx="${s * 0.12}" fill="none" stroke="#${C.NAVY_40}" stroke-width="${s * 0.08}"/>
    <circle cx="${cx - s * 0.45}" cy="${cy - s * 0.25}" r="${s * 0.17}" fill="#${C.MUST}"/>
    <polygon points="${cx - s * 0.85},${cy + s * 0.55} ${cx - s * 0.2},${cy - s * 0.05} ${cx + s * 0.15},${cy + s * 0.3} ${cx + s * 0.45},${cy - s * 0.2} ${cx + s * 0.85},${cy + s * 0.55}" fill="#${C.NAVY_40}"/>
    <text x="${cx}" y="${cy + s * 1.25}" font-family="Arial, sans-serif" font-size="${Math.max(14, s * 0.28)}" fill="#${C.NAVY_70}" text-anchor="middle">Clic derecho › Cambiar imagen</text>
  </svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "MVPG – Mérito Valor Público y Gestión";
  pres.company = "MVPG";
  pres.title = "Plantilla corporativa MVPG";
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  const I = {
    shield: await icon(fa.FaShieldAlt, C.MUST), award: await icon(fa.FaAward, C.MUST),
    chart: await icon(fa.FaChartLine, C.MUST), target: await icon(fa.FaBullseye, C.MUST),
    globe: await icon(fa.FaGlobe, C.NAVY), mail: await icon(fa.FaEnvelope, C.NAVY),
    phone: await icon(fa.FaPhoneAlt, C.NAVY), pin: await icon(fa.FaMapMarkerAlt, C.NAVY),
    users: await icon(fa.FaUsers, C.NAVY),
    checkM: await icon(fa.FaCheck, C.MUST), checkN: await icon(fa.FaCheck, C.NAVY),
    quote: await icon(fa.FaQuoteLeft, C.MUST), bulb: await icon(fa.FaLightbulb, C.MUST),
    bulbN: await icon(fa.FaLightbulb, C.NAVY),
  };

  // ---------- Patrones (masters) ----------
  const footer = (dark) => [
    { text: { text: "Título de la charla  |  Nombre del docente", options: {
      x: M + 0.55, y: 6.85, w: 7, h: 0.35, margin: 0, fontFace: BODY, fontSize: 10,
      color: dark ? C.NAVY_40 : C.NAVY_70, valign: "middle" } } },
    { image: { path: dark ? LOGO_DARK : LOGO_LIGHT, x: W - M - LOGO_W, y: 6.8, w: LOGO_W, h: LOGO_H } },
  ];
  const titlePh = (dark) => ({ placeholder: { options: {
    name: "title", type: "title", x: M, y: 0.78, w: CW, h: 0.85, margin: 0,
    fontFace: HEAD, fontSize: 32, bold: true, color: dark ? C.CREAM : C.NAVY, valign: "top", align: "left" },
    text: "Título de la diapositiva" } });
  const slideNum = (dark) => ({ x: M, y: 6.85, w: 0.5, h: 0.35, fontFace: HEAD, fontSize: 10,
    bold: true, color: dark ? C.MUST : C.MUST_TEXT, margin: 0, valign: "middle" });

  pres.defineSlideMaster({ title: "MVPG_PORTADA", background: { color: C.NAVY }, objects: [] });
  pres.defineSlideMaster({ title: "MVPG_CONTENIDO", background: { color: C.CREAM },
    objects: [...footer(false), titlePh(false)], slideNumber: slideNum(false) });
  pres.defineSlideMaster({ title: "MVPG_TEXTO", background: { color: C.CREAM },
    objects: [...footer(false), titlePh(false),
      { placeholder: { options: { name: "body", type: "body", x: M, y: TOP, w: CW, h: BOTTOM - TOP,
        fontFace: BODY, fontSize: 16, color: C.NAVY, valign: "top", margin: 0 }, text: "" } }],
    slideNumber: slideNum(false) });
  pres.defineSlideMaster({ title: "MVPG_OSCURO", background: { color: C.NAVY },
    objects: [...footer(true), titlePh(true)], slideNumber: slideNum(true) });
  pres.defineSlideMaster({ title: "MVPG_SECCION", background: { color: C.NAVY },
    objects: footer(true), slideNumber: slideNum(true) });

  // ---------- Ayudantes ----------
  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: BODY,
    color: C.NAVY, valign: "top" }, o));
  function header(s, kicker, title, dark = false) {
    T(s, kicker.toUpperCase(), { x: M, y: 0.45, w: 9, h: 0.3, fontFace: HEAD, fontSize: 12, bold: true,
      color: dark ? C.MUST : C.MUST_TEXT, charSpacing: 3 });
    s.addText(title, { placeholder: "title", align: "left" });
  }
  const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE,
    Object.assign({ x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: 0.75 } : { type: "none" },
      rectRadius: o.round ? 0.08 : undefined }, o.extra || {}));
  const circle = (s, x, y, d, fill, lineColor, lw) => s.addShape(pres.shapes.OVAL,
    { x, y, w: d, h: d, fill: { color: fill }, line: lineColor ? { color: lineColor, width: lw || 1 } : { type: "none" } });
  const line = (s, x, y, w, h, color, o = {}) => s.addShape(pres.shapes.LINE, Object.assign(
    { x, y, w, h, line: Object.assign({ color, width: o.width || 1.5 }, o.arrow ? { endArrowType: "triangle" } : {},
      o.dash ? { dashType: o.dash } : {}) }, o.flipV ? { flipV: true } : {}));
  async function img(s, x, y, w, h) { s.addImage({ data: await placeholderImg(w, h), x, y, w, h }); }
  function iconCircle(s, x, y, d, fill, data) {
    circle(s, x, y, d, fill);
    const p = d * 0.25;
    s.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  const notes = (s, t) => s.addNotes(t);

  // ============ 1. PORTADA ============
  {
    const s = pres.addSlide({ masterName: "MVPG_PORTADA" });
    s.addImage({ path: LOGO_DARK, x: 0.8, y: 0.75, w: 3.8, h: 3.8 / 2.134 });
    T(s, "MÉRITO VALOR PÚBLICO Y GESTIÓN", { x: 0.8, y: 2.75, w: 7.5, h: 0.35, fontFace: HEAD, fontSize: 15,
      bold: true, color: C.CREAM, charSpacing: 4 });
    T(s, "Asesoría en procesos de selección por mérito público", { x: 0.8, y: 3.12, w: 7.5, h: 0.35,
      fontSize: 15, color: C.MUST_LIGHT, italic: true });
    T(s, "Título de la charla", { x: 0.8, y: 4.05, w: 7.6, h: 0.9, fontFace: HEAD, fontSize: 40, bold: true,
      color: C.CREAM, valign: "middle" });
    T(s, "Subtítulo o tema central de la sesión", { x: 0.8, y: 4.95, w: 7.6, h: 0.45, fontSize: 20, color: C.NAVY_40 });
    T(s, "Nombre del docente   ·   Fecha   ·   Ciudad", { x: 0.8, y: 5.85, w: 7.6, h: 0.4, fontFace: HEAD,
      fontSize: 13, bold: true, color: C.MUST, charSpacing: 1 });

    rect(s, 9.2, 0, W - 9.2, H, C.NAVY_900);
    T(s, "CONTACTO", { x: 9.7, y: 0.8, w: 3, h: 0.3, fontFace: HEAD, fontSize: 12, bold: true, color: C.MUST, charSpacing: 3 });
    const rows = [[I.globe, "Sitio web", "www.mvpg.com.co"], [I.mail, "Correo", "correo@dominio.com"],
      [I.phone, "Teléfono / WhatsApp", "+57 000 000 0000"], [I.users, "Redes sociales", "@usuario"],
      [I.pin, "Dirección", "Dirección, ciudad"]];
    rows.forEach(([ic, lab, val], i) => {
      const y = 1.45 + i * 0.95;
      iconCircle(s, 9.7, y, 0.55, C.MUST, ic);
      T(s, lab, { x: 10.45, y: y - 0.02, w: 2.4, h: 0.28, fontSize: 11, color: C.NAVY_40 });
      T(s, val, { x: 10.45, y: y + 0.25, w: 2.5, h: 0.32, fontSize: 14, bold: true, color: C.CREAM });
    });
    notes(s, "PORTADA. Edite el título, subtítulo, docente, fecha y los datos de contacto. No mueva ni deforme el logo. Recuerde actualizar el pie de página una sola vez en Ver › Patrón de diapositivas.");
  }

  // ============ 2. AGENDA ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Agenda", "Contenido de la sesión");
    const items = [["Contexto", "Por qué el mérito es la base del servicio público."],
      ["Marco normativo", "Reglas, entidades y etapas del concurso."],
      ["Preparación", "Estrategias para pruebas y valoración de antecedentes."],
      ["Cierre y preguntas", "Conclusiones y espacio de diálogo."]];
    items.forEach(([t, d], i) => {
      const y = TOP + i * 1.15;
      T(s, String(i + 1).padStart(2, "0"), { x: M, y, w: 0.9, h: 0.6, fontFace: HEAD, fontSize: 30, bold: true, color: C.MUST_TEXT });
      T(s, t, { x: M + 1.1, y: y + 0.02, w: 6.2, h: 0.4, fontFace: HEAD, fontSize: 20, bold: true });
      T(s, d, { x: M + 1.1, y: y + 0.45, w: 6.2, h: 0.35, fontSize: 14, color: C.NAVY_70 });
      if (i < 3) line(s, M, y + 1.02, 7.3, 0, C.CREAM_DEEP, { width: 0.75 });
    });
    await img(s, 8.4, TOP, W - M - 8.4, 4.55);
    notes(s, "AGENDA. Hasta 5 puntos. Si necesita más, duplique la diapositiva. Para cambiar la imagen: clic derecho › Cambiar imagen.");
  }

  // ============ 3. GUÍA: TIPOGRAFÍA ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Guía de estilos · 1", "Tipografía y jerarquía");
    rect(s, M, TOP, 4.6, 4.55, C.NAVY, { round: true });
    [[HEAD, "Arial", "Títulos y etiquetas\nNegrita", C.CREAM, 0.9], [BODY, "Calibri", "Textos y notas\nRegular / Negrita", C.MUST_LIGHT, 3.0]]
      .forEach(([f, n, u, col, x]) => {
        T(s, "Aa", { x, y: 2.05, w: 2.0, h: 1.4, fontFace: f, fontSize: 88, bold: f === HEAD, color: col });
        T(s, n, { x, y: 3.5, w: 2.0, h: 0.4, fontFace: f, fontSize: 20, bold: true, color: C.CREAM });
        T(s, u, { x, y: 3.9, w: 2.0, h: 0.55, fontSize: 12, color: C.NAVY_40 });
      });
    T(s, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ\nabcdefghijklmnñopqrstuvwxyz\n0123456789  ¿? ¡! áéíóú", { x: 0.9, y: 4.8, w: 4.0, h: 1.2,
      fontSize: 13, color: C.CREAM, lineSpacingMultiple: 1.15 });

    const rows = [
      [0.62, "ETIQUETA DE SECCIÓN", { fontFace: HEAD, fontSize: 12, bold: true, color: C.MUST_TEXT, charSpacing: 3 }, "Arial Negrita · 12 pt\nMayúsculas · Mostaza"],
      [0.95, "Título principal", { fontFace: HEAD, fontSize: 32, bold: true }, "Arial Negrita · 32–40 pt\nAzul oscuro"],
      [0.8, "Subtítulo de bloque", { fontFace: HEAD, fontSize: 20, bold: true }, "Arial Negrita · 18–20 pt\nAzul oscuro"],
      [0.8, "Texto de párrafo para ideas y explicaciones.", { fontSize: 16 }, "Calibri · 14–16 pt\nAzul oscuro"],
      [0.62, "Notas, fuentes y pies de imagen.", { fontSize: 11, color: C.NAVY_70 }, "Calibri · 10–11 pt\nAzul 70 %"],
    ];
    let y = TOP;
    rows.forEach(([h, sample, o, spec], i) => {
      T(s, sample, Object.assign({ x: 5.6, y, w: 4.9, h, valign: "middle" }, o));
      T(s, spec, { x: 10.7, y, w: W - M - 10.7, h, fontSize: 11, color: C.NAVY_70, valign: "middle" });
      y += h;
      line(s, 5.6, y + 0.04, W - M - 5.6, 0, C.CREAM_DEEP, { width: 0.75 });
      y += 0.08;
    });
    T(s, "Reglas: alinear textos a la izquierda (centrar solo títulos de portada y cierre) · máximo tres niveles por diapositiva · interlineado 1,1–1,2 · no usar cursiva en bloques largos.",
      { x: 5.6, y: y + 0.2, w: W - M - 5.6, h: 0.6, fontSize: 12, color: C.NAVY_70 });
    notes(s, "GUÍA DE TIPOGRAFÍA. Arial (títulos) y Calibri (textos) vienen instaladas en todos los equipos con Office, por lo que la plantilla se verá igual en cualquier computador. Esta diapositiva es de referencia: puede ocultarla antes de presentar.");
  }

  // ============ 4. GUÍA: COLOR ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Guía de estilos · 2", "Colores corporativos");
    const cw = (CW - 0.7) / 3;
    const cards = [
      ["Azul oscuro", C.NAVY, "27 · 42 · 74", "Títulos, textos, fondos de portada y secciones.", C.CREAM, [C.NAVY_900, C.NAVY_600, C.NAVY_70, C.NAVY_40]],
      ["Mostaza", C.MUST, "201 · 162 · 63", "Acentos: etiquetas, números, iconos y datos clave.", C.NAVY, [C.MUST_TEXT, C.MUST_LIGHT, C.MUST_PALE, C.MUST]],
      ["Crema", C.CREAM, "245 · 242 · 234", "Fondo de contenido y texto sobre fondos oscuros.", C.NAVY, [C.CREAM_LIGHT, C.CREAM_DARK, C.CREAM_DEEP, C.CREAM]],
    ];
    cards.forEach(([n, hex, rgb, use, tc, tints], i) => {
      const x = M + i * (cw + 0.35);
      rect(s, x, TOP, cw, 1.75, hex, { round: true, line: hex === C.CREAM ? C.CREAM_DEEP : null });
      T(s, n, { x: x + 0.3, y: TOP + 1.05, w: cw - 0.6, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: tc });
      T(s, [{ text: "HEX #" + hex, options: { bold: true, breakLine: true } }, { text: "RGB " + rgb }],
        { x, y: 3.8, w: cw, h: 0.5, fontSize: 12 });
      T(s, use, { x, y: 4.33, w: cw, h: 0.5, fontSize: 12, color: C.NAVY_70 });
      const tw = (cw - 0.3) / 4;
      tints.slice(0, 4).forEach((t, j) => {
        const tx = x + j * (tw + 0.1);
        rect(s, tx, 5.0, tw, 0.5, t, { round: true, line: C.CREAM_DEEP });
        T(s, "#" + t, { x: tx, y: 5.55, w: tw, h: 0.22, fontSize: 9, color: C.NAVY_70 });
      });
    });
    T(s, "Proporción", { x: M, y: 6.05, w: 1.5, h: 0.35, fontFace: HEAD, fontSize: 12, bold: true, valign: "middle" });
    const bx = M + 1.6, bw = CW - 1.6;
    [[0.6, C.CREAM_DARK, "Crema 60 %", C.NAVY], [0.3, C.NAVY, "Azul oscuro 30 %", C.CREAM], [0.1, C.MUST, "10 %", C.NAVY]]
      .reduce((acc, [p, col, lab, tc]) => {
        rect(s, acc, 6.05, bw * p, 0.35, col);
        T(s, lab, { x: acc, y: 6.05, w: bw * p, h: 0.35, fontSize: 10, bold: true, color: tc, align: "center", valign: "middle" });
        return acc + bw * p;
      }, bx);
    notes(s, "GUÍA DE COLOR. Use únicamente estos tres colores y sus variaciones tonales. Para texto mostaza sobre fondo crema use la variante oscura #8A6A1F (mejor legibilidad). La mostaza base #C9A23F se reserva para acentos sobre fondos azules y para formas.");
  }

  // ============ 5. GUÍA: RETÍCULA Y LOGO ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Guía de estilos · 3", "Retícula, márgenes y uso del logo");
    const mx = M, my = TOP, mw = 6.4, mh = 3.6, k = mw / W;
    rect(s, mx, my, mw, mh, C.CREAM_LIGHT, { line: C.CREAM_DEEP });
    rect(s, mx + M * k, my + 0.45 * k, mw - 2 * M * k, (6.85 + 0.35 - 0.45) * k, C.CREAM_LIGHT,
      { line: C.MUST, extra: { line: { color: C.MUST, width: 1, dashType: "dash" } } });
    rect(s, mx + M * k, my + 0.45 * k, 1.1, 0.12, C.MUST_PALE);
    rect(s, mx + M * k, my + 0.8 * k, 3.2, 0.3, C.NAVY_40);
    rect(s, mx + M * k, my + TOP * k, mw - 2 * M * k, (BOTTOM - TOP) * k, C.CREAM_DARK);
    T(s, "Área de contenido", { x: mx + M * k, y: my + TOP * k, w: mw - 2 * M * k, h: (BOTTOM - TOP) * k,
      fontSize: 12, color: C.NAVY_70, align: "center", valign: "middle" });
    T(s, "01  Pie de página", { x: mx + M * k, y: my + 6.85 * k, w: 2, h: 0.18, fontSize: 7, color: C.NAVY_70, valign: "middle" });
    s.addImage({ path: LOGO_LIGHT, x: mx + mw - M * k - LOGO_W * k, y: my + 6.8 * k, w: LOGO_W * k, h: LOGO_H * k });
    T(s, [
      { text: "Márgenes laterales: 0,6\"", options: { bullet: true, breakLine: true } },
      { text: "Etiqueta 0,45\" · Título 0,78\" · Contenido de 1,9\" a 6,5\"", options: { bullet: true, breakLine: true } },
      { text: "Separación entre bloques: 0,3\" o 0,35\"", options: { bullet: true } },
    ], { x: mx, y: 5.7, w: mw, h: 0.8, fontSize: 12, color: C.NAVY_70, paraSpaceAfter: 2 });

    const rx = 7.5, rw = W - M - rx;
    T(s, "Uso del logo", { x: rx, y: TOP, w: rw, h: 0.4, fontFace: HEAD, fontSize: 20, bold: true });
    const bw = (rw - 0.3) / 2;
    rect(s, rx, 2.45, bw, 1.35, C.CREAM_LIGHT, { round: true, line: C.CREAM_DEEP });
    s.addImage({ path: LOGO_LIGHT, x: rx + (bw - 1.7) / 2, y: 2.45 + (1.35 - 1.7 / 2.134) / 2, w: 1.7, h: 1.7 / 2.134 });
    rect(s, rx + bw + 0.3, 2.45, bw, 1.35, C.NAVY, { round: true });
    s.addImage({ path: LOGO_DARK, x: rx + bw + 0.3 + (bw - 1.7) / 2, y: 2.45 + (1.35 - 1.7 / 2.134) / 2, w: 1.7, h: 1.7 / 2.134 });
    T(s, "Fondos claros: versión azul", { x: rx, y: 3.9, w: bw, h: 0.3, fontSize: 11, color: C.NAVY_70 });
    T(s, "Fondos oscuros: versión blanca", { x: rx + bw + 0.3, y: 3.9, w: bw, h: 0.3, fontSize: 11, color: C.NAVY_70 });
    T(s, [
      { text: "Posición fija: esquina inferior derecha, 0,95\" de ancho.", options: { bullet: true, breakLine: true } },
      { text: "Dejar libre alrededor un espacio igual a la altura de la «P».", options: { bullet: true, breakLine: true } },
      { text: "No deformar, recolorear, girar ni aplicar sombras.", options: { bullet: true, breakLine: true } },
      { text: "No ubicar sobre fotografías sin un fondo sólido.", options: { bullet: true } },
    ], { x: rx, y: 4.45, w: rw, h: 1.9, fontSize: 14, paraSpaceAfter: 8 });
    notes(s, "RETÍCULA Y LOGO. El logo, el número de diapositiva y el pie ya están en el patrón (Ver › Patrón de diapositivas), por eso no se mueven accidentalmente. Diseños disponibles al insertar una diapositiva nueva: MVPG_PORTADA, MVPG_CONTENIDO, MVPG_TEXTO, MVPG_OSCURO y MVPG_SECCION.");
  }

  // ============ 6. SECCIÓN ============
  {
    const s = pres.addSlide({ masterName: "MVPG_SECCION" });
    T(s, "01", { x: M, y: 1.4, w: 4, h: 1.7, fontFace: HEAD, fontSize: 110, bold: true, color: C.MUST });
    T(s, "Título de la sección", { x: M, y: 3.35, w: 9, h: 0.9, fontFace: HEAD, fontSize: 44, bold: true, color: C.CREAM });
    T(s, "Una frase breve que presente lo que se abordará en este bloque de la charla.", { x: M, y: 4.35, w: 7.5, h: 0.9, fontSize: 20, color: C.NAVY_40 });
    notes(s, "SEPARADOR DE SECCIÓN. Duplique esta diapositiva para cada bloque de la charla y cambie el número.");
  }

  // ============ 7. TABLA SIMPLE ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Tablas · básica", "Etapas del concurso de méritos");
    const bB = [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: C.CREAM_DEEP }, { type: "none" }];
    const hd = (t) => ({ text: t, options: { bold: true, color: C.CREAM, fill: { color: C.NAVY }, fontFace: HEAD, fontSize: 12, border: bB } });
    const c = (t, o = {}) => ({ text: t, options: Object.assign({ border: bB }, o) });
    const data = [
      ["Convocatoria", "Entidad", "Semana 1", "Cumplido"],
      ["Inscripción", "Aspirante", "Semanas 2–3", "Cumplido"],
      ["Verificación de requisitos", "Operador", "Semana 5", "En curso"],
      ["Pruebas escritas", "Operador", "Semana 9", "Pendiente"],
      ["Lista de elegibles", "CNSC", "Semana 14", "Pendiente"],
    ];
    const rows = [["Etapa", "Responsable", "Fecha", "Estado"].map(hd)].concat(data.map((r, i) => r.map((t, j) =>
      c(t, { bold: j === 0, fill: { color: i % 2 ? C.CREAM : C.CREAM_LIGHT },
        color: j === 3 ? (t === "Cumplido" ? C.NAVY : t === "En curso" ? C.MUST_TEXT : C.NAVY_70) : C.NAVY }))));
    s.addTable(rows, { x: M, y: TOP, w: 8.1, colW: [3.3, 1.7, 1.7, 1.4], rowH: 0.55, fontFace: BODY, fontSize: 14,
      color: C.NAVY, valign: "middle", margin: [0, 0.15, 0, 0.15] });
    T(s, "Fuente: cronograma de referencia. Reemplace con los datos de su convocatoria.", { x: M, y: TOP + 3.45, w: 8.1, h: 0.3, fontSize: 11, color: C.NAVY_70 });
    const cx = 9.1, cwid = W - M - cx;
    rect(s, cx, TOP, cwid, 3.3, C.NAVY, { round: true });
    T(s, "DATO CLAVE", { x: cx + 0.35, y: TOP + 0.35, w: 3, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: C.MUST, charSpacing: 3 });
    T(s, "14 semanas", { x: cx + 0.35, y: TOP + 0.75, w: cwid - 0.7, h: 0.8, fontFace: HEAD, fontSize: 36, bold: true, color: C.CREAM });
    T(s, "Duración estimada desde la convocatoria hasta la publicación de la lista de elegibles.", { x: cx + 0.35, y: TOP + 1.65, w: cwid - 0.7, h: 1.3, fontSize: 14, color: C.CREAM_DARK });
    notes(s, "TABLA BÁSICA. Para añadir filas: clic en la última fila › Presentación (Diseño de tabla) › Insertar abajo. Mantenga encabezado azul y filas alternas crema.");
  }

  // ============ 8. TABLA COMPLEJA ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Tablas · avanzada", "Matriz de evaluación por criterios");
    const bB = [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: C.CREAM_DEEP }, { type: "none" }];
    const hdO = { bold: true, color: C.CREAM, fill: { color: C.NAVY }, fontFace: HEAD, fontSize: 12, border: bB, align: "center" };
    const subO = { bold: true, color: C.NAVY, fill: { color: C.MUST_PALE }, fontFace: HEAD, fontSize: 12, border: bB, align: "center" };
    const rows = [
      [{ text: "Criterio", options: Object.assign({}, hdO, { rowspan: 2, align: "left" }) },
       { text: "Peso", options: Object.assign({}, hdO, { rowspan: 2 }) },
       { text: "Puntaje por aspirante", options: Object.assign({}, hdO, { colspan: 3 }) },
       { text: "Observaciones", options: Object.assign({}, hdO, { rowspan: 2, align: "left" }) }],
      ["Aspirante A", "Aspirante B", "Aspirante C"].map((t) => ({ text: t, options: subO })),
    ];
    const data = [
      ["Competencias básicas", "40 %", 78, 85, 71, "Eliminatoria (mín. 65)"],
      ["Prueba comportamental", "20 %", 80, 74, 88, "Clasificatoria"],
      ["Educación", "15 %", 70, 90, 60, "Títulos adicionales"],
      ["Experiencia", "15 %", 85, 65, 75, "Relacionada y profesional"],
      ["Entrevista", "10 %", 90, 80, 85, "Si aplica"],
    ];
    const wts = [0.4, 0.2, 0.15, 0.15, 0.1];
    data.forEach((r, i) => {
      const scores = r.slice(2, 5), best = Math.max(...scores);
      rows.push([
        { text: r[0], options: { bold: true, border: bB, fill: { color: i % 2 ? C.CREAM : C.CREAM_LIGHT } } },
        { text: r[1], options: { align: "center", border: bB, fill: { color: i % 2 ? C.CREAM : C.CREAM_LIGHT } } },
        ...scores.map((v) => ({ text: String(v), options: { align: "center", border: bB, bold: v === best,
          fill: { color: v === best ? C.MUST_PALE : i % 2 ? C.CREAM : C.CREAM_LIGHT } } })),
        { text: r[5], options: { border: bB, color: C.NAVY_70, fill: { color: i % 2 ? C.CREAM : C.CREAM_LIGHT } } },
      ]);
    });
    const tot = [0, 1, 2].map((k) => data.reduce((a, r, i) => a + r[2 + k] * wts[i], 0).toFixed(1).replace(".", ","));
    const tO = { bold: true, color: C.CREAM, fill: { color: C.NAVY_600 }, border: bB };
    rows.push([{ text: "Puntaje ponderado", options: tO }, { text: "100 %", options: Object.assign({ align: "center" }, tO) },
      ...tot.map((t) => ({ text: t, options: Object.assign({ align: "center" }, tO) })),
      { text: "Resultado final", options: Object.assign({}, tO, { bold: false, color: C.CREAM_DARK }) }]);
    s.addTable(rows, { x: M, y: TOP, w: CW, colW: [3.4, 1.2, 1.6, 1.6, 1.6, CW - 9.4], rowH: 0.47, fontFace: BODY,
      fontSize: 13, color: C.NAVY, valign: "middle", margin: [0, 0.15, 0, 0.15] });
    rect(s, M, 5.95, 0.3, 0.22, C.MUST_PALE, { line: C.CREAM_DEEP });
    T(s, "Mejor puntaje por criterio. Puntaje ponderado = Σ (puntaje × peso). Datos de ejemplo.", { x: M + 0.45, y: 5.93, w: 9, h: 0.26, fontSize: 11, color: C.NAVY_70, valign: "middle" });
    notes(s, "TABLA AVANZADA con encabezado agrupado (celdas combinadas), resaltado del mejor valor en mostaza pálido y fila de totales. Si modifica los datos, recalcule manualmente el puntaje ponderado.");
  }

  // ============ 9. DIAGRAMA DE FLUJO ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Diagramas · flujo", "Flujo de verificación de requisitos");
    const cy = 3.0;
    const xs = [0.6, 2.708, 5.516, 8.424, 11.233];
    const pill = (x, t) => { s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy - 0.4, w: 1.5, h: 0.8, rectRadius: 0.4, fill: { color: C.NAVY }, line: { type: "none" } });
      T(s, t, { x, y: cy - 0.4, w: 1.5, h: 0.8, fontFace: HEAD, fontSize: 14, bold: true, color: C.CREAM, align: "center", valign: "middle" }); };
    const box = (x, y, t, fill = C.CREAM_LIGHT) => { rect(s, x, y, 2.2, 1.0, fill, { round: true, line: C.NAVY_40 });
      T(s, t, { x: x + 0.1, y, w: 2.0, h: 1.0, fontSize: 14, bold: true, align: "center", valign: "middle" }); };
    pill(xs[0], "Inicio");
    box(xs[1], cy - 0.5, "Recepción de documentos");
    s.addShape(pres.shapes.DIAMOND, { x: xs[2], y: cy - 0.75, w: 2.3, h: 1.5, fill: { color: C.MUST }, line: { type: "none" } });
    T(s, "¿Cumple requisitos?", { x: xs[2] + 0.45, y: cy - 0.5, w: 1.4, h: 1.0, fontSize: 13, bold: true, align: "center", valign: "middle" });
    box(xs[3], cy - 0.5, "Admitido a pruebas");
    pill(xs[4], "Fin");
    [[xs[0] + 1.5, xs[1]], [xs[1] + 2.2, xs[2]], [xs[2] + 2.3, xs[3]], [xs[3] + 2.2, xs[4]]].forEach(([a, b]) =>
      line(s, a + 0.05, cy, b - a - 0.1, 0, C.NAVY, { arrow: true }));
    T(s, "Sí", { x: xs[2] + 2.3, y: cy - 0.4, w: 0.6, h: 0.3, fontSize: 12, bold: true, color: C.MUST_TEXT, align: "center" });
    const dcx = xs[2] + 1.15;
    line(s, dcx, cy + 0.78, 0, 0.77, C.NAVY, { arrow: true });
    T(s, "No", { x: dcx + 0.12, y: cy + 0.95, w: 0.5, h: 0.3, fontSize: 12, bold: true, color: C.MUST_TEXT });
    box(dcx - 1.1, 4.6, "Subsanación y reclamación", C.CREAM_DARK);
    const r1x = xs[1] + 1.1;
    line(s, r1x, 5.1, dcx - 1.1 - r1x, 0, C.NAVY);
    line(s, r1x, cy + 0.55, 0, 5.1 - cy - 0.55, C.NAVY, { arrow: true, flipV: true });
    // leyenda
    const ly = 6.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: ly, w: 0.45, h: 0.25, rectRadius: 0.12, fill: { color: C.NAVY }, line: { type: "none" } });
    T(s, "Inicio / fin", { x: M + 0.55, y: ly, w: 1.3, h: 0.25, fontSize: 11, color: C.NAVY_70, valign: "middle" });
    rect(s, M + 2.0, ly, 0.45, 0.25, C.CREAM_LIGHT, { line: C.NAVY_40 });
    T(s, "Actividad", { x: M + 2.55, y: ly, w: 1.2, h: 0.25, fontSize: 11, color: C.NAVY_70, valign: "middle" });
    s.addShape(pres.shapes.DIAMOND, { x: M + 3.8, y: ly - 0.05, w: 0.45, h: 0.35, fill: { color: C.MUST }, line: { type: "none" } });
    T(s, "Decisión", { x: M + 4.35, y: ly, w: 1.2, h: 0.25, fontSize: 11, color: C.NAVY_70, valign: "middle" });
    notes(s, "DIAGRAMA DE FLUJO. Formas nativas editables: píldora azul = inicio/fin, rectángulo = actividad, rombo mostaza = decisión. Para mover todo junto, seleccione las formas y agrúpelas (Ctrl+G).");
  }

  // ============ 10. ORGANIGRAMA ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Diagramas · organigrama", "Estructura del equipo");
    const node = (x, y, w, h, t, sub, fill, tc, sc, fs) => {
      rect(s, x, y, w, h, fill, { round: true, line: fill === C.CREAM_LIGHT ? C.CREAM_DEEP : null });
      T(s, [{ text: t, options: { bold: true, fontFace: HEAD, fontSize: fs, color: tc, breakLine: true } },
        { text: sub, options: { fontSize: fs - 3, color: sc } }], { x: x + 0.1, y, w: w - 0.2, h, align: "center", valign: "middle" });
    };
    node(5.167, TOP, 3.0, 0.85, "Dirección general", "Nombre y apellido", C.NAVY, C.CREAM, C.MUST_LIGHT, 16);
    line(s, W / 2, TOP + 0.85, 0, 0.3, C.NAVY_70, { width: 1 });
    const l2 = [0.6, 4.867, 9.133];
    line(s, l2[0] + 1.8, TOP + 1.15, l2[2] - l2[0], 0, C.NAVY_70, { width: 1 });
    const names = [["Área académica", ["Coordinación", "Docencia"]], ["Área técnica", ["Pruebas", "Análisis"]], ["Área de gestión", ["Comunicaciones", "Atención"]]];
    l2.forEach((x, i) => {
      line(s, x + 1.8, TOP + 1.15, 0, 0.3, C.NAVY_70, { width: 1 });
      node(x, TOP + 1.45, 3.6, 0.8, names[i][0], "Nombre y apellido", C.MUST, C.NAVY, C.NAVY, 15);
      line(s, x + 1.8, TOP + 2.25, 0, 0.3, C.NAVY_70, { width: 1 });
      line(s, x + 0.85, TOP + 2.55, 1.9, 0, C.NAVY_70, { width: 1 });
      [0, 1].forEach((j) => {
        const nx = x + j * 1.9;
        line(s, nx + 0.85, TOP + 2.55, 0, 0.3, C.NAVY_70, { width: 1 });
        node(nx, TOP + 2.85, 1.7, 0.95, names[i][1][j], "Nombre", C.CREAM_LIGHT, C.NAVY, C.NAVY_70, 13);
      });
    });
    notes(s, "ORGANIGRAMA. Tres niveles con colores jerárquicos: azul (dirección), mostaza (áreas) y crema (equipos). Para agregar un cargo, duplique una caja (Ctrl+D) y su conector.");
  }

  // ============ 11. PROCESO ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Diagramas · proceso", "Ruta de preparación en cinco pasos");
    const steps = [["Diagnóstico", "Identificar el empleo, el nivel y los requisitos de la OPEC."],
      ["Documentación", "Organizar certificados de estudio y experiencia."],
      ["Estudio", "Plan de repaso por ejes temáticos de la convocatoria."],
      ["Simulacros", "Pruebas cronometradas y análisis de resultados."],
      ["Presentación", "Logística del día de la prueba y control de ansiedad."]];
    const fills = [C.NAVY_900, C.NAVY, C.NAVY_600, C.NAVY_70, C.MUST];
    steps.forEach(([t, d], i) => {
      const x = M + i * 2.408;
      s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x, y: 2.05, w: 2.5, h: 1.1, fill: { color: fills[i] }, line: { type: "none" } });
      T(s, String(i + 1).padStart(2, "0"), { x: x + (i === 0 ? 0.3 : 0.55), y: 2.05, w: 1.2, h: 1.1, fontFace: HEAD, fontSize: 26, bold: true,
        color: i === 4 ? C.NAVY : C.CREAM, valign: "middle" });
      T(s, t, { x: x + 0.05, y: 3.45, w: 2.2, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true });
      T(s, d, { x: x + 0.05, y: 3.9, w: 2.15, h: 1.2, fontSize: 13, color: C.NAVY_70 });
    });
    rect(s, M, 5.55, CW, 0.9, C.CREAM_DARK, { round: true });
    s.addImage({ data: I.bulbN, x: M + 0.35, y: 5.77, w: 0.45, h: 0.45 });
    T(s, [{ text: "Resultado esperado: ", options: { bold: true } }, { text: "un aspirante preparado para superar cada etapa del concurso." }],
      { x: M + 1.05, y: 5.55, w: CW - 1.4, h: 0.9, fontSize: 16, valign: "middle" });
    notes(s, "PROCESO. El paso final se destaca en mostaza. Si usa menos pasos, elimine uno y reparta el espacio para mantener los márgenes de 0,6\".");
  }

  // ============ 12. LÍNEA DE TIEMPO HORIZONTAL ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Líneas de tiempo · horizontal", "Cronograma de la convocatoria");
    const ly = 3.55;
    line(s, M, ly, CW, 0, C.NAVY, { width: 2 });
    const ev = [["Feb", "Apertura", "Publicación del acuerdo de convocatoria."], ["Mar", "Inscripciones", "Registro y pago en la plataforma."],
      ["May", "Verificación", "Revisión de requisitos mínimos."], ["Jul", "Pruebas", "Aplicación de pruebas escritas."], ["Oct", "Elegibles", "Publicación de la lista final."]];
    ev.forEach(([d, t, desc], i) => {
      const cx = 1.8 + i * 2.433;
      const now = i === 2;
      const dd = now ? 0.5 : 0.34;
      circle(s, cx - dd / 2, ly - dd / 2, dd, now ? C.NAVY : C.MUST, C.CREAM, 3);
      T(s, d, { x: cx - 1.1, y: 2.5, w: 2.2, h: 0.55, fontFace: HEAD, fontSize: 24, bold: true, color: now ? C.MUST_TEXT : C.NAVY, align: "center", valign: "bottom" });
      T(s, t, { x: cx - 1.1, y: 4.0, w: 2.2, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true, align: "center" });
      T(s, desc, { x: cx - 1.05, y: 4.45, w: 2.1, h: 0.9, fontSize: 13, color: C.NAVY_70, align: "center" });
    });
    T(s, [{ text: "●  ", options: { color: C.NAVY } }, { text: "Hito actual    " }, { text: "●  ", options: { color: C.MUST } }, { text: "Hito programado" }],
      { x: M, y: 6.0, w: 6, h: 0.3, fontSize: 11, color: C.NAVY_70 });
    notes(s, "LÍNEA DE TIEMPO HORIZONTAL. El hito actual se marca con un círculo azul más grande; los demás en mostaza. Ideal para 4 a 6 hitos.");
  }

  // ============ 13. LÍNEA DE TIEMPO VERTICAL ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Líneas de tiempo · vertical", "Hitos del proceso de formación");
    T(s, "Semana a semana", { x: M, y: TOP, w: 4.3, h: 0.4, fontFace: HEAD, fontSize: 20, bold: true });
    T(s, "Use este formato cuando cada hito requiera más texto o cuando la charla siga una secuencia cronológica.", { x: M, y: TOP + 0.5, w: 4.3, h: 0.9, fontSize: 14, color: C.NAVY_70 });
    await img(s, M, 3.45, 4.3, 3.0);
    const lx = 5.9;
    line(s, lx, TOP + 0.15, 0, 4.25, C.NAVY_40, { width: 2 });
    const ev = [["Semana 1", "Inducción", "Presentación del programa, objetivos y metodología."],
      ["Semana 2", "Marco normativo", "Ley de carrera administrativa y acuerdos de convocatoria."],
      ["Semana 3", "Competencias", "Taller de pruebas funcionales y comportamentales."],
      ["Semana 4", "Simulacro final", "Prueba integral con retroalimentación personalizada."]];
    ev.forEach(([d, t, desc], i) => {
      const y = TOP + i * 1.15;
      circle(s, lx - 0.16, y + 0.07, 0.32, i === 3 ? C.NAVY : C.MUST, C.CREAM, 3);
      T(s, d, { x: lx + 0.45, y: y + 0.05, w: 1.5, h: 0.35, fontFace: HEAD, fontSize: 13, bold: true, color: C.MUST_TEXT, valign: "middle" });
      T(s, t, { x: 8.05, y: y + 0.03, w: W - M - 8.05, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, valign: "middle" });
      T(s, desc, { x: 8.05, y: y + 0.45, w: W - M - 8.05, h: 0.55, fontSize: 13, color: C.NAVY_70 });
    });
    notes(s, "LÍNEA DE TIEMPO VERTICAL. Admite descripciones más largas. Para cambiar la imagen: clic derecho › Cambiar imagen.");
  }

  // ============ 14. GALERÍA 3 ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Galería · tres imágenes", "Experiencias de nuestras sesiones");
    const cw = (CW - 0.7) / 3;
    for (let i = 0; i < 3; i++) {
      const x = M + i * (cw + 0.35);
      await img(s, x, TOP, cw, 3.05);
      T(s, ["Título de la imagen 1", "Título de la imagen 2", "Título de la imagen 3"][i], { x, y: 5.15, w: cw, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true });
      T(s, "Breve descripción o pie de foto de una o dos líneas.", { x, y: 5.6, w: cw, h: 0.7, fontSize: 13, color: C.NAVY_70 });
    }
    notes(s, "GALERÍA DE TRES. Clic derecho sobre cada imagen › Cambiar imagen › Desde un archivo. La nueva imagen conserva el tamaño y la posición.");
  }

  // ============ 15. GALERÍA MOSAICO ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Galería · mosaico", "Registro fotográfico");
    const bw = 6.6, bh = BOTTOM - TOP;
    await img(s, M, TOP, bw, bh);
    rect(s, M + 0.25, BOTTOM - 0.8, 3.6, 0.55, C.NAVY, { round: true });
    T(s, "Pie de la imagen principal", { x: M + 0.45, y: BOTTOM - 0.8, w: 3.3, h: 0.55, fontSize: 13, bold: true, color: C.CREAM, valign: "middle" });
    const sx = M + bw + 0.3, sw = (W - M - sx - 0.3) / 2, sh = (bh - 0.3) / 2;
    for (let r = 0; r < 2; r++) for (let c = 0; c < 2; c++) await img(s, sx + c * (sw + 0.3), TOP + r * (sh + 0.3), sw, sh);
    notes(s, "GALERÍA MOSAICO. Una imagen principal y cuatro secundarias. Reemplace con clic derecho › Cambiar imagen. Si la foto tiene otra proporción, use Recortar › Rellenar.");
  }

  // ============ 16. PUNTOS CLAVE (ICONOS) ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Listas · puntos clave", "Nuestros pilares de trabajo");
    const items = [[I.shield, "Confianza", "Acompañamiento transparente y ético en cada etapa del proceso de selección."],
      [I.award, "Excelencia", "Contenidos actualizados y docentes con experiencia en concursos públicos."],
      [I.chart, "Crecimiento", "Seguimiento del progreso de cada aspirante con metas medibles."],
      [I.target, "Enfoque técnico", "Metodología basada en los ejes temáticos oficiales de cada convocatoria."]];
    const cw = (CW - 0.35) / 2, ch = (BOTTOM - TOP - 0.35) / 2;
    items.forEach(([ic, t, d], i) => {
      const x = M + (i % 2) * (cw + 0.35), y = TOP + Math.floor(i / 2) * (ch + 0.35);
      rect(s, x, y, cw, ch, C.CREAM_LIGHT, { round: true, line: C.CREAM_DARK });
      iconCircle(s, x + 0.4, y + 0.4, 0.85, C.NAVY, ic);
      T(s, t, { x: x + 1.55, y: y + 0.42, w: cw - 1.9, h: 0.4, fontFace: HEAD, fontSize: 20, bold: true });
      T(s, d, { x: x + 1.55, y: y + 0.9, w: cw - 1.9, h: 0.9, fontSize: 14, color: C.NAVY_70 });
    });
    notes(s, "PUNTOS CLAVE CON ICONOS. Los iconos son imágenes: para cambiarlos use Insertar › Iconos y coloréelos en mostaza (#C9A23F) sobre el círculo azul.");
  }

  // ============ 17. LISTA CON MARCADOR DE TEXTO ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Listas · viñetas", "Recomendaciones para la prueba escrita");
    const pts = [["Lea la pregunta completa ", "antes de revisar las opciones de respuesta."],
      ["Administre el tiempo: ", "calcule cuántos minutos tiene por pregunta."],
      ["Identifique palabras clave ", "como «excepto», «siempre» o «nunca»."],
      ["Responda primero lo que sabe ", "y regrese luego a las preguntas difíciles."],
      ["Revise la hoja de respuestas ", "antes de entregar."]];
    pts.forEach(([a, b], i) => {
      rect(s, M, TOP + i * 0.9 + 0.1, 0.14, 0.14, C.MUST);
      T(s, [{ text: a, options: { bold: true } }, { text: b }], { x: M + 0.4, y: TOP + i * 0.9, w: 6.9, h: 0.8, fontSize: 18 });
    });
    const cx = 8.4, cwid = W - M - cx;
    rect(s, cx, TOP, cwid, BOTTOM - TOP, C.NAVY, { round: true });
    s.addImage({ data: I.bulb, x: cx + 0.4, y: TOP + 0.4, w: 0.5, h: 0.5 });
    T(s, "IDEA CLAVE", { x: cx + 0.4, y: TOP + 1.1, w: 3, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: C.MUST, charSpacing: 3 });
    T(s, "La preparación constante supera a la memorización de última hora.", { x: cx + 0.4, y: TOP + 1.5, w: cwid - 0.8, h: 2.0,
      fontFace: HEAD, fontSize: 22, bold: true, color: C.CREAM });
    T(s, "Refuerce esta idea al cierre del bloque.", { x: cx + 0.4, y: BOTTOM - 0.8, w: cwid - 0.8, h: 0.4, fontSize: 12, color: C.NAVY_40 });
    notes(s, "LISTA CON VIÑETAS. Pulse Enter para nuevas viñetas. Para una lista a ancho completo inserte una diapositiva nueva con el diseño MVPG_TEXTO, que trae un marcador de texto listo. Use negrita para la idea inicial de cada punto.");
  }

  // ============ 18. COMPARATIVA 2 ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Comparativas · dos elementos", "Preparación individual vs. acompañada");
    const cw = (CW - 0.55) / 2, ch = BOTTOM - TOP;
    const cols = [["OPCIÓN A", "Estudio individual", C.CREAM_DARK, C.NAVY, C.NAVY_70, I.checkN,
      ["Horarios flexibles", "Material disperso y sin filtro", "Sin retroalimentación de simulacros", "Menor costo inicial"]],
      ["OPCIÓN B", "Acompañamiento MVPG", C.NAVY, C.CREAM, C.CREAM_DARK, I.checkM,
      ["Plan de estudio por convocatoria", "Docentes especializados", "Simulacros con análisis de resultados", "Seguimiento personalizado"]]];
    cols.forEach(([k, t, fill, tc, bc, ck, items], i) => {
      const x = M + i * (cw + 0.55);
      rect(s, x, TOP, cw, ch, fill, { round: true });
      T(s, k, { x: x + 0.5, y: TOP + 0.45, w: 3, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: i ? C.MUST : C.MUST_TEXT, charSpacing: 3 });
      T(s, t, { x: x + 0.5, y: TOP + 0.8, w: cw - 1, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true, color: tc });
      items.forEach((it, j) => {
        const y = TOP + 1.7 + j * 0.62;
        s.addImage({ data: ck, x: x + 0.5, y: y + 0.07, w: 0.26, h: 0.26 });
        T(s, it, { x: x + 0.95, y, w: cw - 1.45, h: 0.4, fontSize: 16, color: bc, valign: "middle" });
      });
    });
    circle(s, W / 2 - 0.45, TOP + ch / 2 - 0.45, 0.9, C.MUST, C.CREAM, 4);
    T(s, "VS", { x: W / 2 - 0.45, y: TOP + ch / 2 - 0.45, w: 0.9, h: 0.9, fontFace: HEAD, fontSize: 18, bold: true, align: "center", valign: "middle" });
    notes(s, "COMPARATIVA DE DOS. La opción recomendada va a la derecha en azul. Mantenga el mismo número de puntos en ambas columnas.");
  }

  // ============ 19. COMPARATIVA 3 ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Comparativas · tres elementos", "Modalidades de acompañamiento");
    const cw = (CW - 0.7) / 3, ch = BOTTOM - TOP;
    const cols = [["Básico", "8", "sesiones", ["Material de estudio", "2 simulacros", "Foro de dudas"]],
      ["Integral", "16", "sesiones", ["Material + clases en vivo", "6 simulacros", "Tutor asignado"]],
      ["Intensivo", "24", "sesiones", ["Todo lo del plan integral", "Simulacros ilimitados", "Entrevista de práctica"]]];
    cols.forEach(([t, n, u, items], i) => {
      const x = M + i * (cw + 0.35), hi = i === 1;
      const tc = hi ? C.CREAM : C.NAVY, bc = hi ? C.CREAM_DARK : C.NAVY_70;
      rect(s, x, TOP, cw, ch, hi ? C.NAVY : C.CREAM_LIGHT, { round: true, line: hi ? null : C.CREAM_DEEP });
      if (hi) {
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.4, y: TOP + 0.35, w: 1.7, h: 0.34, rectRadius: 0.17, fill: { color: C.MUST }, line: { type: "none" } });
        T(s, "RECOMENDADO", { x: x + 0.4, y: TOP + 0.35, w: 1.7, h: 0.34, fontFace: HEAD, fontSize: 10, bold: true, align: "center", valign: "middle", charSpacing: 2 });
      }
      T(s, t, { x: x + 0.4, y: TOP + 0.85, w: cw - 0.8, h: 0.45, fontFace: HEAD, fontSize: 22, bold: true, color: tc });
      T(s, [{ text: n, options: { fontSize: 48, bold: true, color: hi ? C.MUST : C.MUST_TEXT } }, { text: "  " + u, options: { fontSize: 14, color: bc } }],
        { x: x + 0.4, y: TOP + 1.4, w: cw - 0.8, h: 0.9, fontFace: HEAD, valign: "bottom" });
      line(s, x + 0.4, TOP + 2.55, cw - 0.8, 0, hi ? C.NAVY_600 : C.CREAM_DEEP, { width: 0.75 });
      items.forEach((it, j) => {
        const y = TOP + 2.8 + j * 0.5;
        s.addImage({ data: hi ? I.checkM : I.checkN, x: x + 0.4, y: y + 0.08, w: 0.22, h: 0.22 });
        T(s, it, { x: x + 0.8, y, w: cw - 1.2, h: 0.38, fontSize: 14, color: bc, valign: "middle" });
      });
    });
    notes(s, "COMPARATIVA DE TRES. La tarjeta central en azul con la etiqueta RECOMENDADO destaca la opción sugerida. Para destacar otra, intercambie los rellenos.");
  }

  // ============ 20. CITA ============
  {
    const s = pres.addSlide({ masterName: "MVPG_SECCION" });
    s.addImage({ data: I.quote, x: M, y: 1.0, w: 0.9, h: 0.9 });
    T(s, "El mérito no es un privilegio: es la garantía de que el servicio público quede en manos de las personas mejor preparadas.",
      { x: M, y: 2.25, w: 10.5, h: 2.4, fontFace: HEAD, fontSize: 32, bold: true, color: C.CREAM, lineSpacingMultiple: 1.1 });
    T(s, "Nombre del autor", { x: M, y: 5.0, w: 6, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true, color: C.MUST });
    T(s, "Cargo, institución o fuente", { x: M, y: 5.4, w: 6, h: 0.35, fontSize: 14, color: C.NAVY_40 });
    notes(s, "CITA DESTACADA. Máximo 30 palabras para que se lea de un vistazo. Cite siempre al autor y la fuente.");
  }

  // ============ 21. DATOS DESTACADOS ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Datos destacados · cifras", "El mérito en números");
    const cw = CW / 3;
    const st = [["+1.200", "aspirantes asesorados", "Personas acompañadas en procesos de selección."],
      ["87 %", "superan la prueba", "Aprobación en pruebas de competencias básicas."],
      ["35", "convocatorias", "Procesos atendidos en entidades del orden nacional y territorial."]];
    st.forEach(([n, l, d], i) => {
      const x = M + i * cw;
      if (i) line(s, x, TOP + 0.3, 0, 2.9, C.CREAM_DEEP, { width: 0.75 });
      const px = x + (i ? 0.45 : 0);
      T(s, n, { x: px, y: TOP + 0.2, w: cw - 0.6, h: 1.3, fontFace: HEAD, fontSize: 60, bold: true, color: i === 1 ? C.MUST_TEXT : C.NAVY, valign: "middle" });
      T(s, l, { x: px, y: TOP + 1.6, w: cw - 0.6, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true });
      T(s, d, { x: px, y: TOP + 2.05, w: cw - 0.8, h: 0.9, fontSize: 13, color: C.NAVY_70 });
    });
    rect(s, M, 5.5, CW, 0.95, C.NAVY, { round: true });
    s.addImage({ data: I.chart, x: M + 0.4, y: 5.72, w: 0.5, h: 0.5 });
    T(s, [{ text: "Conclusión: ", options: { bold: true, color: C.MUST } }, { text: "la preparación estructurada aumenta de forma significativa las probabilidades de éxito.", options: { color: C.CREAM } }],
      { x: M + 1.2, y: 5.5, w: CW - 1.6, h: 0.95, fontSize: 16, valign: "middle" });
    notes(s, "DATOS DESTACADOS. Cifras de ejemplo: reemplácelas por datos reales y cite la fuente en la nota. Máximo tres cifras por diapositiva.");
  }

  // ============ 22. DATO + GRÁFICO ============
  {
    const s = pres.addSlide({ masterName: "MVPG_CONTENIDO" });
    header(s, "Datos destacados · gráfico", "Resultados por tipo de prueba");
    const cwid = 4.0;
    rect(s, M, TOP, cwid, BOTTOM - TOP, C.NAVY, { round: true });
    T(s, "PROMEDIO GENERAL", { x: M + 0.4, y: TOP + 0.45, w: 3.2, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: C.MUST, charSpacing: 3 });
    T(s, "78,4", { x: M + 0.4, y: TOP + 0.85, w: 3.2, h: 1.3, fontFace: HEAD, fontSize: 72, bold: true, color: C.CREAM, valign: "middle" });
    T(s, "puntos sobre 100 en el último simulacro, 6,2 puntos por encima del simulacro anterior.", { x: M + 0.4, y: TOP + 2.3, w: 3.2, h: 1.3, fontSize: 14, color: C.CREAM_DARK });
    s.addChart(pres.charts.BAR, [
      { name: "Simulacro 1", labels: ["Básicas", "Funcionales", "Comportamentales", "Análisis de caso"], values: [68, 71, 80, 70] },
      { name: "Simulacro 2", labels: ["Básicas", "Funcionales", "Comportamentales", "Análisis de caso"], values: [76, 79, 85, 74] },
    ], { x: M + cwid + 0.4, y: TOP, w: CW - cwid - 0.4, h: BOTTOM - TOP, barDir: "col", barGapWidthPct: 60,
      chartColors: [C.NAVY_40, C.NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.NAVY, dataLabelFontSize: 11,
      dataLabelFontFace: BODY, catAxisLabelColor: C.NAVY_70, catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12,
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      catAxisLineColor: C.CREAM_DEEP, showLegend: true, legendPos: "t", legendColor: C.NAVY_70, legendFontFace: BODY, legendFontSize: 12,
      plotArea: { fill: { color: C.CREAM } } });
    notes(s, "DATO + GRÁFICO. El gráfico es nativo: clic derecho › Editar datos para cambiar los valores. Colores: azul oscuro para la serie principal y azul 40 % para la comparación.");
  }

  // ============ 23. CIERRE ============
  {
    const s = pres.addSlide({ masterName: "MVPG_PORTADA" });
    s.addImage({ path: LOGO_DARK, x: 0.8, y: 0.75, w: 2.6, h: 2.6 / 2.134 });
    T(s, "Gracias", { x: 0.8, y: 2.45, w: 7, h: 1.2, fontFace: HEAD, fontSize: 64, bold: true, color: C.CREAM });
    T(s, "¿Preguntas?", { x: 0.8, y: 3.6, w: 7, h: 0.6, fontFace: HEAD, fontSize: 28, bold: true, color: C.MUST });
    T(s, [{ text: "Nombre del docente", options: { bold: true, color: C.CREAM, breakLine: true } },
      { text: "correo@dominio.com   ·   +57 000 000 0000", options: { breakLine: true } },
      { text: "www.mvpg.com.co" }], { x: 0.8, y: 4.65, w: 7, h: 1.2, fontSize: 15, color: C.NAVY_40, lineSpacingMultiple: 1.2 });
    const vals = [[I.shield, "Confianza"], [I.award, "Excelencia"], [I.chart, "Crecimiento"], [I.target, "Enfoque técnico"]];
    rect(s, 8.6, 0.75, W - M - 8.6 + 0.0, 6.0, C.NAVY_900, { round: true });
    T(s, "NUESTROS VALORES", { x: 9.05, y: 1.15, w: 3.5, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: C.MUST, charSpacing: 3 });
    vals.forEach(([ic, t], i) => {
      const y = 1.75 + i * 1.2;
      iconCircle(s, 9.05, y, 0.8, C.NAVY_600, ic);
      T(s, t, { x: 10.1, y, w: 2.5, h: 0.8, fontFace: HEAD, fontSize: 17, bold: true, color: C.CREAM, valign: "middle" });
    });
    notes(s, "CIERRE. Actualice los datos de contacto del docente. Los cuatro valores provienen de la identidad de MVPG.");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
